using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Windows.Data.Pdf;
using Windows.Foundation;
using Windows.Storage.Streams;

namespace PrintSelector
{
    internal sealed class PageTrayAssignment
    {
        public int PageIndex { get; set; }
        public string Tray { get; set; }
    }

    internal static class PdfHelper
    {
        private static T Wait<T>(IAsyncOperation<T> op)
        {
            while (op.Status == AsyncStatus.Started)
                Thread.Sleep(2);
            if (op.Status != AsyncStatus.Completed)
                throw op.ErrorCode ?? new InvalidOperationException("PDF operation failed.");
            return op.GetResults();
        }

        private static void Wait(IAsyncAction op)
        {
            while (op.Status == AsyncStatus.Started)
                Thread.Sleep(2);
            if (op.Status != AsyncStatus.Completed)
                throw op.ErrorCode ?? new InvalidOperationException("PDF action failed.");
            op.GetResults();
        }

        public static PdfDocument OpenDocument(string pdfPath, out InMemoryRandomAccessStream backingStream)
        {
            byte[] fileBytes = File.ReadAllBytes(pdfPath);
            backingStream = new InMemoryRandomAccessStream();
            using (var writer = new DataWriter(backingStream.GetOutputStreamAt(0)))
            {
                writer.WriteBytes(fileBytes);
                Wait(writer.StoreAsync());
                Wait(writer.FlushAsync());
                writer.DetachStream();
            }
            backingStream.Seek(0);
            return Wait(PdfDocument.LoadFromStreamAsync(backingStream));
        }

        public static int GetPageCount(string pdfPath)
        {
            InMemoryRandomAccessStream stream = null;
            try
            {
                PdfDocument doc = OpenDocument(pdfPath, out stream);
                return (int)doc.PageCount;
            }
            finally
            {
                if (stream != null) stream.Dispose();
            }
        }

        public static Bitmap RenderPage(PdfDocument doc, int pageIndex, double scale)
        {
            using (PdfPage page = doc.GetPage((uint)pageIndex))
            using (var renderStream = new InMemoryRandomAccessStream())
            {
                var opts = new PdfPageRenderOptions();
                opts.DestinationWidth = Math.Max(1u, (uint)Math.Round(page.Size.Width * scale));
                opts.DestinationHeight = Math.Max(1u, (uint)Math.Round(page.Size.Height * scale));

                Wait(page.RenderToStreamAsync(renderStream, opts));

                uint size = (uint)renderStream.Size;
                using (var reader = new DataReader(renderStream.GetInputStreamAt(0)))
                {
                    Wait(reader.LoadAsync(size));
                    byte[] bytes = new byte[size];
                    reader.ReadBytes(bytes);

                    using (var ms = new MemoryStream(bytes))
                    using (var raw = new Bitmap(ms))
                    {
                        // Flatten onto a white 24-bit RGB surface for GDI printing and previews
                        var rgb = new Bitmap(raw.Width, raw.Height, PixelFormat.Format24bppRgb);
                        using (Graphics g = Graphics.FromImage(rgb))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(raw, 0, 0, raw.Width, raw.Height);
                        }
                        return rgb;
                    }
                }
            }
        }

        public static Bitmap RenderSinglePage(string pdfPath, int pageIndex, double scale)
        {
            InMemoryRandomAccessStream stream = null;
            try
            {
                PdfDocument doc = OpenDocument(pdfPath, out stream);
                return RenderPage(doc, pageIndex, scale);
            }
            finally
            {
                if (stream != null) stream.Dispose();
            }
        }
    }

    internal static class PrinterEngine
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct DOCINFOW
        {
            public int cbSize;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszDocName;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszOutput;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszDatatype;
            public uint fwType;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct BITMAPINFOHEADER
        {
            public uint biSize;
            public int biWidth;
            public int biHeight;
            public ushort biPlanes;
            public ushort biBitCount;
            public uint biCompression;
            public uint biSizeImage;
            public int biXPelsPerMeter;
            public int biYPelsPerMeter;
            public uint biClrUsed;
            public uint biClrImportant;
        }

        [DllImport("winspool.drv", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);

        [DllImport("winspool.drv", SetLastError = true)]
        private static extern bool ClosePrinter(IntPtr hPrinter);

        [DllImport("winspool.drv", EntryPoint = "DocumentPropertiesW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter, string pDeviceName, IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);

        [DllImport("gdi32.dll", EntryPoint = "CreateDCW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateDC(string lpszDriver, string lpszDevice, string lpszOutput, IntPtr lpInitData);

        [DllImport("gdi32.dll", EntryPoint = "ResetDCW", SetLastError = true)]
        private static extern IntPtr ResetDC(IntPtr hdc, IntPtr lpInitData);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern bool DeleteDC(IntPtr hdc);

        [DllImport("gdi32.dll", EntryPoint = "StartDocW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int StartDoc(IntPtr hdc, [In] ref DOCINFOW lpdi);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int EndDoc(IntPtr hdc);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int StartPage(IntPtr hdc);

        [DllImport("gdi32.dll", SetLastError = true)]
        private static extern int EndPage(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("gdi32.dll")]
        private static extern int StretchDIBits(
            IntPtr hdc, int xDest, int yDest, int DestWidth, int DestHeight,
            int xSrc, int ySrc, int SrcWidth, int SrcHeight,
            IntPtr lpBits, [In] ref BITMAPINFOHEADER lpBitsInfo, uint iUsage, uint dwRop);

        private const int DM_OUT_BUFFER = 2;
        private const int DM_IN_BUFFER = 8;

        private const int DM_ORIENTATION = 0x00000001;
        private const int DM_DEFAULTSOURCE = 0x00000200;
        private const int DM_COLOR = 0x00000800;

        private const short DMORIENT_PORTRAIT = 1;
        private const short DMORIENT_LANDSCAPE = 2;

        private const short DMCOLOR_MONOCHROME = 1;
        private const short DMCOLOR_COLOR = 2;

        private const short DMBIN_UPPER = 1;
        private const short DMBIN_LOWER = 2;
        private const short DMBIN_MIDDLE = 3;
        private const short DMBIN_MANUAL = 4;
        private const short DMBIN_AUTO = 7;

        private const int HORZRES = 8;
        private const int VERTRES = 10;
        private const uint SRCCOPY = 0x00CC0020;

        // Byte offsets inside the standard Unicode DEVMODEW header
        private const int OFFSET_DM_FIELDS = 72;
        private const int OFFSET_DM_ORIENTATION = 76;
        private const int OFFSET_DM_DEFAULTSOURCE = 88;
        private const int OFFSET_DM_COLOR = 92;

        public static List<string> GetInstalledPrinters()
        {
            var list = new List<string>();
            try
            {
                foreach (string name in PrinterSettings.InstalledPrinters)
                    list.Add(name);
            }
            catch (Exception ex)
            {
                CrashLogger.Log("printer_enumeration", ex);
            }
            return list;
        }

        public static short GetTrayId(string printerName, string logicalName)
        {
            try
            {
                var ps = new PrinterSettings { PrinterName = printerName };
                var bins = new Dictionary<string, short>(StringComparer.OrdinalIgnoreCase);

                foreach (PaperSource src in ps.PaperSources)
                {
                    if (!string.IsNullOrEmpty(src.SourceName))
                        bins[src.SourceName.Trim()] = (short)src.RawKind;
                }

                if (bins.Count > 0)
                {
                    if (logicalName == "Auto Select")
                    {
                        short id;
                        if (bins.TryGetValue("Auto Select", out id)) return id;
                        if (bins.TryGetValue("Automatically Select", out id)) return id;
                        return DMBIN_AUTO;
                    }
                    if (logicalName == "Tray 1")
                    {
                        short id;
                        if (bins.TryGetValue("Paper Cassette 1", out id)) return id;
                        if (bins.TryGetValue("Tray 1", out id)) return id;
                        if (bins.TryGetValue("Cassette 1", out id)) return id;
                        return 258;
                    }
                    if (logicalName == "Tray 2")
                    {
                        short id;
                        if (bins.TryGetValue("Paper Cassette 2", out id)) return id;
                        if (bins.TryGetValue("Tray 2", out id)) return id;
                        if (bins.TryGetValue("Cassette 2", out id)) return id;
                        return 259;
                    }
                    if (logicalName == "Tray 3")
                    {
                        short id;
                        if (bins.TryGetValue("Paper Cassette 3", out id)) return id;
                        if (bins.TryGetValue("Tray 3", out id)) return id;
                        if (bins.TryGetValue("Cassette 3", out id)) return id;
                        return 260;
                    }
                    if (logicalName == "Tray 4")
                    {
                        short id;
                        if (bins.TryGetValue("Paper Cassette 4", out id)) return id;
                        if (bins.TryGetValue("Tray 4", out id)) return id;
                        if (bins.TryGetValue("Cassette 4", out id)) return id;
                        return 262;
                    }
                    if (logicalName == "MP Tray")
                    {
                        short id;
                        if (bins.TryGetValue("Paper Tray", out id)) return id;
                        if (bins.TryGetValue("Rear Paper Feed", out id)) return id;
                        if (bins.TryGetValue("MP Tray", out id)) return id;
                        if (bins.TryGetValue("Multi-Purpose Tray", out id)) return id;
                        if (bins.TryGetValue("Manual Feed", out id)) return id;
                    }
                }
            }
            catch
            {
                // Fall through to standard DMBIN defaults below
            }

            switch (logicalName)
            {
                case "Tray 1": return DMBIN_UPPER;
                case "Tray 2": return DMBIN_LOWER;
                case "Tray 3": return DMBIN_MIDDLE;
                case "Tray 4": return 8;
                case "MP Tray": return DMBIN_MANUAL;
                default: return DMBIN_AUTO;
            }
        }

        private static IntPtr BuildConfiguredDevMode(
            IntPtr hPrinter,
            string printerName,
            string trayName,
            string orientation,
            string colorMode)
        {
            int size = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
            if (size <= 0)
                throw new InvalidOperationException("Could not query DEVMODE size for printer: " + printerName);

            IntPtr pDevMode = Marshal.AllocHGlobal(size);
            int res = DocumentProperties(IntPtr.Zero, hPrinter, printerName, pDevMode, IntPtr.Zero, DM_OUT_BUFFER);
            if (res < 0)
            {
                Marshal.FreeHGlobal(pDevMode);
                throw new InvalidOperationException("Failed to initialize printer DEVMODE.");
            }

            short trayId = GetTrayId(printerName, trayName);
            short orientVal = string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase)
                ? DMORIENT_LANDSCAPE
                : DMORIENT_PORTRAIT;
            short colorVal = string.Equals(colorMode, "color", StringComparison.OrdinalIgnoreCase)
                ? DMCOLOR_COLOR
                : DMCOLOR_MONOCHROME;

            int fields = Marshal.ReadInt32(pDevMode, OFFSET_DM_FIELDS);
            fields |= DM_DEFAULTSOURCE | DM_ORIENTATION | DM_COLOR;

            Marshal.WriteInt32(pDevMode, OFFSET_DM_FIELDS, fields);
            Marshal.WriteInt16(pDevMode, OFFSET_DM_ORIENTATION, orientVal);
            Marshal.WriteInt16(pDevMode, OFFSET_DM_DEFAULTSOURCE, trayId);
            Marshal.WriteInt16(pDevMode, OFFSET_DM_COLOR, colorVal);

            // Merge public fields into the driver's private DEVMODE data so the hardware honours tray selection
            DocumentProperties(IntPtr.Zero, hPrinter, printerName, pDevMode, pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);
            return pDevMode;
        }

        public static bool ExecutePrintJob(
            string pdfPath,
            string printerName,
            List<PageTrayAssignment> pageTrayAssignments,
            string colorMode = "color",
            int copies = 1,
            string orientation = "Portrait",
            bool reverseOrder = false)
        {
            if (pageTrayAssignments == null || pageTrayAssignments.Count == 0)
                return false;

            var assignments = new List<PageTrayAssignment>(pageTrayAssignments);
            if (reverseOrder)
                assignments.Reverse();

            IntPtr hPrinter = IntPtr.Zero;
            IntPtr hdc = IntPtr.Zero;
            InMemoryRandomAccessStream pdfStream = null;

            try
            {
                PdfDocument doc = PdfHelper.OpenDocument(pdfPath, out pdfStream);

                if (!OpenPrinter(printerName, out hPrinter, IntPtr.Zero) || hPrinter == IntPtr.Zero)
                    throw new InvalidOperationException("Unable to open printer: " + printerName);

                for (int copyIdx = 0; copyIdx < copies; copyIdx++)
                {
                    string currentTray = assignments[0].Tray;
                    IntPtr pDevMode = BuildConfiguredDevMode(hPrinter, printerName, currentTray, orientation, colorMode);

                    try
                    {
                        hdc = CreateDC("WINSPOOL", printerName, null, pDevMode);
                        if (hdc == IntPtr.Zero)
                            throw new InvalidOperationException("Failed to create printer device context.");
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(pDevMode);
                    }

                    var di = new DOCINFOW
                    {
                        cbSize = Marshal.SizeOf(typeof(DOCINFOW)),
                        lpszDocName = string.Format("Print Selector - Copy {0}", copyIdx + 1)
                    };

                    if (StartDoc(hdc, ref di) <= 0)
                        throw new InvalidOperationException("StartDoc failed.");

                    foreach (PageTrayAssignment pt in assignments)
                    {
                        if (!string.Equals(pt.Tray, currentTray, StringComparison.Ordinal))
                        {
                            IntPtr newDevMode = BuildConfiguredDevMode(hPrinter, printerName, pt.Tray, orientation, colorMode);
                            try
                            {
                                ResetDC(hdc, newDevMode);
                                currentTray = pt.Tray;
                            }
                            finally
                            {
                                Marshal.FreeHGlobal(newDevMode);
                            }
                        }

                        StartPage(hdc);

                        int horzRes = GetDeviceCaps(hdc, HORZRES);
                        int vertRes = GetDeviceCaps(hdc, VERTRES);

                        // Rasterise at ~300 DPI (3.125x over 96 DPI base)
                        using (Bitmap pageBmp = PdfHelper.RenderPage(doc, pt.PageIndex, 3.125))
                        {
                            double imgRatio = (double)pageBmp.Width / pageBmp.Height;
                            double pageRatio = (double)horzRes / vertRes;

                            int newW, newH, xOffset, yOffset;
                            if (imgRatio > pageRatio)
                            {
                                newW = horzRes;
                                newH = (int)(horzRes / imgRatio);
                                xOffset = 0;
                                yOffset = (vertRes - newH) / 2;
                            }
                            else
                            {
                                newH = vertRes;
                                newW = (int)(vertRes * imgRatio);
                                yOffset = 0;
                                xOffset = (horzRes - newW) / 2;
                            }

                            DrawBitmapToDc(hdc, pageBmp, xOffset, yOffset, newW, newH);
                        }

                        EndPage(hdc);
                    }

                    EndDoc(hdc);
                    DeleteDC(hdc);
                    hdc = IntPtr.Zero;
                }

                return true;
            }
            catch (Exception ex)
            {
                var ctx = new Dictionary<string, string>
                {
                    { "color_mode", colorMode },
                    { "copies", copies.ToString() },
                    { "orientation", orientation },
                    { "reverse_order", reverseOrder.ToString() }
                };
                CrashLogger.Log("print_job", ex, pdfPath, printerName, ctx);
                throw new InvalidOperationException("Error executing print job: " + ex.Message, ex);
            }
            finally
            {
                if (hdc != IntPtr.Zero)
                {
                    try { DeleteDC(hdc); } catch { }
                }
                if (hPrinter != IntPtr.Zero)
                {
                    try { ClosePrinter(hPrinter); } catch { }
                }
                if (pdfStream != null)
                {
                    try { pdfStream.Dispose(); } catch { }
                }
            }
        }

        private static void DrawBitmapToDc(IntPtr hdc, Bitmap bmp, int destX, int destY, int destW, int destH)
        {
            BitmapData bd = bmp.LockBits(
                new Rectangle(0, 0, bmp.Width, bmp.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format24bppRgb);

            try
            {
                var bmi = new BITMAPINFOHEADER
                {
                    biSize = (uint)Marshal.SizeOf(typeof(BITMAPINFOHEADER)),
                    biWidth = bmp.Width,
                    biHeight = -bmp.Height, // Negative height indicates a top-down DIB
                    biPlanes = 1,
                    biBitCount = 24,
                    biCompression = 0
                };

                StretchDIBits(
                    hdc,
                    destX, destY, destW, destH,
                    0, 0, bmp.Width, bmp.Height,
                    bd.Scan0, ref bmi, 0, SRCCOPY);
            }
            finally
            {
                bmp.UnlockBits(bd);
            }
        }
    }
}
