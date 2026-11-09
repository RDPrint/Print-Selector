using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace PrintSelector
{
    internal static class Converter
    {
        public static bool ConvertToPdf(string inputPath, string outputPath)
        {
            string ext = Path.GetExtension(inputPath).ToLowerInvariant();
            string fullIn = Path.GetFullPath(inputPath);
            string fullOut = Path.GetFullPath(outputPath);

            try
            {
                if (ext == ".pdf")
                {
                    if (!string.Equals(fullIn, fullOut, StringComparison.OrdinalIgnoreCase))
                        File.Copy(fullIn, fullOut, true);
                    return true;
                }

                if (ext == ".docx" || ext == ".doc" || ext == ".odt" || ext == ".rtf" || ext == ".txt")
                {
                    return ConvertWithWord(fullIn, fullOut);
                }

                if (ext == ".xls" || ext == ".xlsx" || ext == ".ods" || ext == ".csv")
                {
                    return ConvertWithExcel(fullIn, fullOut);
                }

                if (ext == ".ppt" || ext == ".pptx" || ext == ".odp")
                {
                    return ConvertWithPowerPoint(fullIn, fullOut);
                }

                if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".tiff" || ext == ".tif")
                {
                    SaveImageAsPdf(fullIn, fullOut);
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                CrashLogger.Log("conversion", ex, inputPath, null, new Dictionary<string, string> { { "target", outputPath } });
                return false;
            }
        }

        private static bool ConvertWithWord(string inputPath, string outputPath)
        {
            dynamic word = null;
            dynamic docs = null;
            dynamic doc = null;
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) throw new InvalidOperationException("Microsoft Word is not installed.");

                word = Activator.CreateInstance(wordType);
                word.Visible = false;
                word.DisplayAlerts = 0;

                docs = word.Documents;
                doc = docs.Open(inputPath, false, true, false);
                doc.SaveAs(outputPath, 17); // 17 = wdFormatPDF
                doc.Close(0);
                return true;
            }
            catch (Exception ex)
            {
                CrashLogger.Log("conversion", ex, inputPath, null,
                    new Dictionary<string, string> { { "converter", "Word COM" }, { "target", outputPath } });
                throw;
            }
            finally
            {
                ReleaseCom(doc);
                ReleaseCom(docs);
                if (word != null)
                {
                    try { word.Quit(0); } catch { }
                    ReleaseCom(word);
                }
            }
        }

        private static bool ConvertWithExcel(string inputPath, string outputPath)
        {
            dynamic excel = null;
            dynamic workbooks = null;
            dynamic wb = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) throw new InvalidOperationException("Microsoft Excel is not installed.");

                excel = Activator.CreateInstance(excelType);
                excel.Visible = false;
                excel.DisplayAlerts = false;

                workbooks = excel.Workbooks;
                wb = workbooks.Open(inputPath, Type.Missing, true);
                wb.ExportAsFixedFormat(0, outputPath); // 0 = xlTypePDF
                wb.Close(false);
                return true;
            }
            catch (Exception ex)
            {
                CrashLogger.Log("conversion", ex, inputPath, null,
                    new Dictionary<string, string> { { "converter", "Excel COM" }, { "target", outputPath } });
                throw;
            }
            finally
            {
                ReleaseCom(wb);
                ReleaseCom(workbooks);
                if (excel != null)
                {
                    try { excel.Quit(); } catch { }
                    ReleaseCom(excel);
                }
            }
        }

        private static bool ConvertWithPowerPoint(string inputPath, string outputPath)
        {
            dynamic ppt = null;
            dynamic presentations = null;
            dynamic pres = null;
            try
            {
                Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
                if (pptType == null) throw new InvalidOperationException("Microsoft PowerPoint is not installed.");

                ppt = Activator.CreateInstance(pptType);
                presentations = ppt.Presentations;
                // ReadOnly = msoTrue (-1), Untitled = msoFalse (0), WithWindow = msoFalse (0)
                pres = presentations.Open(inputPath, -1, 0, 0);
                pres.SaveAs(outputPath, 32); // 32 = ppSaveAsPDF
                pres.Close();
                return true;
            }
            catch (Exception ex)
            {
                CrashLogger.Log("conversion", ex, inputPath, null,
                    new Dictionary<string, string> { { "converter", "PowerPoint COM" }, { "target", outputPath } });
                throw;
            }
            finally
            {
                ReleaseCom(pres);
                ReleaseCom(presentations);
                if (ppt != null)
                {
                    try { ppt.Quit(); } catch { }
                    ReleaseCom(ppt);
                }
            }
        }

        private static void ReleaseCom(object obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                try { Marshal.FinalReleaseComObject(obj); } catch { }
            }
        }

        // Wraps a raster image into a single-page PDF at 100 DPI so the rest of the
        // pipeline can treat images and documents identically.
        private static void SaveImageAsPdf(string imagePath, string pdfOutputPath)
        {
            byte[] jpegBytes;
            int w, h;

            using (Image img = Image.FromFile(imagePath))
            using (Bitmap bmp = new Bitmap(img.Width, img.Height, PixelFormat.Format24bppRgb))
            {
                w = bmp.Width;
                h = bmp.Height;
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.White);
                    g.DrawImage(img, 0, 0, w, h);
                }
                using (MemoryStream ms = new MemoryStream())
                {
                    bmp.Save(ms, ImageFormat.Jpeg);
                    jpegBytes = ms.ToArray();
                }
            }

            double ptW = w * 72.0 / 100.0;
            double ptH = h * 72.0 / 100.0;
            string wStr = ptW.ToString("0.##", CultureInfo.InvariantCulture);
            string hStr = ptH.ToString("0.##", CultureInfo.InvariantCulture);

            using (FileStream fs = new FileStream(pdfOutputPath, FileMode.Create, FileAccess.Write))
            {
                Action<string> writeAscii = s =>
                {
                    byte[] b = Encoding.ASCII.GetBytes(s);
                    fs.Write(b, 0, b.Length);
                };

                long[] offsets = new long[6];
                writeAscii("%PDF-1.4\n");

                offsets[1] = fs.Position;
                writeAscii("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

                offsets[2] = fs.Position;
                writeAscii("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n");

                offsets[3] = fs.Position;
                writeAscii(string.Format(
                    "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im0 4 0 R >> >> /MediaBox [0 0 {0} {1}] /Contents 5 0 R >>\nendobj\n",
                    wStr, hStr));

                offsets[4] = fs.Position;
                writeAscii(string.Format(
                    "4 0 obj\n<< /Type /XObject /Subtype /Image /Width {0} /Height {1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {2} >>\nstream\n",
                    w, h, jpegBytes.Length));
                fs.Write(jpegBytes, 0, jpegBytes.Length);
                writeAscii("\nendstream\nendobj\n");

                string content = string.Format("q {0} 0 0 {1} 0 0 cm /Im0 Do Q\n", wStr, hStr);
                offsets[5] = fs.Position;
                writeAscii(string.Format("5 0 obj\n<< /Length {0} >>\nstream\n{1}endstream\nendobj\n", content.Length, content));

                long xrefPos = fs.Position;
                writeAscii("xref\n0 6\n0000000000 65535 f \n");
                for (int i = 1; i <= 5; i++)
                    writeAscii(string.Format("{0:0000000000} 00000 n \n", offsets[i]));

                writeAscii(string.Format("trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{0}\n%%EOF\n", xrefPos));
            }
        }
    }
}
