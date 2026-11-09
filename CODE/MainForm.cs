using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

namespace PrintSelector
{
    internal sealed class MainForm : Form
    {
        private static readonly Color BgWindow = Color.FromArgb(36, 36, 36);
        private static readonly Color BgSidebar = Color.FromArgb(43, 43, 43);
        private static readonly Color BgPanel = Color.FromArgb(43, 43, 43);
        private static readonly Color BgCard = Color.FromArgb(51, 51, 51);
        private static readonly Color BgInput = Color.FromArgb(58, 58, 58);
        private static readonly Color AccentBlue = Color.FromArgb(31, 106, 165);
        private static readonly Color AccentBlueHover = Color.FromArgb(24, 84, 132);
        private static readonly Color AccentGreen = Color.FromArgb(34, 139, 34);
        private static readonly Color AccentGreenHover = Color.FromArgb(24, 105, 24);
        private static readonly Color AccentRed = Color.FromArgb(184, 53, 53);
        private static readonly Color AccentRedHover = Color.FromArgb(140, 39, 39);
        private static readonly Color TextPrimary = Color.FromArgb(238, 238, 238);
        private static readonly Color TextMuted = Color.FromArgb(160, 160, 160);

        private static readonly string[] AllTrays = new[]
        {
            "Tray 1", "Tray 2", "Tray 3", "Tray 4", "MP Tray", "Auto Select"
        };

        private static readonly string[] CarbonTrays = new[]
        {
            "Tray 1", "Tray 2", "Tray 3", "Tray 4", "MP Tray"
        };

        private string pdfPath;
        private int pageCount;
        private bool isBatchMode;
        private List<string> batchFiles = new List<string>();

        private string profilesPath;
        private Dictionary<string, List<string>> profiles = new Dictionary<string, List<string>>();
        private string currentLanguage = "English";

        private readonly List<PageRowControl> pageRows = new List<PageRowControl>();
        private readonly Dictionary<string, CheckBox> carbonCheckboxes = new Dictionary<string, CheckBox>();
        private readonly Dictionary<string, bool> carbonTrayState = new Dictionary<string, bool>();
        private readonly List<Image> activeThumbnails = new List<Image>();

        // Sidebar widgets
        private Label lblLogo;
        private Button btnSelectFile;
        private Label lblFileInfo;
        private ToggleSwitch swColor;
        private Label lblOrientation;
        private ComboBox optOrientation;
        private Label lblPages;
        private TextBox entryPages;
        private Label lblCopies;
        private TextBox entryCopies;
        private CheckBox chkReverse;
        private ToggleSwitch swCarbon;
        private Label lblCarbonInfo;
        private Button btnPrint;
        private ComboBox optLanguage;

        // Main panel widgets
        private Label lblPrinter;
        private ComboBox printerDropdown;
        private Button btnRefresh;
        private Label lblPageAssignHeader;
        private Panel seqScrollPanel;
        private Label lblBulk;
        private ComboBox bulkTray;
        private Button btnApplyBulk;
        private Label lblProfile;
        private ComboBox optProfile;
        private Button btnSaveProfile;
        private Button btnDeleteProfile;
        private SmoothProgressBar progressBar;

        private bool suppressPageValidation;
        private bool suppressProfileEvent;
        private bool suppressLanguageEvent;

        public MainForm()
        {
            Text = "Print Selector";
            ClientSize = new Size(920, 680);
            MinimumSize = new Size(900, 680);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = BgWindow;
            ForeColor = TextPrimary;
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
            DoubleBuffered = true;
            AllowDrop = true;

            try
            {
                Icon = Icon.ExtractAssociatedIcon(typeof(MainForm).Assembly.Location);
            }
            catch { }

            LoadProfiles();
            BuildUi();
            RefreshPrinters();
            ApplyTranslations();

            DragEnter += OnFormDragEnter;
            DragDrop += OnFormDragDrop;
        }

        private static string GetConfigPath()
        {
            try
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                if (!string.IsNullOrEmpty(appData))
                {
                    string dir = Path.Combine(appData, "Print Selector");
                    Directory.CreateDirectory(dir);
                    return Path.Combine(dir, "PrintSelector_Profiles.json");
                }
            }
            catch
            {
                // Fall back to the executable's directory when running in portable/restricted environments
            }

            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PrintSelector_Profiles.json");
        }

        private void LoadProfiles()
        {
            profilesPath = GetConfigPath();
            profiles.Clear();
            currentLanguage = "English";

            try
            {
                if (File.Exists(profilesPath))
                {
                    string json = File.ReadAllText(profilesPath);
                    var serializer = new JavaScriptSerializer();
                    var raw = serializer.Deserialize<Dictionary<string, object>>(json);
                    if (raw != null)
                    {
                        foreach (var kv in raw)
                        {
                            if (kv.Key == "_LANGUAGE_" && kv.Value is string)
                            {
                                string lang = (string)kv.Value;
                                if (Translations.Dictionary.ContainsKey(lang))
                                    currentLanguage = lang;
                            }
                            else if (kv.Value is System.Collections.IEnumerable && !(kv.Value is string))
                            {
                                var seq = new List<string>();
                                foreach (object item in (System.Collections.IEnumerable)kv.Value)
                                {
                                    if (item != null) seq.Add(item.ToString());
                                }
                                if (seq.Count > 0)
                                    profiles[kv.Key] = seq;
                            }
                        }
                    }
                }
            }
            catch
            {
                // Ignore corrupt profile data and restore default below
            }

            if (!profiles.ContainsKey("Default"))
                profiles["Default"] = new List<string> { "Auto Select" };
        }

        private void SaveProfilesToDisk()
        {
            var payload = new Dictionary<string, object>();
            foreach (var kv in profiles)
                payload[kv.Key] = kv.Value;
            payload["_LANGUAGE_"] = currentLanguage;

            var serializer = new JavaScriptSerializer();
            File.WriteAllText(profilesPath, serializer.Serialize(payload));
        }

        private void BuildUi()
        {
            // Left sidebar
            var sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 220,
                BackColor = BgSidebar,
                Padding = new Padding(20)
            };
            Controls.Add(sidebar);

            lblLogo = new Label
            {
                Text = "Print Selector",
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleCenter,
                Location = new Point(20, 18),
                Size = new Size(180, 34)
            };
            sidebar.Controls.Add(lblLogo);

            btnSelectFile = CreateStyledButton("Select Document", AccentBlue, AccentBlueHover, new Point(20, 62), new Size(180, 34));
            btnSelectFile.Click += (s, e) => SelectFile();
            sidebar.Controls.Add(btnSelectFile);

            lblFileInfo = new Label
            {
                Text = "No file selected\nDrag and drop is supported.",
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 8.75f),
                TextAlign = ContentAlignment.TopCenter,
                Location = new Point(14, 104),
                Size = new Size(192, 42)
            };
            sidebar.Controls.Add(lblFileInfo);

            swColor = new ToggleSwitch
            {
                Text = "Print in Color",
                Checked = true,
                Location = new Point(20, 156),
                Size = new Size(180, 26)
            };
            sidebar.Controls.Add(swColor);

            lblOrientation = new Label
            {
                Text = "Orientation:",
                ForeColor = TextPrimary,
                AutoSize = true,
                Location = new Point(20, 196)
            };
            sidebar.Controls.Add(lblOrientation);

            optOrientation = CreateStyledCombo(new Point(20, 218), 180, new[] { "Portrait", "Landscape" });
            sidebar.Controls.Add(optOrientation);

            lblPages = new Label
            {
                Text = "Pages (e.g. 1-3, 5):",
                ForeColor = TextPrimary,
                AutoSize = true,
                Location = new Point(20, 260)
            };
            sidebar.Controls.Add(lblPages);

            entryPages = new TextBox
            {
                Text = "All",
                Location = new Point(20, 282),
                Size = new Size(180, 26),
                BackColor = BgInput,
                ForeColor = TextPrimary,
                BorderStyle = BorderStyle.FixedSingle
            };
            entryPages.TextChanged += OnPagesTextChanged;
            entryPages.GotFocus += OnPagesGotFocus;
            entryPages.LostFocus += OnPagesLostFocus;
            sidebar.Controls.Add(entryPages);

            lblCopies = new Label
            {
                Text = "Copies:",
                ForeColor = TextPrimary,
                AutoSize = true,
                Location = new Point(20, 322)
            };
            sidebar.Controls.Add(lblCopies);

            entryCopies = new TextBox
            {
                Text = "1",
                Location = new Point(20, 344),
                Size = new Size(85, 26),
                BackColor = BgInput,
                ForeColor = TextPrimary,
                BorderStyle = BorderStyle.FixedSingle
            };
            sidebar.Controls.Add(entryCopies);

            chkReverse = new CheckBox
            {
                Text = "Reverse Order",
                ForeColor = TextPrimary,
                AutoSize = true,
                Location = new Point(20, 384)
            };
            sidebar.Controls.Add(chkReverse);

            swCarbon = new ToggleSwitch
            {
                Text = "Carbon Copy Mode",
                Checked = false,
                Enabled = false,
                Location = new Point(20, 420),
                Size = new Size(180, 26)
            };
            swCarbon.CheckedChanged += (s, e) => ToggleCarbonCopy();
            sidebar.Controls.Add(swCarbon);

            lblCarbonInfo = new Label
            {
                Text = "Same page printed from each selected tray",
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 8f),
                Location = new Point(20, 450),
                Size = new Size(180, 34)
            };
            sidebar.Controls.Add(lblCarbonInfo);

            btnPrint = CreateStyledButton("PRINT", AccentGreen, AccentGreenHover, new Point(20, 502), new Size(180, 38));
            btnPrint.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            btnPrint.Click += (s, e) => StartPrint();
            sidebar.Controls.Add(btnPrint);

            optLanguage = CreateStyledCombo(new Point(20, 556), 180, Translations.Languages);
            suppressLanguageEvent = true;
            optLanguage.SelectedItem = currentLanguage;
            suppressLanguageEvent = false;
            optLanguage.SelectedIndexChanged += (s, e) =>
            {
                if (!suppressLanguageEvent && optLanguage.SelectedItem != null)
                    ChangeLanguage(optLanguage.SelectedItem.ToString());
            };
            sidebar.Controls.Add(optLanguage);

            // Right main content area
            var mainContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(18),
                BackColor = BgWindow
            };
            Controls.Add(mainContainer);
            mainContainer.BringToFront();

            var mainFrame = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = BgPanel,
                Padding = new Padding(14)
            };
            mainContainer.Controls.Add(mainFrame);

            // Top printer bar
            var topBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 46,
                BackColor = Color.Transparent
            };
            mainFrame.Controls.Add(topBar);

            lblPrinter = new Label
            {
                Text = "Select Printer:",
                AutoSize = true,
                Location = new Point(4, 13),
                ForeColor = TextPrimary
            };
            topBar.Controls.Add(lblPrinter);

            btnRefresh = CreateStyledButton("Refresh", AccentBlue, AccentBlueHover, new Point(0, 7), new Size(92, 30));
            btnRefresh.Click += (s, e) => RefreshPrinters();
            topBar.Controls.Add(btnRefresh);

            printerDropdown = CreateStyledCombo(new Point(116, 10), 300, new[] { "No Printers Found" });
            topBar.Controls.Add(printerDropdown);

            topBar.Resize += (s, e) => LayoutTopBar(topBar);

            // Bottom progress & toolbar
            var bottomSection = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 76,
                BackColor = Color.Transparent
            };
            mainFrame.Controls.Add(bottomSection);

            var toolbar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 48,
                BackColor = BgCard,
                Padding = new Padding(8)
            };
            bottomSection.Controls.Add(toolbar);

            lblBulk = new Label
            {
                Text = "Fill all as:",
                AutoSize = true,
                Location = new Point(10, 15),
                ForeColor = TextPrimary
            };
            toolbar.Controls.Add(lblBulk);

            bulkTray = CreateStyledCombo(new Point(85, 11), 105, AllTrays);
            toolbar.Controls.Add(bulkTray);

            btnApplyBulk = CreateStyledButton("Apply", AccentBlue, AccentBlueHover, new Point(196, 9), new Size(68, 30));
            btnApplyBulk.Click += (s, e) => ApplyBulk();
            toolbar.Controls.Add(btnApplyBulk);

            lblProfile = new Label
            {
                Text = "Template:",
                AutoSize = true,
                ForeColor = TextPrimary
            };
            toolbar.Controls.Add(lblProfile);

            optProfile = CreateStyledCombo(new Point(0, 11), 115, profiles.Keys.ToArray());
            suppressProfileEvent = true;
            optProfile.SelectedItem = "Default";
            suppressProfileEvent = false;
            optProfile.SelectedIndexChanged += (s, e) =>
            {
                if (!suppressProfileEvent && optProfile.SelectedItem != null)
                    ApplyProfile(optProfile.SelectedItem.ToString());
            };
            toolbar.Controls.Add(optProfile);

            btnSaveProfile = CreateStyledButton("Save As", AccentBlue, AccentBlueHover, new Point(0, 9), new Size(80, 30));
            btnSaveProfile.Click += (s, e) => SaveProfile();
            toolbar.Controls.Add(btnSaveProfile);

            btnDeleteProfile = CreateStyledButton("Delete", AccentRed, AccentRedHover, new Point(0, 9), new Size(72, 30));
            btnDeleteProfile.Click += (s, e) => DeleteProfile();
            toolbar.Controls.Add(btnDeleteProfile);

            toolbar.Resize += (s, e) => LayoutBottomToolbar(toolbar);

            progressBar = new SmoothProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 12,
                Value = 0f
            };
            bottomSection.Controls.Add(progressBar);

            // Center scrollable page assignments box
            var centerCard = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = BgCard,
                Padding = new Padding(10)
            };
            mainFrame.Controls.Add(centerCard);
            centerCard.BringToFront();

            lblPageAssignHeader = new Label
            {
                Text = "Page Assignments",
                Dock = DockStyle.Top,
                Height = 28,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                ForeColor = TextPrimary,
                TextAlign = ContentAlignment.MiddleLeft
            };
            centerCard.Controls.Add(lblPageAssignHeader);

            seqScrollPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(46, 46, 46)
            };
            centerCard.Controls.Add(seqScrollPanel);
            seqScrollPanel.BringToFront();

            LayoutTopBar(topBar);
            LayoutBottomToolbar(toolbar);
        }

        private void LayoutTopBar(Control topBar)
        {
            if (topBar == null || lblPrinter == null || printerDropdown == null || btnRefresh == null) return;
            int refreshW = Math.Max(84, TextRenderer.MeasureText(btnRefresh.Text, btnRefresh.Font).Width + 24);
            btnRefresh.Width = refreshW;
            btnRefresh.Location = new Point(topBar.ClientSize.Width - refreshW - 4, 7);

            int dropX = lblPrinter.Right + 10;
            int dropW = Math.Max(140, btnRefresh.Left - dropX - 10);
            printerDropdown.Location = new Point(dropX, 10);
            printerDropdown.Width = dropW;
        }

        private void LayoutBottomToolbar(Control toolbar)
        {
            if (toolbar == null || lblBulk == null || bulkTray == null || btnApplyBulk == null) return;

            lblBulk.Location = new Point(10, 15);
            bulkTray.Location = new Point(lblBulk.Right + 6, 11);

            int applyW = Math.Max(64, TextRenderer.MeasureText(btnApplyBulk.Text, btnApplyBulk.Font).Width + 20);
            btnApplyBulk.Width = applyW;
            btnApplyBulk.Location = new Point(bulkTray.Right + 6, 9);

            int delW = Math.Max(68, TextRenderer.MeasureText(btnDeleteProfile.Text, btnDeleteProfile.Font).Width + 20);
            btnDeleteProfile.Width = delW;
            btnDeleteProfile.Location = new Point(toolbar.ClientSize.Width - delW - 8, 9);

            int saveW = Math.Max(74, TextRenderer.MeasureText(btnSaveProfile.Text, btnSaveProfile.Font).Width + 20);
            btnSaveProfile.Width = saveW;
            btnSaveProfile.Location = new Point(btnDeleteProfile.Left - saveW - 6, 9);

            optProfile.Location = new Point(btnSaveProfile.Left - optProfile.Width - 6, 11);
            lblProfile.Location = new Point(optProfile.Left - lblProfile.Width - 6, 15);
        }

        private static Button CreateStyledButton(string text, Color backColor, Color hoverColor, Point location, Size size)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                BackColor = backColor,
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = hoverColor;
            return btn;
        }

        private static ComboBox CreateStyledCombo(Point location, int width, string[] items)
        {
            var cb = new ComboBox
            {
                Location = location,
                Width = width,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                BackColor = BgInput,
                ForeColor = TextPrimary
            };
            cb.Items.AddRange(items);
            if (items.Length > 0) cb.SelectedIndex = 0;
            return cb;
        }

        // Page range input filtering
        private void OnPagesTextChanged(object sender, EventArgs e)
        {
            if (suppressPageValidation) return;
            string val = entryPages.Text;
            if (val == "All") return;
            if (val == "")
            {
                BeginInvoke(new Action(() =>
                {
                    if (entryPages.Text == "" && !entryPages.Focused)
                    {
                        suppressPageValidation = true;
                        entryPages.Text = "All";
                        suppressPageValidation = false;
                    }
                }));
                return;
            }
            string lower = val.ToLowerInvariant();
            if (lower == "a" || lower == "al" || lower == "all") return;

            var cleanedChars = val.Where(c => (c >= '0' && c <= '9') || c == '-' || c == ',' || c == ' ').ToArray();
            string cleaned = new string(cleanedChars);
            if (cleaned != val)
            {
                int sel = entryPages.SelectionStart;
                suppressPageValidation = true;
                entryPages.Text = cleaned;
                entryPages.SelectionStart = Math.Min(sel, cleaned.Length);
                suppressPageValidation = false;
            }
        }

        private void OnPagesGotFocus(object sender, EventArgs e)
        {
            if (entryPages.Text == "All")
            {
                suppressPageValidation = true;
                entryPages.Text = "";
                suppressPageValidation = false;
            }
        }

        private void OnPagesLostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(entryPages.Text))
            {
                suppressPageValidation = true;
                entryPages.Text = "All";
                suppressPageValidation = false;
            }
        }

        // Printer discovery
        private void RefreshPrinters()
        {
            List<string> printers = PrinterEngine.GetInstalledPrinters();
            printerDropdown.Items.Clear();

            if (printers.Count > 0)
            {
                printerDropdown.Items.AddRange(printers.ToArray());
                printerDropdown.SelectedIndex = 0;
            }
            else
            {
                string noneText = D("NO_PRINTER_FRD");
                printerDropdown.Items.Add(noneText);
                printerDropdown.SelectedIndex = 0;
            }
        }

        // Drag and drop
        private void OnFormDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }

        private void OnFormDragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (files == null || files.Length == 0) return;

            if (files.Length > 1)
            {
                isBatchMode = true;
                batchFiles = new List<string>(files);
                lblFileInfo.Text = D("BATCH_MODE").Replace("{x}", files.Length.ToString());
                btnPrint.Text = D("BATCH_PRINT");
                entryPages.Enabled = false;
                ClearSequenceRows();
            }
            else
            {
                isBatchMode = false;
                batchFiles.Clear();
                btnPrint.Text = D("PRINT");
                entryPages.Enabled = true;
                LoadFile(files[0]);
            }
        }

        // File picker & conversion
        private void SelectFile()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Documents|*.pdf;*.docx;*.doc;*.odt;*.rtf;*.txt;*.xlsx;*.xls;*.csv;*.ods;*.pptx;*.ppt;*.odp;*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.tif|All Files|*.*";
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    isBatchMode = false;
                    batchFiles.Clear();
                    btnPrint.Text = D("PRINT");
                    entryPages.Enabled = true;
                    LoadFile(dlg.FileName);
                }
            }
        }

        private void LoadFile(string filePath)
        {
            lblFileInfo.Text = D("CONVERTING").Replace("{file}", Path.GetFileName(filePath));
            btnPrint.Enabled = false;
            Update();

            var worker = new Thread(() =>
            {
                string outPdf = Path.Combine(Path.GetTempPath(), "selector_temp.pdf");
                bool ok;
                try
                {
                    ok = Converter.ConvertToPdf(filePath, outPdf);
                }
                catch (Exception ex)
                {
                    CrashLogger.Log("conversion", ex, filePath);
                    ok = false;
                }

                BeginInvoke(new Action(() => FinishLoadFile(filePath, outPdf, ok)));
            });
            worker.IsBackground = true;
            // Office COM interop works best on an STA worker thread
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void FinishLoadFile(string filePath, string outPdf, bool success)
        {
            btnPrint.Enabled = true;
            if (success && File.Exists(outPdf))
            {
                pdfPath = outPdf;
                try
                {
                    pageCount = PdfHelper.GetPageCount(pdfPath);
                    lblFileInfo.Text = D("FILE_PAGES")
                        .Replace("{file}", Path.GetFileName(filePath))
                        .Replace("{pc}", pageCount.ToString());

                    if (pageCount == 1)
                    {
                        swCarbon.Enabled = true;
                    }
                    else
                    {
                        swCarbon.Checked = false;
                        swCarbon.Enabled = false;
                    }

                    if (swCarbon.Checked)
                        entryPages.Enabled = false;

                    BuildSequenceUi();
                }
                catch (Exception ex)
                {
                    CrashLogger.Log("file_load", ex, filePath);
                    MessageBox.Show(this, D("PDF_ERR").Replace("{e}", ex.Message), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                lblFileInfo.Text = D("LOAD_ERR");
                MessageBox.Show(this, D("LOAD_ERR_DET"), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Profile templates
        private void RefreshProfileDropdown(string selectedName)
        {
            suppressProfileEvent = true;
            optProfile.Items.Clear();
            foreach (string key in profiles.Keys)
                optProfile.Items.Add(key);
            if (optProfile.Items.Contains(selectedName))
                optProfile.SelectedItem = selectedName;
            else if (optProfile.Items.Count > 0)
                optProfile.SelectedIndex = 0;
            suppressProfileEvent = false;
        }

        private void SaveProfile()
        {
            if (pageRows.Count == 0 && !isBatchMode && !(swCarbon.Checked && carbonCheckboxes.Count > 0))
            {
                MessageBox.Show(this, D("NO_SEQ_SAVE"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string name = PromptForTemplateName(D("TITLE_SAVE_PROFILE"), D("ENTER_NAME"));
            if (string.IsNullOrWhiteSpace(name)) return;

            List<string> seq;
            if (isBatchMode)
            {
                string cur = optProfile.SelectedItem != null ? optProfile.SelectedItem.ToString() : "Default";
                seq = profiles.ContainsKey(cur) ? new List<string>(profiles[cur]) : new List<string> { "Auto Select" };
            }
            else if (swCarbon.Checked && pageCount == 1)
            {
                seq = carbonCheckboxes.Where(kv => kv.Value.Checked).Select(kv => kv.Key).ToList();
                if (seq.Count == 0)
                {
                    MessageBox.Show(this, D("NO_TRAY_SEL"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            else
            {
                seq = pageRows.Select(r => r.SelectedTray).ToList();
            }

            profiles[name] = seq;
            try
            {
                SaveProfilesToDisk();
                RefreshProfileDropdown(name);
                MessageBox.Show(this, D("SAVE_OK").Replace("{name}", name), D("SUCCESS"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, D("SAVE_FAIL").Replace("{e}", ex.Message), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteProfile()
        {
            string name = optProfile.SelectedItem != null ? optProfile.SelectedItem.ToString() : "Default";
            if (name == "Default")
            {
                MessageBox.Show(this, D("NO_DEL_DEF"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                D("CONFIRM_DEL").Replace("{name}", name),
                D("TITLE_CONFIRM_DEL"),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (confirm == DialogResult.Yes && profiles.ContainsKey(name))
            {
                profiles.Remove(name);
                try
                {
                    SaveProfilesToDisk();
                    RefreshProfileDropdown("Default");
                    ApplyProfile("Default");
                    MessageBox.Show(this, D("DEL_OK").Replace("{name}", name), D("SUCCESS"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, D("DEL_FAIL").Replace("{e}", ex.Message), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ApplyProfile(string profileName)
        {
            List<string> seq;
            if (!profiles.TryGetValue(profileName, out seq) || seq == null || seq.Count == 0)
                seq = new List<string> { "Auto Select" };

            if (swCarbon.Checked && pageCount == 1)
            {
                foreach (var kv in carbonCheckboxes)
                    kv.Value.Checked = seq.Contains(kv.Key);
            }
            else
            {
                for (int i = 0; i < pageRows.Count; i++)
                {
                    string tray = i < seq.Count ? seq[i] : seq[seq.Count - 1];
                    pageRows[i].SelectedTray = tray;
                }
            }
        }

        private void ApplyBulk()
        {
            if (bulkTray.SelectedItem == null) return;
            string t = bulkTray.SelectedItem.ToString();
            foreach (PageRowControl row in pageRows)
                row.SelectedTray = t;
        }

        // Page preview window
        private void ShowPreview(int pageIndex)
        {
            var previewForm = new Form
            {
                Text = D("TITLE_PREVIEW").Replace("{i}", (pageIndex + 1).ToString()),
                Size = new Size(720, 860),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = BgWindow,
                ForeColor = TextPrimary,
                Icon = Icon
            };

            try
            {
                Bitmap bmp = PdfHelper.RenderSinglePage(pdfPath, pageIndex, 1.5);
                previewForm.FormClosed += (s, e) => bmp.Dispose();

                var scroll = new Panel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    BackColor = BgWindow
                };
                var pic = new PictureBox
                {
                    Image = bmp,
                    SizeMode = PictureBoxSizeMode.AutoSize,
                    Location = new Point(20, 20)
                };
                scroll.Controls.Add(pic);
                scroll.Resize += (s, e) =>
                {
                    int x = Math.Max(20, (scroll.ClientSize.Width - pic.Width) / 2);
                    pic.Location = new Point(x, 20);
                };
                previewForm.Controls.Add(scroll);
                previewForm.Show(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, D("PREVIEW_ERR").Replace("{e}", ex.Message), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                previewForm.Dispose();
            }
        }

        // Per-page & Carbon Copy UI construction
        private void ClearSequenceRows()
        {
            foreach (var kv in carbonCheckboxes)
                carbonTrayState[kv.Key] = kv.Value.Checked;

            seqScrollPanel.SuspendLayout();
            seqScrollPanel.Controls.Clear();
            pageRows.Clear();
            carbonCheckboxes.Clear();

            foreach (Image img in activeThumbnails)
            {
                try { img.Dispose(); } catch { }
            }
            activeThumbnails.Clear();
            seqScrollPanel.ResumeLayout();
        }

        private void BuildSequenceUi()
        {
            ClearSequenceRows();
            if (isBatchMode || string.IsNullOrEmpty(pdfPath) || pageCount <= 0)
                return;

            if (swCarbon.Checked && pageCount == 1)
            {
                BuildCarbonCopyUi();
                return;
            }

            string curProfile = optProfile.SelectedItem != null ? optProfile.SelectedItem.ToString() : "Default";
            List<string> activeSeq;
            if (!profiles.TryGetValue(curProfile, out activeSeq) || activeSeq == null || activeSeq.Count == 0)
                activeSeq = new List<string> { "Auto Select" };

            seqScrollPanel.SuspendLayout();
            InMemoryRandomAccessStream stream = null;
            try
            {
                PdfDocument doc = PdfHelper.OpenDocument(pdfPath, out stream);
                int y = 8;

                for (int i = 0; i < pageCount; i++)
                {
                    Bitmap thumb = null;
                    try
                    {
                        thumb = PdfHelper.RenderPage(doc, i, 0.06);
                        activeThumbnails.Add(thumb);
                    }
                    catch { }

                    string initialTray = i < activeSeq.Count ? activeSeq[i] : activeSeq[activeSeq.Count - 1];
                    int capturedPage = i;
                    var row = new PageRowControl(i + 1, thumb, AllTrays, initialTray, () => ShowPreview(capturedPage))
                    {
                        Location = new Point(8, y),
                        Width = Math.Max(380, seqScrollPanel.ClientSize.Width - 24)
                    };
                    seqScrollPanel.Controls.Add(row);
                    pageRows.Add(row);
                    y += row.Height + 6;
                }
            }
            finally
            {
                if (stream != null) stream.Dispose();
                seqScrollPanel.ResumeLayout();
            }
        }

        private void ToggleCarbonCopy()
        {
            if (swCarbon.Checked)
            {
                entryPages.Enabled = false;
                suppressPageValidation = true;
                entryPages.Text = "All";
                suppressPageValidation = false;
            }
            else
            {
                entryPages.Enabled = true;
            }
            BuildSequenceUi();
        }

        private void BuildCarbonCopyUi()
        {
            seqScrollPanel.SuspendLayout();
            try
            {
                int y = 12;
                var headerPanel = new Panel
                {
                    Location = new Point(10, y),
                    Size = new Size(360, 80),
                    BackColor = Color.Transparent
                };

                try
                {
                    Bitmap thumb = PdfHelper.RenderSinglePage(pdfPath, 0, 0.09);
                    activeThumbnails.Add(thumb);
                    var pic = new PictureBox
                    {
                        Image = thumb,
                        SizeMode = PictureBoxSizeMode.AutoSize,
                        Location = new Point(4, 4),
                        Cursor = Cursors.Hand
                    };
                    pic.Click += (s, e) => ShowPreview(0);
                    headerPanel.Controls.Add(pic);
                    headerPanel.Height = Math.Max(60, pic.Bottom + 8);
                }
                catch { }

                var lblSelectTrays = new Label
                {
                    Text = D("SELECT_TRAYS"),
                    Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                    ForeColor = TextPrimary,
                    AutoSize = true,
                    Location = new Point(78, 22)
                };
                headerPanel.Controls.Add(lblSelectTrays);
                seqScrollPanel.Controls.Add(headerPanel);

                y += headerPanel.Height + 6;

                for (int i = 0; i < CarbonTrays.Length; i++)
                {
                    string trayName = CarbonTrays[i];
                    bool isChecked;
                    if (!carbonTrayState.TryGetValue(trayName, out isChecked))
                        isChecked = (i == 0);

                    var chk = new CheckBox
                    {
                        Text = trayName,
                        Checked = isChecked,
                        ForeColor = TextPrimary,
                        AutoSize = true,
                        Location = new Point(32, y)
                    };
                    seqScrollPanel.Controls.Add(chk);
                    carbonCheckboxes[trayName] = chk;
                    y += 30;
                }
            }
            finally
            {
                seqScrollPanel.ResumeLayout();
            }
        }

        // Print job dispatch
        private void StartPrint()
        {
            string printer = printerDropdown.SelectedItem != null ? printerDropdown.SelectedItem.ToString() : "None";
            if (printer == "None" || printer == D("NO_PRINTER_FRD") || string.IsNullOrWhiteSpace(printer))
            {
                MessageBox.Show(this, D("NO_PRINTER_SEL"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string colorMode = swColor.Checked ? "color" : "mono";
            string selectedOrient = optOrientation.SelectedItem != null ? optOrientation.SelectedItem.ToString() : "Portrait";
            bool isLandscape = Translations.Dictionary.Values.Any(d =>
            {
                string land;
                return d.TryGetValue("LANDSCAPE", out land) && land == selectedOrient;
            });
            string orientation = isLandscape ? "Landscape" : "Portrait";
            bool reverseOrder = chkReverse.Checked;

            int copies;
            if (!int.TryParse(entryCopies.Text.Trim(), out copies) || copies < 1)
                copies = 1;

            if (isBatchMode)
            {
                btnPrint.Enabled = false;
                btnPrint.Text = D("PRINTING_BATCH");
                progressBar.Value = 0.5f;
                StartBatchPrint(printer, colorMode, copies, orientation, reverseOrder);
                return;
            }

            if (string.IsNullOrEmpty(pdfPath) || pageCount == 0)
            {
                MessageBox.Show(this, D("REQ_DOC"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Single-page carbon copy mode
            if (swCarbon.Checked && pageCount == 1)
            {
                List<string> selectedTrays = carbonCheckboxes.Where(kv => kv.Value.Checked).Select(kv => kv.Key).ToList();
                if (selectedTrays.Count == 0)
                {
                    MessageBox.Show(this, D("NO_TRAY_SEL"), D("WARN"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sequence = selectedTrays.Select(t => new PageTrayAssignment { PageIndex = 0, Tray = t }).ToList();
                btnPrint.Enabled = false;
                btnPrint.Text = D("GEN_PCL");
                progressBar.Value = 0.5f;

                var thread = new Thread(() =>
                {
                    try
                    {
                        bool ok = PrinterEngine.ExecutePrintJob(pdfPath, printer, sequence, colorMode, copies, orientation, reverseOrder);
                        BeginInvoke(new Action(() => FinishPrint(ok, null)));
                    }
                    catch (Exception ex)
                    {
                        CrashLogger.Log("print_job", ex, pdfPath, printer,
                            new Dictionary<string, string> { { "mode", "carbon_copy" }, { "trays", string.Join(", ", selectedTrays) } });
                        BeginInvoke(new Action(() => FinishPrint(false, ex.Message)));
                    }
                }) { IsBackground = true };
                thread.Start();
                return;
            }

            // Standard per-page mode
            string pageStr = entryPages.Text.Trim();
            List<int> pagesToPrint = ParsePageRange(pageStr, pageCount);
            if (pagesToPrint.Count == 0)
            {
                MessageBox.Show(this, D("REQ_NO_PAGES"), D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var pageSequence = new List<PageTrayAssignment>();
            foreach (int pIdx in pagesToPrint)
            {
                string tray = pIdx < pageRows.Count ? pageRows[pIdx].SelectedTray : "Auto Select";
                pageSequence.Add(new PageTrayAssignment { PageIndex = pIdx, Tray = tray });
            }

            btnPrint.Enabled = false;
            btnPrint.Text = D("GEN_PCL");
            progressBar.Value = 0.5f;

            var printThread = new Thread(() =>
            {
                try
                {
                    bool ok = PrinterEngine.ExecutePrintJob(pdfPath, printer, pageSequence, colorMode, copies, orientation, reverseOrder);
                    BeginInvoke(new Action(() => FinishPrint(ok, null)));
                }
                catch (Exception ex)
                {
                    CrashLogger.Log("print_job", ex, pdfPath, printer,
                        new Dictionary<string, string> { { "mode", "standard" }, { "page_range", pageStr } });
                    BeginInvoke(new Action(() => FinishPrint(false, ex.Message)));
                }
            }) { IsBackground = true };
            printThread.Start();
        }

        private static List<int> ParsePageRange(string pageStr, int totalPages)
        {
            if (string.IsNullOrWhiteSpace(pageStr) || string.Equals(pageStr, "all", StringComparison.OrdinalIgnoreCase))
                return Enumerable.Range(0, totalPages).ToList();

            var pages = new SortedSet<int>();
            foreach (string rawPart in pageStr.Split(','))
            {
                string part = rawPart.Trim();
                if (string.IsNullOrEmpty(part)) continue;

                if (part.Contains("-"))
                {
                    string[] bounds = part.Split('-');
                    int start, end;
                    if (bounds.Length == 2 && int.TryParse(bounds[0].Trim(), out start) && int.TryParse(bounds[1].Trim(), out end))
                    {
                        if (start >= 1 && start <= totalPages && end >= 1 && end <= totalPages && start <= end)
                        {
                            for (int p = start; p <= end; p++)
                                pages.Add(p - 1);
                        }
                    }
                }
                else
                {
                    int p;
                    if (int.TryParse(part, out p) && p >= 1 && p <= totalPages)
                        pages.Add(p - 1);
                }
            }
            return pages.ToList();
        }

        private void StartBatchPrint(string printer, string colorMode, int copies, string orientation, bool reverseOrder)
        {
            string curProfile = optProfile.SelectedItem != null ? optProfile.SelectedItem.ToString() : "Default";
            List<string> activeSeq;
            if (!profiles.TryGetValue(curProfile, out activeSeq) || activeSeq == null || activeSeq.Count == 0)
                activeSeq = new List<string> { "Auto Select" };

            var worker = new Thread(() =>
            {
                for (int i = 0; i < batchFiles.Count; i++)
                {
                    string fp = batchFiles[i];
                    int indexDisplay = i + 1;
                    BeginInvoke(new Action(() =>
                    {
                        lblFileInfo.Text = D("BATCH_PROC")
                            .Replace("{file}", Path.GetFileName(fp))
                            .Replace("{i}", indexDisplay.ToString())
                            .Replace("{total}", batchFiles.Count.ToString());
                    }));

                    string outPdf = Path.Combine(Path.GetTempPath(), "selector_temp.pdf");
                    if (Converter.ConvertToPdf(fp, outPdf))
                    {
                        try
                        {
                            int pc = PdfHelper.GetPageCount(outPdf);
                            var seqConfig = new List<PageTrayAssignment>();
                            for (int pg = 0; pg < pc; pg++)
                            {
                                string tr = pg < activeSeq.Count ? activeSeq[pg] : activeSeq[activeSeq.Count - 1];
                                seqConfig.Add(new PageTrayAssignment { PageIndex = pg, Tray = tr });
                            }

                            bool ok = PrinterEngine.ExecutePrintJob(outPdf, printer, seqConfig, colorMode, copies, orientation, reverseOrder);
                            if (!ok)
                            {
                                CrashLogger.Log("batch_print", new InvalidOperationException("Batch print returned failure for " + fp), fp, printer);
                            }
                        }
                        catch (Exception ex)
                        {
                            CrashLogger.Log("batch_print", ex, fp, printer,
                                new Dictionary<string, string>
                                {
                                    { "mode", "batch" },
                                    { "batch_index", indexDisplay.ToString() },
                                    { "batch_total", batchFiles.Count.ToString() }
                                });
                        }
                    }
                }
                BeginInvoke(new Action(() => FinishPrint(true, null)));
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void FinishPrint(bool success, string errorDetails)
        {
            btnPrint.Enabled = true;
            btnPrint.Text = isBatchMode ? D("BATCH_PRINT") : D("PRINT");

            if (success)
            {
                progressBar.Value = 1.0f;
                MessageBox.Show(this, D("PRINT_OK"), D("SUCCESS"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                progressBar.Value = 0f;
            }
            else
            {
                progressBar.Value = 0f;
                string msg = D("SPOOL_ERR");
                if (!string.IsNullOrEmpty(errorDetails))
                    msg += "\n\nDetails:\n" + errorDetails;
                MessageBox.Show(this, msg, D("ERR"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Localisation
        private string D(string key)
        {
            Dictionary<string, string> langDict;
            if (!Translations.Dictionary.TryGetValue(currentLanguage, out langDict))
                langDict = Translations.Dictionary["English"];

            string val;
            return langDict.TryGetValue(key, out val) ? val : key;
        }

        private void ChangeLanguage(string newLang)
        {
            if (!Translations.Dictionary.ContainsKey(newLang)) return;
            currentLanguage = newLang;
            try { SaveProfilesToDisk(); } catch { }
            ApplyTranslations();
        }

        private void ApplyTranslations()
        {
            Dictionary<string, string> d = Translations.Dictionary[currentLanguage];
            Text = d["TITLE"];
            lblLogo.Text = d["TITLE"];
            btnSelectFile.Text = d["SELECT_DOC"];

            if (string.IsNullOrEmpty(pdfPath) && !isBatchMode)
                lblFileInfo.Text = d["NO_FILE"];

            swColor.Text = d["COLOR"];
            lblOrientation.Text = d["ORIENTATION"];

            string curOrient = optOrientation.SelectedItem != null ? optOrientation.SelectedItem.ToString() : "Portrait";
            bool isLandscape = Translations.Dictionary.Values.Any(langDict =>
            {
                string land;
                return langDict.TryGetValue("LANDSCAPE", out land) && land == curOrient;
            });

            optOrientation.Items.Clear();
            optOrientation.Items.Add(d["PORTRAIT"]);
            optOrientation.Items.Add(d["LANDSCAPE"]);
            optOrientation.SelectedItem = isLandscape ? d["LANDSCAPE"] : d["PORTRAIT"];

            lblPages.Text = d["PAGES"];
            lblCopies.Text = d["COPIES"];
            chkReverse.Text = d["REVERSE"];
            btnPrint.Text = isBatchMode ? d["BATCH_PRINT"] : d["PRINT"];
            lblPrinter.Text = d["PRINTER_SELECT"];

            if (printerDropdown.Items.Count == 1)
            {
                string first = printerDropdown.Items[0].ToString();
                bool isPlaceholder = Translations.Dictionary.Values.Any(langDict =>
                {
                    string nf;
                    return langDict.TryGetValue("NO_PRINTER_FRD", out nf) && nf == first;
                });
                if (isPlaceholder)
                {
                    printerDropdown.Items.Clear();
                    printerDropdown.Items.Add(d["NO_PRINTER_FRD"]);
                    printerDropdown.SelectedIndex = 0;
                }
            }

            btnRefresh.Text = d["REFRESH"];
            lblPageAssignHeader.Text = d["PAGE_ASSIGN"];
            lblBulk.Text = d["BULK_FILL"];
            btnApplyBulk.Text = d["APPLY"];
            lblProfile.Text = d["TEMPLATE"];
            btnSaveProfile.Text = d["SAVE_AS"];
            btnDeleteProfile.Text = d["DELETE"];
            swCarbon.Text = d["CARBON_COPY"];
            lblCarbonInfo.Text = d["CARBON_INFO"];

            if (lblPrinter.Parent != null) LayoutTopBar(lblPrinter.Parent);
            if (lblBulk.Parent != null) LayoutBottomToolbar(lblBulk.Parent);
        }

        private string PromptForTemplateName(string title, string prompt)
        {
            using (var dlg = new Form())
            {
                dlg.Text = title;
                dlg.ClientSize = new Size(340, 145);
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.BackColor = BgPanel;
                dlg.ForeColor = TextPrimary;

                var lbl = new Label
                {
                    Text = prompt,
                    Location = new Point(18, 18),
                    Size = new Size(304, 24),
                    ForeColor = TextPrimary
                };
                var txt = new TextBox
                {
                    Location = new Point(18, 48),
                    Size = new Size(304, 26),
                    BackColor = BgInput,
                    ForeColor = TextPrimary,
                    BorderStyle = BorderStyle.FixedSingle
                };
                var btnOk = CreateStyledButton("OK", AccentBlue, AccentBlueHover, new Point(156, 92), new Size(80, 30));
                btnOk.DialogResult = DialogResult.OK;
                var btnCancel = CreateStyledButton("Cancel", BgCard, BgInput, new Point(242, 92), new Size(80, 30));
                btnCancel.DialogResult = DialogResult.Cancel;

                dlg.Controls.AddRange(new Control[] { lbl, txt, btnOk, btnCancel });
                dlg.AcceptButton = btnOk;
                dlg.CancelButton = btnCancel;

                return dlg.ShowDialog(this) == DialogResult.OK ? txt.Text.Trim() : null;
            }
        }
    }

    // Individual page row with thumbnail and tray selector
    internal sealed class PageRowControl : Panel
    {
        private readonly ComboBox comboTray;

        public string SelectedTray
        {
            get { return comboTray.SelectedItem != null ? comboTray.SelectedItem.ToString() : "Auto Select"; }
            set
            {
                if (comboTray.Items.Contains(value))
                    comboTray.SelectedItem = value;
            }
        }

        public PageRowControl(int pageNumber, Image thumbnail, string[] trays, string initialTray, Action onThumbnailClick)
        {
            Height = 52;
            BackColor = Color.FromArgb(54, 54, 54);

            if (thumbnail != null)
            {
                var pic = new PictureBox
                {
                    Image = thumbnail,
                    SizeMode = PictureBoxSizeMode.Zoom,
                    Location = new Point(10, 4),
                    Size = new Size(34, 44),
                    Cursor = Cursors.Hand
                };
                pic.Click += (s, e) => { if (onThumbnailClick != null) onThumbnailClick(); };
                Controls.Add(pic);
            }

            var lbl = new Label
            {
                Text = "Page " + pageNumber,
                ForeColor = Color.FromArgb(238, 238, 238),
                AutoSize = true,
                Location = new Point(56, 16)
            };
            Controls.Add(lbl);

            comboTray = new ComboBox
            {
                Location = new Point(136, 13),
                Width = 140,
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(64, 64, 64),
                ForeColor = Color.FromArgb(238, 238, 238)
            };
            comboTray.Items.AddRange(trays);
            if (comboTray.Items.Contains(initialTray))
                comboTray.SelectedItem = initialTray;
            else if (comboTray.Items.Count > 0)
                comboTray.SelectedIndex = 0;

            Controls.Add(comboTray);
        }
    }

    // Pill-style toggle switch matching the CustomTkinter switch appearance
    internal sealed class ToggleSwitch : Control
    {
        private bool isChecked;
        public event EventHandler CheckedChanged;

        public bool Checked
        {
            get { return isChecked; }
            set
            {
                if (isChecked != value)
                {
                    isChecked = value;
                    Invalidate();
                    if (CheckedChanged != null)
                        CheckedChanged(this, EventArgs.Empty);
                }
            }
        }

        public ToggleSwitch()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.UserPaint | ControlStyles.SupportsTransparentBackColor, true);
            Cursor = Cursors.Hand;
            BackColor = Color.Transparent;
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            if (Enabled)
                Checked = !Checked;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            int trackW = 38;
            int trackH = 20;
            int trackY = (Height - trackH) / 2;

            Color trackColor = !Enabled
                ? Color.FromArgb(70, 70, 70)
                : (isChecked ? Color.FromArgb(31, 106, 165) : Color.FromArgb(90, 90, 90));

            using (GraphicsPath path = RoundedRect(new Rectangle(0, trackY, trackW, trackH), trackH))
            using (Brush trackBrush = new SolidBrush(trackColor))
            {
                g.FillPath(trackBrush, path);
            }

            int knobSize = 14;
            int knobX = isChecked ? (trackW - knobSize - 3) : 3;
            int knobY = trackY + 3;
            Color knobColor = Enabled ? Color.White : Color.FromArgb(140, 140, 140);
            using (Brush knobBrush = new SolidBrush(knobColor))
            {
                g.FillEllipse(knobBrush, knobX, knobY, knobSize, knobSize);
            }

            Color textColor = Enabled ? Color.FromArgb(238, 238, 238) : Color.FromArgb(130, 130, 130);
            TextRenderer.DrawText(
                g,
                Text,
                Font,
                new Rectangle(trackW + 8, 0, Width - trackW - 8, Height),
                textColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int diameter)
        {
            var path = new GraphicsPath();
            var arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal sealed class SmoothProgressBar : Control
    {
        private float val;

        public float Value
        {
            get { return val; }
            set
            {
                val = Math.Max(0f, Math.Min(1f, value));
                Invalidate();
            }
        }

        public SmoothProgressBar()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.Clear(Color.FromArgb(58, 58, 58));
            if (val > 0f)
            {
                int fillW = (int)(Width * val);
                using (Brush b = new SolidBrush(Color.FromArgb(31, 106, 165)))
                {
                    e.Graphics.FillRectangle(b, 0, 0, fillW, Height);
                }
            }
        }
    }
}
