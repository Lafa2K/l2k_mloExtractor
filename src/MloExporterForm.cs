using CodeWalker.GameFiles;
using CodeWalker.Properties;
using CodeWalker.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CodeWalker.MloExporter
{
    public class MloExporterForm : Form
    {
        private const string AppTitle = "Blender MLO Extractor";
        private const string IdleBannerFileName = "banner_idle.png";
        private const string LoadingBannerFileName = "banner_loading.gif";
        private const string LegacyBannerFileName = "banner.png";
        private const string IconFileName = "icon.ico";
        private const int DesktopMargin = 12;

        private static readonly Color ThemeBase = Color.FromArgb(0x12, 0x18, 0x2E);
        private static readonly Color ThemeBackground = Color.FromArgb(0x0B, 0x0F, 0x1D);
        private static readonly Color ThemeSurface = Color.FromArgb(0x16, 0x1E, 0x38);
        private static readonly Color ThemeSurfaceAlt = Color.FromArgb(0x23, 0x30, 0x54);
        private static readonly Color ThemeBorder = Color.FromArgb(0x43, 0x59, 0x8B);
        private static readonly Color ThemeTextPrimary = Color.FromArgb(0xF1, 0xF6, 0xFF);
        private static readonly Color ThemeTextSecondary = Color.FromArgb(0xB8, 0xC9, 0xE5);

        private sealed class SelectionListItem
        {
            public YtypPropSelectionItem Item { get; set; }

            public override string ToString()
            {
                return Item?.Label ?? string.Empty;
            }
        }

        private sealed class YmapFileListItem
        {
            public YmapExteriorFileInfo Item { get; set; }

            public override string ToString()
            {
                return Item?.Label ?? string.Empty;
            }
        }

        private GameFileCache GameFileCache;
        private readonly YtypPropExporter Exporter = new YtypPropExporter();

        private TabControl WorkflowTabControl;
        private TabPage MloTabPage;
        private TabPage YmapExteriorTabPage;
        private Button OpenButton;
        private Button ExportButton;
        private PictureBox BannerPictureBox;
        private CheckBox ExportTexturesCheckBox;
        private CheckBox OpenFolderCheckBox;
        private CheckBox ImportAllMloCheckBox;
        private GroupBox RoomsGroupBox;
        private GroupBox EntitySetsGroupBox;
        private CheckedListBox RoomsCheckedListBox;
        private CheckedListBox EntitySetsCheckedListBox;
        private Label CacheStatusLabel;
        private Label InputPathLabel;
        private Label OutputPathLabel;
        private Label AddonRpfLabel;
        private Label StatusLabel;
        private ProgressBar ExportProgressBar;
        private TextBox SummaryTextBox;
        private TextBox AddonRpfTextBox;
        private OpenFileDialog OpenFileDialog;
        private Button OpenYmapButton;
        private Button ExportYmapButton;
        private CheckBox ExportYmapTexturesCheckBox;
        private CheckBox OpenYmapFolderCheckBox;
        private Label YmapInputPathLabel;
        private Label YmapOutputPathLabel;
        private Label YmapAddonRpfLabel;
        private Label YmapStatusLabel;
        private TextBox YmapAddonRpfTextBox;
        private TextBox YmapSummaryTextBox;
        private GroupBox YmapFilesGroupBox;
        private ListBox YmapFilesListBox;
        private ProgressBar YmapExportProgressBar;
        private OpenFileDialog OpenYmapFileDialog;

        private volatile bool CacheReady = false;
        private volatile bool CacheInitializing = true;
        private volatile bool ExportInProgress = false;
        private CancellationTokenSource CacheContentLoopCancellation;
        private string CurrentBannerAssetPath;
        private Icon LoadedAppIcon;

        private string LoadedInputPath;
        private string LoadedOutputPath;
        private YtypPropSelectionInfo LoadedSelectionInfo;
        private string[] LoadedYmapInputPaths;
        private string LoadedYmapOutputPath;
        private YmapExteriorSelectionInfo LoadedYmapSelectionInfo;

        public MloExporterForm()
        {
            InitializeUi();
            AllowDrop = true;
            DragEnter += Form_DragEnter;
            DragDrop += Form_DragDrop;
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            FitWindowToWorkingArea();

            if (!GTAFolder.UpdateGTAFolder(true))
            {
                Close();
                return;
            }

            GameFileCache = GameFileCacheFactory.Create();
            GTAFolder.UpdateEnhancedFormTitle(this);
            await InitializeCacheAsync();
        }

        private void InitializeUi()
        {
            Text = AppTitle;
            Width = 800;
            Height = 860;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(800, 700);
            AutoScroll = true;
            AutoScrollMinSize = new Size(0, 1035);
            LoadWindowIcon();

            OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.Filter = "YTYP files|*.ytyp;*.ytyp.xml|Binary YTYP|*.ytyp|YTYP XML|*.ytyp.xml";
            OpenFileDialog.Multiselect = false;
            OpenFileDialog.Title = "Open a YTYP or YTYP.XML";

            OpenYmapFileDialog = new OpenFileDialog();
            OpenYmapFileDialog.Filter = "YMAP files|*.ymap;*.ymap.xml|Binary YMAP|*.ymap|YMAP XML|*.ymap.xml";
            OpenYmapFileDialog.Multiselect = true;
            OpenYmapFileDialog.Title = "Open one or more YMAP or YMAP.XML files";

            BannerPictureBox = new PictureBox()
            {
                Left = 20,
                Top = 20,
                Width = 740,
                Height = 250,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle
            };
            Controls.Add(BannerPictureBox);

            WorkflowTabControl = new TabControl()
            {
                Left = 20,
                Top = 285,
                Width = 740,
                Height = 700,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(WorkflowTabControl);

            MloTabPage = new TabPage("MLO EXTRACTION")
            {
                BackColor = ThemeBackground,
                ForeColor = ThemeTextPrimary
            };
            WorkflowTabControl.Controls.Add(MloTabPage);

            YmapExteriorTabPage = new TabPage("YMAP EXTERIOR EXTRACTION")
            {
                BackColor = ThemeBackground,
                ForeColor = ThemeTextPrimary
            };
            WorkflowTabControl.Controls.Add(YmapExteriorTabPage);

            CreateMloExtractionTab();
            CreateYmapExteriorExtractionTab();

            ApplyTheme();
            UpdateActivityVisualState();
        }

        private void FitWindowToWorkingArea()
        {
            var workingArea = Screen.FromControl(this).WorkingArea;
            int maxWidth = Math.Max(MinimumSize.Width, workingArea.Width - (DesktopMargin * 2));
            int maxHeight = Math.Max(MinimumSize.Height, workingArea.Height - (DesktopMargin * 2));

            if (Width > maxWidth)
            {
                Width = maxWidth;
            }

            if (Height > maxHeight)
            {
                Height = maxHeight;
            }

            int left = workingArea.Left + ((workingArea.Width - Width) / 2);
            int top = workingArea.Top + ((workingArea.Height - Height) / 2);

            left = Math.Max(workingArea.Left + DesktopMargin, Math.Min(left, workingArea.Right - Width - DesktopMargin));
            top = Math.Max(workingArea.Top + DesktopMargin, Math.Min(top, workingArea.Bottom - Height - DesktopMargin));

            Location = new Point(left, top);
        }

        private void CreateMloExtractionTab()
        {
            var titleLabel = new Label()
            {
                Left = 12,
                Top = 16,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Open a YTYP or YTYP.XML, export Drawable assets first, then import the exported YTYP XML in Blender."
            };
            MloTabPage.Controls.Add(titleLabel);

            ExportTexturesCheckBox = new CheckBox()
            {
                Left = 12,
                Top = 50,
                Width = 260,
                Height = 24,
                Checked = true,
                Text = "Export related and shared textures"
            };
            MloTabPage.Controls.Add(ExportTexturesCheckBox);

            OpenFolderCheckBox = new CheckBox()
            {
                Left = 292,
                Top = 50,
                Width = 220,
                Height = 24,
                Checked = true,
                Text = "Open output folder when done"
            };
            MloTabPage.Controls.Add(OpenFolderCheckBox);

            OpenButton = new Button()
            {
                Left = 12,
                Top = 87,
                Width = 160,
                Height = 32,
                Text = "Open YTYP...",
                Enabled = false
            };
            OpenButton.Click += OpenButton_Click;
            MloTabPage.Controls.Add(OpenButton);

            ExportButton = new Button()
            {
                Left = 187,
                Top = 87,
                Width = 160,
                Height = 32,
                Text = "Export Selected",
                Enabled = false
            };
            ExportButton.Click += ExportButton_Click;
            MloTabPage.Controls.Add(ExportButton);

            CacheStatusLabel = new Label()
            {
                Left = 367,
                Top = 94,
                Width = 345,
                Height = 20,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Waiting to initialize GTA file cache..."
            };
            MloTabPage.Controls.Add(CacheStatusLabel);

            var inputCaption = new Label()
            {
                Left = 12,
                Top = 137,
                Width = 90,
                Height = 20,
                Text = "Input:"
            };
            MloTabPage.Controls.Add(inputCaption);

            InputPathLabel = new Label()
            {
                Left = 72,
                Top = 137,
                Width = 640,
                Height = 36,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            MloTabPage.Controls.Add(InputPathLabel);

            var outputCaption = new Label()
            {
                Left = 12,
                Top = 177,
                Width = 90,
                Height = 20,
                Text = "Output:"
            };
            MloTabPage.Controls.Add(outputCaption);

            OutputPathLabel = new Label()
            {
                Left = 72,
                Top = 177,
                Width = 640,
                Height = 36,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            MloTabPage.Controls.Add(OutputPathLabel);

            AddonRpfLabel = new Label()
            {
                Left = 12,
                Top = 217,
                Width = 90,
                Height = 20,
                Text = "Addon RPF:"
            };
            MloTabPage.Controls.Add(AddonRpfLabel);

            AddonRpfTextBox = new TextBox()
            {
                Left = 102,
                Top = 213,
                Width = 610,
                Height = 24,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            MloTabPage.Controls.Add(AddonRpfTextBox);

            ImportAllMloCheckBox = new CheckBox()
            {
                Left = 12,
                Top = 253,
                Width = 200,
                Height = 24,
                Checked = true,
                Text = "Import All MLO",
                Enabled = false
            };
            ImportAllMloCheckBox.CheckedChanged += ImportAllMloCheckBox_CheckedChanged;
            MloTabPage.Controls.Add(ImportAllMloCheckBox);

            RoomsGroupBox = new GroupBox()
            {
                Left = 12,
                Top = 286,
                Width = 340,
                Height = 175,
                Text = "Rooms (0)"
            };
            MloTabPage.Controls.Add(RoomsGroupBox);

            RoomsCheckedListBox = new CheckedListBox()
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            RoomsGroupBox.Controls.Add(RoomsCheckedListBox);

            EntitySetsGroupBox = new GroupBox()
            {
                Left = 372,
                Top = 286,
                Width = 340,
                Height = 175,
                Text = "Entity Sets (0)"
            };
            MloTabPage.Controls.Add(EntitySetsGroupBox);

            EntitySetsCheckedListBox = new CheckedListBox()
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            EntitySetsGroupBox.Controls.Add(EntitySetsCheckedListBox);

            ExportProgressBar = new ProgressBar()
            {
                Left = 12,
                Top = 476,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Minimum = 0,
                Maximum = 1000,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };
            MloTabPage.Controls.Add(ExportProgressBar);

            StatusLabel = new Label()
            {
                Left = 12,
                Top = 506,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Select a YTYP when the cache is ready."
            };
            MloTabPage.Controls.Add(StatusLabel);

            SummaryTextBox = new TextBox()
            {
                Left = 12,
                Top = 536,
                Width = 700,
                Height = 120,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
            MloTabPage.Controls.Add(SummaryTextBox);
        }

        private void CreateYmapExteriorExtractionTab()
        {
            var titleLabel = new Label()
            {
                Left = 12,
                Top = 16,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Open one or more YMAP/YMAP.XML files. YMAP files containing interiors are ignored."
            };
            YmapExteriorTabPage.Controls.Add(titleLabel);

            ExportYmapTexturesCheckBox = new CheckBox()
            {
                Left = 12,
                Top = 50,
                Width = 260,
                Height = 24,
                Checked = true,
                Text = "Export related and shared textures"
            };
            YmapExteriorTabPage.Controls.Add(ExportYmapTexturesCheckBox);

            OpenYmapFolderCheckBox = new CheckBox()
            {
                Left = 292,
                Top = 50,
                Width = 220,
                Height = 24,
                Checked = true,
                Text = "Open output folder when done"
            };
            YmapExteriorTabPage.Controls.Add(OpenYmapFolderCheckBox);

            OpenYmapButton = new Button()
            {
                Left = 12,
                Top = 87,
                Width = 170,
                Height = 32,
                Text = "Open YMAPs...",
                Enabled = false
            };
            OpenYmapButton.Click += OpenYmapButton_Click;
            YmapExteriorTabPage.Controls.Add(OpenYmapButton);

            ExportYmapButton = new Button()
            {
                Left = 197,
                Top = 87,
                Width = 170,
                Height = 32,
                Text = "Export Exterior",
                Enabled = false
            };
            ExportYmapButton.Click += ExportYmapButton_Click;
            YmapExteriorTabPage.Controls.Add(ExportYmapButton);

            var inputCaption = new Label()
            {
                Left = 12,
                Top = 137,
                Width = 90,
                Height = 20,
                Text = "Input:"
            };
            YmapExteriorTabPage.Controls.Add(inputCaption);

            YmapInputPathLabel = new Label()
            {
                Left = 72,
                Top = 137,
                Width = 640,
                Height = 50,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            YmapExteriorTabPage.Controls.Add(YmapInputPathLabel);

            var outputCaption = new Label()
            {
                Left = 12,
                Top = 193,
                Width = 90,
                Height = 20,
                Text = "Output:"
            };
            YmapExteriorTabPage.Controls.Add(outputCaption);

            YmapOutputPathLabel = new Label()
            {
                Left = 72,
                Top = 193,
                Width = 640,
                Height = 36,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            YmapExteriorTabPage.Controls.Add(YmapOutputPathLabel);

            YmapAddonRpfLabel = new Label()
            {
                Left = 12,
                Top = 233,
                Width = 90,
                Height = 20,
                Text = "Addon RPF:"
            };
            YmapExteriorTabPage.Controls.Add(YmapAddonRpfLabel);

            YmapAddonRpfTextBox = new TextBox()
            {
                Left = 102,
                Top = 229,
                Width = 610,
                Height = 24,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            YmapExteriorTabPage.Controls.Add(YmapAddonRpfTextBox);

            YmapFilesGroupBox = new GroupBox()
            {
                Left = 12,
                Top = 268,
                Width = 700,
                Height = 190,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "YMAP Files (0)"
            };
            YmapExteriorTabPage.Controls.Add(YmapFilesGroupBox);

            YmapFilesListBox = new ListBox()
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true
            };
            YmapFilesGroupBox.Controls.Add(YmapFilesListBox);

            YmapExportProgressBar = new ProgressBar()
            {
                Left = 12,
                Top = 476,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Minimum = 0,
                Maximum = 1000,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };
            YmapExteriorTabPage.Controls.Add(YmapExportProgressBar);

            YmapStatusLabel = new Label()
            {
                Left = 12,
                Top = 506,
                Width = 700,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Select YMAP files when the cache is ready."
            };
            YmapExteriorTabPage.Controls.Add(YmapStatusLabel);

            YmapSummaryTextBox = new TextBox()
            {
                Left = 12,
                Top = 536,
                Width = 700,
                Height = 120,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
            YmapExteriorTabPage.Controls.Add(YmapSummaryTextBox);
        }

        private void LoadWindowIcon()
        {
            try
            {
                var iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", IconFileName);
                if (!File.Exists(iconPath))
                {
                    return;
                }

                using (var stream = new FileStream(iconPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var icon = new Icon(stream))
                {
                    LoadedAppIcon?.Dispose();
                    LoadedAppIcon = (Icon)icon.Clone();
                    Icon = LoadedAppIcon;
                }
            }
            catch
            {
            }
        }

        private string GetBannerAssetPath(bool loading)
        {
            string fileName = loading ? LoadingBannerFileName : IdleBannerFileName;
            var bannerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", fileName);
            if (File.Exists(bannerPath))
            {
                return bannerPath;
            }

            if (!loading)
            {
                var legacyBannerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "img", LegacyBannerFileName);
                if (File.Exists(legacyBannerPath))
                {
                    return legacyBannerPath;
                }
            }

            return null;
        }

        private void UpdateActivityVisualState()
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(UpdateActivityVisualState));
                return;
            }

            bool showLoadingBanner = ExportInProgress;
            bool showBusyProgress = CacheInitializing || ExportInProgress;
            var bannerPath = GetBannerAssetPath(showLoadingBanner);
            if (!string.IsNullOrWhiteSpace(bannerPath) && !string.Equals(CurrentBannerAssetPath, bannerPath, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    BannerPictureBox.Image = null;
                    BannerPictureBox.ImageLocation = null;
                    CurrentBannerAssetPath = bannerPath;
                    BannerPictureBox.ImageLocation = bannerPath;
                    BannerPictureBox.Load();
                }
                catch
                {
                    CurrentBannerAssetPath = null;
                }
            }

            if (showBusyProgress)
            {
                SetProgressBarBusy(ExportProgressBar);
                SetProgressBarBusy(YmapExportProgressBar);
            }
            else
            {
                SetProgressBarIdle(ExportProgressBar);
                SetProgressBarIdle(YmapExportProgressBar);
            }
        }

        private void SetProgressBarBusy(ProgressBar progressBar)
        {
            if (progressBar == null)
            {
                return;
            }

            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;
        }

        private void SetProgressBarIdle(ProgressBar progressBar)
        {
            if (progressBar == null)
            {
                return;
            }

            progressBar.MarqueeAnimationSpeed = 0;
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
        }

        private void ApplyTheme()
        {
            BackColor = ThemeBackground;
            ForeColor = ThemeTextPrimary;

            BannerPictureBox.BackColor = ThemeSurface;

            ApplyThemeToControl(this);
        }

        private void ApplyThemeToControl(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                if (control is Button button)
                {
                    button.FlatStyle = FlatStyle.Flat;
                    button.UseVisualStyleBackColor = false;
                    button.BackColor = ThemeBase;
                    button.ForeColor = ThemeTextPrimary;
                    button.FlatAppearance.BorderColor = ThemeBorder;
                    button.FlatAppearance.MouseOverBackColor = ThemeSurfaceAlt;
                    button.FlatAppearance.MouseDownBackColor = ThemeSurface;
                }
                else if (control is CheckBox checkBox)
                {
                    checkBox.ForeColor = ThemeTextPrimary;
                    checkBox.BackColor = ThemeBackground;
                }
                else if (control is GroupBox groupBox)
                {
                    groupBox.ForeColor = ThemeTextPrimary;
                    groupBox.BackColor = ThemeSurface;
                }
                else if (control is CheckedListBox checkedListBox)
                {
                    checkedListBox.BackColor = ThemeSurface;
                    checkedListBox.ForeColor = ThemeTextPrimary;
                    checkedListBox.BorderStyle = BorderStyle.None;
                }
                else if (control is ListBox listBox)
                {
                    listBox.BackColor = ThemeSurface;
                    listBox.ForeColor = ThemeTextPrimary;
                    listBox.BorderStyle = BorderStyle.None;
                }
                else if (control is TextBox textBox)
                {
                    textBox.BackColor = ThemeSurface;
                    textBox.ForeColor = ThemeTextPrimary;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (control is Label label)
                {
                    label.ForeColor = ThemeTextSecondary;
                    label.BackColor = Color.Transparent;
                }
                else if (control is PictureBox pictureBox)
                {
                    pictureBox.BackColor = ThemeSurface;
                }
                else if (control is TabPage tabPage)
                {
                    tabPage.BackColor = ThemeBackground;
                    tabPage.ForeColor = ThemeTextPrimary;
                }

                if (control.HasChildren)
                {
                    ApplyThemeToControl(control);
                }
            }

            CacheStatusLabel.ForeColor = ThemeTextPrimary;
            InputPathLabel.ForeColor = ThemeTextPrimary;
            OutputPathLabel.ForeColor = ThemeTextPrimary;
            StatusLabel.ForeColor = ThemeTextPrimary;
            YmapInputPathLabel.ForeColor = ThemeTextPrimary;
            YmapOutputPathLabel.ForeColor = ThemeTextPrimary;
            YmapStatusLabel.ForeColor = ThemeTextPrimary;
        }

        private async Task InitializeCacheAsync()
        {
            CacheInitializing = true;
            UpdateActivityVisualState();
            SetCacheStatus("Loading GTA keys...");
            try
            {
                await Task.Run(() =>
                {
                    GTA5Keys.LoadFromPath(GTAFolder.CurrentGTAFolder, GTAFolder.IsGen9, Settings.Default.Key);

                    GameFileCache.EnableDlc = true;
                    GameFileCache.EnableMods = true;
                    GameFileCache.LoadArchetypes = true;
                    GameFileCache.LoadVehicles = false;
                    GameFileCache.LoadPeds = false;
                    GameFileCache.LoadAudio = false;
                    GameFileCache.BuildExtendedJenkIndex = false;
                    GameFileCache.DoFullStringIndex = false;
                    GameFileCache.Init(UpdateStatusSafe, UpdateStatusSafe);
                });

                StartCacheContentLoop();
                CacheReady = true;
                OpenButton.Enabled = true;
                OpenYmapButton.Enabled = true;
                SetCacheStatus("Ready. Open a YTYP/YMAP or drop files on this window.");
                UpdateStatusSafe("Ready to load files.");
                UpdateSelectionUiState();
            }
            catch (Exception ex)
            {
                SetCacheStatus("Unable to initialize GTA files.");
                MessageBox.Show(this, "Failed to initialize the GTA file cache:\n" + ex, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CacheInitializing = false;
                UpdateActivityVisualState();
            }
        }

        private async void OpenButton_Click(object sender, EventArgs e)
        {
            if (ExportInProgress)
            {
                return;
            }

            if (OpenFileDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            await LoadSelectedFileAsync(OpenFileDialog.FileName);
        }

        private async void ExportButton_Click(object sender, EventArgs e)
        {
            if (ExportInProgress)
            {
                return;
            }

            await ExportLoadedFileAsync();
        }

        private async void OpenYmapButton_Click(object sender, EventArgs e)
        {
            if (ExportInProgress)
            {
                return;
            }

            if (OpenYmapFileDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }

            await LoadSelectedYmapFilesAsync(OpenYmapFileDialog.FileNames);
        }

        private async void ExportYmapButton_Click(object sender, EventArgs e)
        {
            if (ExportInProgress)
            {
                return;
            }

            await ExportLoadedYmapFilesAsync();
        }

        private void Form_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if ((files != null) && (files.Length == 1) && YtypPropExporter.SupportsInputPath(files[0]))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }

                if ((files != null) && (files.Length > 0) && files.All(YtypPropExporter.SupportsYmapInputPath))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }

            e.Effect = DragDropEffects.None;
        }

        private async void Form_DragDrop(object sender, DragEventArgs e)
        {
            if (ExportInProgress || !CacheReady)
            {
                return;
            }

            var files = e.Data.GetData(DataFormats.FileDrop) as string[];
            if ((files == null) || (files.Length == 0))
            {
                return;
            }

            if ((files.Length == 1) && YtypPropExporter.SupportsInputPath(files[0]))
            {
                WorkflowTabControl.SelectedTab = MloTabPage;
                await LoadSelectedFileAsync(files[0]);
                return;
            }

            if (files.All(YtypPropExporter.SupportsYmapInputPath))
            {
                WorkflowTabControl.SelectedTab = YmapExteriorTabPage;
                await LoadSelectedYmapFilesAsync(files);
            }
        }

        private async Task LoadSelectedFileAsync(string inputPath)
        {
            if (!CacheReady)
            {
                MessageBox.Show(this, "The GTA file cache is still loading. Please wait a moment and try again.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!YtypPropExporter.SupportsInputPath(inputPath))
            {
                MessageBox.Show(this, "Only .ytyp and .ytyp.xml files are supported.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            InputPathLabel.Text = inputPath;
            OutputPathLabel.Text = string.Empty;
            SummaryTextBox.Clear();

            SetBusyUiState(true);
            UpdateStatusSafe("Loading MLO rooms and entity sets...");

            Exception loadException = null;
            YtypPropSelectionInfo selectionInfo = null;

            try
            {
                selectionInfo = await Task.Run(() => Exporter.LoadSelectionInfo(inputPath));
            }
            catch (Exception ex)
            {
                loadException = ex;
            }
            finally
            {
                SetBusyUiState(false);
            }

            if (loadException != null)
            {
                ResetLoadedSelection();
                SummaryTextBox.Text = loadException.ToString();
                MessageBox.Show(this, "Unable to load the selected YTYP:\n" + loadException.Message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var outputFolder = YtypPropExporter.GetSuggestedOutputFolderPath(inputPath, selectionInfo?.PrimaryMloName);
            LoadedInputPath = inputPath;
            LoadedOutputPath = outputFolder;
            LoadedSelectionInfo = selectionInfo;
            OutputPathLabel.Text = outputFolder;
            PopulateSelectionControls(selectionInfo);
            UpdateStatusSafe("Choose rooms and entity sets, then click Export Selected.");
        }

        private async Task ExportLoadedFileAsync()
        {
            if (string.IsNullOrWhiteSpace(LoadedInputPath) || (LoadedSelectionInfo == null))
            {
                MessageBox.Show(this, "Open a YTYP first so the rooms and entity sets can be loaded.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selection = BuildExportSelection();
            if (!selection.ImportAllMlo && (selection.RoomKeys.Count == 0) && (selection.EntitySetKeys.Count == 0))
            {
                MessageBox.Show(this, "Select at least one room or one entity set before exporting.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SummaryTextBox.Clear();

            SetBusyUiState(true);
            Exception exportException = null;
            YtypPropExportResult result = null;

            try
            {
                result = await Task.Run(() => Exporter.Export(
                    GameFileCache,
                    LoadedInputPath,
                    LoadedOutputPath,
                    ExportTexturesCheckBox.Checked,
                    selection,
                    UpdateProgressSafe,
                    UpdateStatusSafe));
            }
            catch (Exception ex)
            {
                exportException = ex;
            }
            finally
            {
                SetBusyUiState(false);
            }

            if (exportException != null)
            {
                SummaryTextBox.Text = exportException.ToString();
                MessageBox.Show(this, "Export failed:\n" + exportException.Message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (result == null)
            {
                return;
            }

            var summary = BuildSummary(result, LoadedOutputPath);
            SummaryTextBox.Text = summary;
            UpdateStatusSafe("Export complete.");

            if (OpenFolderCheckBox.Checked && Directory.Exists(LoadedOutputPath))
            {
                try
                {
                    Process.Start("explorer", "\"" + LoadedOutputPath + "\"");
                }
                catch
                {
                }
            }
        }

        private async Task LoadSelectedYmapFilesAsync(string[] inputPaths)
        {
            if (!CacheReady)
            {
                MessageBox.Show(this, "The GTA file cache is still loading. Please wait a moment and try again.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var paths = inputPaths?
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.InvariantCultureIgnoreCase)
                .ToArray();

            if ((paths == null) || (paths.Length == 0) || !paths.All(YtypPropExporter.SupportsYmapInputPath))
            {
                MessageBox.Show(this, "Only .ymap and .ymap.xml files are supported.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            YmapInputPathLabel.Text = BuildYmapInputLabel(paths);
            YmapOutputPathLabel.Text = string.Empty;
            YmapSummaryTextBox.Clear();
            YmapFilesListBox.Items.Clear();
            YmapFilesGroupBox.Text = "YMAP Files (0)";

            SetBusyUiState(true);
            UpdateStatusSafe("Loading exterior YMAP files...");

            Exception loadException = null;
            YmapExteriorSelectionInfo selectionInfo = null;

            try
            {
                selectionInfo = await Task.Run(() => Exporter.LoadYmapExteriorSelectionInfo(paths));
            }
            catch (Exception ex)
            {
                loadException = ex;
            }
            finally
            {
                SetBusyUiState(false);
            }

            if (loadException != null)
            {
                ResetLoadedYmapSelection();
                YmapSummaryTextBox.Text = loadException.ToString();
                MessageBox.Show(this, "Unable to load the selected YMAP files:\n" + loadException.Message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var outputFolder = YtypPropExporter.GetSuggestedYmapOutputFolderPath(paths);
            LoadedYmapInputPaths = paths;
            LoadedYmapOutputPath = outputFolder;
            LoadedYmapSelectionInfo = selectionInfo;
            YmapOutputPathLabel.Text = outputFolder;
            PopulateYmapControls(selectionInfo);
            UpdateStatusSafe("Choose Export Exterior to export non-interior YMAP props.");
        }

        private async Task ExportLoadedYmapFilesAsync()
        {
            if ((LoadedYmapInputPaths == null) || (LoadedYmapInputPaths.Length == 0) || (LoadedYmapSelectionInfo == null))
            {
                MessageBox.Show(this, "Open one or more YMAP files first.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (LoadedYmapSelectionInfo.ExportableFiles == 0)
            {
                MessageBox.Show(this, "Every selected YMAP contains an interior or has no exterior entities to export.", AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            YmapSummaryTextBox.Clear();

            SetBusyUiState(true);
            Exception exportException = null;
            YtypPropExportResult result = null;

            try
            {
                result = await Task.Run(() => Exporter.ExportYmapExterior(
                    GameFileCache,
                    LoadedYmapInputPaths,
                    LoadedYmapOutputPath,
                    ExportYmapTexturesCheckBox.Checked,
                    YmapAddonRpfTextBox.Text,
                    UpdateProgressSafe,
                    UpdateStatusSafe));
            }
            catch (Exception ex)
            {
                exportException = ex;
            }
            finally
            {
                SetBusyUiState(false);
            }

            if (exportException != null)
            {
                YmapSummaryTextBox.Text = exportException.ToString();
                MessageBox.Show(this, "YMAP export failed:\n" + exportException.Message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (result == null)
            {
                return;
            }

            YmapSummaryTextBox.Text = BuildYmapSummary(result, LoadedYmapOutputPath);
            UpdateStatusSafe("Exterior YMAP export complete.");

            if (OpenYmapFolderCheckBox.Checked && Directory.Exists(LoadedYmapOutputPath))
            {
                try
                {
                    Process.Start("explorer", "\"" + LoadedYmapOutputPath + "\"");
                }
                catch
                {
                }
            }
        }

        private void PopulateSelectionControls(YtypPropSelectionInfo selectionInfo)
        {
            RoomsCheckedListBox.Items.Clear();
            EntitySetsCheckedListBox.Items.Clear();

            if (selectionInfo == null)
            {
                UpdateSelectionUiState();
                return;
            }

            foreach (var room in selectionInfo.Rooms)
            {
                RoomsCheckedListBox.Items.Add(new SelectionListItem() { Item = room }, true);
            }

            foreach (var entitySet in selectionInfo.EntitySets)
            {
                EntitySetsCheckedListBox.Items.Add(new SelectionListItem() { Item = entitySet }, false);
            }

            RoomsGroupBox.Text = "Rooms (" + selectionInfo.Rooms.Count.ToString() + ")";
            EntitySetsGroupBox.Text = "Entity Sets (" + selectionInfo.EntitySets.Count.ToString() + ")";
            ImportAllMloCheckBox.Checked = true;
            SummaryTextBox.Text = BuildLoadedSummary(selectionInfo);
            UpdateSelectionUiState();
        }

        private void ResetLoadedSelection()
        {
            LoadedInputPath = null;
            LoadedOutputPath = null;
            LoadedSelectionInfo = null;
            RoomsCheckedListBox.Items.Clear();
            EntitySetsCheckedListBox.Items.Clear();
            RoomsGroupBox.Text = "Rooms (0)";
            EntitySetsGroupBox.Text = "Entity Sets (0)";
            ImportAllMloCheckBox.Checked = true;
            UpdateSelectionUiState();
        }

        private void PopulateYmapControls(YmapExteriorSelectionInfo selectionInfo)
        {
            YmapFilesListBox.Items.Clear();

            if (selectionInfo == null)
            {
                UpdateSelectionUiState();
                return;
            }

            foreach (var file in selectionInfo.Files)
            {
                YmapFilesListBox.Items.Add(new YmapFileListItem() { Item = file });
            }

            YmapFilesGroupBox.Text = "YMAP Files (" + selectionInfo.TotalFiles.ToString() + ")";
            YmapSummaryTextBox.Text = BuildYmapLoadedSummary(selectionInfo);
            UpdateSelectionUiState();
        }

        private void ResetLoadedYmapSelection()
        {
            LoadedYmapInputPaths = null;
            LoadedYmapOutputPath = null;
            LoadedYmapSelectionInfo = null;
            YmapFilesListBox.Items.Clear();
            YmapFilesGroupBox.Text = "YMAP Files (0)";
            UpdateSelectionUiState();
        }

        private string BuildLoadedSummary(YtypPropSelectionInfo selectionInfo)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Loaded " + selectionInfo.MloCount.ToString() + " MLO archetype(s).");
            sb.AppendLine(selectionInfo.Rooms.Count.ToString() + " room(s) found.");
            sb.AppendLine(selectionInfo.EntitySets.Count.ToString() + " entity set(s) found.");
            sb.AppendLine();
            sb.AppendLine("Output layout:");
            sb.AppendLine("- <source>.ytyp.xml in the export root");
            sb.AppendLine("- Drawable\\ for props and textures");
            sb.AppendLine();
            sb.AppendLine("Rooms are checked by default.");
            sb.AppendLine("Entity sets are optional and start unchecked.");
            return sb.ToString();
        }

        private string BuildYmapInputLabel(string[] inputPaths)
        {
            if ((inputPaths == null) || (inputPaths.Length == 0))
            {
                return "-";
            }

            if (inputPaths.Length == 1)
            {
                return inputPaths[0];
            }

            return inputPaths.Length.ToString(CultureInfo.InvariantCulture) + " files selected. First: " + inputPaths[0];
        }

        private string BuildYmapLoadedSummary(YmapExteriorSelectionInfo selectionInfo)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Loaded " + selectionInfo.TotalFiles.ToString() + " YMAP file(s).");
            sb.AppendLine(selectionInfo.ExportableFiles.ToString() + " exportable exterior YMAP file(s).");
            sb.AppendLine(selectionInfo.IgnoredInteriorFiles.ToString() + " YMAP file(s) ignored because they contain interiors.");
            sb.AppendLine(selectionInfo.TotalExteriorEntities.ToString() + " exterior entit" + (selectionInfo.TotalExteriorEntities == 1 ? "y" : "ies") + " found.");
            sb.AppendLine();
            sb.AppendLine("Output layout:");
            sb.AppendLine("- <source>.ymap.xml in the export root");
            sb.AppendLine("- Drawable\\ for props and textures");
            sb.AppendLine();
            sb.AppendLine("Interior YMAP status text:");
            sb.AppendLine("This YMAP contains an interior > ignored");
            return sb.ToString();
        }

        private YtypPropExportSelection BuildExportSelection()
        {
            var selection = new YtypPropExportSelection()
            {
                ImportAllMlo = ImportAllMloCheckBox.Checked,
                PreferredRpfName = AddonRpfTextBox.Text
            };

            foreach (var checkedItem in RoomsCheckedListBox.CheckedItems)
            {
                var item = checkedItem as SelectionListItem;
                if (!string.IsNullOrWhiteSpace(item?.Item?.Key))
                {
                    selection.RoomKeys.Add(item.Item.Key);
                }
            }

            foreach (var checkedItem in EntitySetsCheckedListBox.CheckedItems)
            {
                var item = checkedItem as SelectionListItem;
                if (!string.IsNullOrWhiteSpace(item?.Item?.Key))
                {
                    selection.EntitySetKeys.Add(item.Item.Key);
                }
            }

            return selection;
        }

        private string BuildSummary(YtypPropExportResult result, string outputFolder)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Output: " + outputFolder);
            sb.AppendLine("MLO XML: " + Path.Combine(outputFolder, result.ExportedYtypXmlFileName ?? "export.ytyp.xml"));
            sb.AppendLine("Drawable folder: " + Path.Combine(outputFolder, YtypPropExporter.DrawableFolderName));
            if (!string.IsNullOrWhiteSpace(AddonRpfTextBox.Text))
            {
                sb.AppendLine("Scoped addon RPF: " + AddonRpfTextBox.Text.Trim());
            }
            sb.AppendLine(result.ExportedTargets.ToString() + " of " + result.TotalTargets.ToString() + " prop files exported.");

            if (result.ExportedTextures > 0)
            {
                sb.AppendLine(result.ExportedTextures.ToString() + " textures exported.");
            }
            if (result.MissingTextures > 0)
            {
                sb.AppendLine(result.MissingTextures.ToString() + " referenced textures were not found.");
                AppendGroupedSummaryEntries(sb, "Missing textures:", result.MissingTextureNames);
            }
            if (result.MissingArchetypes > 0)
            {
                sb.AppendLine(result.MissingArchetypes.ToString() + " prop archetypes could not be resolved.");
                AppendGroupedSummaryEntries(sb, "Missing archetypes:", result.MissingArchetypeNames);
            }
            if (result.MissingResources > 0)
            {
                sb.AppendLine(result.MissingResources.ToString() + " prop resources were not found.");
                AppendGroupedSummaryEntries(sb, "Missing resources:", result.MissingResourceNames);
            }
            if (result.Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Errors:");
                foreach (var error in result.Errors)
                {
                    sb.AppendLine(error);
                }
            }

            return sb.ToString();
        }

        private string BuildYmapSummary(YtypPropExportResult result, string outputFolder)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Output: " + outputFolder);
            sb.AppendLine("YMAP XML files: " + result.ExportedSourceXmlFileNames.Count.ToString());
            foreach (var fileName in result.ExportedSourceXmlFileNames)
            {
                sb.AppendLine(" - " + Path.Combine(outputFolder, fileName));
            }
            sb.AppendLine("Drawable folder: " + Path.Combine(outputFolder, YtypPropExporter.DrawableFolderName));
            if (!string.IsNullOrWhiteSpace(YmapAddonRpfTextBox.Text))
            {
                sb.AppendLine("Scoped addon RPF: " + YmapAddonRpfTextBox.Text.Trim());
            }
            if (result.IgnoredYmapNames.Count > 0)
            {
                sb.AppendLine(result.IgnoredYmapNames.Count.ToString() + " YMAP file(s) ignored.");
                AppendGroupedSummaryEntries(sb, "Ignored YMAP files:", result.IgnoredYmapNames);
            }
            sb.AppendLine(result.ExportedTargets.ToString() + " of " + result.TotalTargets.ToString() + " prop files exported.");

            if (result.ExportedTextures > 0)
            {
                sb.AppendLine(result.ExportedTextures.ToString() + " textures exported.");
            }
            if (result.MissingTextures > 0)
            {
                sb.AppendLine(result.MissingTextures.ToString() + " referenced textures were not found.");
                AppendGroupedSummaryEntries(sb, "Missing textures:", result.MissingTextureNames);
            }
            if (result.MissingArchetypes > 0)
            {
                sb.AppendLine(result.MissingArchetypes.ToString() + " prop archetypes could not be resolved.");
                AppendGroupedSummaryEntries(sb, "Missing archetypes:", result.MissingArchetypeNames);
            }
            if (result.MissingResources > 0)
            {
                sb.AppendLine(result.MissingResources.ToString() + " prop resources were not found.");
                AppendGroupedSummaryEntries(sb, "Missing resources:", result.MissingResourceNames);
            }
            if (result.Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Errors:");
                foreach (var error in result.Errors)
                {
                    sb.AppendLine(error);
                }
            }

            return sb.ToString();
        }

        private void AppendGroupedSummaryEntries(StringBuilder sb, string title, IEnumerable<string> values)
        {
            var groupedValues = values?
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .GroupBy(value => value, StringComparer.InvariantCultureIgnoreCase)
                .OrderByDescending(group => group.Count())
                .ThenBy(group => group.Key, StringComparer.InvariantCultureIgnoreCase)
                .ToArray();

            if ((groupedValues == null) || (groupedValues.Length == 0))
            {
                return;
            }

            sb.AppendLine(title);
            foreach (var group in groupedValues)
            {
                var line = " - " + group.Key;
                if (group.Count() > 1)
                {
                    line += " x" + group.Count().ToString();
                }

                sb.AppendLine(line);
            }
        }

        private void SetBusyUiState(bool busy)
        {
            ExportInProgress = busy;

            if (InvokeRequired)
            {
                BeginInvoke(new Action<bool>(SetBusyUiState), busy);
                return;
            }

            OpenButton.Enabled = CacheReady && !busy;
            ExportButton.Enabled = CacheReady && (LoadedSelectionInfo != null) && !busy;
            ExportTexturesCheckBox.Enabled = !busy;
            OpenFolderCheckBox.Enabled = !busy;
            AddonRpfTextBox.Enabled = !busy;
            ImportAllMloCheckBox.Enabled = (LoadedSelectionInfo != null) && !busy;
            RoomsCheckedListBox.Enabled = (LoadedSelectionInfo != null) && !ImportAllMloCheckBox.Checked && !busy;
            EntitySetsCheckedListBox.Enabled = (LoadedSelectionInfo != null) && !busy;
            OpenYmapButton.Enabled = CacheReady && !busy;
            ExportYmapButton.Enabled = CacheReady && (LoadedYmapSelectionInfo != null) && (LoadedYmapSelectionInfo.ExportableFiles > 0) && !busy;
            ExportYmapTexturesCheckBox.Enabled = !busy;
            OpenYmapFolderCheckBox.Enabled = !busy;
            YmapAddonRpfTextBox.Enabled = !busy;
            YmapFilesListBox.Enabled = !busy;

            if (busy)
            {
                ExportProgressBar.Value = 0;
                YmapExportProgressBar.Value = 0;
            }

            UpdateActivityVisualState();
        }

        private void UpdateSelectionUiState()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(UpdateSelectionUiState));
                return;
            }

            bool hasSelection = LoadedSelectionInfo != null;
            ImportAllMloCheckBox.Enabled = hasSelection && !ExportInProgress;
            RoomsCheckedListBox.Enabled = hasSelection && !ImportAllMloCheckBox.Checked && !ExportInProgress;
            EntitySetsCheckedListBox.Enabled = hasSelection && !ExportInProgress;
            ExportButton.Enabled = CacheReady && hasSelection && !ExportInProgress;

            bool hasYmapSelection = LoadedYmapSelectionInfo != null;
            OpenYmapButton.Enabled = CacheReady && !ExportInProgress;
            ExportYmapButton.Enabled = CacheReady && hasYmapSelection && (LoadedYmapSelectionInfo.ExportableFiles > 0) && !ExportInProgress;
        }

        private void ImportAllMloCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            UpdateSelectionUiState();
        }

        private void UpdateProgressSafe(YtypPropExportProgress progress)
        {
            if (IsDisposed || (progress == null))
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<YtypPropExportProgress>(UpdateProgressSafe), progress);
                return;
            }

            StatusLabel.Text = progress.Status ?? string.Empty;
            YmapStatusLabel.Text = progress.Status ?? string.Empty;
            UpdateProgressBarValue(ExportProgressBar, progress);
            UpdateProgressBarValue(YmapExportProgressBar, progress);
        }

        private void UpdateProgressBarValue(ProgressBar progressBar, YtypPropExportProgress progress)
        {
            if ((progressBar == null) || (progress == null))
            {
                return;
            }

            if (progress.Total > 0)
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = Math.Max(0, Math.Min((progress.Current * 1000) / progress.Total, 1000));
            }
            else
            {
                progressBar.Style = ProgressBarStyle.Marquee;
                progressBar.MarqueeAnimationSpeed = 30;
            }
        }

        private void UpdateStatusSafe(string text)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(UpdateStatusSafe), text);
                return;
            }

            StatusLabel.Text = text;
            YmapStatusLabel.Text = text;
        }

        private void SetCacheStatus(string text)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(SetCacheStatus), text);
                return;
            }

            CacheStatusLabel.Text = text;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            CacheContentLoopCancellation?.Cancel();
            CacheContentLoopCancellation?.Dispose();
            CacheContentLoopCancellation = null;
            BannerPictureBox.ImageLocation = null;
            BannerPictureBox?.Image?.Dispose();
            LoadedAppIcon?.Dispose();
            LoadedAppIcon = null;
            base.OnFormClosed(e);
        }

        private void StartCacheContentLoop()
        {
            CacheContentLoopCancellation?.Cancel();
            CacheContentLoopCancellation?.Dispose();
            CacheContentLoopCancellation = new CancellationTokenSource();
            var token = CacheContentLoopCancellation.Token;

            Task.Run(() =>
            {
                while (!token.IsCancellationRequested && !IsDisposed)
                {
                    if (GameFileCache?.IsInited == true)
                    {
                        GameFileCache.BeginFrame();
                        var itemsPending = GameFileCache.ContentThreadProc();
                        if (!itemsPending)
                        {
                            Thread.Sleep(10);
                        }
                    }
                    else
                    {
                        Thread.Sleep(50);
                    }
                }
            }, token);
        }
    }
}
