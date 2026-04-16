using CodeWalker.GameFiles;
using CodeWalker.Properties;
using CodeWalker.Tools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
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

        private GameFileCache GameFileCache;
        private readonly YtypPropExporter Exporter = new YtypPropExporter();

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

        private volatile bool CacheReady = false;
        private volatile bool CacheInitializing = true;
        private volatile bool ExportInProgress = false;
        private CancellationTokenSource CacheContentLoopCancellation;
        private string CurrentBannerAssetPath;
        private Icon LoadedAppIcon;

        private string LoadedInputPath;
        private string LoadedOutputPath;
        private YtypPropSelectionInfo LoadedSelectionInfo;

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
            Height = 820;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(800, 700);
            AutoScroll = true;
            AutoScrollMinSize = new Size(0, 980);
            LoadWindowIcon();

            OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.Filter = "YTYP files|*.ytyp;*.ytyp.xml|Binary YTYP|*.ytyp|YTYP XML|*.ytyp.xml";
            OpenFileDialog.Multiselect = false;
            OpenFileDialog.Title = "Open a YTYP or YTYP.XML";

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

            var titleLabel = new Label()
            {
                Left = 20,
                Top = 280,
                Width = 740,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Open a YTYP or YTYP.XML, export Drawable assets first, then import the exported YTYP XML in Blender."
            };
            Controls.Add(titleLabel);

            ExportTexturesCheckBox = new CheckBox()
            {
                Left = 20,
                Top = 315,
                Width = 260,
                Height = 24,
                Checked = true,
                Text = "Export related and shared textures"
            };
            Controls.Add(ExportTexturesCheckBox);

            OpenFolderCheckBox = new CheckBox()
            {
                Left = 300,
                Top = 315,
                Width = 220,
                Height = 24,
                Checked = true,
                Text = "Open output folder when done"
            };
            Controls.Add(OpenFolderCheckBox);

            OpenButton = new Button()
            {
                Left = 20,
                Top = 352,
                Width = 160,
                Height = 32,
                Text = "Open YTYP...",
                Enabled = false
            };
            OpenButton.Click += OpenButton_Click;
            Controls.Add(OpenButton);

            ExportButton = new Button()
            {
                Left = 195,
                Top = 352,
                Width = 160,
                Height = 32,
                Text = "Export Selected",
                Enabled = false
            };
            ExportButton.Click += ExportButton_Click;
            Controls.Add(ExportButton);

            CacheStatusLabel = new Label()
            {
                Left = 375,
                Top = 359,
                Width = 385,
                Height = 20,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Waiting to initialize GTA file cache..."
            };
            Controls.Add(CacheStatusLabel);

            var inputCaption = new Label()
            {
                Left = 20,
                Top = 402,
                Width = 90,
                Height = 20,
                Text = "Input:"
            };
            Controls.Add(inputCaption);

            InputPathLabel = new Label()
            {
                Left = 80,
                Top = 402,
                Width = 680,
                Height = 36,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            Controls.Add(InputPathLabel);

            var outputCaption = new Label()
            {
                Left = 20,
                Top = 442,
                Width = 90,
                Height = 20,
                Text = "Output:"
            };
            Controls.Add(outputCaption);

            OutputPathLabel = new Label()
            {
                Left = 80,
                Top = 442,
                Width = 680,
                Height = 36,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "-"
            };
            Controls.Add(OutputPathLabel);

            AddonRpfLabel = new Label()
            {
                Left = 20,
                Top = 482,
                Width = 90,
                Height = 20,
                Text = "Addon RPF:"
            };
            Controls.Add(AddonRpfLabel);

            AddonRpfTextBox = new TextBox()
            {
                Left = 110,
                Top = 478,
                Width = 650,
                Height = 24,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            Controls.Add(AddonRpfTextBox);

            ImportAllMloCheckBox = new CheckBox()
            {
                Left = 20,
                Top = 518,
                Width = 200,
                Height = 24,
                Checked = true,
                Text = "Import All MLO",
                Enabled = false
            };
            ImportAllMloCheckBox.CheckedChanged += ImportAllMloCheckBox_CheckedChanged;
            Controls.Add(ImportAllMloCheckBox);

            RoomsGroupBox = new GroupBox()
            {
                Left = 20,
                Top = 551,
                Width = 360,
                Height = 190,
                Text = "Rooms (0)"
            };
            Controls.Add(RoomsGroupBox);

            RoomsCheckedListBox = new CheckedListBox()
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            RoomsGroupBox.Controls.Add(RoomsCheckedListBox);

            EntitySetsGroupBox = new GroupBox()
            {
                Left = 400,
                Top = 551,
                Width = 360,
                Height = 190,
                Text = "Entity Sets (0)"
            };
            Controls.Add(EntitySetsGroupBox);

            EntitySetsCheckedListBox = new CheckedListBox()
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                HorizontalScrollbar = true
            };
            EntitySetsGroupBox.Controls.Add(EntitySetsCheckedListBox);

            ExportProgressBar = new ProgressBar()
            {
                Left = 20,
                Top = 756,
                Width = 740,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Minimum = 0,
                Maximum = 1000,
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30
            };
            Controls.Add(ExportProgressBar);

            StatusLabel = new Label()
            {
                Left = 20,
                Top = 786,
                Width = 740,
                Height = 22,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Text = "Select a YTYP when the cache is ready."
            };
            Controls.Add(StatusLabel);

            SummaryTextBox = new TextBox()
            {
                Left = 20,
                Top = 816,
                Width = 740,
                Height = 120,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical
            };
            Controls.Add(SummaryTextBox);

            ApplyTheme();
            UpdateActivityVisualState();
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
                ExportProgressBar.Style = ProgressBarStyle.Marquee;
                ExportProgressBar.MarqueeAnimationSpeed = 30;
            }
            else
            {
                ExportProgressBar.MarqueeAnimationSpeed = 0;
                ExportProgressBar.Style = ProgressBarStyle.Continuous;
                ExportProgressBar.Value = 0;
            }
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

                if (control.HasChildren)
                {
                    ApplyThemeToControl(control);
                }
            }

            CacheStatusLabel.ForeColor = ThemeTextPrimary;
            InputPathLabel.ForeColor = ThemeTextPrimary;
            OutputPathLabel.ForeColor = ThemeTextPrimary;
            StatusLabel.ForeColor = ThemeTextPrimary;
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
                SetCacheStatus("Ready. Open a YTYP or drop one on this window.");
                UpdateStatusSafe("Ready to load an MLO.");
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
            if ((files == null) || (files.Length != 1))
            {
                return;
            }

            await LoadSelectedFileAsync(files[0]);
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

            if (busy)
            {
                ExportProgressBar.Value = 0;
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
            if (progress.Total > 0)
            {
                ExportProgressBar.Style = ProgressBarStyle.Continuous;
                ExportProgressBar.Value = Math.Max(0, Math.Min((progress.Current * 1000) / progress.Total, 1000));
            }
            else
            {
                ExportProgressBar.Style = ProgressBarStyle.Marquee;
                ExportProgressBar.MarqueeAnimationSpeed = 30;
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
