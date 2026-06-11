using System.Diagnostics;
using System.Drawing;
using System.Text;
using Md2Doc.Core;

namespace Md2Doc;

public sealed class MainForm : Form
{
    // 入力方式
    private readonly RadioButton _textInputRadio = new() { Text = "テキスト入力", Checked = true, AutoSize = true };
    private readonly RadioButton _fileInputRadio = new() { Text = "ファイル入力", AutoSize = true };

    // Markdown 入力エリア
    private readonly TextBox _markdownTextBox = new()
    {
        Multiline = true, ScrollBars = ScrollBars.Both, WordWrap = false, Dock = DockStyle.Fill
    };
    private readonly Button _clearButton = new() { Text = "クリア", AutoSize = true };

    // ファイル入力
    private readonly TextBox _inputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _inputBrowseButton = new() { Text = "参照..." };

    // 出力
    private readonly TextBox _outputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _outputBrowseButton = new() { Text = "参照..." };

    // フォント（本文）
    private readonly string[] _allFonts;
    private readonly ComboBox _bodyRecentFontCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
    private readonly ComboBox _bodyFontCombo;
    private readonly NumericUpDown _bodyFontSizeNumeric = new()
    {
        Minimum = 8, Maximum = 72, DecimalPlaces = 1, Value = 11, Width = 70
    };

    // 見出しオプション
    private readonly CheckBox _headingNumberCheck = new()
    {
        Text = "見出しに番号を付与（例: 1.  1.1  1.1.1）", AutoSize = true
    };

    // ヘッダー内容
    private readonly RadioButton _headerNoneRadio = new() { Text = "なし", Checked = true, AutoSize = true };
    private readonly RadioButton _headerFileNameRadio = new() { Text = "元ファイル名", AutoSize = true };
    private readonly RadioButton _headerCustomRadio = new() { Text = "自由記入:", AutoSize = true };
    private readonly TextBox _headerCustomTextBox = new() { Width = 220, Enabled = false };

    // ヘッダー配置
    private readonly RadioButton _headerAlignLeftRadio = new() { Text = "左寄り", Checked = true, AutoSize = true };
    private readonly RadioButton _headerAlignCenterRadio = new() { Text = "中央", AutoSize = true };
    private readonly RadioButton _headerAlignRightRadio = new() { Text = "右寄り", AutoSize = true };

    // フッター内容
    private readonly CheckBox _footerPageNumberCheck = new()
    {
        Text = "ページ番号を挿入（ページ番号/ページ数）", AutoSize = true
    };

    // フッター配置
    private readonly RadioButton _footerAlignLeftRadio = new() { Text = "左寄り", AutoSize = true };
    private readonly RadioButton _footerAlignCenterRadio = new() { Text = "中央", Checked = true, AutoSize = true };
    private readonly RadioButton _footerAlignRightRadio = new() { Text = "右寄り", AutoSize = true };

    // 変換エンジン選択（既定: Open XML方式）
    private readonly RadioButton _engineOpenXmlRadio = new()
    {
        Text = "Open XML方式（標準・高速）",
        Checked = true,
        AutoSize = true,
    };
    private readonly RadioButton _engineWordComRadio = new()
    {
        Text = "Word COM方式（Word必須・互換確認用）",
        AutoSize = true,
    };

    // 実行・変換後アクション・結果
    private readonly Button _convertButton = new()
    {
        Text = "変換実行",
        AutoSize = false,
        MinimumSize = new Size(110, 34),
        Padding = new Padding(12, 0, 12, 0),
    };
    private readonly Button _openFileButton = new() { Text = "Wordファイルを開く", AutoSize = true, Enabled = false };
    private readonly Button _openFolderButton = new() { Text = "保存先フォルダを開く", AutoSize = true, Enabled = false };
    private readonly Label _resultLabel = new() { AutoSize = true };

    // 状態
    private string _lastOutputPath = "";
    private string _lastOutputFolder = "";
    private string _suggestedOutputPath = "";
    private bool _suppressFontSync = false;

    public MainForm()
    {
        Text = "Markdown変換ツール";
        Width = 960;
        Height = 860;
        MinimumSize = new Size(760, 600);
        Icon = AppIcon.Create();

        _allFonts = FontFamily.Families
            .Select(f => f.Name)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        _bodyFontCombo = CreateFontComboBox("MS Gothic");
        _markdownTextBox.MinimumSize = new Size(100, 80);
        _convertButton.Font = new Font(Font, FontStyle.Bold);

        // ── 入力セクション（Markdown テキスト・伸縮可能）─────────────────────
        // Percent,100 行を最終行にすることで、AutoSize 行が押し出されない構造を保証する。
        var inputModePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        inputModePanel.Controls.Add(new Label { Text = "入力方式:", AutoSize = true, Margin = new Padding(0, 4, 6, 0) });
        inputModePanel.Controls.Add(_textInputRadio);
        inputModePanel.Controls.Add(_fileInputRadio);

        var mdLabelPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        mdLabelPanel.Controls.Add(new Label { Text = "Markdown本文（テキスト入力）", AutoSize = true });
        mdLabelPanel.Controls.Add(_clearButton);

        var inputSection = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
        };
        inputSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // inputModePanel
        inputSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // mdLabelPanel
        inputSection.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // markdownTextBox（最終行）
        inputSection.Controls.Add(inputModePanel);
        inputSection.Controls.Add(mdLabelPanel);
        inputSection.Controls.Add(_markdownTextBox);

        // ── ファイル入出力セクション（常に表示・AutoSize）─────────────────────
        // root の AutoSize 行に配置することで、Markdown 伸縮の影響を受けない。
        var inputFilePanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        inputFilePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        inputFilePanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        inputFilePanel.Controls.Add(_inputFilePathTextBox, 0, 0);
        inputFilePanel.Controls.Add(_inputBrowseButton, 1, 0);

        var outputPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
        outputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        outputPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        outputPanel.Controls.Add(_outputFilePathTextBox, 0, 0);
        outputPanel.Controls.Add(_outputBrowseButton, 1, 0);

        var ioSection = new TableLayoutPanel
        {
            ColumnCount = 1,
            RowCount = 4,
            AutoSize = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
        };
        ioSection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        ioSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // 入力ファイルlabel
        ioSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // inputFilePanel
        ioSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // 出力ファイルlabel
        ioSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // outputPanel
        ioSection.Controls.Add(new Label { Text = "入力ファイル（ファイル入力）", AutoSize = true });
        ioSection.Controls.Add(inputFilePanel);
        ioSection.Controls.Add(new Label { Text = "出力ファイル（.docx）", AutoSize = true });
        ioSection.Controls.Add(outputPanel);

        // ── 設定セクション（下部・高さ固定）──────────────────────────────
        var buttonPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        buttonPanel.Controls.Add(_convertButton);
        buttonPanel.Controls.Add(_openFileButton);
        buttonPanel.Controls.Add(_openFolderButton);

        _resultLabel.Margin = new Padding(0, 4, 0, 0);

        var settingsSection = new TableLayoutPanel
        {
            ColumnCount = 1,
            RowCount = 5,
            AutoSize = true,
            Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top,
            Margin = new Padding(0, 6, 0, 0),
        };
        settingsSection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        settingsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // fontGroupBox
        settingsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // optionsGroupBox
        settingsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // engineGroupBox
        settingsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // buttonPanel
        settingsSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));  // resultLabel
        settingsSection.Controls.Add(BuildFontGroupBox());
        settingsSection.Controls.Add(BuildOptionsGroupBox());
        settingsSection.Controls.Add(BuildEngineGroupBox());
        settingsSection.Controls.Add(buttonPanel);
        settingsSection.Controls.Add(_resultLabel);

        // ── ルートレイアウト（3段構成）───────────────────────────────────
        // Row 0 (Percent): Markdown テキスト入力（伸縮可・Percent 行が最終行）
        // Row 1 (AutoSize): ファイル入出力パス（常に表示）
        // Row 2 (AutoSize): 各種設定・実行ボタン（常に表示）
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(12),
        };
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));  // inputSection（伸縮）
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // ioSection（常時表示）
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // settingsSection（常時表示）
        root.Controls.Add(inputSection);
        root.Controls.Add(ioSection);
        root.Controls.Add(settingsSection);

        Controls.Add(root);

        AllowDrop = true;
        _markdownTextBox.AllowDrop = true;
        DragEnter += OnFileDragEnter;
        DragDrop += OnFileDragDrop;
        _markdownTextBox.DragEnter += OnFileDragEnter;
        _markdownTextBox.DragDrop += OnFileDragDrop;

        _textInputRadio.CheckedChanged += (_, _) => { if (_textInputRadio.Checked) RefreshInputMode(); };
        _fileInputRadio.CheckedChanged += (_, _) => { if (_fileInputRadio.Checked) RefreshInputMode(); };
        _clearButton.Click += (_, _) => _markdownTextBox.Clear();
        _inputFilePathTextBox.TextChanged += (_, _) => SuggestOutputPath();
        _inputBrowseButton.Click += (_, _) => BrowseInputFile();
        _outputBrowseButton.Click += (_, _) => BrowseOutputFile();
        _headerNoneRadio.CheckedChanged += (_, _) => { if (_headerNoneRadio.Checked) RefreshHeaderMode(); };
        _headerFileNameRadio.CheckedChanged += (_, _) => { if (_headerFileNameRadio.Checked) RefreshHeaderMode(); };
        _headerCustomRadio.CheckedChanged += (_, _) => { if (_headerCustomRadio.Checked) RefreshHeaderMode(); };
        _bodyRecentFontCombo.SelectedIndexChanged += (_, _) => SyncRecentToMain(_bodyRecentFontCombo, _bodyFontCombo);
        _openFileButton.Click += (_, _) => OpenLastOutputFile();
        _openFolderButton.Click += (_, _) => OpenLastOutputFolder();
        _convertButton.Click += async (_, _) => await ConvertAsync();

        RefreshInputMode();
        RefreshHeaderMode();

        // 保存済み設定を復元（フォント・各種設定・ウィンドウサイズ）
        var settings = UserSettings.Load();
        _lastOutputFolder = settings.LastOutputFolder;
        ApplySettings(settings);

        // ApplySettings でメインコンボの選択が確定してから候補を初期化する
        var history = FontHistory.Load();
        PopulateRecentFontCombo(_bodyRecentFontCombo, _bodyFontCombo, history);

        SuggestOutputPath();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        CaptureSettings().Save();
        base.OnFormClosing(e);
    }

    private GroupBox BuildFontGroupBox()
    {
        const int SizeLabelWidth = 74;
        const int RowHeight = 24;

        static Label RowLabel(string text, int width) => new()
        {
            Text = text, Width = width, Height = RowHeight,
            AutoSize = false, TextAlign = ContentAlignment.MiddleLeft,
        };

        var bodyRow = new TableLayoutPanel
        {
            ColumnCount = 4, RowCount = 1, AutoSize = true, Dock = DockStyle.Fill, Margin = Padding.Empty,
        };
        bodyRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bodyRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bodyRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bodyRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bodyRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bodyRow.Controls.Add(_bodyRecentFontCombo, 0, 0);
        bodyRow.Controls.Add(_bodyFontCombo, 1, 0);
        bodyRow.Controls.Add(RowLabel("サイズ(pt):", SizeLabelWidth), 2, 0);
        bodyRow.Controls.Add(_bodyFontSizeNumeric, 3, 0);

        var box = new GroupBox { Text = "フォント設定", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 6) };
        box.Controls.Add(bodyRow);
        return box;
    }

    private GroupBox BuildOptionsGroupBox()
    {
        var headerAlignPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        headerAlignPanel.Controls.Add(_headerAlignLeftRadio);
        headerAlignPanel.Controls.Add(_headerAlignCenterRadio);
        headerAlignPanel.Controls.Add(_headerAlignRightRadio);

        var footerAlignPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        footerAlignPanel.Controls.Add(_footerAlignLeftRadio);
        footerAlignPanel.Controls.Add(_footerAlignCenterRadio);
        footerAlignPanel.Controls.Add(_footerAlignRightRadio);

        var headerRowPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        headerRowPanel.Controls.Add(_headerNoneRadio);
        headerRowPanel.Controls.Add(_headerFileNameRadio);
        headerRowPanel.Controls.Add(_headerCustomRadio);
        headerRowPanel.Controls.Add(_headerCustomTextBox);
        headerRowPanel.Controls.Add(new Label { Text = "  ", AutoSize = true });
        headerRowPanel.Controls.Add(headerAlignPanel);

        var footerRowPanel = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        footerRowPanel.Controls.Add(_footerPageNumberCheck);
        footerRowPanel.Controls.Add(new Label { Text = "  ", AutoSize = true });
        footerRowPanel.Controls.Add(footerAlignPanel);

        var table = new TableLayoutPanel { ColumnCount = 2, RowCount = 3, AutoSize = true, Dock = DockStyle.Fill };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "見出し:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 0);
        table.Controls.Add(_headingNumberCheck, 1, 0);
        table.Controls.Add(new Label { Text = "ヘッダー:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 1);
        table.Controls.Add(headerRowPanel, 1, 1);
        table.Controls.Add(new Label { Text = "フッター:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 2);
        table.Controls.Add(footerRowPanel, 1, 2);

        var box = new GroupBox { Text = "文書オプション", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 6) };
        box.Controls.Add(table);
        return box;
    }

    private GroupBox BuildEngineGroupBox()
    {
        var table = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, AutoSize = true, Dock = DockStyle.Fill };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.Controls.Add(_engineOpenXmlRadio, 0, 0);
        table.Controls.Add(_engineWordComRadio, 1, 0);

        var box = new GroupBox { Text = "変換エンジン", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(6, 2, 6, 6) };
        box.Controls.Add(table);
        return box;
    }

    private ComboBox CreateFontComboBox(string defaultFont)
    {
        var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        combo.Items.AddRange(_allFonts);
        var idx = Array.IndexOf(_allFonts, defaultFont);
        combo.SelectedIndex = idx >= 0 ? idx : 0;
        return combo;
    }

    private const string RecentFontPlaceholder = "── 履歴から選択 ──";

    private void PopulateRecentFontCombo(ComboBox recentCombo, ComboBox mainCombo, IReadOnlyList<string> history)
    {
        _suppressFontSync = true;
        try
        {
            recentCombo.Items.Clear();
            var valid = history.Where(f => Array.IndexOf(_allFonts, f) >= 0).ToList();
            if (valid.Count == 0)
            {
                recentCombo.Items.Add("（履歴なし）");
                recentCombo.SelectedIndex = 0;
                recentCombo.Enabled = false;
            }
            else
            {
                recentCombo.Items.Add(RecentFontPlaceholder);
                recentCombo.Items.AddRange(valid.ToArray<object>());
                recentCombo.SelectedIndex = 0;
                recentCombo.Enabled = true;
            }
        }
        finally
        {
            _suppressFontSync = false;
        }
    }

    private void SyncRecentToMain(ComboBox recentCombo, ComboBox mainCombo)
    {
        if (_suppressFontSync) return;
        if (!recentCombo.Enabled) return;
        var selected = recentCombo.SelectedItem as string;
        // プレースホルダー行は何もしない
        if (selected is null || selected == RecentFontPlaceholder) return;
        var idx = mainCombo.FindStringExact(selected);
        if (idx >= 0)
            mainCombo.SelectedIndex = idx;
    }

    private static int GetSelectedAlignment(RadioButton leftRadio, RadioButton centerRadio)
    {
        if (leftRadio.Checked) return 0;
        if (centerRadio.Checked) return 1;
        return 2;
    }

    private void ApplySettings(UserSettings s)
    {
        if (s.WindowWidth >= 500) Width = s.WindowWidth;
        if (s.WindowHeight >= 400) Height = s.WindowHeight;

        var bIdx = _bodyFontCombo.FindStringExact(s.BodyFontName);
        if (bIdx >= 0) _bodyFontCombo.SelectedIndex = bIdx;
        _bodyFontSizeNumeric.Value = (decimal)Math.Clamp(s.BodyFontSize, 8, 72);

        _headingNumberCheck.Checked = s.HeadingNumbering;

        if (s.HeaderMode == 1) _headerFileNameRadio.Checked = true;
        else if (s.HeaderMode == 2) _headerCustomRadio.Checked = true;
        _headerCustomTextBox.Text = s.HeaderCustomText;

        switch (s.HeaderAlignment)
        {
            case 1: _headerAlignCenterRadio.Checked = true; break;
            case 2: _headerAlignRightRadio.Checked = true; break;
            default: _headerAlignLeftRadio.Checked = true; break;
        }

        _footerPageNumberCheck.Checked = s.FooterPageNumber;
        switch (s.FooterAlignment)
        {
            case 0: _footerAlignLeftRadio.Checked = true; break;
            case 2: _footerAlignRightRadio.Checked = true; break;
            default: _footerAlignCenterRadio.Checked = true; break;
        }

        // 既存設定に ConversionEngine がない場合は "OpenXml"（UserSettings 既定値）になる
        if (s.ConversionEngine == "WordCom")
            _engineWordComRadio.Checked = true;
        else
            _engineOpenXmlRadio.Checked = true;
    }

    private UserSettings CaptureSettings() => new()
    {
        BodyFontName = (string?)_bodyFontCombo.SelectedItem ?? "MS Gothic",
        BodyFontSize = (double)_bodyFontSizeNumeric.Value,
        HeadingNumbering = _headingNumberCheck.Checked,
        HeaderMode = _headerNoneRadio.Checked ? 0 : _headerFileNameRadio.Checked ? 1 : 2,
        HeaderCustomText = _headerCustomTextBox.Text,
        HeaderAlignment = GetSelectedAlignment(_headerAlignLeftRadio, _headerAlignCenterRadio),
        FooterPageNumber = _footerPageNumberCheck.Checked,
        FooterAlignment = GetSelectedAlignment(_footerAlignLeftRadio, _footerAlignCenterRadio),
        ConversionEngine = _engineOpenXmlRadio.Checked ? "OpenXml" : "WordCom",
        LastOutputFolder = _lastOutputFolder,
        WindowWidth = Width,
        WindowHeight = Height,
    };

    private void SuggestOutputPath()
    {
        var current = _outputFilePathTextBox.Text.Trim();
        // ユーザーが手動入力した場合は上書きしない
        if (!string.IsNullOrEmpty(current) && current != _suggestedOutputPath)
            return;

        string suggestion;
        if (_fileInputRadio.Checked)
        {
            var inputPath = _inputFilePathTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                _suggestedOutputPath = "";
                _outputFilePathTextBox.Text = "";
                return;
            }
            var dir = Path.GetDirectoryName(inputPath) ?? "";
            var name = Path.GetFileNameWithoutExtension(inputPath);
            suggestion = Path.Combine(dir, name + ".docx");
        }
        else
        {
            var folder = string.IsNullOrWhiteSpace(_lastOutputFolder)
                ? Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                : _lastOutputFolder;
            suggestion = Path.Combine(folder, "output.docx");
        }

        _suggestedOutputPath = suggestion;
        _outputFilePathTextBox.Text = suggestion;
    }

    private static void OnFileDragEnter(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        if (files?.Length > 0 && IsAcceptedFile(files[0]))
            e.Effect = DragDropEffects.Copy;
    }

    private void OnFileDragDrop(object? sender, DragEventArgs e)
    {
        var files = e.Data?.GetData(DataFormats.FileDrop) as string[];
        if (files?.Length > 0 && IsAcceptedFile(files[0]))
        {
            try
            {
                _markdownTextBox.Text = File.ReadAllText(files[0], Encoding.UTF8);
                _textInputRadio.Checked = true; // RefreshInputMode → SuggestOutputPath も呼ばれる
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"ファイルの読み込みに失敗しました: {ex.Message}",
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private static bool IsAcceptedFile(string path)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext is ".md" or ".txt";
    }

    private void OpenLastOutputFile()
    {
        if (string.IsNullOrEmpty(_lastOutputPath)) return;
        if (!File.Exists(_lastOutputPath))
        {
            MessageBox.Show(this, "ファイルが見つかりません。", "エラー",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        Process.Start(new ProcessStartInfo { FileName = _lastOutputPath, UseShellExecute = true });
    }

    private void OpenLastOutputFolder()
    {
        if (string.IsNullOrEmpty(_lastOutputPath)) return;
        var folder = Path.GetDirectoryName(_lastOutputPath);
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder)) return;
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,\"{_lastOutputPath}\""
        });
    }

    private void RefreshInputMode()
    {
        var useText = _textInputRadio.Checked;
        _markdownTextBox.Enabled = useText;
        _clearButton.Enabled = useText;
        _inputFilePathTextBox.Enabled = !useText;
        _inputBrowseButton.Enabled = !useText;
        _headerFileNameRadio.Enabled = !useText;
        if (useText && _headerFileNameRadio.Checked)
            _headerNoneRadio.Checked = true;
        SuggestOutputPath();
    }

    private void RefreshHeaderMode()
    {
        _headerCustomTextBox.Enabled = _headerCustomRadio.Checked;
    }

    private void BrowseInputFile()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Markdown/Text (*.md;*.txt)|*.md;*.txt|All files (*.*)|*.*",
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            _inputFilePathTextBox.Text = dialog.FileName; // TextChanged → SuggestOutputPath
    }

    private void BrowseOutputFile()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Word Document (*.docx)|*.docx",
            DefaultExt = "docx",
            InitialDirectory = !string.IsNullOrWhiteSpace(_lastOutputFolder)
                ? _lastOutputFolder
                : Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _suggestedOutputPath = ""; // 手動選択済みとしてマーク
            _outputFilePathTextBox.Text = dialog.FileName;
        }
    }

    private async Task ConvertAsync()
    {
        try
        {
            _convertButton.Enabled = false;
            _resultLabel.Text = "変換中... 0%";

            var markdown = GetMarkdownInput();
            var outputPath = _outputFilePathTextBox.Text.Trim();
            var bodyFontName = (string?)_bodyFontCombo.SelectedItem ?? "MS Gothic";
            var bodyFontSize = (double)_bodyFontSizeNumeric.Value;
            var numberHeadings = _headingNumberCheck.Checked;
            var headerText = GetHeaderText();
            var addPageNumbers = _footerPageNumberCheck.Checked;
            var headerAlignment = GetSelectedAlignment(_headerAlignLeftRadio, _headerAlignCenterRadio);
            var footerAlignment = GetSelectedAlignment(_footerAlignLeftRadio, _footerAlignCenterRadio);

            ValidateInput(markdown, outputPath);

            if (!ConfirmOverwrite(outputPath))
            {
                _resultLabel.Text = "変換をキャンセルしました。";
                return;
            }

            var progress = new Progress<int>(pct => _resultLabel.Text = $"変換中... {pct}%");
            bool useOpenXml = _engineOpenXmlRadio.Checked;

            if (useOpenXml)
            {
                await Task.Run(() => OpenXmlConverter.ConvertToDocx(
                    markdown, outputPath,
                    bodyFontName, bodyFontSize,
                    numberHeadings,
                    headerText, headerAlignment,
                    addPageNumbers, footerAlignment,
                    progress));
            }
            else
            {
                await Task.Run(() => WordInteropConverter.ConvertToDocx(
                    markdown, outputPath,
                    bodyFontName, bodyFontSize,
                    numberHeadings,
                    headerText, headerAlignment,
                    addPageNumbers, footerAlignment,
                    progress));
            }

            _lastOutputPath = outputPath;
            _lastOutputFolder = Path.GetDirectoryName(outputPath) ?? _lastOutputFolder;
            _openFileButton.Enabled = true;
            _openFolderButton.Enabled = true;

            var history = FontHistory.Update(bodyFontName);
            PopulateRecentFontCombo(_bodyRecentFontCombo, _bodyFontCombo, history);

            CaptureSettings().Save();
            var engineLabel = useOpenXml ? "[Open XML]" : "[Word COM]";
            _resultLabel.Text = $"変換完了 {engineLabel}: {outputPath}";
        }
        catch (Exception ex)
        {
            AppLog.Error("変換失敗", ex);
            _resultLabel.Text = $"変換失敗: {ex.Message}";
        }
        finally
        {
            _convertButton.Enabled = true;
        }
    }

    private string GetMarkdownInput()
    {
        if (_textInputRadio.Checked)
            return _markdownTextBox.Text;

        var inputPath = _inputFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(inputPath))
            throw new InvalidOperationException("入力ファイルを指定してください。");
        if (!File.Exists(inputPath))
            throw new FileNotFoundException("入力ファイルが見つかりません。", inputPath);

        return File.ReadAllText(inputPath, Encoding.UTF8);
    }

    private string? GetHeaderText()
    {
        if (_headerNoneRadio.Checked) return null;
        if (_headerFileNameRadio.Checked)
        {
            var path = _inputFilePathTextBox.Text.Trim();
            return string.IsNullOrWhiteSpace(path) ? null : Path.GetFileName(path);
        }
        var custom = _headerCustomTextBox.Text.Trim();
        return string.IsNullOrEmpty(custom) ? null : custom;
    }

    private static void ValidateInput(string markdown, string outputPath)
    {
        if (string.IsNullOrWhiteSpace(markdown))
            throw new InvalidOperationException("Markdown本文が空です。");
        if (string.IsNullOrWhiteSpace(outputPath))
            throw new InvalidOperationException("出力ファイルを指定してください。");

        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory))
            throw new InvalidOperationException("出力先フォルダが存在しません。");
    }

    private bool ConfirmOverwrite(string outputPath)
    {
        if (!File.Exists(outputPath)) return true;

        var result = MessageBox.Show(
            this,
            "出力ファイルは既に存在します。上書きしますか？",
            "上書き確認",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        return result == DialogResult.Yes;
    }
}
