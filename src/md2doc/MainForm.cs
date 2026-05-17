using System.Drawing;
using System.Text;

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

    // フォント（見出し / 本文）
    private readonly string[] _allFonts;
    private readonly ComboBox _headingRecentFontCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
    private readonly ComboBox _headingFontCombo;
    private readonly NumericUpDown _headingFontSizeNumeric = new()
    {
        Minimum = 8, Maximum = 72, DecimalPlaces = 1, Value = 14, Width = 70
    };
    private readonly ComboBox _bodyRecentFontCombo = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
    private readonly ComboBox _bodyFontCombo;
    private readonly NumericUpDown _bodyFontSizeNumeric = new()
    {
        Minimum = 8, Maximum = 72, DecimalPlaces = 1, Value = 11, Width = 70
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

    // 実行・結果
    private readonly Button _convertButton = new() { Text = "変換実行", AutoSize = true };
    private readonly Label _resultLabel = new() { AutoSize = true };

    public MainForm()
    {
        Text = "Markdown変換ツール（Word）";
        Width = 960;
        Height = 860;

        _allFonts = FontFamily.Families
            .Select(f => f.Name)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        _headingFontCombo = CreateFontComboBox("MS Gothic");
        _bodyFontCombo = CreateFontComboBox("MS Gothic");

        var history = FontHistory.Load();
        PopulateRecentFontCombo(_headingRecentFontCombo, history);
        PopulateRecentFontCombo(_bodyRecentFontCombo, history);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 12,
            Padding = new Padding(10),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 入力方式 label
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // inputModePanel
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // Markdown本文 label + clearButton
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // markdownTextBox
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 入力ファイル label
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // inputFilePanel
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 出力ファイル label
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // outputPanel
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // fontGroupBox
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // optionsGroupBox
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // convertButton
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // resultLabel

        var inputModePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        inputModePanel.Controls.Add(_textInputRadio);
        inputModePanel.Controls.Add(_fileInputRadio);

        var mdLabelPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        mdLabelPanel.Controls.Add(new Label { Text = "Markdown本文（テキスト入力）", AutoSize = true });
        mdLabelPanel.Controls.Add(_clearButton);

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

        root.Controls.Add(new Label { Text = "入力方式", AutoSize = true });
        root.Controls.Add(inputModePanel);
        root.Controls.Add(mdLabelPanel);
        root.Controls.Add(_markdownTextBox);
        root.Controls.Add(new Label { Text = "入力ファイル（ファイル入力）", AutoSize = true });
        root.Controls.Add(inputFilePanel);
        root.Controls.Add(new Label { Text = "出力ファイル（.docx）", AutoSize = true });
        root.Controls.Add(outputPanel);
        root.Controls.Add(BuildFontGroupBox());
        root.Controls.Add(BuildOptionsGroupBox());
        root.Controls.Add(_convertButton);
        root.Controls.Add(_resultLabel);

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
        _inputBrowseButton.Click += (_, _) => BrowseInputFile();
        _outputBrowseButton.Click += (_, _) => BrowseOutputFile();
        _headerNoneRadio.CheckedChanged += (_, _) => { if (_headerNoneRadio.Checked) RefreshHeaderMode(); };
        _headerFileNameRadio.CheckedChanged += (_, _) => { if (_headerFileNameRadio.Checked) RefreshHeaderMode(); };
        _headerCustomRadio.CheckedChanged += (_, _) => { if (_headerCustomRadio.Checked) RefreshHeaderMode(); };
        _headingRecentFontCombo.SelectedIndexChanged += (_, _) => SyncRecentToMain(_headingRecentFontCombo, _headingFontCombo);
        _bodyRecentFontCombo.SelectedIndexChanged += (_, _) => SyncRecentToMain(_bodyRecentFontCombo, _bodyFontCombo);
        _convertButton.Click += async (_, _) => await ConvertAsync();

        RefreshInputMode();
        RefreshHeaderMode();
        SyncRecentToMain(_headingRecentFontCombo, _headingFontCombo);
        SyncRecentToMain(_bodyRecentFontCombo, _bodyFontCombo);
    }

    private GroupBox BuildFontGroupBox()
    {
        // ラベルを固定幅・固定高さにすることで FlowLayoutPanel 内でコンボと縦位置を揃える
        const int LabelWidth = 50;
        const int SizeLabelWidth = 74;
        const int RowHeight = 24;

        static Label RowLabel(string text, int width) => new()
        {
            Text = text, Width = width, Height = RowHeight,
            AutoSize = false, TextAlign = ContentAlignment.MiddleLeft,
        };

        var headingRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
        headingRow.Controls.Add(RowLabel("見出し:", LabelWidth));
        headingRow.Controls.Add(_headingRecentFontCombo);
        headingRow.Controls.Add(_headingFontCombo);
        headingRow.Controls.Add(RowLabel("サイズ(pt):", SizeLabelWidth));
        headingRow.Controls.Add(_headingFontSizeNumeric);

        var bodyRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = Padding.Empty };
        bodyRow.Controls.Add(RowLabel("本文:", LabelWidth));
        bodyRow.Controls.Add(_bodyRecentFontCombo);
        bodyRow.Controls.Add(_bodyFontCombo);
        bodyRow.Controls.Add(RowLabel("サイズ(pt):", SizeLabelWidth));
        bodyRow.Controls.Add(_bodyFontSizeNumeric);

        // 1列2行の TableLayoutPanel でヘッダーとボディ行を縦積みする
        var stack = new TableLayoutPanel
        {
            ColumnCount = 1, RowCount = 2, AutoSize = true, Margin = Padding.Empty
        };
        stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        stack.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        stack.Controls.Add(headingRow, 0, 0);
        stack.Controls.Add(bodyRow, 0, 1);

        var box = new GroupBox { Text = "フォント設定", Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(6) };
        box.Controls.Add(stack);
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

        var table = new TableLayoutPanel { ColumnCount = 2, RowCount = 2, AutoSize = true };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        table.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "ヘッダー:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 0);
        table.Controls.Add(headerRowPanel, 1, 0);
        table.Controls.Add(new Label { Text = "フッター:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 1);
        table.Controls.Add(footerRowPanel, 1, 1);

        var box = new GroupBox { Text = "オプション", Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(6) };
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

    private void PopulateRecentFontCombo(ComboBox combo, IReadOnlyList<string> history)
    {
        combo.Items.Clear();
        var valid = history.Where(f => Array.IndexOf(_allFonts, f) >= 0).ToList();
        if (valid.Count == 0)
        {
            combo.Items.Add("（なし）");
            combo.SelectedIndex = 0;
            combo.Enabled = false;
        }
        else
        {
            combo.Items.AddRange(valid.ToArray<object>());
            combo.SelectedIndex = 0;
            combo.Enabled = true;
        }
    }

    private void SyncRecentToMain(ComboBox recentCombo, ComboBox mainCombo)
    {
        if (!recentCombo.Enabled) return;
        var selected = recentCombo.SelectedItem as string;
        if (selected is null) return;
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
                _textInputRadio.Checked = true;
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
            _inputFilePathTextBox.Text = dialog.FileName;
    }

    private void BrowseOutputFile()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Word Document (*.docx)|*.docx",
            DefaultExt = "docx",
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
            _outputFilePathTextBox.Text = dialog.FileName;
    }

    private async Task ConvertAsync()
    {
        try
        {
            _convertButton.Enabled = false;
            _resultLabel.Text = "変換中... 0%";

            var markdown = GetMarkdownInput();
            var outputPath = _outputFilePathTextBox.Text.Trim();
            var headingFontName = (string?)_headingFontCombo.SelectedItem ?? "MS Gothic";
            var headingFontSize = (double)_headingFontSizeNumeric.Value;
            var bodyFontName = (string?)_bodyFontCombo.SelectedItem ?? "MS Gothic";
            var bodyFontSize = (double)_bodyFontSizeNumeric.Value;
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

            await Task.Run(() => WordInteropConverter.ConvertToDocx(
                markdown, outputPath,
                headingFontName, headingFontSize,
                bodyFontName, bodyFontSize,
                headerText, headerAlignment,
                addPageNumbers, footerAlignment,
                progress));

            var history = FontHistory.Update(headingFontName, bodyFontName);
            PopulateRecentFontCombo(_headingRecentFontCombo, history);
            PopulateRecentFontCombo(_bodyRecentFontCombo, history);

            _resultLabel.Text = $"変換完了: {outputPath}";
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
