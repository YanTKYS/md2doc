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

    // ファイル入力
    private readonly TextBox _inputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _inputBrowseButton = new() { Text = "参照..." };

    // 出力
    private readonly TextBox _outputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _outputBrowseButton = new() { Text = "参照..." };

    // フォント（見出し / 本文）
    private readonly ComboBox _headingFontCombo;
    private readonly NumericUpDown _headingFontSizeNumeric = new()
    {
        Minimum = 8, Maximum = 72, DecimalPlaces = 1, Value = 14, Width = 70
    };
    private readonly ComboBox _bodyFontCombo;
    private readonly NumericUpDown _bodyFontSizeNumeric = new()
    {
        Minimum = 8, Maximum = 72, DecimalPlaces = 1, Value = 11, Width = 70
    };

    // ヘッダー
    private readonly RadioButton _headerNoneRadio = new() { Text = "なし", Checked = true, AutoSize = true };
    private readonly RadioButton _headerFileNameRadio = new() { Text = "元ファイル名", AutoSize = true };
    private readonly RadioButton _headerCustomRadio = new() { Text = "自由記入:", AutoSize = true };
    private readonly TextBox _headerCustomTextBox = new() { Width = 220, Enabled = false };

    // フッター
    private readonly CheckBox _footerPageNumberCheck = new()
    {
        Text = "ページ番号を挿入（ページ番号/ページ数）", AutoSize = true
    };

    // 実行・結果
    private readonly Button _convertButton = new() { Text = "変換実行", AutoSize = true };
    private readonly Label _resultLabel = new() { AutoSize = true };

    public MainForm()
    {
        Text = "Markdown変換ツール（Word）";
        Width = 960;
        Height = 800;

        _headingFontCombo = CreateFontComboBox("MS Gothic");
        _bodyFontCombo = CreateFontComboBox("MS Gothic");

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 12,
            Padding = new Padding(10),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // 入力方式 label
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // inputModePanel
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // Markdown本文 label
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
        root.Controls.Add(new Label { Text = "Markdown本文（テキスト入力）", AutoSize = true });
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

        _textInputRadio.CheckedChanged += (_, _) => { if (_textInputRadio.Checked) RefreshInputMode(); };
        _fileInputRadio.CheckedChanged += (_, _) => { if (_fileInputRadio.Checked) RefreshInputMode(); };
        _inputBrowseButton.Click += (_, _) => BrowseInputFile();
        _outputBrowseButton.Click += (_, _) => BrowseOutputFile();
        _headerNoneRadio.CheckedChanged += (_, _) => { if (_headerNoneRadio.Checked) RefreshHeaderMode(); };
        _headerFileNameRadio.CheckedChanged += (_, _) => { if (_headerFileNameRadio.Checked) RefreshHeaderMode(); };
        _headerCustomRadio.CheckedChanged += (_, _) => { if (_headerCustomRadio.Checked) RefreshHeaderMode(); };
        _convertButton.Click += async (_, _) => await ConvertAsync();

        RefreshInputMode();
        RefreshHeaderMode();
    }

    private GroupBox BuildFontGroupBox()
    {
        var table = new TableLayoutPanel { ColumnCount = 4, AutoSize = true };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        table.Controls.Add(new Label { Text = "見出し:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        table.Controls.Add(_headingFontCombo, 1, 0);
        table.Controls.Add(new Label { Text = "サイズ(pt):", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
        table.Controls.Add(_headingFontSizeNumeric, 3, 0);

        table.Controls.Add(new Label { Text = "本文:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        table.Controls.Add(_bodyFontCombo, 1, 1);
        table.Controls.Add(new Label { Text = "サイズ(pt):", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 1);
        table.Controls.Add(_bodyFontSizeNumeric, 3, 1);

        var box = new GroupBox { Text = "フォント設定", Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(6) };
        box.Controls.Add(table);
        return box;
    }

    private GroupBox BuildOptionsGroupBox()
    {
        var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        var headerPanel = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        headerPanel.Controls.Add(_headerNoneRadio);
        headerPanel.Controls.Add(_headerFileNameRadio);
        headerPanel.Controls.Add(_headerCustomRadio);
        headerPanel.Controls.Add(_headerCustomTextBox);

        var footerPanel = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill };
        footerPanel.Controls.Add(_footerPageNumberCheck);

        table.Controls.Add(new Label { Text = "ヘッダー:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 0);
        table.Controls.Add(headerPanel, 1, 0);
        table.Controls.Add(new Label { Text = "フッター:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top }, 0, 1);
        table.Controls.Add(footerPanel, 1, 1);

        var box = new GroupBox { Text = "オプション", Dock = DockStyle.Fill, AutoSize = true, Padding = new Padding(6) };
        box.Controls.Add(table);
        return box;
    }

    private static ComboBox CreateFontComboBox(string defaultFont)
    {
        var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220 };
        var fonts = FontFamily.Families
            .Select(f => f.Name)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        combo.Items.AddRange(fonts);
        var idx = Array.IndexOf(fonts, defaultFont);
        combo.SelectedIndex = idx >= 0 ? idx : 0;
        return combo;
    }

    private void RefreshInputMode()
    {
        var useText = _textInputRadio.Checked;
        _markdownTextBox.Enabled = useText;
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
            _resultLabel.Text = string.Empty;

            var markdown = GetMarkdownInput();
            var outputPath = _outputFilePathTextBox.Text.Trim();
            var headingFontName = (string?)_headingFontCombo.SelectedItem ?? "MS Gothic";
            var headingFontSize = (double)_headingFontSizeNumeric.Value;
            var bodyFontName = (string?)_bodyFontCombo.SelectedItem ?? "MS Gothic";
            var bodyFontSize = (double)_bodyFontSizeNumeric.Value;
            var headerText = GetHeaderText();
            var addPageNumbers = _footerPageNumberCheck.Checked;

            ValidateInput(markdown, outputPath);

            if (!ConfirmOverwrite(outputPath))
            {
                _resultLabel.Text = "変換をキャンセルしました。";
                return;
            }

            await Task.Run(() => WordInteropConverter.ConvertToDocx(
                markdown, outputPath,
                headingFontName, headingFontSize,
                bodyFontName, bodyFontSize,
                headerText, addPageNumbers));

            _resultLabel.Text = $"変換完了: {outputPath}";
        }
        catch (Exception ex)
        {
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
