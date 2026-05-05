using System.Text;

namespace Md2Doc;

public sealed class MainForm : Form
{
    private readonly RadioButton _textInputRadio = new()
    {
        Text = "テキスト入力",
        Checked = true,
        AutoSize = true
    };
    private readonly RadioButton _fileInputRadio = new() { Text = "ファイル入力", AutoSize = true };
    private readonly TextBox _markdownTextBox = new()
    {
        Multiline = true,
        ScrollBars = ScrollBars.Both,
        WordWrap = false,
        Dock = DockStyle.Fill
    };
    private readonly TextBox _inputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _inputBrowseButton = new() { Text = "参照..." };
    private readonly TextBox _outputFilePathTextBox = new() { Dock = DockStyle.Fill };
    private readonly Button _outputBrowseButton = new() { Text = "参照..." };
    private readonly TextBox _fontNameTextBox = new() { Text = "MS Gothic", Width = 200 };
    private readonly NumericUpDown _fontSizeNumeric = new()
    {
        Minimum = 8,
        Maximum = 72,
        DecimalPlaces = 1,
        Value = 11
    };
    private readonly Button _convertButton = new() { Text = "変換実行", AutoSize = true };
    private readonly Label _resultLabel = new() { AutoSize = true };

    public MainForm()
    {
        Text = "Markdown変換ツール（Word）";
        Width = 920;
        Height = 720;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 11,
            Padding = new Padding(10),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

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

        var fontPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        fontPanel.Controls.Add(new Label { Text = "フォント名:" });
        fontPanel.Controls.Add(_fontNameTextBox);
        fontPanel.Controls.Add(new Label { Text = "フォントサイズ:" });
        fontPanel.Controls.Add(_fontSizeNumeric);

        root.Controls.Add(new Label { Text = "入力方式", AutoSize = true });
        root.Controls.Add(inputModePanel);
        root.Controls.Add(new Label { Text = "Markdown本文（テキスト入力）", AutoSize = true });
        root.Controls.Add(_markdownTextBox);
        root.Controls.Add(new Label { Text = "入力ファイル（ファイル入力）", AutoSize = true });
        root.Controls.Add(inputFilePanel);
        root.Controls.Add(new Label { Text = "出力ファイル（.docx）", AutoSize = true });
        root.Controls.Add(outputPanel);
        root.Controls.Add(fontPanel);
        root.Controls.Add(_convertButton);
        root.Controls.Add(_resultLabel);

        Controls.Add(root);

        _textInputRadio.CheckedChanged += (_, _) => RefreshInputMode();
        _fileInputRadio.CheckedChanged += (_, _) => RefreshInputMode();
        _inputBrowseButton.Click += (_, _) => BrowseInputFile();
        _outputBrowseButton.Click += (_, _) => BrowseOutputFile();
        _convertButton.Click += async (_, _) => await ConvertAsync();

        RefreshInputMode();
    }

    private void RefreshInputMode()
    {
        var useText = _textInputRadio.Checked;
        _markdownTextBox.Enabled = useText;
        _inputFilePathTextBox.Enabled = !useText;
        _inputBrowseButton.Enabled = !useText;
    }

    private void BrowseInputFile()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Markdown/Text (*.md;*.txt)|*.md;*.txt|All files (*.*)|*.*",
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _inputFilePathTextBox.Text = dialog.FileName;
        }
    }

    private void BrowseOutputFile()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Word Document (*.docx)|*.docx",
            DefaultExt = "docx",
        };

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            _outputFilePathTextBox.Text = dialog.FileName;
        }
    }

    private async Task ConvertAsync()
    {
        try
        {
            _convertButton.Enabled = false;
            _resultLabel.Text = string.Empty;

            var markdown = GetMarkdownInput();
            var outputPath = _outputFilePathTextBox.Text.Trim();
            var fontName = _fontNameTextBox.Text.Trim();
            var fontSize = (double)_fontSizeNumeric.Value;

            ValidateInput(markdown, outputPath, fontName);
            EnsureOverwriteConfirmed(outputPath);

            await Task.Run(
                () => WordInteropConverter.ConvertToDocx(
                    markdown,
                    outputPath,
                    fontName,
                    fontSize));

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
        {
            return _markdownTextBox.Text;
        }

        var inputPath = _inputFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(inputPath))
        {
            throw new InvalidOperationException("入力ファイルを指定してください。");
        }

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("入力ファイルが見つかりません。", inputPath);
        }

        return File.ReadAllText(inputPath, Encoding.UTF8);
    }

    private static void ValidateInput(string markdown, string outputPath, string fontName)
    {
        if (string.IsNullOrWhiteSpace(markdown))
        {
            throw new InvalidOperationException("Markdown本文が空です。");
        }

        if (string.IsNullOrWhiteSpace(outputPath))
        {
            throw new InvalidOperationException("出力ファイルを指定してください。");
        }

        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (string.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory))
        {
            throw new InvalidOperationException("出力先フォルダが存在しません。");
        }

        if (string.IsNullOrWhiteSpace(fontName))
        {
            throw new InvalidOperationException("フォント名を指定してください。");
        }
    }

    private void EnsureOverwriteConfirmed(string outputPath)
    {
        if (!File.Exists(outputPath))
        {
            return;
        }

        var result = MessageBox.Show(
            this,
            "出力ファイルは既に存在します。上書きしますか？",
            "上書き確認",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (result != DialogResult.Yes)
        {
            throw new InvalidOperationException("上書きがキャンセルされました。");
        }
    }
}
