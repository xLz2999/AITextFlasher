using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using HtmlToOpenXml;
using Markdig;
using System.IO;

namespace AITextFlasher
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        // 按钮1：转纯文本并复制到剪贴板
        private async void BtnToText_Click(object sender, RoutedEventArgs e)
        {
            string mdText = InputTextBox.Text;
            if (string.IsNullOrWhiteSpace(mdText)) return;

            // 使用 Markdig 把 Markdown 转成纯文本
            string plainText = Markdown.ToPlainText(mdText);

            // 调用 Avalonia 的剪贴板功能
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard != null)
            {
                await clipboard.SetTextAsync(plainText);
                StatusText.Text = "✅ 纯文本已复制！";
            }
        }

        // 按钮2：转成 Word 并保存
        private async void BtnToWord_Click(object sender, RoutedEventArgs e)
        {
            string mdText = InputTextBox.Text;
            if (string.IsNullOrWhiteSpace(mdText)) return;

            // 1. 弹出保存文件对话框，让用户选择保存位置
            var topLevel = TopLevel.GetTopLevel(this);
            var file = await topLevel.StorageProvider.SaveFilePickerAsync(new Avalonia.Platform.Storage.FilePickerSaveOptions
            {
                Title = "保存Word文档",
                DefaultExtension = "docx",
                SuggestedFileName = "AI导出文本.docx"
            });

            if (file != null)
            {
                // 2. 将 Markdown 转为 HTML (作为中间媒介)
                var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
                string html = Markdown.ToHtml(mdText, pipeline);

                // 3. 将 HTML 写入 Word 文档
                using (MemoryStream generatedDocument = new MemoryStream())
                {
                    using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = package.MainDocumentPart ?? package.AddMainDocumentPart();
                        HtmlConverter converter = new HtmlConverter(mainPart);
                        converter.ParseHtml(html);
                        mainPart.Document.Save();
                    }

                    // 把在内存中生成的Word保存到用户选择的文件路径里
                    using (var fileStream = await file.OpenWriteAsync())
                    {
                        generatedDocument.Position = 0;
                        await generatedDocument.CopyToAsync(fileStream);
                    }
                }
                StatusText.Text = "✅ Word已成功导出！";
            }
        }
    }
}
