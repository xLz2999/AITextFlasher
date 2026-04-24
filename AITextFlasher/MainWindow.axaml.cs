using Avalonia.Controls;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using HtmlToOpenXml;
using System.Threading.Tasks;
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
        // 🌟 新功能 1：一键清空按钮
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            InputTextBox.Text = "";
            StatusText.Text = ""; // 同时清空提示语
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
                // 🌟 新功能 2：阅后即焚提示语
                StatusText.Text = "✅ 纯文本已复制！";
                await Task.Delay(3000); // 魔法指令：让程序偷偷等待 3 秒
                if (StatusText.Text == "✅ 纯文本已复制！") // 防止这 3 秒内你点了别的按钮
                {
                    StatusText.Text = "";
                }
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
                // 🌟 新功能 2：阅后即焚提示语
                StatusText.Text = "✅ Word已成功导出！";
                await Task.Delay(3000); // 等待 3 秒
                if (StatusText.Text == "✅ Word已成功导出！")
                {
                    StatusText.Text = "";
                }
            }
        }
    }
}
