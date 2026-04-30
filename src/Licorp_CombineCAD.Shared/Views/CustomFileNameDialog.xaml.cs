using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace Licorp_CombineCAD.Views
{
    public partial class CustomFileNameDialog : Window
    {
        private readonly string _projectNumber;
        private readonly string _projectName;

        public CustomFileNameDialog(string template, string projectNumber, string projectName)
        {
            InitializeComponent();
            _projectNumber = string.IsNullOrWhiteSpace(projectNumber) ? "PRJ-001" : projectNumber;
            _projectName = string.IsNullOrWhiteSpace(projectName) ? "Project" : projectName;
            TemplateTextBox.Text = string.IsNullOrWhiteSpace(template)
                ? "{SheetNumber} - {SheetName}"
                : template;

            Loaded += (s, e) =>
            {
                TemplateTextBox.Focus();
                TemplateTextBox.CaretIndex = TemplateTextBox.Text.Length;
                UpdatePreview();
            };
        }

        public string FileNameTemplate { get; private set; }

        private void TokenButton_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button button) || button.Tag == null)
                return;

            var token = button.Tag.ToString();
            var index = TemplateTextBox.CaretIndex;
            TemplateTextBox.Text = TemplateTextBox.Text.Insert(index, token);
            TemplateTextBox.CaretIndex = index + token.Length;
            TemplateTextBox.Focus();
            UpdatePreview();
        }

        private void TemplateTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePreview();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            FileNameTemplate = string.IsNullOrWhiteSpace(TemplateTextBox.Text)
                ? "{SheetNumber} - {SheetName}"
                : TemplateTextBox.Text.Trim();
            DialogResult = true;
        }

        private void UpdatePreview()
        {
            if (PreviewTextBlock == null || TemplateTextBox == null)
                return;

            var preview = (TemplateTextBox.Text ?? "")
                .Replace("{SheetNumber}", "A101")
                .Replace("{SheetName}", "Floor Plan")
                .Replace("{PaperSize}", "A1")
                .Replace("{ProjectNumber}", _projectNumber)
                .Replace("{ProjectName}", _projectName);

            foreach (var c in Path.GetInvalidFileNameChars())
                preview = preview.Replace(c, '-');

            preview = preview.Trim();
            if (string.IsNullOrWhiteSpace(preview))
                preview = "A101";

            PreviewTextBlock.Text = preview + ".dwg";
        }
    }
}
