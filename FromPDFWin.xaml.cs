using Aspose.Pdf.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;
using Aspose.Pdf;

namespace Converter.Pages
{
    /// <summary>
    /// Логика взаимодействия для FromPDFWin.xaml
    /// </summary>
    public partial class FromPDFWin : System.Windows.Controls.Page
    {
        public FromPDFWin()
        {
            InitializeComponent();
        }

        private void ConvertPdfToDocx(string pdfFilePath, string docxFilePath)
        {
            try
            {
                Document pdfDocument = new Document(pdfFilePath);

                TextAbsorber textAbsorber = new TextAbsorber();

                pdfDocument.Pages.Accept(textAbsorber);

                string extractedText = textAbsorber.Text;

                Document doc = new Document();
                Aspose.Pdf.Page page = doc.Pages.Add();
                page.Paragraphs.Add(new TextFragment(extractedText));
                doc.Save(docxFilePath);

                MessageBox.Show("Конвертация завершена.", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при конвертации: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ConvertToDoc_Click(object sender, RoutedEventArgs e)
        {
            string pdfFilePath = txtPath.Text;
            string docxFilePath = System.IO.Path.ChangeExtension(pdfFilePath, ".docx");

            ConvertPdfToDocx(pdfFilePath, docxFilePath);
        }

        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "PDF файлы (*.pdf)|*.pdf";
            openFileDialog.CheckFileExists = true;
            openFileDialog.CheckPathExists = true;
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (openFileDialog.ShowDialog() == true)
            {
                txtPath.Text = openFileDialog.FileName;
            }
        }
    }
}
