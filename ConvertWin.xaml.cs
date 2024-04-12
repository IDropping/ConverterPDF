using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System.Reflection;
using Aspose.Slides;
using Aspose.Cells;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace Converter
{
    /// <summary>
    /// Логика взаимодействия для ConvertWin.xaml
    /// </summary>
    public partial class ConvertWin : System.Windows.Controls.Page
    {
        public ConvertWin()
        {
            InitializeComponent();
        }

        private void ConvertPNG_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Документы (*.pnf, *.jpg, *.bmp, *.tiff)|*.png;*.jpg;*.bmp;*.tiff|Все файлы (*.*)|*.*";
            dialog.FileName = "";

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                txtPath.Text = dialog.FileName;
            }
        }

        private void ConvertPresentation_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Документы (*.pptx, *.ppt, *.odp)|*.pptx;*.ppt;*.odp|Все файлы (*.*)|*.*";
            dialog.FileName = "";

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                txtPath.Text = dialog.FileName;
            }
        }

        private void ConvertDOCX_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Документы (*.docx, *.doc, *.odt)|*.docx;*.doc;*.odt|Все файлы (*.*)|*.*";
            dialog.FileName = "";

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                txtPath.Text = dialog.FileName;
            }
        }

        private void ConvertXLSX_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Документы (*.xlsx, *.xls, *.ods)|*.xlsx;*.xls;*.ods|Все файлы (*.*)|*.*";
            dialog.FileName = "";

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                txtPath.Text = dialog.FileName;
            }
        }

        private void ConvertText_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object oMissing = System.Reflection.Missing.Value;

            try
            {
                FileInfo inputFile = new FileInfo(txtPath.Text);

                if (!inputFile.Exists)
                {
                    MessageBox.Show("Файл не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string extension = inputFile.Extension.ToLower();
                if (extension != ".docx" && extension != ".doc" && extension != ".odt")
                {
                    MessageBox.Show("Невозможно конвертировать данный формат файла в PDF.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                word.Visible = false;
                word.ScreenUpdating = false;

                object inputFilename = (object)inputFile.FullName;

                Document doc = word.Documents.Open(ref inputFilename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFilename = System.IO.Path.ChangeExtension(inputFile.FullName, ".pdf");
                object fileformat = WdSaveFormat.wdFormatPDF;

                doc.SaveAs2(ref outputFilename,
                    ref fileformat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                object savechanges = WdSaveOptions.wdSaveChanges;
                ((_Document)doc).Close(ref savechanges, ref oMissing, ref oMissing);
                doc = null;

                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                word = null;

                MessageBox.Show("Файл был успешно конвертирован в PDF");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при конвертации файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (word != null)
                {
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                    word = null;
                }
            }
        }

        private void ConvertPresentationToPdf(string presentationFilePath)
        {
            try
            {
                Presentation pres = new Presentation(presentationFilePath);

                string outputFilename = System.IO.Path.ChangeExtension(presentationFilePath, ".pdf");
                pres.Save(outputFilename, Aspose.Slides.Export.SaveFormat.Pdf);

                MessageBox.Show("Файл был успешно конвертирован в PDF");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при конвертации файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertPresent_Click(object sender, RoutedEventArgs e)
        {
            string presentationFilePath = txtPath.Text;
            ConvertPresentationToPdf(presentationFilePath);
        }

        private void ConvertTableToPDF(string excelFilePath)
        {
            try
            {
                Workbook workbook = new Workbook(excelFilePath);

                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

                workbook.Save(System.IO.Path.ChangeExtension(excelFilePath, ".pdf"), pdfSaveOptions);

                MessageBox.Show("Файл был успешно конвертирован в PDF");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при конвертации файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertTable_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = txtPath.Text;

            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                MessageBox.Show("Пожалуйста, введите путь к файлу Excel.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ConvertTableToPDF(excelFilePath);
        }

        private void ConvertImageToPDF(string imagePath)
        {
            try
            {
                PdfDocument document = new PdfDocument();

                PdfPage page = document.AddPage();

                XGraphics gfx = XGraphics.FromPdfPage(page);

                XImage image = XImage.FromFile(imagePath);

                double width = image.PixelWidth * 72 / image.HorizontalResolution;
                double height = image.PixelHeight * 72 / image.VerticalResolution;

                gfx.DrawImage(image, 0, 0, width, height);

                document.Save(System.IO.Path.ChangeExtension(imagePath, ".pdf"));

                MessageBox.Show("Изображение успешно конвертировано в PDF");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при конвертации изображения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertImage_Click(object sender, RoutedEventArgs e)
        {
            string imagePath = txtPath.Text;

            if (string.IsNullOrWhiteSpace(imagePath))
            {
                MessageBox.Show("Пожалуйста, введите путь к изображению.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            ConvertImageToPDF(imagePath);
        }
    }
}
