using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.IO;
using SkiaSharp;
using Aspose.Words;
using DocumentFormat.OpenXml.Office2010.CustomUI;
using Aspose.Slides;

namespace Converter
{
    /// <summary>
    /// Логика взаимодействия для PreviewWin.xaml
    /// </summary>
    public partial class PreviewWin : Page
    {
        private Document doc;
        private int currentPageIndex = 0;
        public PreviewWin()
        {
            InitializeComponent();
        }

        private void PreviewDocument(string filePath)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл не найден.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                doc = new Document(filePath);
                RenderPage(currentPageIndex);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при открытии документа: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenDocumentButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Документы (*.docx, *.doc, *.odt)|*.docx;*.doc;*.odt|Все файлы (*.*)|*.*";

            if (openFileDialog.ShowDialog() == true)
            {
                PreviewDocument(openFileDialog.FileName);
            }
        }
        private void RenderPage(int pageIndex)
        {
            if (doc != null && pageIndex >= 0 && pageIndex < doc.PageCount)
            {
                using (SKBitmap bitmap = new SKBitmap(800, 600))
                {
                    using (SKCanvas canvas = new SKCanvas(bitmap))
                    {
                        doc.RenderToSize(pageIndex, canvas, 0f, 0f, 1200f, 650f);

                        BitmapImage bitmapImage = ConvertToBitmapImage(bitmap);
                        pictureBox.Source = bitmapImage;
                    }
                }
            }
        }
        private BitmapImage ConvertToBitmapImage(SKBitmap bitmap)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                bitmap.Encode(stream, SKEncodedImageFormat.Png, 200);
                stream.Position = 0;

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = stream;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        private void NextPage_Click(object sender, RoutedEventArgs e)
        {
            currentPageIndex++;
            if (currentPageIndex >= doc.PageCount)
                currentPageIndex = doc.PageCount - 1;

            RenderPage(currentPageIndex);
        }

        private void PreviousPage_Click(object sender, RoutedEventArgs e)
        {
            currentPageIndex--;
            if (currentPageIndex < 0)
                currentPageIndex = 0;

            RenderPage(currentPageIndex);
        }
    }
}
