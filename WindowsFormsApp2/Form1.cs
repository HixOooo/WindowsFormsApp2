using System;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1;
using Newtonsoft.Json;
using System.IO;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element; // Для Text в iTextSharp
using iText.Kernel.Font;
using iText.Kernel.Pdf.Canvas.Parser;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing; // Для Text в Open XML SDK
using System.Collections.Generic;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private BookManager _bookManager;

        public Form1()
        {
            InitializeComponent();
            _bookManager = new BookManager();
            openFileDialog = new System.Windows.Forms.OpenFileDialog();
        }

        // Добавление книги
        private void btnAdd_Click(object sender, EventArgs e)
        {
            string title = txtTitle.Text;
            string author = txtAuthor.Text;
            int year;

            if (string.IsNullOrEmpty(title) || string.IsNullOrEmpty(author) || !int.TryParse(txtYear.Text, out year))
            {
                MessageBox.Show("Пожалуйста, заполните все поля корректно.");
                return;
            }

            _bookManager.AddBook(title, author, year);
            UpdateBookList();
            ClearInputFields();
        }

        // Удаление книги по ID
        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (lstBooks.SelectedItem == null)
            {
                MessageBox.Show("Выберите книгу для удаления.");
                return;
            }

            var selectedBook = (Book)lstBooks.SelectedItem;
            _bookManager.RemoveBook(selectedBook.Id);
            UpdateBookList();
        }

        // Поиск книги по названию
        private void btnSearchTitle_Click(object sender, EventArgs e)
        {
            string title = txtTitle.Text;
            if (!string.IsNullOrEmpty(title))
            {
                var books = _bookManager.FindBookByName(title);
                lstBooks.DataSource = books;
            }
        }

        // Поиск книги по автору
        private void btnSearchAuthor_Click(object sender, EventArgs e)
        {
            string author = txtAuthor.Text;
            if (!string.IsNullOrEmpty(author))
            {
                var books = _bookManager.FindBookByAuthor(author);
                lstBooks.DataSource = books;
            }
        }

        // Обновление списка книг в ListBox
        private void UpdateBookList()
        {
            lstBooks.DataSource = null; // Очистка DataSource
            lstBooks.DataSource = _bookManager.GetAllBooks();
        }
        // Метод для обработки нажатия на кнопку "Показать все книги"
        private void btnShowAllBooks_Click(object sender, EventArgs e)
        {
            // Получаем все книги из менеджера
            var allBooks = _bookManager.GetAllBooks();

            // Обновляем список на форме, отображая все книги
            lstBooks.DataSource = null; // Сначала очищаем текущий список
            lstBooks.DataSource = allBooks; // Затем заполняем его заново
        }

        // Очистка текстовых полей
        private void ClearInputFields()
        {
            txtTitle.Clear();
            txtAuthor.Clear();
            txtYear.Clear();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _bookManager.ExportBooksToJson(saveFileDialog.FileName);
                    MessageBox.Show("Книги успешно экспортированы!");
                }
            }
        }
        private void btnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _bookManager.ImportBooksFromJson(openFileDialog.FileName);
                    UpdateBookList();
                    MessageBox.Show("Книги успешно импортированы!");
                }
            }
        }
        // Обработчик кнопки для экспорта в DOCX
        private void btnExportToDocx_Click(object sender, EventArgs e)
        {
            var books = _bookManager.GetAllBooks();  // Получаем все книги
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Word Document|*.docx",
                FileName = "BooksList.docx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                _bookManager.ExportBooksToDocx(filePath, books);  // Экспортируем в DOCX
                MessageBox.Show("Книги экспортированы в Word!");
            }
        }
        // Обработчик кнопки для экспорта в PDF
        private void btnExportToPdf_Click(object sender, EventArgs e)
        {
            var books = _bookManager.GetAllBooks();  // Получаем все книги
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "PDF File|*.pdf",
                FileName = "BooksList.pdf"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;
                _bookManager.ExportBooksToPdf(filePath, books);  // Экспортируем в PDF
                MessageBox.Show("Книги экспортированы в PDF!");
            }
        }
        private void btnImportFromPdf_Click(object sender, EventArgs e)
        {
            openFileDialog.Filter = "PDF Files|*.pdf";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Не удалось найти файл.");
                    return;
                }

                List<Book> books = _bookManager.ImportBooksFromPdf(filePath);
                lstBooks.DataSource = books;
                MessageBox.Show($"Импортировано {books.Count} книг из PDF.");
            }
        }

        // Импорт из DOCX
        private void btnImportFromDocx_Click(object sender, EventArgs e)
        {
            // Настройка фильтра для диалогового окна
            openFileDialog.Filter = "Word Files|*.docx";

            // Открытие диалогового окна для выбора файла
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Проверка на существование файла
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Не удалось найти файл.");
                    return;
                }

                // Вызов метода импорта для DOCX
                List<Book> books = _bookManager.ImportBooksFromDocx(filePath);

                // Установка источника данных для DataGridView (предполагается, что lstBooks это DataGridView)
                lstBooks.DataSource = books;

                // Показ сообщения о количестве импортированных книг
                MessageBox.Show($"Импортировано {books.Count} книг из DOCX.");
            }
        }
        private void ConvertDocument(string inputFilePath)
        {
            try
            {
                // Определение выходного файла на основе расширения входного
                string outputFilePath = Path.ChangeExtension(inputFilePath,
                    Path.GetExtension(inputFilePath).ToLower() == ".pdf" ? ".docx" : ".pdf");

                _bookManager.ConvertFile(inputFilePath, outputFilePath);

                MessageBox.Show($"Файл успешно конвертирован: {outputFilePath}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка конвертации: {ex.Message}");
            }
        }
        private void btnConvert_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF and DOCX Files|*.pdf;*.docx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ConvertDocument(openFileDialog.FileName);
                }
            }
        }
    }
}

