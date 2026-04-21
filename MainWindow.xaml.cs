using System;
using System.Data;
using System.IO;
using System.Windows;
using Microsoft.Win32;        // OpenFileDialog, SaveFileDialog
using OfficeOpenXml;          // EPPlus
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace FilesDemo
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Источник данных для DataGrid
        private DataTable currentDataTable;

        public MainWindow()
        {
            InitializeComponent();
            // Устанавливаем лицензионный контекст EPPlus (для некоммерческого использования)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Инициализируем пустую таблицу со столбцами по умолчанию
            InitializeEmptyTable();
        }

        // Создаёт пустую таблицу с примерной структурой
        private void InitializeEmptyTable()
        {
            currentDataTable = new DataTable("Products");
            currentDataTable.Columns.Add("ID", typeof(int));
            currentDataTable.Columns.Add("Наименование", typeof(string));
            currentDataTable.Columns.Add("Количество", typeof(int));
            currentDataTable.Columns.Add("Цена", typeof(decimal));

            // Добавим тестовую строку для наглядности
            currentDataTable.Rows.Add(1, "Образец", 10, 99.99m);

            dataGrid.ItemsSource = currentDataTable.DefaultView;
            statusText.Text = "Новая таблица создана";
        }

        // 1. Импорт из Excel
        private void BtnLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDlg = new OpenFileDialog();
            openDlg.Filter = "Excel файлы|*.xlsx|Все файлы|*.*";
            if (openDlg.ShowDialog() == true)
            {
                try
                {
                    using (var package = new ExcelPackage(new FileInfo(openDlg.FileName)))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // первый лист
                        if (worksheet.Dimension == null)
                        {
                            MessageBox.Show("Файл пуст или не содержит данных.", "Ошибка");
                            return;
                        }

                        // Создаём DataTable на основе заголовков (первая строка)
                        var newTable = new DataTable();
                        int colCount = worksheet.Dimension.Columns;
                        for (int col = 1; col <= colCount; col++)
                        {
                            string header = worksheet.Cells[1, col].Text;
                            if (string.IsNullOrWhiteSpace(header))
                                header = $"Column{col}";
                            newTable.Columns.Add(header, typeof(string)); // временно все строки
                        }

                        // Читаем данные, начиная со второй строки
                        int rowCount = worksheet.Dimension.Rows;
                        for (int row = 2; row <= rowCount; row++)
                        {
                            var newRow = newTable.NewRow();
                            for (int col = 1; col <= colCount; col++)
                            {
                                newRow[col - 1] = worksheet.Cells[row, col].Text;
                            }
                            newTable.Rows.Add(newRow);
                        }

                        currentDataTable = newTable;
                        dataGrid.ItemsSource = currentDataTable.DefaultView;
                        statusText.Text = $"Загружено {rowCount - 1} строк из {openDlg.FileName}";
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при загрузке Excel: {ex.Message}", "Ошибка");
                }
            }
        }

        // 2. Добавление новой пустой строки
        private void BtnAddRow_Click(object sender, RoutedEventArgs e)
        {
            if (currentDataTable != null)
            {
                DataRow newRow = currentDataTable.NewRow();
                // Заполним значениями по умолчанию (типы столбцов)
                for (int i = 0; i < currentDataTable.Columns.Count; i++)
                {
                    if (currentDataTable.Columns[i].DataType == typeof(string))
                        newRow[i] = "";
                    else if (currentDataTable.Columns[i].DataType == typeof(int))
                        newRow[i] = 0;
                    else if (currentDataTable.Columns[i].DataType == typeof(decimal))
                        newRow[i] = 0.00m;
                    else
                        newRow[i] = DBNull.Value;
                }
                currentDataTable.Rows.Add(newRow);
                statusText.Text = "Добавлена новая строка";
            }
        }

        // 3. Удаление выбранной строки
        private void BtnDeleteRow_Click(object sender, RoutedEventArgs e)
        {
            if (dataGrid.SelectedItem != null && currentDataTable != null)
            {
                DataRowView selectedRow = (DataRowView)dataGrid.SelectedItem;
                currentDataTable.Rows.Remove(selectedRow.Row);
                statusText.Text = "Строка удалена";
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления.", "Информация");
            }
        }

        // 4. Экспорт в PDF
        private void BtnSavePdf_Click(object sender, RoutedEventArgs e)
        {
            if (currentDataTable == null || currentDataTable.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта.", "Ошибка");
                return;
            }

            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "PDF файл|*.pdf";
            if (saveDlg.ShowDialog() == true)
            {
                try
                {
                    using (var writer = new PdfWriter(saveDlg.FileName))
                    using (var pdf = new PdfDocument(writer))
                    {
                        var document = new Document(pdf);
                        document.Add(new Paragraph("Отчёт по данным")
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetFontSize(18)
                            .SetMarginBottom(20));

                        // Создаём таблицу с количеством столбцов DataTable
                        Table table = new Table(currentDataTable.Columns.Count, false);
                        // Заголовки
                        foreach (DataColumn column in currentDataTable.Columns)
                        {
                            table.AddHeaderCell(new Cell().Add(new Paragraph(column.ColumnName)));
                        }
                        // Данные
                        foreach (DataRow row in currentDataTable.Rows)
                        {
                            foreach (var item in row.ItemArray)
                            {
                                table.AddCell(new Cell().Add(new Paragraph(item?.ToString() ?? "")));
                            }
                        }
                        document.Add(table);
                    }
                    statusText.Text = $"Данные сохранены в PDF: {saveDlg.FileName}";
                    MessageBox.Show("PDF успешно создан!", "Успех");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании PDF: {ex.Message}", "Ошибка");
                }
            }
        }
    }
}
