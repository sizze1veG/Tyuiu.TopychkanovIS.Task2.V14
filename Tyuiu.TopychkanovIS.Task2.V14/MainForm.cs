using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml; 
using iTextSharp.text; 
using iTextSharp.text.pdf; 
using Spire.Doc;
using System.Linq;
using System.Collections.Generic;

namespace Tyuiu.TopychkanovIS.Task2.V14
{
    public partial class MainForm : MaterialForm
    {
        private MaterialTabControl tabControl;
        private MaterialTabSelector tabSelector;
        private OleDbConnection connection;
        private bool isInitialized = false;
        public MainForm()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue500, Primary.Blue500, Accent.LightBlue200, TextShade.WHITE);

            this.WindowState = FormWindowState.Maximized;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=lab2_TopychkanovIS_bd.accdb");
            InitializeControls();
            isInitialized = true;
        }

        private void InitializeControls()
        {
            tabControl = new MaterialTabControl
            {
                Dock = DockStyle.Fill
            };

            tabSelector = new MaterialTabSelector
            {
                BaseTabControl = tabControl,
                Dock = DockStyle.Top
            };

            this.Controls.Add(tabControl);
            this.Controls.Add(tabSelector);

            AddTabPage("Сотрудники");
            AddTabPage("Заказы");
            AddTabPage("Туры");
            AddTabPage("Клиенты");
            AddTabPage("Услуги");
            AddTabPage("Страна");
            AddTabPage("Транспорт");
            AddTabPage("Город");
            AddTabPage("Аккаунты");

            AddQueryTabPage("Запросы", new string[]
            {
                "Запрос по турам",
                "Клиенты Запрос",
                "Поставщики Запрос",
                "Сотрудники Запрос",
                "Страна Запрос",
                "Туры в Мальдивы на 6 дней",
                "Услуги Запрос"
            });

            AddReportTabPage("Отчеты", new string[]
            {
                "Клиенты",
                "Поставщики",
                "Сотрудники",
                "Туры",
                "Туры в Мальдивы на 6 дней",
                "Услуги"
            });

            LoadData();
        }

        private void AddTabPage(string tabName)
        {
            var tabPage = new TabPage(tabName)
            {
                Name = tabName
            };

            var dataGridView = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };

            var panel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50
            };

            var searchBox = new MaterialSingleLineTextField
            {
                Hint = "Search",
                Dock = DockStyle.Left,
                Width = 200
            };
            searchBox.TextChanged += (sender, e) => ApplyFilter(dataGridView, searchBox.Text);

            var exportCsvButton = new MaterialRaisedButton
            {
                Text = "Export CSV",
                Dock = DockStyle.Right,
                Width = 100
            };
            exportCsvButton.Click += (sender, e) => ExportDataToCsv(dataGridView);

            var exportExcelButton = new MaterialRaisedButton
            {
                Text = "Export Excel",
                Dock = DockStyle.Right,
                Width = 100
            };
            exportExcelButton.Click += (sender, e) => ExportDataToExcel(dataGridView);

            var exportPdfButton = new MaterialRaisedButton
            {
                Text = "Export PDF",
                Dock = DockStyle.Right,
                Width = 100
            };
            exportPdfButton.Click += (sender, e) => ExportDataToPdf(dataGridView);

            var exportWordButton = new MaterialRaisedButton
            {
                Text = "Export Word",
                Dock = DockStyle.Right,
                Width = 100
            };
            exportWordButton.Click += (sender, e) => ExportDataToWord(dataGridView);

            var addButton = new MaterialRaisedButton
            {
                Text = "Добавить",
                Dock = DockStyle.Right,
                Width = 100,
                Enabled = false
            };
            addButton.Click += (sender, e) => AddRow(dataGridView, tabName);

            var editButton = new MaterialRaisedButton
            {
                Text = "Редактировать",
                Dock = DockStyle.Right,
                Width = 100,
                Enabled = false
            };
            editButton.Click += (sender, e) => EditRow(dataGridView, tabName);

            var deleteButton = new MaterialRaisedButton
            {
                Text = "Удалить",
                Dock = DockStyle.Right,
                Width = 100,
                Enabled = false
            };
            deleteButton.Click += (sender, e) => DeleteRow(dataGridView, tabName);

            var aboutButton = new MaterialRaisedButton
            {
                Text = "О программе",
                Dock = DockStyle.Right,
                Width = 100
            };
            aboutButton.Click += (sender, e) => InfoProg();

            panel.Controls.Add(addButton);
            panel.Controls.Add(editButton);
            panel.Controls.Add(deleteButton);
            panel.Controls.Add(searchBox);
            panel.Controls.Add(exportCsvButton);
            panel.Controls.Add(exportExcelButton);
            panel.Controls.Add(exportPdfButton);
            panel.Controls.Add(exportWordButton);
            panel.Controls.Add(aboutButton);

            tabPage.Controls.Add(dataGridView);
            tabPage.Controls.Add(panel);

            dataGridView.SelectionChanged += (sender, e) =>
            {
                bool isLastRowSelected = dataGridView.CurrentRow != null &&
                                         dataGridView.CurrentRow.Index == dataGridView.Rows.Count - 1;

                bool isRowEmpty = true;
                if (dataGridView.CurrentRow != null)
                {
                    foreach (DataGridViewCell cell in dataGridView.CurrentRow.Cells)
                    {
                        if (!string.IsNullOrWhiteSpace(cell.Value?.ToString()))
                        {
                            isRowEmpty = false;
                            break;
                        }
                    }
                }

                addButton.Enabled = isLastRowSelected && isRowEmpty;
                editButton.Enabled = dataGridView.CurrentRow != null && !isRowEmpty;
                deleteButton.Enabled = dataGridView.CurrentRow != null && !isRowEmpty;
            };

            tabControl.TabPages.Add(tabPage);
        }

        private void InfoProg()
        {
            AboutForm aboutForm = new AboutForm();
            aboutForm.ShowDialog();
        }

        private void AddRow(DataGridView dataGridView, string tableName)
        {
            try
            {
                Form inputForm = new Form();
                inputForm.Text = $"Добавление записи в {tableName}";
                inputForm.StartPosition = FormStartPosition.CenterScreen;

                List<TextBox> textBoxes = new List<TextBox>();
                int columnCount = dataGridView.Columns.Count - 1; 
                for (int i = 1; i < dataGridView.Columns.Count; i++)
                {
                    TextBox textBox = new TextBox();
                    textBox.Width = 200;
                    textBoxes.Add(textBox);
                }

                Button saveButton = new Button();
                saveButton.Text = "Сохранить";
                saveButton.Click += (sender, e) =>
                {
                    string columnNames = string.Join(", ", dataGridView.Columns.Cast<DataGridViewColumn>().Skip(1).Select(c => $"[{c.Name}]"));
                    string values = string.Join(", ", Enumerable.Repeat("@val", columnCount));
                    string query = $"INSERT INTO [{tableName}] ({columnNames}) VALUES ({values})";

                    try
                    {
                        connection.Open();
                        OleDbCommand command = new OleDbCommand(query, connection);

                        for (int i = 0; i < columnCount; i++)
                        {
                            command.Parameters.AddWithValue($"@val{i + 1}", textBoxes[i].Text);
                        }

                        int rowsAffected = command.ExecuteNonQuery();
                        if (rowsAffected > 0)
                        {
                            MessageBox.Show("Запись успешно добавлена в таблицу.");

                            LoadTableData(dataGridView, tableName);
                        }
                        else
                        {
                            MessageBox.Show("Ошибка при добавлении записи.");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при добавлении записи: {ex.Message}");
                    }
                    finally
                    {
                        connection.Close();
                    }

                    inputForm.Close();
                };

                FlowLayoutPanel panel = new FlowLayoutPanel();
                panel.FlowDirection = FlowDirection.TopDown;
                panel.AutoSize = true;
                panel.WrapContents = false;

                foreach (TextBox textBox in textBoxes)
                {
                    panel.Controls.Add(textBox);
                }
                panel.Controls.Add(saveButton);

                inputForm.Controls.Add(panel);
                inputForm.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении строки: {ex.Message}");
            }
        }

        private void EditRow(DataGridView dataGridView, string tableName)
        {
            if (dataGridView.CurrentRow != null)
            {
                try
                {
                    Form editForm = new Form();
                    editForm.Text = $"Редактирование записи в {tableName}";
                    editForm.StartPosition = FormStartPosition.CenterScreen;

                    List<TextBox> textBoxes = new List<TextBox>();
                    int columnCount = dataGridView.Columns.Count - 1; 
                    for (int i = 1; i < dataGridView.Columns.Count; i++) 
                    {
                        TextBox textBox = new TextBox();
                        textBox.Width = 200;
                        textBox.Text = dataGridView.CurrentRow.Cells[i].Value?.ToString(); 
                        textBoxes.Add(textBox);
                    }

                    Button saveButton = new Button();
                    saveButton.Text = "Сохранить";
                    saveButton.Click += (sender, e) =>
                    {
                        string columnNames = string.Join(", ", dataGridView.Columns.Cast<DataGridViewColumn>().Skip(1).Select(c => $"[{c.Name}] = @val{c.Index}"));
                        string query = $"UPDATE [{tableName}] SET {columnNames} WHERE [{dataGridView.Columns[0].Name}] = @PrimaryKey";

                        try
                        {
                            connection.Open();
                            OleDbCommand command = new OleDbCommand(query, connection);

                            for (int i = 0; i < columnCount; i++)
                            {
                                command.Parameters.AddWithValue($"@val{i + 1}", textBoxes[i].Text);
                            }
                            command.Parameters.AddWithValue("@PrimaryKey", dataGridView.CurrentRow.Cells[0].Value);

                            int rowsAffected = command.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно обновлена в таблице.");
                                LoadTableData(dataGridView, tableName); 
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при обновлении записи.");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ошибка при обновлении записи: {ex.Message}");
                        }
                        finally
                        {
                            connection.Close();
                        }

                        editForm.Close();
                    };

                    FlowLayoutPanel panel = new FlowLayoutPanel();
                    panel.FlowDirection = FlowDirection.TopDown;
                    panel.AutoSize = true;
                    panel.WrapContents = false;

                    foreach (TextBox textBox in textBoxes)
                    {
                        panel.Controls.Add(textBox);
                    }
                    panel.Controls.Add(saveButton);

                    editForm.Controls.Add(panel);
                    editForm.ShowDialog();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при редактировании строки: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Выберите строку для редактирования.");
            }
        }
        private void DeleteRow(DataGridView dataGridView, string tableName)
        {
            if (dataGridView.CurrentRow != null)
            {
                try
                {
                    object primaryKeyValue = dataGridView.CurrentRow.Cells[0].Value;

                    DialogResult result = MessageBox.Show($"Вы действительно хотите удалить запись с {dataGridView.Columns[0].Name} = {primaryKeyValue}?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        try
                        {
                            connection.Open();

                            string query = $"DELETE FROM [{tableName}] WHERE [{dataGridView.Columns[0].Name}] = @PrimaryKey";
                            OleDbCommand command = new OleDbCommand(query, connection);
                            command.Parameters.AddWithValue("@PrimaryKey", primaryKeyValue);

                            int rowsAffected = command.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно удалена из таблицы.");
                                LoadTableData(dataGridView, tableName);
                            }
                            else
                            {
                                MessageBox.Show("Ошибка при удалении записи.");
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Ошибка при удалении записи: {ex.Message}");
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при удалении строки: {ex.Message}");
                }
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления.");
            }
        }

        private void AddQueryTabPage(string tabName, string[] queries)
        {
            var tabPage = new TabPage(tabName)
            {
                Name = tabName
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DataSource = queries
            };
            listBox.SelectedIndexChanged += (sender, e) => {
                if (isInitialized)
                {
                    LoadQueryData(listBox.SelectedItem.ToString());
                }
            };

            tabPage.Controls.Add(listBox);
            tabControl.TabPages.Add(tabPage);
        }

        private void AddReportTabPage(string tabName, string[] reports)
        {
            var tabPage = new TabPage(tabName)
            {
                Name = tabName
            };

            var listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DataSource = reports
            };
            listBox.SelectedIndexChanged += (sender, e) => {
                if (isInitialized)
                {
                    ExportReportToPdf(listBox.SelectedItem.ToString());
                }
            };

            tabPage.Controls.Add(listBox);
            tabControl.TabPages.Add(tabPage);
        }

        private void LoadData()
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Name == "Запросы" || tabPage.Name == "Отчеты")
                    continue;

                var dataGridView = tabPage.Controls[0] as DataGridView;
                var tableName = tabPage.Name;
                LoadTableData(dataGridView, tableName);
            }
        }

        private void LoadTableData(DataGridView dataGridView, string tableName)
        {
            try
            {
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                string query = $"SELECT * FROM [{tableName}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                dataGridView.DataSource = dataTable;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }

        private void LoadQueryData(string queryName)
        {
            try
            {
                connection.Open();
                string query = $"SELECT * FROM [{queryName}]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                var queryResultsForm = new Form
                {
                    Text = queryName,
                    Width = 800,
                    Height = 600
                };

                var dataGridView = new DataGridView
                {
                    Dock = DockStyle.Fill,
                    DataSource = dataTable,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                };

                queryResultsForm.Controls.Add(dataGridView);
                queryResultsForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void ExportReportToPdf(string reportName)
        {
            try
            {
                string databasePath = @"lab2_TopychkanovIS_bd.accdb";
                string query = $"SELECT * FROM [{reportName}]";

                using (OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + databasePath))
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                {
                    DataTable dataTable = new DataTable();
                    adapter.Fill(dataTable);

                    string pdfOutputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), $"{reportName}.pdf");

                    using (iTextSharp.text.Document document = new iTextSharp.text.Document())
                    {
                        PdfWriter.GetInstance(document, new FileStream(pdfOutputPath, FileMode.Create));
                        document.Open();

                        BaseFont baseFont = BaseFont.CreateFont("c:/windows/fonts/arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                        iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 12);

                        PdfPTable pdfTable = new PdfPTable(dataTable.Columns.Count);

                        for (int i = 0; i < dataTable.Columns.Count; i++)
                        {
                            PdfPCell cell = new PdfPCell(new Phrase(dataTable.Columns[i].ColumnName, font));
                            pdfTable.AddCell(cell);
                        }

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataTable.Columns.Count; j++)
                            {
                                PdfPCell cell = new PdfPCell(new Phrase(dataTable.Rows[i][j].ToString(), font));
                                pdfTable.AddCell(cell);
                            }
                        }

                        document.Add(pdfTable);
                    }

                    MessageBox.Show("Report exported to PDF successfully.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting report to PDF: {ex.Message}");
            }
        }

        private void ApplyFilter(DataGridView dataGridView, string filterText)
        {
            if (dataGridView.DataSource is DataTable dataTable)
            {
                string filterExpression = string.Join(" OR ", dataTable.Columns.Cast<DataColumn>()
                .Select(c => $"CONVERT([{c.ColumnName}], System.String) LIKE '%{filterText}%'"));
                dataTable.DefaultView.RowFilter = filterExpression;
            }
        }

        private void ExportDataToCsv(DataGridView dataGridView)
        {
            if (dataGridView.DataSource is DataTable dataTable)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName, false, System.Text.Encoding.UTF8))
                        {
                            foreach (DataColumn column in dataTable.Columns)
                            {
                                writer.Write(column.ColumnName + ",");
                            }
                            writer.WriteLine();

                            foreach (DataRow row in dataTable.Rows)
                            {
                                for (int i = 0; i < dataTable.Columns.Count; i++)
                                {
                                    writer.Write(row[i].ToString());
                                    if (i < dataTable.Columns.Count - 1) writer.Write(",");
                                }
                                writer.WriteLine();
                            }
                        }
                        MessageBox.Show("Export to CSV successful.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data to CSV: {ex.Message}");
                    }
                }
            }
        }

        private void ExportDataToExcel(DataGridView dataGridView)
        {
            if (dataGridView.DataSource is DataTable dataTable)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var package = new ExcelPackage())
                        {
                            var worksheet = package.Workbook.Worksheets.Add("Data");
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                            }
                            for (int i = 0; i < dataTable.Rows.Count; i++)
                            {
                                for (int j = 0; j < dataTable.Columns.Count; j++)
                                {
                                    worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                                }
                            }
                            package.SaveAs(new FileInfo(saveFileDialog.FileName));
                        }
                        MessageBox.Show("Export to Excel successful.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data to Excel: {ex.Message}");
                    }
                }
            }
        }

        private void ExportDataToPdf(DataGridView dataGridView)
        {
            if (dataGridView.DataSource is DataTable dataTable)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var document = new iTextSharp.text.Document())
                        {
                            PdfWriter.GetInstance(document, new FileStream(saveFileDialog.FileName, FileMode.Create));
                            document.Open();

                            var baseFont = BaseFont.CreateFont("c:/windows/fonts/arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            var font = new iTextSharp.text.Font(baseFont, 12, iTextSharp.text.Font.NORMAL);

                            PdfPTable pdfTable = new PdfPTable(dataTable.Columns.Count);

                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                pdfTable.AddCell(new PdfPCell(new Phrase(dataTable.Columns[i].ColumnName, font)));
                            }

                            foreach (DataRow row in dataTable.Rows)
                            {
                                for (int j = 0; j < dataTable.Columns.Count; j++)
                                {
                                    pdfTable.AddCell(new PdfPCell(new Phrase(row[j].ToString(), font)));
                                }
                            }
                            document.Add(pdfTable);
                            document.Close();
                        }
                        MessageBox.Show("Export to PDF successful.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data to PDF: {ex.Message}");
                    }
                }
            }
        }

        private void ExportDataToWord(DataGridView dataGridView)
        {
            if (dataGridView.DataSource is DataTable dataTable)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word files (*.docx)|*.docx|All files (*.*)|*.*"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (var document = new Spire.Doc.Document())
                        {
                            var section = document.AddSection();
                            var table = section.AddTable(true);

                            table.ResetCells(dataTable.Rows.Count + 1, dataTable.Columns.Count);

                            TableRow row = table.Rows[0];
                            row.IsHeader = true;
                            for (int i = 0; i < dataTable.Columns.Count; i++)
                            {
                                row.Cells[i].AddParagraph().AppendText(dataTable.Columns[i].ColumnName);
                            }

                            for (int i = 0; i < dataTable.Rows.Count; i++)
                            {
                                TableRow dataRow = table.Rows[i + 1];
                                for (int j = 0; j < dataTable.Columns.Count; j++)
                                {
                                    dataRow.Cells[j].AddParagraph().AppendText(dataTable.Rows[i][j].ToString());
                                }
                            }

                            document.SaveToFile(saveFileDialog.FileName, FileFormat.Docx);
                        }
                        MessageBox.Show("Export to Word successful.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error exporting data to Word: {ex.Message}");
                    }
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (connection != null && connection.State != ConnectionState.Closed)
            {
                connection.Close();
                connection.Dispose();
            }

            Application.Exit();
        }
    }
}

