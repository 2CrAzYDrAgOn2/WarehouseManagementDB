using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout.Properties;
using Microsoft.Office.Interop.Word;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace test_DataBase
{
    internal enum RowState
    {
        Existed,
        New,
        Modified,
        ModifiedNew,
        Deleted
    }

    public partial class Form1 : Form
    {
        private readonly DataBase dataBase = new DataBase();
        private bool admin;
        private int selectedRow;

        public Form1()
        {
            try
            {
                InitializeComponent();
                StartPosition = FormStartPosition.CenterScreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// SetAdminStatus проверяет доступ
        /// </summary>
        /// <param name="isAdmin"></param>
        public void SetAdminStatus(bool isAdmin)
        {
            admin = isAdmin;
        }

        /// <summary>
        /// CreateColumns вызывается при создании колонок
        /// </summary>
        private void CreateColumns()
        {
            try
            {
                dataGridViewEquipment.Columns.Add("EquipmentID", "Номер");
                dataGridViewEquipment.Columns.Add("Name", "Название");
                dataGridViewEquipment.Columns.Add("Category", "Категория");
                dataGridViewEquipment.Columns.Add("PurchaseData", "Дата покупки");
                dataGridViewEquipment.Columns.Add("Price", "Цена");
                dataGridViewEquipment.Columns.Add("Quantity", "Количество");
                dataGridViewEquipment.Columns.Add("Location", "Расположение");
                dataGridViewEquipment.Columns.Add("IsNew", String.Empty);
                dataGridViewEquipmentMovement.Columns.Add("MovementID", "Номер");
                dataGridViewEquipmentMovement.Columns.Add("EquipmentID", "Номер оборудования");
                dataGridViewEquipmentMovement.Columns.Add("MovementDate", "Дата передвижения");
                dataGridViewEquipmentMovement.Columns.Add("MovementType", "Тип передвижения");
                dataGridViewEquipmentMovement.Columns.Add("Quantity", "Количество");
                dataGridViewEquipmentMovement.Columns.Add("IsNew", String.Empty);
                dataGridViewSupplier.Columns.Add("SupplierID", "Номер");
                dataGridViewSupplier.Columns.Add("Name", "Имя");
                dataGridViewSupplier.Columns.Add("ContactPerson", "Контактное лицо");
                dataGridViewSupplier.Columns.Add("Phone", "Телефон");
                dataGridViewSupplier.Columns.Add("Email", "Email");
                dataGridViewSupplier.Columns.Add("IsNew", String.Empty);
                dataGridViewEquipmentSupplier.Columns.Add("EquipmentID", "Номер оборудования");
                dataGridViewEquipmentSupplier.Columns.Add("SupplierID", "Номер поставщика");
                dataGridViewEquipmentSupplier.Columns.Add("IsNew", String.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// CreateColumns вызывается при очистке полей
        /// </summary>
        private void ClearFields()
        {
            try
            {
                textBoxEquipmentID.Text = "";
                textBoxName.Text = "";
                textBoxCategory.Text = "";
                textBoxPurchaseData.Text = "";
                textBoxPrice.Text = "";
                textBoxQuantinity.Text = "";
                textBoxLocation.Text = "";
                textBoxMovementID.Text = "";
                textBoxEquipmentIDEquipmentMovement.Text = "";
                textBoxMovementDate.Text = "";
                textBoxMovementType.Text = "";
                textBoxQuantinityEquipmentMovement.Text = "";
                textBoxSupplierID.Text = "";
                textBoxNameSupplier.Text = "";
                textBoxContactPerson.Text = "";
                textBoxPhone.Text = "";
                textBoxEmail.Text = "";
                textBoxEquipmentIDEquipmentSupplier.Text = "";
                textBoxSupplierIDEquipmentSupplier.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ReadSingleRow вызывается при чтении строк
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="iDataRecord"></param>
        private void ReadSingleRow(DataGridView dataGridView, IDataRecord iDataRecord)
        {
            try
            {
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetDateTime(3).ToString("yyyy-MM-dd"), iDataRecord.GetInt32(4), iDataRecord.GetInt32(5), iDataRecord.GetString(6), RowState.Modified);
                        break;

                    case "dataGridViewEquipmentMovement":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), iDataRecord.GetDateTime(2).ToString("yyyy-MM-dd"), iDataRecord.GetString(3), iDataRecord.GetInt32(4), RowState.Modified);
                        break;

                    case "dataGridViewSupplier":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewEquipmentSupplier":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), RowState.Modified);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// RefreshDataGrid вызывается при обновлении dataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="tableName"></param>
        private void RefreshDataGrid(DataGridView dataGridView, string tableName)
        {
            try
            {
                dataGridView.Rows.Clear();
                string queryString = $"select * from {tableName}";
                SqlCommand sqlCommand = new SqlCommand(queryString, dataBase.GetConnection());
                dataBase.OpenConnection();
                SqlDataReader sqlDataReader = sqlCommand.ExecuteReader();
                while (sqlDataReader.Read())
                {
                    ReadSingleRow(dataGridView, sqlDataReader);
                }
                sqlDataReader.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Form1_Load вызывается при загрузке формы "Form1"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                CreateColumns();
                RefreshDataGrid(dataGridViewEquipment, "Equipment");
                RefreshDataGrid(dataGridViewEquipmentMovement, "EquipmentMovement");
                RefreshDataGrid(dataGridViewSupplier, "Supplier");
                RefreshDataGrid(dataGridViewEquipmentSupplier, "EquipmentSupplier");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// DataGridView_CellClick вызывается при нажатии на ячейку в DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        /// <param name="selectedRow"></param>
        private void DataGridView_CellClick(DataGridView dataGridView, int selectedRow)
        {
            try
            {
                DataGridViewRow dataGridViewRow = dataGridView.Rows[selectedRow];
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        textBoxEquipmentID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxName.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxCategory.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxPurchaseData.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxPrice.Text = dataGridViewRow.Cells[4].Value.ToString();
                        textBoxQuantinity.Text = dataGridViewRow.Cells[5].Value.ToString();
                        textBoxLocation.Text = dataGridViewRow.Cells[6].Value.ToString();
                        break;

                    case "dataGridViewEquipmentMovement":
                        textBoxMovementID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxEquipmentIDEquipmentMovement.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxMovementDate.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxMovementType.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxQuantinityEquipmentMovement.Text = dataGridViewRow.Cells[4].Value.ToString();
                        break;

                    case "dataGridViewSupplier":
                        textBoxSupplierID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxNameSupplier.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxContactPerson.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxPhone.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxEmail.Text = dataGridViewRow.Cells[4].Value.ToString();
                        break;

                    case "dataGridViewEquipmentSupplier":
                        textBoxEquipmentIDEquipmentSupplier.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxSupplierIDEquipmentSupplier.Text = dataGridViewRow.Cells[1].Value.ToString();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Search вызывается при поиске данных в DataGridView
        /// </summary>
        /// <param name="dataGridView"></param>
        private void Search(DataGridView dataGridView)
        {
            try
            {
                dataGridView.Rows.Clear();
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        string searchStringEquipment = $"select * from Equipment where concat (EquipmentID, Name, Category, PurchaseDate, Price, Quantity, Location) like '%" + textBoxSearchEquipment.Text + "%'";
                        SqlCommand sqlCommandEquipment = new SqlCommand(searchStringEquipment, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderEquipment = sqlCommandEquipment.ExecuteReader();
                        while (sqlDataReaderEquipment.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderEquipment);
                        }
                        sqlDataReaderEquipment.Close();
                        break;

                    case "dataGridViewEquipmentMovement":
                        string searchStringEquipmentMovement = $"select * from EquipmentMovement where concat (MovementID, EquipmentID, MovementDate, MovementType, Quantity) like '%" + textBoxSearchEquipmentMovement.Text + "%'";
                        SqlCommand sqlCommandEquipmentMovement = new SqlCommand(searchStringEquipmentMovement, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderEquipmentMovement = sqlCommandEquipmentMovement.ExecuteReader();
                        while (sqlDataReaderEquipmentMovement.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderEquipmentMovement);
                        }
                        sqlDataReaderEquipmentMovement.Close();
                        break;

                    case "dataGridViewSupplier":
                        string searchStringSupplier = $"select * from Supplier where concat (SupplierID, Name, ContactPerson, Phone, Email) like '%" + textBoxSearchSupplier.Text + "%'";
                        SqlCommand sqlCommandSupplier = new SqlCommand(searchStringSupplier, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderSupplier = sqlCommandSupplier.ExecuteReader();
                        while (sqlDataReaderSupplier.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderSupplier);
                        }
                        sqlDataReaderSupplier.Close();
                        break;

                    case "dataGridViewEquipmentSupplier":
                        string searchStringEquipmentSupplier = $"select * from EquipmentSupplier where concat (EquipmentID, SupplierID) like '%" + textBoxSearchEquipmentSupplier.Text + "%'";
                        SqlCommand sqlCommandEquipmentSupplier = new SqlCommand(searchStringEquipmentSupplier, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderEquipmentSupplier = sqlCommandEquipmentSupplier.ExecuteReader();
                        while (sqlDataReaderEquipmentSupplier.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderEquipmentSupplier);
                        }
                        sqlDataReaderEquipmentSupplier.Close();
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// DeleteRow вызывается при удалении строки
        /// </summary>
        /// <param name="dataGridView"></param>
        private void DeleteRow(DataGridView dataGridView)
        {
            try
            {
                int index = dataGridView.CurrentCell.RowIndex;
                dataGridView.Rows[index].Visible = false;
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[7].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[7].Value = RowState.Deleted;
                        break;

                    case "dataGridViewEquipmentMovement":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                        break;

                    case "dataGridViewSupplier":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                        break;

                    case "dataGridViewEquipmentSupplier":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[2].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[2].Value = RowState.Deleted;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// UpdateBase вызывается при обновлении базы данных
        /// </summary>
        /// <param name="dataGridView"></param>
        private void UpdateBase(DataGridView dataGridView)
        {
            try
            {
                dataBase.OpenConnection();
                for (int index = 0; index < dataGridView.Rows.Count; index++)
                {
                    switch (dataGridView.Name)
                    {
                        case "dataGridViewEquipment":
                            var rowStateEquipment = (RowState)dataGridView.Rows[index].Cells[7].Value;
                            if (rowStateEquipment == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEquipment == RowState.Deleted)
                            {
                                var equipmentID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Equipment where EquipmentID = '{equipmentID}'";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipment == RowState.Modified)
                            {
                                var equipmentID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var name = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var category = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var purchaseData = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var price = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var quantity = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var location = dataGridView.Rows[index].Cells[6].Value.ToString();
                                var changeQuery = $"update Equipment set Name = '{name}', Category = '{category}', PurchaseDate = '{purchaseData}', Price = '{price}', Quantity = '{quantity}', Location = '{location}' where EquipmentID = '{equipmentID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewEquipmentMovement":
                            var rowStateEquipmentMovement = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateEquipmentMovement == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEquipmentMovement == RowState.Deleted)
                            {
                                var movementID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from EquipmentMovement where MovementID = '{movementID}'";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipmentMovement == RowState.Modified)
                            {
                                var movementID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var equipmentID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var movementDate = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var movementType = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var quantity = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update EquipmentMovement set EquipmentID = '{equipmentID}', MovementDate = '{movementDate}', MovementType = '{movementType}', Quantity = '{quantity}' where MovementID = '{movementID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewSupplier":
                            var rowStateSupplier = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateSupplier == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateSupplier == RowState.Deleted)
                            {
                                var supplierID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Supplier where SupplierID = '{supplierID}'";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateSupplier == RowState.Modified)
                            {
                                var supplierID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var name = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var contactPerson = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var phone = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var email = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update Supplier set Name = '{name}', ContactPerson = '{contactPerson}', Phone = '{phone}', Email = '{email}' where SupplierID = '{supplierID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewEquipmentSupplier":
                            var rowStateEquipmentSupplier = (RowState)dataGridView.Rows[index].Cells[2].Value;
                            if (rowStateEquipmentSupplier == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEquipmentSupplier == RowState.Deleted)
                            {
                                var equipmentID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var supplierID = Convert.ToInt32(dataGridView.Rows[index].Cells[1].Value);
                                var deleteQuery = $"delete from EquipmentSupplier where EquipmentID = '{equipmentID}' and SupplierID = '{supplierID}'";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipmentSupplier == RowState.Modified)
                            {
                                var equipmentID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var supplierID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var changeQuery = $"update EquipmentSupplier set EquipmentID = '{equipmentID}', SupplierID = '{supplierID}' where EquipmentID = '{equipmentID}' and SupplierID = '{supplierID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                dataBase.CloseConnection();
            }
        }

        /// <summary>
        /// Change вызывается при изменении данных в базе данных
        /// </summary>
        /// <param name="dataGridView"></param>
        private void Change(DataGridView dataGridView)
        {
            try
            {
                var selectedRowIndex = dataGridView.CurrentCell.RowIndex;
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        var equipmentID = textBoxEquipmentID.Text;
                        var name = textBoxName.Text;
                        var category = textBoxCategory.Text;
                        var purchaseData = textBoxPurchaseData.Value;
                        var price = textBoxPrice.Text;
                        var quantity = textBoxQuantinity.Text;
                        var location = textBoxLocation.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(equipmentID, name, category, purchaseData, price, quantity, location);
                        dataGridView.Rows[selectedRowIndex].Cells[7].Value = RowState.Modified;
                        break;

                    case "dataGridViewEquipmentMovement":
                        var movementID = textBoxMovementID.Text;
                        var equipmentIDEquipmentMovement = textBoxEquipmentIDEquipmentMovement.Text;
                        var movementDate = textBoxMovementDate.Value;
                        var movementType = textBoxMovementType.Text;
                        var quantityEquipmentMovement = textBoxQuantinityEquipmentMovement.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(movementID, equipmentIDEquipmentMovement, movementDate, movementType, quantityEquipmentMovement);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewSupplier":
                        var supplierID = textBoxSupplierID.Text;
                        var nameSupplier = textBoxNameSupplier.Text;
                        var contactPerson = textBoxContactPerson.Text;
                        var phone = textBoxPhone.Text;
                        var email = textBoxEmail.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(supplierID, nameSupplier, contactPerson, phone, email);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewEquipmentSupplier":
                        var equipmentIDEquipmentSupplier = textBoxEquipmentIDEquipmentSupplier.Text;
                        var supplierIDEquipmentSupplier = textBoxSupplierIDEquipmentSupplier.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(equipmentIDEquipmentSupplier, supplierIDEquipmentSupplier);
                        dataGridView.Rows[selectedRowIndex].Cells[2].Value = RowState.Modified;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToWord вызывается при экспорте данных в Word
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToWord(DataGridView dataGridView)
        {
            try
            {
                var wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = true;
                Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Add();
                Paragraph title = doc.Paragraphs.Add();
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        title.Range.Text = "Данные оборудования";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title.Range.Text = "Данные передвижения оборудования";
                        break;

                    case "dataGridViewSupplier":
                        title.Range.Text = "Данные поставщиков";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title.Range.Text = "Данные поставок оборудования";
                        break;
                }
                title.Range.Font.Bold = 1;
                title.Range.Font.Size = 14;
                title.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();
                Table table = doc.Tables.Add(title.Range, dataGridView.RowCount + 1, dataGridView.ColumnCount - 1);
                for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                {
                    table.Cell(1, col + 1).Range.Text = dataGridView.Columns[col].HeaderText;
                }
                for (int row = 0; row < dataGridView.RowCount; row++)
                {
                    for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                    {
                        table.Cell(row + 2, col + 1).Range.Text = dataGridView[col, row].Value.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToExcel вызывается при экспорте данных в Excel
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToExcel(DataGridView dataGridView)
        {
            try
            {
                var excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                string title = "";
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        title = "Данные оборудования";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title = "Данные передвижения оборудования";
                        break;

                    case "dataGridViewSupplier":
                        title = "Данные поставщиков";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title = "Данные поставок оборудования";
                        break;
                }
                Excel.Range titleRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, dataGridView.ColumnCount - 1]];
                titleRange.Merge();
                titleRange.Value = title;
                titleRange.Font.Bold = true;
                titleRange.Font.Size = 14;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                for (int col = 0; col < dataGridView.ColumnCount; col++)
                {
                    worksheet.Cells[2, col + 1] = dataGridView.Columns[col].HeaderText;
                }
                for (int row = 0; row < dataGridView.RowCount; row++)
                {
                    for (int col = 0; col < dataGridView.ColumnCount - 1; col++)
                    {
                        worksheet.Cells[row + 3, col + 1] = dataGridView[col, row].Value.ToString();
                        Excel.Range dataRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[dataGridView.RowCount + 2, dataGridView.ColumnCount]];
                        dataRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        dataRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                }
                worksheet.Columns.AutoFit();
                worksheet.Rows.AutoFit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ExportToPDF вызывается при экспорте данных в PDF
        /// </summary>
        /// <param name="dataGridView"></param>
        private void ExportToPDF(DataGridView dataGridView)
        {
            try
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.pdf");
                var pdfWriter = new PdfWriter(filePath);
                var pdfDocument = new PdfDocument(pdfWriter);
                var pdfDoc = new iText.Layout.Document(pdfDocument);
                PdfFont timesFont = PdfFontFactory.CreateFont("c:/windows/fonts/times.ttf", PdfEncodings.IDENTITY_H, true);
                string title = "";
                switch (dataGridView.Name)
                {
                    case "dataGridViewEquipment":
                        title = "Данные оборудования";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title = "Данные передвижения оборудования";
                        break;

                    case "dataGridViewSupplier":
                        title = "Данные поставщиков";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title = "Данные поставок оборудования";
                        break;
                }
                pdfDoc.Add(new iText.Layout.Element.Paragraph(title).SetFont(timesFont).SetTextAlignment(TextAlignment.CENTER));
                iText.Layout.Element.Table table = new iText.Layout.Element.Table(dataGridView.Columns.Count - 1);
                table.UseAllAvailableWidth();
                var columnsList = dataGridView.Columns.Cast<DataGridViewColumn>().ToList();
                foreach (DataGridViewColumn column in columnsList.Take(dataGridView.Columns.Count - 1))
                {
                    iText.Layout.Element.Cell headerCell = new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(column.HeaderText).SetFont(timesFont));
                    table.AddHeaderCell(headerCell);
                }
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells.Cast<DataGridViewCell>().Take(dataGridView.Columns.Count - 1))
                    {
                        table.AddCell(new iText.Layout.Element.Cell().Add(new iText.Layout.Element.Paragraph(cell.Value.ToString()).SetFont(timesFont)));
                    }
                }
                pdfDoc.Add(table);
                pdfDoc.Close();
                MessageBox.Show("PDF успешно экспортирован.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ImportToExcel(DataGridView dataGridView)
        {
            switch (dataGridView.Name)
            {
                case "dataGridViewEquipment":

                    break;

                case "dataGridViewEquipmentMovement":

                    break;

                case "dataGridViewSupplier":

                    break;

                case "dataGridViewEquipmentSupplier":

                    break;
            }
        }

        /// <summary>
        /// ButtonRefresh_Click вызывается при нажатии на кнопку обновления
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonRefresh_Click(object sender, EventArgs e)
        {
            try
            {
                RefreshDataGrid(dataGridViewEquipment, "Equipment");
                RefreshDataGrid(dataGridViewEquipmentMovement, "EquipmentMovement");
                RefreshDataGrid(dataGridViewSupplier, "Supplier");
                RefreshDataGrid(dataGridViewEquipmentSupplier, "EquipmentSupplier");
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// ButtonClear_Click вызывается при нажатии на кнопку "Изменить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClear_Click(object sender, EventArgs e)
        {
            try
            {
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                AddFormEquipment addForm = new AddFormEquipment();
                addForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                AddFormEquipmentMovement addForm = new AddFormEquipmentMovement();
                addForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                if (admin)
                {
                    AddFormSupplier addForm = new AddFormSupplier();
                    addForm.Show();
                }
                else
                {
                    MessageBox.Show("У вас недостаточно прав!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                AddFormEquipmentSupplier addForm = new AddFormEquipmentSupplier();
                addForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonDeleteEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteRow(dataGridViewEquipment);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonDeleteEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteRow(dataGridViewEquipmentMovement);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonDeleteSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteRow(dataGridViewSupplier);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonDeleteEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                DeleteRow(dataGridViewEquipmentSupplier);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonChangeEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                Change(dataGridViewEquipment);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonChangeEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                Change(dataGridViewEquipmentMovement);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonChangeSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                Change(dataGridViewSupplier);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonChangeEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                Change(dataGridViewEquipmentSupplier);
                ClearFields();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateBase(dataGridViewEquipment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateBase(dataGridViewEquipmentMovement);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                if (admin)
                {
                    UpdateBase(dataGridViewSupplier);
                }
                else
                {
                    MessageBox.Show("У вас недостаточно прав!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateBase(dataGridViewEquipmentSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonWordEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToWord(dataGridViewEquipment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonWordEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToWord(dataGridViewEquipmentMovement);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonWordSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToWord(dataGridViewSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonWordEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToWord(dataGridViewEquipmentSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonExcelEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToExcel(dataGridViewEquipment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonExcelEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToExcel(dataGridViewEquipmentMovement);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonExcelSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToExcel(dataGridViewSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonExcelEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToExcel(dataGridViewEquipmentSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonPDFEquipment_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToPDF(dataGridViewEquipment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonPDFEquipmentMovement_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToPDF(dataGridViewEquipmentMovement);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonPDFSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToPDF(dataGridViewSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonPDFEquipmentSupplier_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToPDF(dataGridViewEquipmentSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGridViewEquipment_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                selectedRow = e.RowIndex;
                if (e.RowIndex >= 0)
                {
                    DataGridView_CellClick(dataGridViewEquipment, selectedRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGridViewEquipmentMovement_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                selectedRow = e.RowIndex;
                if (e.RowIndex >= 0)
                {
                    DataGridView_CellClick(dataGridViewEquipmentMovement, selectedRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGridViewSupplier_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                selectedRow = e.RowIndex;
                if (e.RowIndex >= 0)
                {
                    DataGridView_CellClick(dataGridViewSupplier, selectedRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DataGridViewEquipmentSupplier_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                selectedRow = e.RowIndex;
                if (e.RowIndex >= 0)
                {
                    DataGridView_CellClick(dataGridViewEquipmentSupplier, selectedRow);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBoxSearchEquipment_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Search(dataGridViewEquipment);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBoxSearchEquipmentMovement_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Search(dataGridViewEquipmentMovement);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBoxSearchSupplier_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Search(dataGridViewSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBoxSearchEquipmentSupplier_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Search(dataGridViewEquipmentSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}