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
                dataGridViewEquipment.Columns.Add("ClientID", "Номер");
                dataGridViewEquipment.Columns.Add("FirstName", "Имя");
                dataGridViewEquipment.Columns.Add("LastName", "Фамилия");
                dataGridViewEquipment.Columns.Add("PhoneNumber", "Телефон");
                dataGridViewEquipment.Columns.Add("Email", "Email");
                dataGridViewEquipment.Columns.Add("IsNew", String.Empty);
                dataGridViewEquipmentMovement.Columns.Add("TechnicianID", "Номер");
                dataGridViewEquipmentMovement.Columns.Add("FistName", "Имя");
                dataGridViewEquipmentMovement.Columns.Add("LastName", "Фамилия");
                dataGridViewEquipmentMovement.Columns.Add("PhoneNumber", "Телефон");
                dataGridViewEquipmentMovement.Columns.Add("Email", "Email");
                dataGridViewEquipmentMovement.Columns.Add("IsNew", String.Empty);
                dataGridViewSupplier.Columns.Add("OrderID", "Номер");
                dataGridViewSupplier.Columns.Add("ClientID", "Номер клиента");
                dataGridViewSupplier.Columns.Add("TechnicianID", "Номер техника");
                dataGridViewSupplier.Columns.Add("OrderDate", "Дата заказа");
                dataGridViewSupplier.Columns.Add("Description", "Описание");
                dataGridViewSupplier.Columns.Add("Status", "Статус");
                dataGridViewSupplier.Columns.Add("IsNew", String.Empty);
                dataGridViewEquipmentSupplier.Columns.Add("InstallationID", "Номер");
                dataGridViewEquipmentSupplier.Columns.Add("ClientID", "Номер клиента");
                dataGridViewEquipmentSupplier.Columns.Add("TechnicianID", "Номер техника");
                dataGridViewEquipmentSupplier.Columns.Add("InstallationDate", "Дата установки");
                dataGridViewEquipmentSupplier.Columns.Add("SoftwareName", "Название ПО");
                dataGridViewEquipmentSupplier.Columns.Add("LicenseKey", "Лицензионный ключ");
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
                textBoxStatus.Text = "";
                textBoxEquipmentIDEquipmentSupplier.Text = "";
                textBoxSupplierIDEquipmentSupplier.Text = "";
                textBoxTechnicianIDEquipmentSupplier.Text = "";
                textBoxInstallationDate.Text = "";
                textBoxSoftwareName.Text = "";
                textBoxLicenseKey.Text = "";
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
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewEquipmentMovement":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewSupplier":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), iDataRecord.GetInt32(2), iDataRecord.GetDateTime(3), iDataRecord.GetString(4), iDataRecord.GetString(5), RowState.Modified);
                        break;

                    case "dataGridViewEquipmentSupplier":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), iDataRecord.GetInt32(2), iDataRecord.GetDateTime(3), iDataRecord.GetString(4), iDataRecord.GetString(5), RowState.Modified);
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
                        textBoxStatus.Text = dataGridViewRow.Cells[5].Value.ToString();
                        break;

                    case "dataGridViewEquipmentSupplier":
                        textBoxEquipmentIDEquipmentSupplier.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxSupplierIDEquipmentSupplier.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxTechnicianIDEquipmentSupplier.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxInstallationDate.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxSoftwareName.Text = dataGridViewRow.Cells[4].Value.ToString();
                        textBoxLicenseKey.Text = dataGridViewRow.Cells[5].Value.ToString();
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
                        string searchStringEquipment = $"select * from Equipment where concat (ClientID, FirstName, LastName, PhoneNumber, Email) like '%" + textBoxSearchEquipment.Text + "%'";
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
                        string searchStringEquipmentMovement = $"select * from EquipmentMovement where concat (TechnicianID, FirstName, LastName, PhoneNumber, Email) like '%" + textBoxSearchEquipmentMovement.Text + "%'";
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
                        string searchStringSupplier = $"select * from Supplier where concat (OrderID, ClientID, TechnicianID, OrderDate, Description, Status) like '%" + textBoxSearchSupplier.Text + "%'";
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
                        string searchStringEquipmentSupplier = $"select * from EquipmentSupplier where concat (IstallationID, ClientID, TechnicianID, InstallationDate, SoftwareName, LicenseKey) like '%" + textBoxSearchEquipmentSupplier.Text + "%'";
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
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
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
                            dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                        break;

                    case "dataGridViewEquipmentSupplier":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
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
                            var rowStateEquipment = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateEquipment == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEquipment == RowState.Deleted)
                            {
                                var clientID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Equipment where ClientID = {clientID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipment == RowState.Modified)
                            {
                                var clientID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var firstName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var lastName = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var phoneNumber = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var email = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update Equipment set FirstName = '{firstName}', LastName = '{lastName}', PhoneNumber = '{phoneNumber}', Email = '{email}' where ClientID = '{clientID}'";
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
                                var technicianID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from EquipmentMovement where TechnicianID = {technicianID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipmentMovement == RowState.Modified)
                            {
                                var technicianID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var firstName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var lastName = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var phoneNumber = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var email = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update EquipmentMovement set FirstName = '{firstName}', LastName = '{lastName}', PhoneNumber = '{phoneNumber}', Email = '{email}' where TechnicianID = '{technicianID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewSupplier":
                            var rowStateSupplier = (RowState)dataGridView.Rows[index].Cells[6].Value;
                            if (rowStateSupplier == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateSupplier == RowState.Deleted)
                            {
                                var orderID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Supplier where OrderID = {orderID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateSupplier == RowState.Modified)
                            {
                                var orderID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var clientID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var technicianID = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var orderDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var description = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var status = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var changeQuery = $"update Supplier set ClientID = '{clientID}', TechnicianID = '{technicianID}', OrderDate = '{orderDate}', Description = '{description}', Status = '{status}' where OrderID = '{orderID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewEquipmentSupplier":
                            var rowStateEquipmentSupplier = (RowState)dataGridView.Rows[index].Cells[6].Value;
                            if (rowStateEquipmentSupplier == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateEquipmentSupplier == RowState.Deleted)
                            {
                                var installationID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from EquipmentSupplier where InstallationID = {installationID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateEquipmentSupplier == RowState.Modified)
                            {
                                var installationID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var clientID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var technicianID = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var installationDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var softwareName = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var licenseKey = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var changeQuery = $"update EquipmentSupplier set ClientID = '{clientID}', TechnicianID = '{technicianID}', InstallationDate = '{installationDate}', SoftwareName = '{softwareName}', LicenseKey = '{licenseKey}' where InstallationID = '{installationID}'";
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
                        var clientID = textBoxEquipmentID.Text;
                        var firstName = textBoxName.Text;
                        var lastName = textBoxCategory.Text;
                        var phoneNumber = textBoxPurchaseData.Text;
                        var email = textBoxPrice.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(clientID, firstName, lastName, phoneNumber, email);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewEquipmentMovement":
                        var technicianID = textBoxMovementID.Text;
                        var firstNameEquipmentMovement = textBoxEquipmentIDEquipmentMovement.Text;
                        var lastNameEquipmentMovement = textBoxMovementDate.Text;
                        var phoneNumberEquipmentMovement = textBoxMovementType.Text;
                        var emailEquipmentMovement = textBoxQuantinityEquipmentMovement.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(technicianID, firstNameEquipmentMovement, lastNameEquipmentMovement, phoneNumberEquipmentMovement, emailEquipmentMovement);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewSupplier":
                        var orderID = textBoxSupplierID.Text;
                        var clientIDSupplier = textBoxNameSupplier.Text;
                        var technicianIDSupplier = textBoxContactPerson.Text;
                        var orderDate = textBoxPhone.Text;
                        var description = textBoxEmail.Text;
                        var status = textBoxStatus.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(orderID, clientIDSupplier, technicianIDSupplier, orderDate, description, status);
                        dataGridView.Rows[selectedRowIndex].Cells[6].Value = RowState.Modified;
                        break;

                    case "dataGridViewEquipmentSupplier":
                        var installationID = textBoxEquipmentIDEquipmentSupplier.Text;
                        var clientIDEquipmentSupplier = textBoxSupplierIDEquipmentSupplier.Text;
                        var technicianIDEquipmentSupplier = textBoxTechnicianIDEquipmentSupplier.Text;
                        var installationDate = textBoxInstallationDate.Text;
                        var softwareName = textBoxSoftwareName.Text;
                        var licenseKey = textBoxLicenseKey.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(installationID, clientIDEquipmentSupplier, technicianIDEquipmentSupplier, installationDate, softwareName, licenseKey);
                        dataGridView.Rows[selectedRowIndex].Cells[6].Value = RowState.Modified;
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
                        title.Range.Text = "Данные клиентов";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title.Range.Text = "Данные техников";
                        break;

                    case "dataGridViewSupplier":
                        title.Range.Text = "Данные заказов";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title.Range.Text = "Данные установок ПО";
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
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title = "Данные техников";
                        break;

                    case "dataGridViewSupplier":
                        title = "Данные заказов";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title = "Данные установок ПО";
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
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewEquipmentMovement":
                        title = "Данные техников";
                        break;

                    case "dataGridViewSupplier":
                        title = "Данные заказов";
                        break;

                    case "dataGridViewEquipmentSupplier":
                        title = "Данные установок ПО";
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

        private void ButtonNewClients_Click(object sender, EventArgs e)
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

        private void ButtonNewTechnicians_Click(object sender, EventArgs e)
        {
            try
            {
                if (admin)
                {
                    AddFormEquipmentMovement addForm = new AddFormEquipmentMovement();
                    addForm.Show();
                }
                else
                {
                    MessageBox.Show("У вас недостаточно прав");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewRepairOrders_Click(object sender, EventArgs e)
        {
            try
            {
                AddFormSupplier addForm = new AddFormSupplier();
                addForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonNewSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonDeleteClients_Click(object sender, EventArgs e)
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

        private void ButtonDeleteTechnicians_Click(object sender, EventArgs e)
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

        private void ButtonDeleteRepairOrders_Click(object sender, EventArgs e)
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

        private void ButtonDeleteSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonChangeClients_Click(object sender, EventArgs e)
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

        private void ButtonChangeTechnicians_Click(object sender, EventArgs e)
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

        private void ButtonChangeRepairOrders_Click(object sender, EventArgs e)
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

        private void ButtonChangeSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonSaveClients_Click(object sender, EventArgs e)
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

        private void ButtonSaveTechnicians_Click(object sender, EventArgs e)
        {
            try
            {
                if (admin)
                {
                    UpdateBase(dataGridViewEquipmentMovement);
                }
                else
                {
                    MessageBox.Show("У вас недостаточно прав");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveRepairOrders_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateBase(dataGridViewSupplier);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ButtonSaveSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonWordClients_Click(object sender, EventArgs e)
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

        private void ButtonWordTechnicians_Click(object sender, EventArgs e)
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

        private void ButtonWordRepairOrders_Click(object sender, EventArgs e)
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

        private void ButtonWordSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonExcelClients_Click(object sender, EventArgs e)
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

        private void ButtonExcelTechnicians_Click(object sender, EventArgs e)
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

        private void ButtonExcelRepairOrders_Click(object sender, EventArgs e)
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

        private void ButtonExcelSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void ButtonPDFClients_Click(object sender, EventArgs e)
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

        private void ButtonPDFTechnicians_Click(object sender, EventArgs e)
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

        private void ButtonPDFRepairOrders_Click(object sender, EventArgs e)
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

        private void ButtonPDFSoftwareInstallation_Click(object sender, EventArgs e)
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

        private void DataGridViewClients_CellClick(object sender, DataGridViewCellEventArgs e)
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

        private void DataGridViewTechnicians_CellClick(object sender, DataGridViewCellEventArgs e)
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

        private void DataGridViewRepairOrders_CellClick(object sender, DataGridViewCellEventArgs e)
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

        private void DataGridViewSoftwareInstallation_CellClick(object sender, DataGridViewCellEventArgs e)
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

        private void TextBoxSearchClients_TextChanged(object sender, EventArgs e)
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

        private void TextBoxSearchTechnicians_TextChanged(object sender, EventArgs e)
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

        private void TextBoxSearchRepairOrders_TextChanged(object sender, EventArgs e)
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

        private void TextBoxSearchSoftwareInstallation_TextChanged(object sender, EventArgs e)
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