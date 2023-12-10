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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
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
                dataGridViewClients.Columns.Add("ClientID", "Номер");
                dataGridViewClients.Columns.Add("FirstName", "Имя");
                dataGridViewClients.Columns.Add("LastName", "Фамилия");
                dataGridViewClients.Columns.Add("PhoneNumber", "Телефон");
                dataGridViewClients.Columns.Add("Email", "Email");
                dataGridViewClients.Columns.Add("IsNew", String.Empty);
                dataGridViewTechnicians.Columns.Add("TechnicianID", "Номер");
                dataGridViewTechnicians.Columns.Add("FistName", "Имя");
                dataGridViewTechnicians.Columns.Add("LastName", "Фамилия");
                dataGridViewTechnicians.Columns.Add("PhoneNumber", "Телефон");
                dataGridViewTechnicians.Columns.Add("Email", "Email");
                dataGridViewTechnicians.Columns.Add("IsNew", String.Empty);
                dataGridViewRepairOrders.Columns.Add("OrderID", "Номер");
                dataGridViewRepairOrders.Columns.Add("ClientID", "Номер клиента");
                dataGridViewRepairOrders.Columns.Add("TechnicianID", "Номер техника");
                dataGridViewRepairOrders.Columns.Add("OrderDate", "Дата заказа");
                dataGridViewRepairOrders.Columns.Add("Description", "Описание");
                dataGridViewRepairOrders.Columns.Add("Status", "Статус");
                dataGridViewRepairOrders.Columns.Add("IsNew", String.Empty);
                dataGridViewSoftwareInstallation.Columns.Add("InstallationID", "Номер");
                dataGridViewSoftwareInstallation.Columns.Add("ClientID", "Номер клиента");
                dataGridViewSoftwareInstallation.Columns.Add("TechnicianID", "Номер техника");
                dataGridViewSoftwareInstallation.Columns.Add("InstallationDate", "Дата установки");
                dataGridViewSoftwareInstallation.Columns.Add("SoftwareName", "Название ПО");
                dataGridViewSoftwareInstallation.Columns.Add("LicenseKey", "Лицензионный ключ");
                dataGridViewSoftwareInstallation.Columns.Add("IsNew", String.Empty);
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
                textBoxClientIDClients.Text = "";
                textBoxFirstNameClients.Text = "";
                textBoxLastNameClients.Text = "";
                textBoxPhoneNumberClients.Text = "";
                textBoxEmailClients.Text = "";
                textBoxTechnicianID.Text = "";
                textBoxFirstNameTechnicians.Text = "";
                textBoxLastNameTechnicians.Text = "";
                textBoxPhoneNumberTechnicians.Text = "";
                textBoxEmailTechnicians.Text = "";
                textBoxOrderID.Text = "";
                textBoxClientIDRepairOrders.Text = "";
                textBoxTechnicianIDRepairOrders.Text = "";
                textBoxOrderDate.Text = "";
                textBoxDescription.Text = "";
                textBoxStatus.Text = "";
                textBoxInstallationID.Text = "";
                textBoxClientIDSoftwareInstallation.Text = "";
                textBoxTechnicianIDSoftwareInstallation.Text = "";
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
                    case "dataGridViewClients":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewTechnicians":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetString(1), iDataRecord.GetString(2), iDataRecord.GetString(3), iDataRecord.GetString(4), RowState.Modified);
                        break;

                    case "dataGridViewRepairOrders":
                        dataGridView.Rows.Add(iDataRecord.GetInt32(0), iDataRecord.GetInt32(1), iDataRecord.GetInt32(2), iDataRecord.GetDateTime(3), iDataRecord.GetString(4), iDataRecord.GetString(5), RowState.Modified);
                        break;

                    case "dataGridViewSoftwareInstallation":
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
                RefreshDataGrid(dataGridViewClients, "Clients");
                RefreshDataGrid(dataGridViewTechnicians, "Technicians");
                RefreshDataGrid(dataGridViewRepairOrders, "RepairOrders");
                RefreshDataGrid(dataGridViewSoftwareInstallation, "SoftwareInstallation");
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
                    case "dataGridViewClients":
                        textBoxClientIDClients.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxFirstNameClients.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxLastNameClients.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxPhoneNumberClients.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxEmailClients.Text = dataGridViewRow.Cells[4].Value.ToString();
                        break;

                    case "dataGridViewTechnicians":
                        textBoxTechnicianID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxFirstNameTechnicians.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxLastNameTechnicians.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxPhoneNumberTechnicians.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxEmailTechnicians.Text = dataGridViewRow.Cells[4].Value.ToString();
                        break;

                    case "dataGridViewRepairOrders":
                        textBoxOrderID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxClientIDRepairOrders.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxTechnicianIDRepairOrders.Text = dataGridViewRow.Cells[2].Value.ToString();
                        textBoxOrderDate.Text = dataGridViewRow.Cells[3].Value.ToString();
                        textBoxDescription.Text = dataGridViewRow.Cells[4].Value.ToString();
                        textBoxStatus.Text = dataGridViewRow.Cells[5].Value.ToString();
                        break;

                    case "dataGridViewSoftwareInstallation":
                        textBoxInstallationID.Text = dataGridViewRow.Cells[0].Value.ToString();
                        textBoxClientIDSoftwareInstallation.Text = dataGridViewRow.Cells[1].Value.ToString();
                        textBoxTechnicianIDSoftwareInstallation.Text = dataGridViewRow.Cells[2].Value.ToString();
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
                    case "dataGridViewClients":
                        string searchStringClients = $"select * from Clients where concat (ClientID, FirstName, LastName, PhoneNumber, Email) like '%" + textBoxSearchClients.Text + "%'";
                        SqlCommand sqlCommandClients = new SqlCommand(searchStringClients, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderClients = sqlCommandClients.ExecuteReader();
                        while (sqlDataReaderClients.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderClients);
                        }
                        sqlDataReaderClients.Close();
                        break;

                    case "dataGridViewTechnicians":
                        string searchStringTechnicians = $"select * from Technicians where concat (TechnicianID, FirstName, LastName, PhoneNumber, Email) like '%" + textBoxSearchTechnicians.Text + "%'";
                        SqlCommand sqlCommandTechnicians = new SqlCommand(searchStringTechnicians, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderTechnicians = sqlCommandTechnicians.ExecuteReader();
                        while (sqlDataReaderTechnicians.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderTechnicians);
                        }
                        sqlDataReaderTechnicians.Close();
                        break;

                    case "dataGridViewRepairOrders":
                        string searchStringRepairOrders = $"select * from RepairOrders where concat (OrderID, ClientID, TechnicianID, OrderDate, Description, Status) like '%" + textBoxSearchRepairOrders.Text + "%'";
                        SqlCommand sqlCommandRepairOrders = new SqlCommand(searchStringRepairOrders, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderRepairOrders = sqlCommandRepairOrders.ExecuteReader();
                        while (sqlDataReaderRepairOrders.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderRepairOrders);
                        }
                        sqlDataReaderRepairOrders.Close();
                        break;

                    case "dataGridViewSoftwareInstallation":
                        string searchStringSoftwareInstallation = $"select * from SoftwareInstallation where concat (IstallationID, ClientID, TechnicianID, InstallationDate, SoftwareName, LicenseKey) like '%" + textBoxSearchSoftwareInstallation.Text + "%'";
                        SqlCommand sqlCommandSoftwareInstallation = new SqlCommand(searchStringSoftwareInstallation, dataBase.GetConnection());
                        dataBase.OpenConnection();
                        SqlDataReader sqlDataReaderSoftwareInstallation = sqlCommandSoftwareInstallation.ExecuteReader();
                        while (sqlDataReaderSoftwareInstallation.Read())
                        {
                            ReadSingleRow(dataGridView, sqlDataReaderSoftwareInstallation);
                        }
                        sqlDataReaderSoftwareInstallation.Close();
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
                    case "dataGridViewClients":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                        break;

                    case "dataGridViewTechnicians":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[5].Value = RowState.Deleted;
                        break;

                    case "dataGridViewRepairOrders":
                        if (dataGridView.Rows[index].Cells[0].Value.ToString() == string.Empty)
                        {
                            dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                            return;
                        }
                        dataGridView.Rows[index].Cells[6].Value = RowState.Deleted;
                        break;

                    case "dataGridViewSoftwareInstallation":
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
                        case "dataGridViewClients":
                            var rowStateClients = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateClients == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateClients == RowState.Deleted)
                            {
                                var clientID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Clients where ClientID = {clientID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateClients == RowState.Modified)
                            {
                                var clientID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var firstName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var lastName = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var phoneNumber = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var email = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update Clients set FirstName = '{firstName}', LastName = '{lastName}', PhoneNumber = '{phoneNumber}', Email = '{email}' where ClientID = '{clientID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewTechnicians":
                            var rowStateTechnicians = (RowState)dataGridView.Rows[index].Cells[5].Value;
                            if (rowStateTechnicians == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateTechnicians == RowState.Deleted)
                            {
                                var technicianID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from Technicians where TechnicianID = {technicianID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateTechnicians == RowState.Modified)
                            {
                                var technicianID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var firstName = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var lastName = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var phoneNumber = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var email = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var changeQuery = $"update Technicians set FirstName = '{firstName}', LastName = '{lastName}', PhoneNumber = '{phoneNumber}', Email = '{email}' where TechnicianID = '{technicianID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewRepairOrders":
                            var rowStateRepairOrders = (RowState)dataGridView.Rows[index].Cells[6].Value;
                            if (rowStateRepairOrders == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateRepairOrders == RowState.Deleted)
                            {
                                var orderID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from RepairOrders where OrderID = {orderID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateRepairOrders == RowState.Modified)
                            {
                                var orderID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var clientID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var technicianID = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var orderDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var description = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var status = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var changeQuery = $"update RepairOrders set ClientID = '{clientID}', TechnicianID = '{technicianID}', OrderDate = '{orderDate}', Description = '{description}', Status = '{status}' where OrderID = '{orderID}'";
                                var sqlCommand = new SqlCommand(changeQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            break;

                        case "dataGridViewSoftwareInstallation":
                            var rowStateSoftwareInstallation = (RowState)dataGridView.Rows[index].Cells[6].Value;
                            if (rowStateSoftwareInstallation == RowState.Existed)
                            {
                                continue;
                            }
                            if (rowStateSoftwareInstallation == RowState.Deleted)
                            {
                                var installationID = Convert.ToInt32(dataGridView.Rows[index].Cells[0].Value);
                                var deleteQuery = $"delete from SoftwareInstallation where InstallationID = {installationID}";
                                var sqlCommand = new SqlCommand(deleteQuery, dataBase.GetConnection());
                                sqlCommand.ExecuteNonQuery();
                            }
                            if (rowStateSoftwareInstallation == RowState.Modified)
                            {
                                var installationID = dataGridView.Rows[index].Cells[0].Value.ToString();
                                var clientID = dataGridView.Rows[index].Cells[1].Value.ToString();
                                var technicianID = dataGridView.Rows[index].Cells[2].Value.ToString();
                                var installationDate = dataGridView.Rows[index].Cells[3].Value.ToString();
                                var softwareName = dataGridView.Rows[index].Cells[4].Value.ToString();
                                var licenseKey = dataGridView.Rows[index].Cells[5].Value.ToString();
                                var changeQuery = $"update SoftwareInstallation set ClientID = '{clientID}', TechnicianID = '{technicianID}', InstallationDate = '{installationDate}', SoftwareName = '{softwareName}', LicenseKey = '{licenseKey}' where InstallationID = '{installationID}'";
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
                    case "dataGridViewClients":
                        var clientID = textBoxClientIDClients.Text;
                        var firstName = textBoxFirstNameClients.Text;
                        var lastName = textBoxLastNameClients.Text;
                        var phoneNumber = textBoxPhoneNumberClients.Text;
                        var email = textBoxEmailClients.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(clientID, firstName, lastName, phoneNumber, email);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewTechnicians":
                        var technicianID = textBoxTechnicianID.Text;
                        var firstNameTechnicians = textBoxFirstNameTechnicians.Text;
                        var lastNameTechnicians = textBoxLastNameTechnicians.Text;
                        var phoneNumberTechnicians = textBoxPhoneNumberTechnicians.Text;
                        var emailTechnicians = textBoxEmailTechnicians.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(technicianID, firstNameTechnicians, lastNameTechnicians, phoneNumberTechnicians, emailTechnicians);
                        dataGridView.Rows[selectedRowIndex].Cells[5].Value = RowState.Modified;
                        break;

                    case "dataGridViewRepairOrders":
                        var orderID = textBoxOrderID.Text;
                        var clientIDRepairOrders = textBoxClientIDRepairOrders.Text;
                        var technicianIDRepairOrders = textBoxTechnicianIDRepairOrders.Text;
                        var orderDate = textBoxOrderDate.Text;
                        var description = textBoxDescription.Text;
                        var status = textBoxStatus.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(orderID, clientIDRepairOrders, technicianIDRepairOrders, orderDate, description, status);
                        dataGridView.Rows[selectedRowIndex].Cells[6].Value = RowState.Modified;
                        break;

                    case "dataGridViewSoftwareInstallation":
                        var installationID = textBoxInstallationID.Text;
                        var clientIDSoftwareInstallation = textBoxClientIDSoftwareInstallation.Text;
                        var technicianIDSoftwareInstallation = textBoxTechnicianIDSoftwareInstallation.Text;
                        var installationDate = textBoxInstallationDate.Text;
                        var softwareName = textBoxSoftwareName.Text;
                        var licenseKey = textBoxLicenseKey.Text;
                        dataGridView.Rows[selectedRowIndex].SetValues(installationID, clientIDSoftwareInstallation, technicianIDSoftwareInstallation, installationDate, softwareName, licenseKey);
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
                    case "dataGridViewClients":
                        title.Range.Text = "Данные клиентов";
                        break;

                    case "dataGridViewTechnicians":
                        title.Range.Text = "Данные техников";
                        break;

                    case "dataGridViewRepairOrders":
                        title.Range.Text = "Данные заказов";
                        break;

                    case "dataGridViewSoftwareInstallation":
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
                    case "dataGridViewClients":
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewTechnicians":
                        title = "Данные техников";
                        break;

                    case "dataGridViewRepairOrders":
                        title = "Данные заказов";
                        break;

                    case "dataGridViewSoftwareInstallation":
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
                    case "dataGridViewClients":
                        title = "Данные клиентов";
                        break;

                    case "dataGridViewTechnicians":
                        title = "Данные техников";
                        break;

                    case "dataGridViewRepairOrders":
                        title = "Данные заказов";
                        break;

                    case "dataGridViewSoftwareInstallation":
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
                RefreshDataGrid(dataGridViewClients, "Clients");
                RefreshDataGrid(dataGridViewTechnicians, "Technicians");
                RefreshDataGrid(dataGridViewRepairOrders, "RepairOrders");
                RefreshDataGrid(dataGridViewSoftwareInstallation, "SoftwareInstallation");
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
                AddFormClients addForm = new AddFormClients();
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
                    AddFormTechnicians addForm = new AddFormTechnicians();
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
                AddFormRepairOrders addForm = new AddFormRepairOrders();
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
                AddFormSoftwareInstallation addForm = new AddFormSoftwareInstallation();
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
                DeleteRow(dataGridViewClients);
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
                DeleteRow(dataGridViewTechnicians);
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
                DeleteRow(dataGridViewRepairOrders);
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
                DeleteRow(dataGridViewSoftwareInstallation);
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
                Change(dataGridViewClients);
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
                Change(dataGridViewTechnicians);
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
                Change(dataGridViewRepairOrders);
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
                Change(dataGridViewSoftwareInstallation);
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
                UpdateBase(dataGridViewClients);
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
                    UpdateBase(dataGridViewTechnicians);
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
                UpdateBase(dataGridViewRepairOrders);
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
                UpdateBase(dataGridViewSoftwareInstallation);
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
                ExportToWord(dataGridViewClients);
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
                ExportToWord(dataGridViewTechnicians);
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
                ExportToWord(dataGridViewRepairOrders);
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
                ExportToWord(dataGridViewSoftwareInstallation);
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
                ExportToExcel(dataGridViewClients);
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
                ExportToExcel(dataGridViewTechnicians);
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
                ExportToExcel(dataGridViewRepairOrders);
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
                ExportToExcel(dataGridViewSoftwareInstallation);
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
                ExportToPDF(dataGridViewClients);
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
                ExportToPDF(dataGridViewTechnicians);
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
                ExportToPDF(dataGridViewRepairOrders);
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
                ExportToPDF(dataGridViewSoftwareInstallation);
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
                    DataGridView_CellClick(dataGridViewClients, selectedRow);
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
                    DataGridView_CellClick(dataGridViewTechnicians, selectedRow);
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
                    DataGridView_CellClick(dataGridViewRepairOrders, selectedRow);
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
                    DataGridView_CellClick(dataGridViewSoftwareInstallation, selectedRow);
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
                Search(dataGridViewClients);
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
                Search(dataGridViewTechnicians);
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
                Search(dataGridViewRepairOrders);
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
                Search(dataGridViewSoftwareInstallation);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}