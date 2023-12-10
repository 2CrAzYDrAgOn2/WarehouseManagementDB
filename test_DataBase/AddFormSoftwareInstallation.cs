using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormSoftwareInstallation : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormSoftwareInstallation()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        /// <summary>
        /// ButtonSave_Click вызывается при нажатии на кнопку "Сохранить"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonSave_Click(object sender, EventArgs e)
        {
            try
            {
                dataBase.OpenConnection();
                var installationDate = textBoxInstallationDate.Text;
                var softwareName = textBoxSoftwareName.Text;
                var licenseKey = textBoxLicenseKey.Text;
                if (int.TryParse(textBoxClientIDSoftwareInstallation.Text, out int clientIDSoftwareInstallation) && int.TryParse(textBoxTechnicianIDSoftwareInstallation.Text, out int technicianIDSoftwareInstallation))
                {
                    var addQuery = $"insert into SoftwareInstallation (ClientID, TechnicianID, InstallationDate, SoftwareName, LicenseKey) values ('{clientIDSoftwareInstallation}', '{technicianIDSoftwareInstallation}', '{installationDate}', '{softwareName}', '{licenseKey}')";
                    var sqlCommand = new SqlCommand(addQuery, dataBase.GetConnection());
                    sqlCommand.ExecuteNonQuery();
                    MessageBox.Show("Запись успешно создана!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Цена должна иметь числовой формат!", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
    }
}