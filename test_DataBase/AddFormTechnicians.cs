using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormTechnicians : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormTechnicians()
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
                var firstNameTechnicians = textBoxFirstNameTechnicians.Text;
                var lastNameTechnicians = textBoxLastNameTechnicians.Text;
                var phoneNumberTechnicians = textBoxPhoneNumberTechnicians.Text;
                var emailTechnicians = textBoxEmailTechnicians.Text;
                var addQuery = $"insert into Technicians (FirstName, LastName, PhoneNumber, Email) values ('{firstNameTechnicians}', '{lastNameTechnicians}', '{phoneNumberTechnicians}', '{emailTechnicians}')";
                var sqlCommand = new SqlCommand(addQuery, dataBase.GetConnection());
                sqlCommand.ExecuteNonQuery();
                MessageBox.Show("Запись успешно создана!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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