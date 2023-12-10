using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormRepairOrders : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormRepairOrders()
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
                var orderDate = textBoxOrderDate.Text;
                var description = textBoxDescription.Text;
                var status = textBoxStatus.Text;
                if (int.TryParse(textBoxClientIDRepairOrders.Text, out int clientIDRepairOrders) && int.TryParse(textBoxTechnicianIDRepairOrders.Text, out int technicianIDRepairOrders))
                {
                    var addQuery = $"insert into RepairOrders (ClientID, TechnicianID, OrderDate, Description, Status) values ('{clientIDRepairOrders}', '{technicianIDRepairOrders}', '{orderDate}', '{description}', '{status}')";
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