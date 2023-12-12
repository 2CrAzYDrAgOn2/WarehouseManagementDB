using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormEquipment : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormEquipment()
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
                var name = textBoxName.Text;
                var category = textBoxCategory.Text;
                var purchaseDate = textBoxPurchaseData.Text;
                var location = textBoxLocation.Text;
                if (int.TryParse(textBoxPrice.Text, out int price) && int.TryParse(textBoxQuantinity.Text, out int quantity))
                {
                    var addQuery = $"insert into Equipment (Name, Category, PurchaseDate, Price, Quantity, Location) values ('{name}', '{category}', '{purchaseDate}', '{price}', '{quantity}', '{location}')";
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