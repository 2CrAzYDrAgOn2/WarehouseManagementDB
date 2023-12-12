using System;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class AddFormEquipmentMovement : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public AddFormEquipmentMovement()
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
                var movementDate = textBoxMovementDate.Text;
                var movementType = textBoxMovementType.Text;
                if (int.TryParse(textBoxEquipmentIDEquipmentMovement.Text, out int equipmentID) && int.TryParse(textBoxQuantinityEquipmentMovement.Text, out int quantity))
                {
                    var addQuery = $"insert into EquipmentMovement (EquipmentID, MovementDate, MovementType, Quantity) values ('{equipmentID}', '{movementDate}', '{movementType}', '{quantity}')";
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