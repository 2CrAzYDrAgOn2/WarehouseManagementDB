﻿using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace test_DataBase
{
    public partial class LogIn : Form
    {
        private readonly DataBase dataBase = new DataBase();

        public LogIn()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        /// <summary>
        /// Form1_Load вызывается при загрузке формы "Form1"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            textBoxPassword.PasswordChar = '•';
            textBoxLogin.MaxLength = 50;
            textBoxPassword.MaxLength = 50;
        }

        /// <summary>
        /// ButtonEnter_Click вызывается при нажатии на кнопку "Войти"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonEnter_Click(object sender, EventArgs e)
        {
            var loginUser = textBoxLogin.Text;
            var passwordUser = textBoxPassword.Text;
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
            DataTable dataTable = new DataTable();
            string querystring = $"select UserID, UserLogin, UserPassword, IsAdmin from Registration where UserLogin = '{loginUser}' and UserPassword = '{passwordUser}'";
            SqlCommand sqlCommand = new SqlCommand(querystring, dataBase.GetConnection());
            sqlDataAdapter.SelectCommand = sqlCommand;
            sqlDataAdapter.Fill(dataTable);
            if (dataTable.Rows.Count == 1)
            {
                bool isAdmin = Convert.ToBoolean(dataTable.Rows[0]["IsAdmin"]);
                MessageBox.Show("Вы успешно вошли!", "Успешно!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Form1 form1 = new Form1();
                form1.SetAdminStatus(isAdmin);
                this.Hide();
                form1.ShowDialog();
                this.Show();
            }
            else
            {
                MessageBox.Show("Такого аккаунта не существует!", "Аккаунта не существует!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// ButtonClear_Click вызывается при нажатии на кнопку очистки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ButtonClear_Click(object sender, EventArgs e)
        {
            textBoxLogin.Text = "";
            textBoxPassword.Text = "";
        }

        /// <summary>
        /// LabelAuth_Click вызывается при нажатии на текст "Ещё нет аккаунта?"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void LabelAuth_Click(object sender, EventArgs e)
        {
            this.Hide();
            SignUp formLogin = new SignUp();
            formLogin.ShowDialog();
        }
    }
}