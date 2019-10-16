using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;

namespace Расчёт_стартапа
{
    public partial class Reg : Form
    {
        public Reg()
        {
            InitializeComponent();
        }

        public static string GetCrypt(string text)
        {
            var crypt = new System.Security.Cryptography.SHA256Managed();
            var hash = new System.Text.StringBuilder();
            byte[] crypto = crypt.ComputeHash(Encoding.UTF8.GetBytes(text));
            foreach (byte theByte in crypto)
            {
                hash.Append(theByte.ToString("x2"));
            }
            return hash.ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != textBox3.Text)
            {
                MessageBox.Show("Введённые пароли не совпадают");
            }
            else
            {
                string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection3 = new MySqlConnection(connectionString3);
                try
                {

                    string command3 = string.Format("Select * From users WHERE login=\"" + textBox1.Text + "\"");
                    MySqlCommand check = new MySqlCommand(command3, connection3);
                    connection3.Open();
                    if (check.ExecuteScalar() == null)
                    {
                        string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        using (MySqlConnection connection = new MySqlConnection(connectionString))
                        {
                           
                            string commText = "Insert Into users  ( login, password, surname, name, patronymic, company_name) VALUES( @login, @password, @surname, @name, @patronymic, @company_name )";
                            MySqlCommand comm = new MySqlCommand(commText, connection);


                            comm.Parameters.AddWithValue("@login", textBox1.Text);
                            comm.Parameters.AddWithValue("@password", GetCrypt(textBox2.Text));
                            comm.Parameters.AddWithValue("@surname", textBox4.Text);
                            comm.Parameters.AddWithValue("@name", textBox5.Text);
                            comm.Parameters.AddWithValue("@patronymic", textBox6.Text);
                            comm.Parameters.AddWithValue("@company_name", textBox7.Text);
                            connection.Open();
                            try
                            {
                                comm.ExecuteNonQuery();

                                Login login = new Login();
                                login.Show();
                                this.Hide();
                            }
                            catch
                            {
                                MessageBox.Show("Добавить не удалось!");
                            }
                           
                        }
                    }
                    else
                    {
                        MessageBox.Show("Такой логин уже существует");
                    }
                }
                finally
                {
                    connection3.Close();
                }
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void Reg_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void Reg_FormClosed(object sender, FormClosedEventArgs e)
        {
            Login form1 = new Login();
            form1.Show();
            this.Hide();

        }
    }
}
