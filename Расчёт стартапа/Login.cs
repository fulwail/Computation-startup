using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.IO;
using System.Windows.Forms;


namespace Расчёт_стартапа
{
    public partial class Login : Form
    {
        string d = "";
        public Login()
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

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            Reg reg = new Reg();
            reg.ShowDialog();
        }
       
      
        private void button1_Click(object sender, EventArgs e)
        {

            d = textBox1.Text;
            string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection = new MySqlConnection(connectionString);
            try
            {
                string command = string.Format("Select * From users WHERE login=\"" + textBox1.Text + "\""
                    + "AND password=\"" + GetCrypt(textBox2.Text) + "\"");
                MySqlCommand check = new MySqlCommand(command, connection);
                connection.Open();
                if (check.ExecuteScalar() != null)
                {
                    Form1 form1 = new Form1();
                    form1.Custom_Initialize2(d.ToString());
                    form1.Show();

                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль");
                }
            }
            finally
            {
                connection.Close();
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Login_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;

            if (l != '"' && l != '\\' && l != '/' && l != '`')
            { }
            else
            {
                e.Handled = true;
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;

            if (l != '"' && l != '\\' && l != '/' && l != '`')
            { }
            else
            {
                e.Handled = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void Login_Load(object sender, EventArgs e)
        {

        }
    }
}
