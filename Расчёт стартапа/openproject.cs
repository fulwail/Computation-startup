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
    public partial class openproject : Form
    {

        public openproject()
        {
            InitializeComponent();
        }
        string login = "";
        public void Custom_Initialize2(string id3)
        {
            login = id3;
            // MessageBox.Show(Convert.ToString(login));
            //Необходимо вывести id из БД!!!!!
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void openproject_Load(object sender, EventArgs e)
        {
            string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            string command = "Select User_ID From users WHERE login=\"" + login + "\"";
            MySqlCommand cmd = new MySqlCommand(command, connection);
            int id = 0;
            using (MySqlDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    id = reader.GetInt32(0);
                }
            }
            string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds = new DataSet();
            MySqlConnection connection2 = new MySqlConnection(connectionString2);
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter("SELECT `project_name` FROM `project` WHERE `User_ID`=\"" + id+ "\""+ "ORDER BY project_name ASC", connection2);

            dataAdapter.Fill(ds, "Name");
            dataGridView1.DataSource = ds.Tables["Name"];
           
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string index = dataGridView1.SelectedCells[0].Value.ToString();
           
                   Form1 form1 = new Form1();
            form1.Text = index;
                    form1.Custom_Initialize3(index.ToString());
                    form1.Custom_Initialize2(login.ToString());
                    form1.Show();
            this.Hide();


        }

        private void openproject_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
    }
}
