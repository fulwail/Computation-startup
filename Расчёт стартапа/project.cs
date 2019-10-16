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
    public partial class project : Form
    {
        public project()
        {
            InitializeComponent();
        }
        string login="";
        string index = "";
        int index5 = 0;
        public void Custom_Initialize2(string id3)
        {
            login = id3;
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

     

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection3 = new MySqlConnection(connectionString3);
            try
            {

                string command3 = string.Format("Select * From project WHERE project_name=\"" + textBox1.Text + "\"");
                MySqlCommand check = new MySqlCommand(command3, connection3);
                connection3.Open();
                if (check.ExecuteScalar() == null)
                {
                   
                    string command9 = string.Format("Select User_ID From users WHERE login=\"" + login+ "\"");
                    MySqlCommand cmd9 = new MySqlCommand(command9, connection3);
                    int P_ID = 0;
                    using (MySqlDataReader reader9 = cmd9.ExecuteReader())
                    {

                        while (reader9.Read())
                        {
                             P_ID= reader9.GetInt32(0);
                            index5 = P_ID;


                        }
                    }
                    string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection = new MySqlConnection(connectionString))
                    {
              
                        string commText = "Insert Into project  (project_name, User_ID) VALUES( @project_name, @User_ID )";
                        MySqlCommand comm = new MySqlCommand(commText, connection);


                        comm.Parameters.AddWithValue("@project_name", textBox1.Text);
                        comm.Parameters.AddWithValue("@User_ID", index5);
                        
                        connection.Open();
                        try
                        {
                            comm.ExecuteNonQuery();

                            MessageBox.Show("Проект был создан");
                            connection.Close();
                            index = textBox1.Text;
                            Form1 form1 = new Form1();
                            
                            form1.Text = index;
                            form1.Custom_Initialize3(index.ToString());
                            form1.Custom_Initialize2(login.ToString());
                            form1.Show();
                            this.Hide();

                        }
                        catch
                        {
                            MessageBox.Show(Convert.ToString(index5));
                            MessageBox.Show("Добавить не удалось!");
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Такой проект уже существует");
                }
            }
            finally
            {
                connection3.Close();
            }
        }

        private void project_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
    }
}
