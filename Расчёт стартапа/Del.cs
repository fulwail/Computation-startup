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
    public partial class Del : Form
    {
        public Del()
        {
            InitializeComponent();
        }
        string login = "";
        public void Custom_Initialize2(string id3)
        {
            login = id3;
        }

        private void Del_Load(object sender, EventArgs e)
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
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter("SELECT `project_name` FROM `project` WHERE `User_ID`=\"" + id + "\"" + "ORDER BY project_name ASC", connection2);

            dataAdapter.Fill(ds, "Name");
            dataGridView1.DataSource = ds.Tables["Name"];
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            string index = dataGridView1.SelectedCells[0].Value.ToString();
            string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
         
            var result = new System.Windows.Forms.DialogResult();
            result = MessageBox.Show("Вы точно хотите удалить проект?", "Внимание",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                string commText2 = "DELETE FROM  project WHERE project_name=\"" + index + "\"";
                MySqlCommand comm2 = new MySqlCommand(commText2, connection);
                try
                {
                    comm2.ExecuteNonQuery();
                   

                }
                catch
                {
                  
                }
                string commText3 = "DELETE FROM  издержки_по_выполнению WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm3 = new MySqlCommand(commText3, connection);
                try
                {
                    comm3.ExecuteNonQuery();
                    

                }
                catch
                {
                   
                }
                string commText4 = "DELETE FROM  пок_основн_средств_и_аморт_отчислений WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm4 = new MySqlCommand(commText4, connection);
                try
                {
                    comm4.ExecuteNonQuery();
                  

                }
                catch
                {
                    
                }

                string commText5 = "DELETE FROM  прочие_расходы WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm5 = new MySqlCommand(commText5, connection);
                try
                {
                    comm5.ExecuteNonQuery();
                   

                }
                catch
                {
                   
                }
                string commText6 = "DELETE FROM   бюджет_рекламы WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm6 = new MySqlCommand(commText6, connection);
                try
                {
                    comm6.ExecuteNonQuery();
                    

                }
                catch
                {
                }

                string commText7 = "DELETE FROM    годовой_баланс_раб_времени WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm7 = new MySqlCommand(commText7, connection);
                try
                {
                    comm7.ExecuteNonQuery();
                    

                }
                catch
                {
                  
                }
                string commText8 = "DELETE FROM    оплата_труда WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm8 = new MySqlCommand(commText8, connection);
                try
                {
                    comm8.ExecuteNonQuery();
                  

                }
                catch
                {
                }
                string commText9 = "DELETE FROM    норма_времени WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm9 = new MySqlCommand(commText9, connection);
                try
                {
                    comm9.ExecuteNonQuery();
                    

                }
                catch
                {
                   
                }
                string commText10 = "DELETE FROM    нормативные_показатели WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm10 = new MySqlCommand(commText10, connection);
                try
                {
                    comm10.ExecuteNonQuery();
                   

                }
                catch
                {
             
                }
                string commText11 = "DELETE FROM    основные_и_вспомогательные_материалы WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm11 = new MySqlCommand(commText11, connection);
                try
                {
                    comm11.ExecuteNonQuery();
               

                }
                catch
                {
                }
                string commText12 = "DELETE FROM     conditionally_const_expenses WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm12 = new MySqlCommand(commText12, connection);
                try
                {
                    comm12.ExecuteNonQuery();
                   

                }
                catch
                {
                   
                }

                string commText13 = "DELETE FROM     conditional_variable_costs WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm13 = new MySqlCommand(commText13, connection);
                try
                {
                    comm13.ExecuteNonQuery();
                  

                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }

                string commText14 = "DELETE FROM     условное_колво_услуги WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm14 = new MySqlCommand(commText14, connection);
                try
                {
                    comm14.ExecuteNonQuery();
   

                }
                catch
                {
                   
                }


                string commText15 = "DELETE FROM     цена_по WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection);
                try
                {
                    comm15.ExecuteNonQuery();
                  

                }
                catch
                {
                    
                }

                string commText16 = "DELETE FROM     технико_экономич_показатели WHERE ID_проекта=\"" + index + "\"";
                MySqlCommand comm16 = new MySqlCommand(commText16, connection);
                try
                {
                    comm16.ExecuteNonQuery();
                    Form1 form1 = new Form1();
                    form1.Custom_Initialize2(login.ToString());
                    form1.Show();
                    this.Hide();

                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
            }

        }

        private void Del_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
    }
}
