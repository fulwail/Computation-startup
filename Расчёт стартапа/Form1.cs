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
using System.Security.Cryptography;
using ZedGraph;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using System.Drawing.Drawing2D;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;


namespace Расчёт_стартапа
{


    public partial class Form1 : Form
    {
        public static int DS_Count(string s)
        {
            string substr = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0].ToString();
            int count = (s.Length - s.Replace(substr, "").Length) / substr.Length;
            return count;

        }


        public static void lockword( KeyPressEventArgs e)
            {
                char l = e.KeyChar;

                if (l != '"' && l != '\\' && l != '/' && l != '`')
                { }
                else
                {
                    e.Handled = true;
                 }
            }
        public static void locknumb(object sender, KeyPressEventArgs e)
        {
            char l = e.KeyChar;
            e.Handled = l != 8 && !(Char.IsDigit(e.KeyChar) ||
                ((e.KeyChar == System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator[0]) &&
               (DS_Count(((TextBox)sender).Text) < 1)));
        }
        double rentab = 0;
         string login = "";
        string project = "none";
        string name_rab_time = "";
        double top = 0;
        double dr = 0;
        double dk = 0;
        double dv = 0;
        double dp = 0;
        double fmg = 0;
        double t = 0;
        double fvg = 0;
        double kv = 0;
        double fvm = 0;
        double drm = 0;
        double nvrsum = 0;
        double qyg = 0;
        double qym = 0;
        double qy = 0;
        double na = 0;
        double ag = 0;
        double topsum = 0;
        double tpzsum = 0;
        double tolnsum = 0;
        double tobssum = 0;
        string none = "";
        double st = 0;
        double stoimingod = 0;
        double stoimingod1 = 0;
        double stoiminmes = 0;
        double stoimin = 0;
        string itogistoim = "Итого на один программный продукт:";
        string itogistoimmes = "Итого за месяц:";
        string itogistoimgod = "Итого за год:";
        double otm = 0;
        double otg = 0;
        double ot1 = 0;
        double nvrsum1 = 0;
        double tpzsum1 = 0;
        double tolnsum1 = 0;
        double tobssum1 = 0;
        double topsum1 = 0;
        double rec = 0;
        double recmes = 0;
        double recgod = 0;
        double aggod = 0;
        double agmes = 0;
        double agin = 0;
        double proc = 0;
        double procmes = 0;
        double procgod = 0;
        double postmes = 0;
        double postgod = 0;
        double post = 0;
        double permes = 0;
        double pergod = 0;
        double per = 0;
        double yslyga = 0;
        double yslygames = 0;
        double yslygagod = 0;
        string nameras = "";
        double procent = 0;
        double razr = 0;
        double fmes = 0;
        double zp = 0;
        double zps = 0;
        double vir = 0;
        double tkb = 0;
        public void Custom_Initialize2(string id3)
        {
            login = id3;
        }

        public void Custom_Initialize3(string id)
        {
            project = id;
        }

        public Form1()
        {
            InitializeComponent();
        }
 
        private void выйтиИзПрограммыToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            z1.GraphPane.Title = "График точки безубыточности";
            GraphPane pane = z1.GraphPane;
            pane.CurveList.Clear();
            pane.YAxis.IsOmitMag = true;
            pane.XAxis.Min = 0;
            pane.XAxis.Max = 30;
            pane.YAxis.Min = 0;
            pane.YAxis.Max = 100000;
            pane.XAxis.Title = "";
            pane.YAxis.Title = "";
            pane.XAxis.IsOmitMag = true;
            pane.XAxis.ScaleMag =0;
            pane.YAxis.ScaleMag =-2 ;
            z1.AxisChange();
            z1.Invalidate();
            tabPage1.AutoScroll = true;
            tabPage2.AutoScroll = true;
            tabPage3.AutoScroll = true;
            tabPage4.AutoScroll = true;
            tabPage7.AutoScroll = true;
            tabPage8.AutoScroll = true;
            tabPage10.AutoScroll = true;
            tabPage12.AutoScroll = true;
            nvrsum = 0;
            tpzsum = 0;
            tolnsum = 0;
            tobssum = 0;
            topsum = 0;
            string connectionStrin = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connectio = new MySqlConnection(connectionStrin);
            connectio.Open();
            string comman = "Select Комплексная_работа From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd = new MySqlCommand(comman, connectio);
            string id = "";
            using (MySqlDataReader per = cmd.ExecuteReader())
            {
                while (per.Read())
                {
                    id = per.GetString(0);
                    name_rab_time = id;
                    textBox1.Text = name_rab_time;
                }
            }
            string comman2 = "Select 	норма_времени From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd2 = new MySqlCommand(comman2, connectio);
            double id2 = 0;
            using (MySqlDataReader per2 = cmd2.ExecuteReader())
            {
                while (per2.Read())
                {
                    string ids2 = per2.GetString(0);

                    id2 = Convert.ToDouble(ids2);
                    nvrsum = id2;
                }
            }
            string comman3 = "Select 	выходные_дни From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd3 = new MySqlCommand(comman3, connectio);
            double id3 = 0;
            using (MySqlDataReader per3 = cmd3.ExecuteReader())
            {
                while (per3.Read())
                {
                    string ids3 = per3.GetString(0);
                    id3 = Convert.ToDouble(ids3);
                    dv = id3;
                    textBox2.Text = Convert.ToString(dv);
                }
            }
            string comman4 = "Select 	праздничные_дни From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd4 = new MySqlCommand(comman4, connectio);
            double id4 = 0;
            using (MySqlDataReader per4 = cmd4.ExecuteReader())
            {
                while (per4.Read())
                {
                    string ids4 = per4.GetString(0);
                    id4 = Convert.ToDouble(ids4);
                    dp = id4;
                    textBox3.Text = Convert.ToString(dp);
                }
            }
            string comman5 = "Select 	продол_работы_день From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd5 = new MySqlCommand(comman5, connectio);
            double id5 = 0;
            using (MySqlDataReader per5 = cmd5.ExecuteReader())
            {
                while (per5.Read())
                {
                    string ids5 = per5.GetString(0);
                    id5 = Convert.ToDouble(ids5);
                    t = id5;
                    textBox4.Text = Convert.ToString(t);
                }
            }
            string comman6 = "Select 	Объём_услуг_год From условное_колво_услуги WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd6 = new MySqlCommand(comman6, connectio);
            double id6 = 0;
            using (MySqlDataReader per6 = cmd6.ExecuteReader())
            {
                while (per6.Read())
                {
                    string ids6 = per6.GetString(0);
                    id6 = Convert.ToDouble(ids6);
                    qyg = id6;
                }
            }
            string comman7 = "Select макс_фонд From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd7 = new MySqlCommand(comman7, connectio);
            double id7 = 0;
            using (MySqlDataReader per7 = cmd7.ExecuteReader())
            {
                while (per7.Read())
                {
                    string ids7 = per7.GetString(0);
                    id7 = Convert.ToDouble(ids7);
                    fmg = id7;
                }
            }
            string comman8 = "Select 	возм_фонд_времени_год From  нормативные_показатели WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd8 = new MySqlCommand(comman7, connectio);
            double id8 = 0;
            using (MySqlDataReader per8 = cmd8.ExecuteReader())
            {
                while (per8.Read())
                {
                    string ids8 = per8.GetString(0);
                    id8 = Convert.ToDouble(ids8);
                    fvg = id8;
                }
            }
            string comman9 = "Select 	Объём_услуг_месяц From условное_колво_услуги WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd9 = new MySqlCommand(comman9, connectio);
            double id9 = 0;
            using (MySqlDataReader per9 = cmd9.ExecuteReader())
            {
                while (per9.Read())
                {
                    string ids9 = per9.GetString(0);
                    id9 = Convert.ToDouble(ids9);
                    qym = id9;
                }
            }
            stoimingod = 0;
            stoiminmes = 0;
            stoimin = 0;
            string comman10 = "Select 	sum From основные_и_вспомогательные_материалы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoim + "\"";
            MySqlCommand cmd10 = new MySqlCommand(comman10, connectio);
            double id10 = 0;
            using (MySqlDataReader per10 = cmd10.ExecuteReader())
            {
                while (per10.Read())
                {
                    string ids10 = per10.GetString(0);
                    id10 = Convert.ToDouble(ids10);
                    stoimin = id10;
                }
            }
            string comman11 = "Select 	sum From основные_и_вспомогательные_материалы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimmes + "\"";
            MySqlCommand cmd11 = new MySqlCommand(comman11, connectio);
            double id11 = 0;
            using (MySqlDataReader per11 = cmd11.ExecuteReader())
            {
                while (per11.Read())
                {
                    string ids11 = per11.GetString(0);
                    id11 = Convert.ToDouble(ids11);
                    stoiminmes = id11;
                }
            }
            string comman12 = "Select 	sum From основные_и_вспомогательные_материалы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimgod + "\"";
            MySqlCommand cmd12 = new MySqlCommand(comman12, connectio);
            double id12 = 0;
            using (MySqlDataReader per12 = cmd12.ExecuteReader())
            {
                while (per12.Read())
                {
                    string ids12 = per12.GetString(0);
                    id12 = Convert.ToDouble(ids12);
                    stoimingod = id12;
                }
            }
            string comman121 = "Select 	sum From funds_and_depreciation  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimgod + "\"";
            MySqlCommand cmd121 = new MySqlCommand(comman121, connectio);
            double id121 = 0;
            using (MySqlDataReader per121 = cmd121.ExecuteReader())
            {
                while (per121.Read())
                {
                    string ids121 = per121.GetString(0);
                    id121 = Convert.ToDouble(ids121);
                    stoimingod1 = id121;
                }
            }
            string comman13 = "Select 	стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimgod + "\"";
            MySqlCommand cmd13 = new MySqlCommand(comman13, connectio);
            double id13 = 0;
            using (MySqlDataReader per13 = cmd13.ExecuteReader())
            {
                while (per13.Read())
                {
                    string ids13 = per13.GetString(0);
                    id13 = Convert.ToDouble(ids13);
                    recgod = id13;

                }
            }
            string comman14 = "Select 	стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimmes + "\"";
            MySqlCommand cmd14 = new MySqlCommand(comman13, connectio);
            double id14 = 0;
            using (MySqlDataReader per14 = cmd14.ExecuteReader())
            {
                while (per14.Read())
                {
                    string ids14 = per14.GetString(0);
                    id14 = Convert.ToDouble(ids14);
                    recmes = id14;
                }
            }
            string comman15 = "Select 	стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoim + "\"";
            MySqlCommand cmd15 = new MySqlCommand(comman15, connectio);
            double id15 = 0;
            using (MySqlDataReader per15 = cmd15.ExecuteReader())
            {
                while (per15.Read())
                {
                    string ids15 = per15.GetString(0);
                    id15 = Convert.ToDouble(ids15);
                    rec = id15;
                }
            }
            string comman16 = "Select 	depreciation_deductions_year From  funds_and_depreciation  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoim + "\"";
            MySqlCommand cmd16 = new MySqlCommand(comman16, connectio);
            double id16 = 0;
            using (MySqlDataReader per16 = cmd16.ExecuteReader())
            {
                while (per16.Read())
                {
                    string ids16 = per16.GetString(0);
                    id16 = Convert.ToDouble(ids16);
                    ag = id16;
                }
            }
            string comman17 = "Select 	depreciation_deductions_year From  funds_and_depreciation  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimmes + "\"";
            MySqlCommand cmd17 = new MySqlCommand(comman16, connectio);
            double id17 = 0;
            using (MySqlDataReader per17 = cmd17.ExecuteReader())
            {
                while (per17.Read())
                {
                    string ids17 = per17.GetString(0);
                    id17 = Convert.ToDouble(ids17);
                    agmes = id17;
                }
            }
            string comman18 = "Select 	depreciation_deductions_year From  funds_and_depreciation  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimgod + "\"";
            MySqlCommand cmd18 = new MySqlCommand(comman18, connectio);
            double id18 = 0;
            using (MySqlDataReader per18 = cmd18.ExecuteReader())
            {
                while (per18.Read())
                {
                    string ids18 = per18.GetString(0);
                    id18 = Convert.ToDouble(ids18);
                    aggod = id18;
                }
            }
            string itogopost = "Итого";
            string comman19 = "Select 	single_product From   conditionally_const_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogopost + "\"";
            MySqlCommand cmd19 = new MySqlCommand(comman19, connectio);
            double id19 = 0;
            using (MySqlDataReader per19 = cmd19.ExecuteReader())
            {
                while (per19.Read())
                {
                    string ids19 = per19.GetString(0);
                    id19 = Convert.ToDouble(ids19);
                    post = id19;
                }
            }
            string comman20 = "Select 	month  From   conditionally_const_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogopost + "\"";
            MySqlCommand cmd20 = new MySqlCommand(comman20, connectio);
            double id20 = 0;
            using (MySqlDataReader per20 = cmd20.ExecuteReader())
            {
                while (per20.Read())
                {
                    string ids20 = per20.GetString(0);
                    id20 = Convert.ToDouble(ids20);
                    postmes = id20;
                }
            }
            string comman21 = "Select 	year  From   conditionally_const_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogopost + "\"";
            MySqlCommand cmd21 = new MySqlCommand(comman21, connectio);
            double id21 = 0;
            using (MySqlDataReader per21 = cmd21.ExecuteReader())
            {
                while (per21.Read())
                {
                    string ids21 = per21.GetString(0);
                    id21 = Convert.ToDouble(ids21);
                    postgod = id21;
                }
            }
            string itogoper = "Итого";
            string comman22 = "Select 	single_product From    conditional_variable_costs  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogoper + "\"";
            MySqlCommand cmd22 = new MySqlCommand(comman22, connectio);
            double id22 = 0;
            using (MySqlDataReader per22 = cmd22.ExecuteReader())
            {
                while (per22.Read())
                {
                    string ids22 = per22.GetString(0);
                    id22 = Convert.ToDouble(ids22);
                    per = id22;
                }
            }
            string comman23 = "Select 	month From    conditional_variable_costs  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogoper + "\"";
            MySqlCommand cmd23 = new MySqlCommand(comman23, connectio);
            double id23 = 0;
            using (MySqlDataReader per23 = cmd23.ExecuteReader())
            {
                while (per23.Read())
                {
                    string ids23 = per23.GetString(0);
                    id23 = Convert.ToDouble(ids23);
                    permes = id23;
                }
            }
            string comman24 = "Select 	year From    conditional_variable_costs  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogoper + "\"";
            MySqlCommand cmd24 = new MySqlCommand(comman24, connectio);
            double id24 = 0;
            using (MySqlDataReader per24 = cmd24.ExecuteReader())
            {
                while (per24.Read())
                {
                    string ids24 = per24.GetString(0);
                    id24 = Convert.ToDouble(ids24);
                    pergod = id24;
                }
            }
            string comman25 = "Select 	сумма From other_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimgod + "\"";
            MySqlCommand cmd25 = new MySqlCommand(comman25, connectio);
            double id25 = 0;
            using (MySqlDataReader per25 = cmd25.ExecuteReader())
            {
                while (per25.Read())
                {
                    string ids25 = per25.GetString(0);
                    id25 = Convert.ToDouble(ids25);
                    procgod = id25;
                }
            }
            string comman26 = "Select 	сумма From other_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoimmes + "\"";
            MySqlCommand cmd26 = new MySqlCommand(comman26, connectio);
            double id26 = 0;
            using (MySqlDataReader per26 = cmd26.ExecuteReader())
            {
                while (per26.Read())
                {
                    string ids26 = per26.GetString(0);
                    id26 = Convert.ToDouble(ids26);
                    procmes = id26;
                }
            }
            string comman27 = "Select 	сумма From other_expenses  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itogistoim + "\"";
            MySqlCommand cmd27 = new MySqlCommand(comman27, connectio);
            double id27 = 0;
            using (MySqlDataReader per27 = cmd27.ExecuteReader())
            {
                while (per27.Read())
                {
                    string ids27 = per27.GetString(0);
                    id27 = Convert.ToDouble(ids27);
                    proc = id27;
                }
            }
            string itog = "Итого по услуге:";

            string comman28 = "Select 	service From costs  WHERE 	ID_project=\"" + project + "\"AND name=\"" + itog + "\"";
            MySqlCommand cmd28 = new MySqlCommand(comman28, connectio);
            double id28 = 0;
            using (MySqlDataReader per28 = cmd28.ExecuteReader())
            {
                while (per28.Read())
                {
                    string ids28 = per28.GetString(0);
                    id28 = Convert.ToDouble(ids28);
                    yslyga = id28;
                }
            }
            string comman30 = "Select name From  цена_по WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd30 = new MySqlCommand(comman30, connectio);
            string id30 = "";
            using (MySqlDataReader per30 = cmd.ExecuteReader())
            {
                while (per30.Read())
                {
                    id30 = per30.GetString(0);
                    nameras = id30;
                    textBox23.Text = nameras;

                }
            }
            string comman31 = "Select процент From  цена_по WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd31 = new MySqlCommand(comman31, connectio);
            using (MySqlDataReader per31 = cmd31.ExecuteReader())
            {
                while (per31.Read())
                {
                    string ids31 = per31.GetString(0);
                   
                    

                    textBox22.Text = ids31;
                }
            }
            string comman32 = "Select оплата_труда_мес From  оплата_труда WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd32 = new MySqlCommand(comman31, connectio);
           double id32 = 0;
            using (MySqlDataReader per32 = cmd32.ExecuteReader())
            {
                while (per32.Read())
                {
                    string ids32 = per32.GetString(0);
                    id32 = Convert.ToDouble(ids32);
                    otm = id32;


                }
            }
            string comman33 = "Select Выручка From  цена_по WHERE 	ID_project=\"" + project + "\"";
            MySqlCommand cmd33 = new MySqlCommand(comman33, connectio);
            double id33 = 0;
            using (MySqlDataReader per33 = cmd33.ExecuteReader())
            {
                while (per33.Read())
                {
                    string ids33 = per33.GetString(0);
                    id33 = Convert.ToDouble(ids33);
                    vir= id33;
                }
            }
            string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds = new DataSet();
            MySqlConnection connection = new MySqlConnection(connectionString);
            MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени From норма_времени WHERE 	ID_project=\"" + project + "\"", connection);

            dataAdapter.Fill(ds, "Name");
            dataGridView1.DataSource = ds.Tables["Name"];
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.Columns[0].HeaderText = "Наименование работ";
            dataGridView1.Columns[1].HeaderText = "Топ,час";
            dataGridView1.Columns[2].HeaderText = "Тпз,час";
            dataGridView1.Columns[3].HeaderText = "Тотл ,час";
            dataGridView1.Columns[4].HeaderText = "Тобс,час";
            dataGridView1.Columns[5].HeaderText = "Нвр,час";
            string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds2 = new DataSet();
            MySqlConnection connection2 = new MySqlConnection(connectionString2);
            MySqlDataAdapter dataAdapter2 = new MySqlDataAdapter("Select Комплексная_работа, 	Объём_услуг_день, 	Объём_услуг_месяц, 	Объём_услуг_год From условное_колво_услуги WHERE 	ID_project=\"" + project + "\"", connection2);
            dataAdapter2.Fill(ds2, "Name");
            dataGridView2.DataSource = ds2.Tables["Name"];
            dataGridView2.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView2.Columns[0].HeaderText = "Комплексная работа";
            dataGridView2.Columns[1].HeaderText = "Объём услуг в день";
            dataGridView2.Columns[2].HeaderText = "Объём услуг в месяц";
            dataGridView2.Columns[3].HeaderText = "Объём услуг в год";
            string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds3 = new DataSet();
            MySqlConnection connection3 = new MySqlConnection(connectionString3);
            MySqlDataAdapter dataAdapter3 = new MySqlDataAdapter("Select  name,amount, price,sum, lifetime, depreciation_rates_year,	depreciation_deductions_year From funds_and_depreciation WHERE 	ID_project=\"" + project + "\"", connection3);
            dataAdapter3.Fill(ds3, "Name");
            dataGridView3.DataSource = ds3.Tables["Name"];
            dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView3.Columns[0].HeaderText = "Наименование";
            dataGridView3.Columns[1].HeaderText = "Количество, шт";
            dataGridView3.Columns[2].HeaderText = "Цена, руб";
            dataGridView3.Columns[3].HeaderText = "Стоимость, руб";
            dataGridView3.Columns[4].HeaderText = "Срок службы, лет";
            dataGridView3.Columns[5].HeaderText = "Годовые нормы амортизации,%";
            dataGridView3.Columns[6].HeaderText = "Годовые амортизационные отчисления, руб";
            string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds12 = new DataSet();
            MySqlConnection connection12 = new MySqlConnection(connectionString12);
            MySqlDataAdapter dataAdapter12 = new MySqlDataAdapter("Select  name, Характеристика, норма_расход, колво_матер, price,sum From  основные_и_вспомогательные_материалы WHERE 	ID_project=\"" + project + "\"", connection12);
            dataAdapter12.Fill(ds12, "Name");
            dataGridView4.DataSource = ds12.Tables["Name"];
            dataGridView4.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView4.Columns[0].HeaderText = "Наименование";
            dataGridView4.Columns[1].HeaderText = "Характеристика";
            dataGridView4.Columns[2].HeaderText = "Норма расходов";
            dataGridView4.Columns[3].HeaderText = "Количество материала в упаковке";
            dataGridView4.Columns[4].HeaderText = "Цена";
            dataGridView4.Columns[5].HeaderText = "Стоимость";
            string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds13 = new DataSet();
            MySqlConnection connection13 = new MySqlConnection(connectionString13);
            MySqlDataAdapter dataAdapter13 = new MySqlDataAdapter("Select  name,amount From  annual_time_balance WHERE 	ID_project=\"" + project + "\"", connection13);
            dataAdapter13.Fill(ds13, "Name");
            dataGridView5.DataSource = ds13.Tables["Name"];
            dataGridView5.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView5.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView5.Columns[0].HeaderText = "Наименование";
            dataGridView5.Columns[1].HeaderText = "Количество";
            string connectionString24 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds24 = new DataSet();
            MySqlConnection connection24 = new MySqlConnection(connectionString24);
            MySqlDataAdapter dataAdapter24 = new MySqlDataAdapter("Select name,single_product, month, year From conditional_variable_costs WHERE 	ID_project=\"" + project + "\"", connection24);
            dataAdapter24.Fill(ds24, "Name");
            dataGridView6.DataSource = ds24.Tables["Name"];
            dataGridView6.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView6.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView6.Columns[0].HeaderText = "Наименование";
            dataGridView6.Columns[1].HeaderText = "Один программный продукт, руб";
            dataGridView6.Columns[2].HeaderText = "Месяц, руб";
            dataGridView6.Columns[3].HeaderText = "Год, руб";
            string connectionString25 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds25 = new DataSet();
            MySqlConnection connection25 = new MySqlConnection(connectionString25);
            MySqlDataAdapter dataAdapter25 = new MySqlDataAdapter("Select  name,amount, price, сумма From other_expenses WHERE 	ID_project=\"" + project + "\"", connection25);
            dataAdapter25.Fill(ds25, "Name");
            dataGridView7.DataSource = ds25.Tables["Name"];
            dataGridView7.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView7.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView7.Columns[0].HeaderText = "Наименование";
            dataGridView7.Columns[1].HeaderText = "Количество (Размер помещения)";
            dataGridView7.Columns[2].HeaderText = "Цена за единицу, руб";
            dataGridView7.Columns[3].HeaderText = "Сумма, руб";
            string connectionString26 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds26 = new DataSet();
            MySqlConnection connection26 = new MySqlConnection(connectionString26);
            MySqlDataAdapter dataAdapter26 = new MySqlDataAdapter("Select  name,amount, price, стоимость_месяц, стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"", connection26);
            dataAdapter26.Fill(ds26, "Name");
            dataGridView8.DataSource = ds26.Tables["Name"];
            dataGridView8.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView8.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView8.Columns[0].HeaderText = "Наименование";
            dataGridView8.Columns[1].HeaderText = "Количество показов, шт";
            dataGridView8.Columns[2].HeaderText = "Цена за один день (услугу), руб";
            dataGridView8.Columns[3].HeaderText = "Стоимость в месяц, руб";
            dataGridView8.Columns[4].HeaderText = "Стоимость в год, руб";
            string connectionString27 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds27 = new DataSet();
            MySqlConnection connection27 = new MySqlConnection(connectionString27);
            MySqlDataAdapter dataAdapter27 = new MySqlDataAdapter("Select name,single_product, month, year From conditionally_const_expenses WHERE 	ID_project=\"" + project + "\"", connection27);
            dataAdapter27.Fill(ds27, "Name");
            dataGridView9.DataSource = ds27.Tables["Name"];
            dataGridView9.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView9.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView9.Columns[0].HeaderText = "Наименование";
            dataGridView9.Columns[1].HeaderText = "Один программный продукт, руб";
            dataGridView9.Columns[2].HeaderText = "Месяц, руб";
            dataGridView9.Columns[3].HeaderText = "Год, руб";
            string connectionString28 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds28 = new DataSet();
            MySqlConnection connection28 = new MySqlConnection(connectionString28);
            MySqlDataAdapter dataAdapter28 = new MySqlDataAdapter("Select name, service, month, year From costs WHERE 	ID_project=\"" + project + "\"", connection28);
            dataAdapter28.Fill(ds28, "Name");
            dataGridView10.DataSource = ds28.Tables["Name"];
            dataGridView10.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView10.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView10.Columns[0].HeaderText = "Наименование расходов";
            dataGridView10.Columns[1].HeaderText = "Услуга, руб";
            dataGridView10.Columns[2].HeaderText = "Месяц, руб";
            dataGridView10.Columns[3].HeaderText = "Год, руб";
            string connectionString29= "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds29 = new DataSet();
            MySqlConnection connection29 = new MySqlConnection(connectionString29);
            MySqlDataAdapter dataAdapter29 = new MySqlDataAdapter("Select name, полная_себ, прибыль, цена_предприятия,  НДС, Цена_НДС,amount, Выручка From  цена_по WHERE 	ID_project=\"" + project + "\"", connection29);
            dataAdapter29.Fill(ds29, "Name");
            dataGridView11.DataSource = ds29.Tables["Name"];
            dataGridView11.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView11.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView11.Columns[0].HeaderText = "Услуга";
            dataGridView11.Columns[1].HeaderText = "Полная себестоимость";
            dataGridView11.Columns[2].HeaderText = "Прибыль";
            dataGridView11.Columns[3].HeaderText = "Цена предприятия";
            dataGridView11.Columns[4].HeaderText = "НДС";
            dataGridView11.Columns[5].HeaderText = "Цена с НДС";
            dataGridView11.Columns[6].HeaderText = "Кол-во услуг";
            dataGridView11.Columns[7].HeaderText = "Выручка без НДС";
            string connectionStrin30 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds30 = new DataSet();
            MySqlConnection connection30 = new MySqlConnection(connectionStrin30);
            MySqlDataAdapter dataAdapter30 = new MySqlDataAdapter("Select  name, Расчёт From  techniс_and_econom_indicators WHERE 	ID_project=\"" + project + "\"", connection30);
            dataAdapter30.Fill(ds30, "Name");
            dataGridView12.DataSource = ds30.Tables["Name"];
            dataGridView12.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView12.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView12.Columns[0].HeaderText = "Наименование";
            dataGridView12.Columns[1].HeaderText = "Расчёт";
        }
        private void создатьНовыйПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            project form1 = new project();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
        private void открытьПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openproject form1 = new openproject();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
        private void выйтиИзПрофиляToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void войтиПодСтарымПользователемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Login form1 = new Login();
            form1.Show();
            this.Hide();
        }
        private void создатьНовогоПользователяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Reg form1 = new Reg();
            form1.Show();
            this.Hide();
        }
        private void удалитьПроектToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Del form1 = new Del();
            form1.Custom_Initialize2(login.ToString());
            form1.Show();
            this.Hide();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox5.Text != "" && textBox6.Text != "")
            {
                nvrsum = 0;
                tpzsum = 0;
                tolnsum = 0;
                tobssum = 0;
                topsum = 0;
                string itogi = "Итого:";
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection15 = new MySqlConnection(connectionString15);
                connection15.Open();
                string commText15 = "DELETE FROM  норма_времени WHERE ID_project=\"" + project + "\" AND Наим_работ=\"" + itogi + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
                try
                {
                    comm15.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
                name_rab_time = textBox5.Text;
                top = Convert.ToDouble(textBox6.Text);
                double tpz = (top / 100) * 5;
                tpz = Math.Round(tpz, 2);
                double toln = (top / 100) * 2;
                toln = Math.Round(toln, 2);
                double tobs = (top / 100) * 5;
                tobs = Math.Round(tobs, 2);
                double nvr = top + tpz + toln + tobs;
                nvr = Math.Round(nvr, 2);
                string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection = new MySqlConnection(connectionString))
                {
                    string commText = "Insert Into норма_времени  ( ID_project, Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени ) VALUES( @ID_project, @Наим_работ, @Время_оп_работы, @Время_отдыха, @время_обсл, @подготов_время, @норма_времени )";
                    MySqlCommand comm = new MySqlCommand(commText, connection);
                    comm.Parameters.AddWithValue("@ID_project", project);
                    comm.Parameters.AddWithValue("@Наим_работ", name_rab_time);
                    comm.Parameters.AddWithValue("@Время_оп_работы", top);
                    comm.Parameters.AddWithValue("@Время_отдыха", tpz);
                    comm.Parameters.AddWithValue("@время_обсл", toln);
                    comm.Parameters.AddWithValue("@подготов_время", tobs);
                    comm.Parameters.AddWithValue("@норма_времени", nvr);
                    connection.Open();
                    try
                    {
                        comm.ExecuteNonQuery();
                        string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds = new DataSet();
                        MySqlConnection connection2 = new MySqlConnection(connectionString2);
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени From норма_времени WHERE 	ID_project=\"" + project + "\"", connection2);

                        dataAdapter.Fill(ds, "Name");
                        dataGridView1.DataSource = ds.Tables["Name"];
                        dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView1.Columns[0].HeaderText = "Наименование работ";
                        dataGridView1.Columns[1].HeaderText = "Топ,час";
                        dataGridView1.Columns[2].HeaderText = "Тпз,час";
                        dataGridView1.Columns[3].HeaderText = "Тотл ,час";
                        dataGridView1.Columns[4].HeaderText = "Тобс,час";
                        dataGridView1.Columns[5].HeaderText = "Нвр,час";
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                    textBox6.Text = "";
                    textBox5.Text = "";
                }
            }

            else
            {
                MessageBox.Show("Вы заполнили не все данные");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (project != "none")
            {
                dataGridView1.DataSource = new object();
            dataGridView2.DataSource = new object();
            if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox1.Text != "")
            {
                    if (radioButton1.Checked == true)
                    { kv = 1.1; }
                    else { kv = 1.2; }
                string itogi = "Итого:";
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection15 = new MySqlConnection(connectionString15);
                connection15.Open();
                string commText15 = "DELETE FROM  норма_времени WHERE ID_project=\"" + project + "\" AND Наим_работ=\"" + itogi + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
                try
                {
                    comm15.ExecuteNonQuery();
                    dk = 365;
                    dv = Convert.ToInt16(textBox2.Text);
                    dp = Convert.ToInt16(textBox3.Text);
                    t = Convert.ToInt16(textBox4.Text);
                    dr = dk - (dv + dp);
                    fmg = dr * t;
                        fmg = Math.Round(fmg, 2);
                    fvg = fmg * kv;
                        fvg = Math.Round(fvg, 2);
                        fvm = fvg / 12;
                        fvm = Math.Round(fvm, 2);
                    drm = fvm / t;
                    drm = Math.Round(drm, 2);
                    nvrsum1 = 0;
                    tpzsum = 0;
                    tolnsum = 0;
                    tobssum = 0;
                    topsum = 0;
                    nvrsum = 0;
                    string connectionString01 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    DataSet ds2 = new DataSet();
                    MySqlConnection connection01 = new MySqlConnection(connectionString01);
                    MySqlDataAdapter dataAdapter2 = new MySqlDataAdapter("Select Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени From норма_времени WHERE 	ID_project=\"" + project + "\"", connection01);
                    dataAdapter2.Fill(ds2, "Name");
                    dataGridView1.DataSource = ds2.Tables["Name"];
                    dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    dataGridView1.Columns[0].HeaderText = "Наименование работ";
                    dataGridView1.Columns[1].HeaderText = "Топ,час";
                    dataGridView1.Columns[2].HeaderText = "Тпз,час";
                    dataGridView1.Columns[3].HeaderText = "Тотл ,час";
                    dataGridView1.Columns[4].HeaderText = "Тобс,час";
                    dataGridView1.Columns[5].HeaderText = "Нвр,час";
                    foreach (DataGridViewRow row0 in dataGridView1.Rows)
                    {
                        double incom;
                        double.TryParse((row0.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                        nvrsum += incom;
                        nvrsum = Math.Round(nvrsum, 2);
                    }
                    nvrsum1 = nvrsum;
                    tpzsum1 = 0;
                    foreach (DataGridViewRow row2 in dataGridView1.Rows)
                    {
                        double incom2;
                        double.TryParse((row2.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom2);
                        tpzsum += incom2;
                        tpzsum = Math.Round(tpzsum, 2);
                    }
                    tpzsum1 = tpzsum;
                    tpzsum1 = Math.Round(top, 2);
                    tolnsum1 = 0;
                    foreach (DataGridViewRow row3 in dataGridView1.Rows)
                    {
                        double incom3;
                        double.TryParse((row3.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom3);
                        tolnsum += incom3;
                        tolnsum = Math.Round(tolnsum, 2);
                    }
                    tolnsum1 = tolnsum;
                    tobssum1 = 0;
                    foreach (DataGridViewRow row4 in dataGridView1.Rows)
                    {
                        double incom4;
                        double.TryParse((row4.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom4);
                        tobssum += incom4;
                        tobssum = Math.Round(tobssum, 2);
                    }
                    tobssum1 = tobssum;
                    topsum1 = 0;
                    foreach (DataGridViewRow row5 in dataGridView1.Rows)
                    {
                        double incom5;
                        double.TryParse((row5.Cells[1].Value ?? "0").ToString().Replace(".", ","), out incom5);
                        topsum += incom5;
                        topsum = Math.Round(topsum, 2);
                    }
                    topsum1 = topsum;
                    string connectionString11 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection11 = new MySqlConnection(connectionString11))
                    {
                        string commText11 = "Insert Into норма_времени  ( ID_project, Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени ) VALUES( @ID_project, @Наим_работ, @Время_оп_работы, @Время_отдыха, @время_обсл, @подготов_время, @норма_времени )";
                        MySqlCommand comm11 = new MySqlCommand(commText11, connection11);
                        comm11.Parameters.AddWithValue("@ID_project", project);
                        comm11.Parameters.AddWithValue("@Наим_работ", itogi);
                        comm11.Parameters.AddWithValue("@Время_оп_работы", topsum);
                        comm11.Parameters.AddWithValue("@Время_отдыха", tpzsum);
                        comm11.Parameters.AddWithValue("@время_обсл", tolnsum);
                        comm11.Parameters.AddWithValue("@подготов_время", tobssum);
                        comm11.Parameters.AddWithValue("@норма_времени", nvrsum);
                        connection11.Open();
                        try
                        {
                            comm11.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                        string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds = new DataSet();
                        MySqlConnection connection2 = new MySqlConnection(connectionString2);
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select Наим_работ, Время_оп_работы, Время_отдыха, время_обсл, подготов_время, норма_времени From норма_времени WHERE 	ID_project=\"" + project + "\"", connection2);
                        dataAdapter.Fill(ds, "Name");
                        dataGridView1.DataSource = ds.Tables["Name"];
                        dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView1.Columns[0].HeaderText = "Наименование работ";
                        dataGridView1.Columns[1].HeaderText = "Топ,час";
                        dataGridView1.Columns[2].HeaderText = "Тпз,час";
                        dataGridView1.Columns[3].HeaderText = "Тотл ,час";
                        dataGridView1.Columns[4].HeaderText = "Тобс,час";
                        dataGridView1.Columns[5].HeaderText = "Нвр,час";
                    }
                    qyg = fvg / nvrsum;
                    qyg = Math.Round(qyg, 2);
                    qym = fvm / nvrsum;
                    qym = Math.Round(qym, 2);
                    qy = qyg / dr;
                    qy = Math.Round(qy, 2);
                        textBox24.Text += "Расчёт нормативных показателей:" + Environment.NewLine + Environment.NewLine;
                        textBox24.Text += "dr=" + dk + "-(" + dv + "+" + dp + ")" + "=" + dr + " – количество рабочих дней в году;" + Environment.NewLine;
                        textBox24.Text +="fmg=" + dr + "*" + t + "=" + fmg + " – максимально возможный фонд времени работы;" + Environment.NewLine;
                        textBox24.Text += "fvg=" + fmg + "*" + kv + "=" + fvg + " – возможный годовой фонд времени;" + Environment.NewLine;
                        textBox24.Text += "fvm=" + fvg + "/" + 12 + "=" + fvm + " – возможный месячный фонд времени работы;" + Environment.NewLine;
                        textBox24.Text += "drm = " + fvm + " / " + t + "=" + drm + " – среднемесячное количество рабочих дней;" + Environment.NewLine;
                        textBox24.Text += "Qyg=" + fvg + "/" + nvrsum + "=" + qyg + " – годовое количество работ;" + Environment.NewLine;
                        textBox24.Text += "qym=" + fvm + "/" + nvrsum + "=" + qym + " – количество услуг за месяц;" + Environment.NewLine;
                        textBox24.Text += "qy=" + qyg + "/" + dr + "=" + qy + " – количество услуг за месяц;" + Environment.NewLine + Environment.NewLine;
                        string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    MySqlConnection connection = new MySqlConnection(connectionString);
                    connection.Open();
                    string commText2 = "DELETE FROM  условное_колво_услуги WHERE ID_project=\"" + project + "\"";
                    MySqlCommand comm2 = new MySqlCommand(commText2, connection);
                    try
                    {
                        comm2.ExecuteNonQuery();
                    }
                    catch
                    {
                        MessageBox.Show("Удалить проект не удалось");
                    }
                    string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection13 = new MySqlConnection(connectionString13))
                    {
                        string commText13 = "Insert Into условное_колво_услуги  ( ID_project, Комплексная_работа, Объём_услуг_день, Объём_услуг_месяц, Объём_услуг_год) VALUES( @ID_project, @Комплексная_работа, @Объём_услуг_день, @Объём_услуг_месяц, @Объём_услуг_год )";
                        MySqlCommand comm13 = new MySqlCommand(commText13, connection13);
                        comm13.Parameters.AddWithValue("@ID_project", project);
                        comm13.Parameters.AddWithValue("@Комплексная_работа", textBox1.Text);
                        comm13.Parameters.AddWithValue("@Объём_услуг_день", qy);
                        comm13.Parameters.AddWithValue("@Объём_услуг_месяц", qym);
                        comm13.Parameters.AddWithValue("@Объём_услуг_год", qyg);
                        connection13.Open();
                        try
                        {
                            comm13.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                        string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds = new DataSet();
                        MySqlConnection connection3 = new MySqlConnection(connectionString3);
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select Комплексная_работа, 	Объём_услуг_день, 	Объём_услуг_месяц, 	Объём_услуг_год From условное_колво_услуги WHERE 	ID_project=\"" + project + "\"", connection3);
                        dataAdapter.Fill(ds, "Name");
                        dataGridView2.DataSource = ds.Tables["Name"];
                        dataGridView2.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView2.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView2.Columns[0].HeaderText = "Комплексная работа";
                        dataGridView2.Columns[1].HeaderText = "Объём услуг в день";
                        dataGridView2.Columns[2].HeaderText = "Объём услуг в месяц";
                        dataGridView2.Columns[3].HeaderText = "Объём услуг в год";
                        string connectionString4 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        MySqlConnection connection4 = new MySqlConnection(connectionString4);
                        connection4.Open();
                        string commText4 = "DELETE FROM  нормативные_показатели WHERE ID_project=\"" + project + "\"";
                        MySqlCommand comm4 = new MySqlCommand(commText4, connection4);
                        try
                        {
                            comm4.ExecuteNonQuery();

                        }
                        catch
                        {
                            MessageBox.Show("Удалить проект не удалось");
                        }
                        string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
                        {

                            string commText5 = "Insert Into нормативные_показатели  ( ID_project, Комплексная_работа, норма_времени,продол_работы_день,возм_фонд_времени_год, выходные_дни, праздничные_дни, макс_фонд) VALUES( @ID_project, @Комплексная_работа, @норма_времени,@продол_работы_день, @возм_фонд_времени_год, @выходные_дни, @праздничные_дни, @макс_фонд )";
                            MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                            comm5.Parameters.AddWithValue("@ID_project", project);
                            comm5.Parameters.AddWithValue("@Комплексная_работа", textBox1.Text);
                            comm5.Parameters.AddWithValue("@норма_времени", nvrsum);
                            comm5.Parameters.AddWithValue("@продол_работы_день", t);
                            comm5.Parameters.AddWithValue("@возм_фонд_времени_год", fvg);
                            comm5.Parameters.AddWithValue("@выходные_дни", dv);
                            comm5.Parameters.AddWithValue("@праздничные_дни", dp);
                            comm5.Parameters.AddWithValue("@макс_фонд", fmg);
                            connection5.Open();
                            try
                            {
                                comm5.ExecuteNonQuery();
                            }
                            catch
                            {
                                MessageBox.Show("Вы не выбрали проект4");
                            }
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
            }
            else
            {
                MessageBox.Show("Вы заполнили не все данные");
            }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (textBox9.Text != "" && textBox10.Text != "" && textBox11.Text != "" && textBox7.Text != "" )
            {
                string itogistoim = "Итого на один программный продукт:";
                string itogistoimmes = "Итого за месяц:";
                string itogistoimgod = "Итого за год:";
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection15 = new MySqlConnection(connectionString15);
                connection15.Open();
                string commText15 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
                try
                {
                    comm15.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }

                string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection16 = new MySqlConnection(connectionString16);
                connection16.Open();
                string commText16 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
                MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
                try
                {
                    comm16.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
                string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection17 = new MySqlConnection(connectionString17);
                connection17.Open();
                string commText17 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
                MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
                try
                {
                    comm17.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }

                double srok_slygb = Convert.ToDouble(textBox7.Text);
                double cena = Convert.ToDouble(textBox11.Text);
                double kolvo = Convert.ToDouble(textBox10.Text);
                double stoim = cena * kolvo;
                na = 100 / srok_slygb;
                na = Math.Round(na, 2);
                ag = (stoim * na) / 100;
                ag = Math.Round(ag, 2);

                string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
                {
                    string commText5 = "Insert Into funds_and_depreciation  ( ID_project, name,amount, price,sum, lifetime, depreciation_rates_year,	depreciation_deductions_year) VALUES(  @ID_project, @name, @колво, @price, @sum, @lifetime, @depreciation_rates_year,	@depreciation_deductions_year )";
                    MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                    comm5.Parameters.AddWithValue("@ID_project", project);
                    comm5.Parameters.AddWithValue("@name", textBox9.Text);
                    comm5.Parameters.AddWithValue("@колво", textBox10.Text);
                    comm5.Parameters.AddWithValue("@price", textBox11.Text);
                    comm5.Parameters.AddWithValue("@sum", stoim);
                    comm5.Parameters.AddWithValue("@lifetime", textBox7.Text);
                    comm5.Parameters.AddWithValue("@depreciation_rates_year", na);
                    comm5.Parameters.AddWithValue("@depreciation_deductions_year", ag);
                    connection5.Open();
                    try
                    {
                        comm5.ExecuteNonQuery();
                        string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds = new DataSet();
                        MySqlConnection connection3 = new MySqlConnection(connectionString3);
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name,amount, price,sum, lifetime, depreciation_rates_year,	depreciation_deductions_year From funds_and_depreciation WHERE 	ID_project=\"" + project + "\"", connection3);
                        dataAdapter.Fill(ds, "Name");
                        dataGridView3.DataSource = ds.Tables["Name"];
                        dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView3.Columns[0].HeaderText = "Наименование";
                        dataGridView3.Columns[1].HeaderText = "Количество, шт";
                        dataGridView3.Columns[2].HeaderText = "Цена, руб";
                        dataGridView3.Columns[3].HeaderText = "Стоимость, руб";
                        dataGridView3.Columns[4].HeaderText = "Срок службы, лет";
                        dataGridView3.Columns[5].HeaderText = "Годовые нормы амортизации,%";
                        dataGridView3.Columns[6].HeaderText = "Годовые амортизационные отчисления, руб";
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                    textBox7.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Вы заполнили не все данные");
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox8.Text != "" && textBox12.Text != "" && textBox13.Text != "" && textBox14.Text != "" && textBox15.Text != "")
            {
            double cena = Convert.ToDouble(textBox12.Text);
            double kolvo = Convert.ToDouble(textBox13.Text);
            double nr = Convert.ToDouble(textBox8.Text);
            double stoim = kolvo / nr;
            stoim = cena / stoim;
            stoim = Math.Round(stoim, 2);
            string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection16 = new MySqlConnection(connectionString16);
            connection16.Open();
            string commText16 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
            MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
            try
            {
                comm16.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection17 = new MySqlConnection(connectionString17);
            connection17.Open();
            string commText17 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
            MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
            try
            {
                comm17.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
            {
                string commText5 = "Insert Into  основные_и_вспомогательные_материалы  ( ID_project, name, Характеристика, норма_расход, колво_матер, price,sum ) VALUES(  @ID_project, @name,  @Характеристика, @норма_расход, @колво_матер, @price, @sum )";
                MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                comm5.Parameters.AddWithValue("@ID_project", project);
                comm5.Parameters.AddWithValue("@name", textBox14.Text);
                comm5.Parameters.AddWithValue("@Характеристика", textBox15.Text);
                comm5.Parameters.AddWithValue("@норма_расход", textBox8.Text);
                comm5.Parameters.AddWithValue("@колво_матер", textBox13.Text);
                comm5.Parameters.AddWithValue("@price", cena);
                comm5.Parameters.AddWithValue("@sum", stoim);
                connection5.Open();
                try
                {
                    comm5.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект4234");
                }
                string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                DataSet ds = new DataSet();
                MySqlConnection connection3 = new MySqlConnection(connectionString3);
                MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name, Характеристика, норма_расход, колво_матер, price,sum FROM основные_и_вспомогательные_материалы WHERE 	ID_project=\"" + project + "\"", connection3);

                dataAdapter.Fill(ds, "Name");
                dataGridView4.DataSource = ds.Tables["Name"];
                dataGridView4.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView4.Columns[0].HeaderText = "Наименование";
                dataGridView4.Columns[1].HeaderText = "Характеристика";
                dataGridView4.Columns[2].HeaderText = "Норма расход, шт";
                dataGridView4.Columns[3].HeaderText = "Количество материала в упаковке, ед";
                dataGridView4.Columns[4].HeaderText = "Цена одной упаковки, руб";
                dataGridView4.Columns[5].HeaderText = "Стоимость материала, руб";
                textBox14.Text = "";
                textBox12.Text = "";
                textBox15.Text = "";
                textBox13.Text = "";
                textBox8.Text = "";
              }
            }
            else
            {
                MessageBox.Show("Вы заполнили не все данные");
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            none = "";
            if (project != "none")
            {
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection15 = new MySqlConnection(connectionString15);
                connection15.Open();
                string commText15 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
                try
                {
                    comm15.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
                string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection16 = new MySqlConnection(connectionString16);
                connection16.Open();
                string commText16 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
                MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
                try
                {
                    comm16.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
                string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection17 = new MySqlConnection(connectionString17);
                connection17.Open();
                string commText17 = "DELETE FROM  funds_and_depreciation WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
                MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
                try
                {
                    comm17.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Удалить проект не удалось");
                }
                dataGridView3.DataSource = new object();
                stoimingod1 = 0;
                stoiminmes = 0;
                stoimin = 0;
                aggod = 0;
                agmes = 0;
                agin = 0;
                string connectionString03 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                DataSet ds03 = new DataSet();
                MySqlConnection connection03 = new MySqlConnection(connectionString03);
                MySqlDataAdapter dataAdapter03 = new MySqlDataAdapter("Select  name,amount, price,sum, lifetime, depreciation_rates_year,	depreciation_deductions_year From funds_and_depreciation WHERE 	ID_project=\"" + project + "\"", connection03);
                dataAdapter03.Fill(ds03, "Name");
                dataGridView3.DataSource = ds03.Tables["Name"];
                dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView3.Columns[0].HeaderText = "Наименование";
                dataGridView3.Columns[1].HeaderText = "Количество, шт";
                dataGridView3.Columns[2].HeaderText = "Цена, руб";
                dataGridView3.Columns[3].HeaderText = "Стоимость, руб";
                dataGridView3.Columns[4].HeaderText = "Срок службы, лет";
                dataGridView3.Columns[5].HeaderText = "Годовые нормы амортизации,%";
                dataGridView3.Columns[6].HeaderText = "Годовые амортизационные отчисления, руб";
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                    stoimingod1 += incom;
                    stoimingod1 = Math.Round(stoimingod1, 2);
                }
                stoiminmes = stoimingod1 / 12;
                stoiminmes = Math.Round(stoiminmes, 2);
                stoimin = stoimingod1 / qyg;
                stoimin = Math.Round(stoimin, 2);
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);
                    aggod += incom;
                    aggod = Math.Round(aggod, 2);
                }
                agmes = aggod / 12;
                agmes = Math.Round(agmes, 2);
                agin = aggod / qyg;
                agin = Math.Round(agin, 2);

                string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection12 = new MySqlConnection(connectionString12))
                {
                    string commText12 = "Insert Into funds_and_depreciation  ( ID_project, name,amount, price,sum, lifetime, depreciation_rates_year,	depreciation_deductions_year) VALUES(  @ID_project, @name, @колво, @price, @sum, @lifetime, @depreciation_rates_year,	@depreciation_deductions_year )";
                    MySqlCommand comm12 = new MySqlCommand(commText12, connection12);
                    comm12.Parameters.AddWithValue("@ID_project", project);
                    comm12.Parameters.AddWithValue("@name", itogistoimgod);
                    comm12.Parameters.AddWithValue("@колво", none);
                    comm12.Parameters.AddWithValue("@price", none);
                    comm12.Parameters.AddWithValue("@sum", stoimingod1);
                    comm12.Parameters.AddWithValue("@lifetime", none);
                    comm12.Parameters.AddWithValue("@depreciation_rates_year", none);
                    comm12.Parameters.AddWithValue("@depreciation_deductions_year", aggod);
                    connection12.Open();
                    try
                    {
                        comm12.ExecuteNonQuery();
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                    string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection13 = new MySqlConnection(connectionString13))
                    {
                        string commText13 = "Insert Into funds_and_depreciation  ( ID_project, name, amount, price, sum, lifetime, depreciation_rates_year,	depreciation_deductions_year) VALUES(  @ID_project, @name, @колво, @price, @sum, @lifetime, @depreciation_rates_year,	@depreciation_deductions_year )";
                        MySqlCommand comm13 = new MySqlCommand(commText13, connection13);
                        comm13.Parameters.AddWithValue("@ID_project", project);
                        comm13.Parameters.AddWithValue("@name", itogistoimmes);
                        comm13.Parameters.AddWithValue("@колво", none);
                        comm13.Parameters.AddWithValue("@price", none);
                        comm13.Parameters.AddWithValue("@sum", stoiminmes);
                        comm13.Parameters.AddWithValue("@lifetime", none);
                        comm13.Parameters.AddWithValue("@depreciation_rates_year", none);
                        comm13.Parameters.AddWithValue("@depreciation_deductions_year", agmes);
                        connection13.Open();
                        try
                        {
                            comm13.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                    }
                    string connectionString18 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection18 = new MySqlConnection(connectionString18))
                    {
                        string commText18 = "Insert Into funds_and_depreciation  ( ID_project, name, amount, price, sum, lifetime, depreciation_rates_year,	depreciation_deductions_year) VALUES(  @ID_project, @name, @колво, @price, @sum, @lifetime, @depreciation_rates_year,	@depreciation_deductions_year )";
                        MySqlCommand comm18 = new MySqlCommand(commText18, connection18);
                        comm18.Parameters.AddWithValue("@ID_project", project);
                        comm18.Parameters.AddWithValue("@name", itogistoim);
                        comm18.Parameters.AddWithValue("@колво", none);
                        comm18.Parameters.AddWithValue("@price", none);
                        comm18.Parameters.AddWithValue("@sum", stoimin);
                        comm18.Parameters.AddWithValue("@lifetime", none);
                        comm18.Parameters.AddWithValue("@depreciation_rates_year", none);
                        comm18.Parameters.AddWithValue("@depreciation_deductions_year", agin);
                        connection18.Open();
                        try
                        {
                            comm18.ExecuteNonQuery();
                            string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                            DataSet ds = new DataSet();
                            MySqlConnection connection3 = new MySqlConnection(connectionString3);
                            MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name, amount, price, sum, lifetime, depreciation_rates_year,	depreciation_deductions_year From funds_and_depreciation WHERE 	ID_project=\"" + project + "\"", connection3);
                            dataAdapter.Fill(ds, "Name");
                            dataGridView3.DataSource = ds.Tables["Name"];
                            dataGridView3.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                            dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            dataGridView3.Columns[0].HeaderText = "Наименование";
                            dataGridView3.Columns[1].HeaderText = "Количество, шт";
                            dataGridView3.Columns[2].HeaderText = "Цена, руб";
                            dataGridView3.Columns[3].HeaderText = "Стоимость, руб";
                            dataGridView3.Columns[4].HeaderText = "Срок службы, лет";
                            dataGridView3.Columns[5].HeaderText = "Годовые нормы амортизации,%";
                            dataGridView3.Columns[6].HeaderText = "Годовые амортизационные отчисления, руб";
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            none = "";
            if (project != "none")
            {
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection16 = new MySqlConnection(connectionString16);
            connection16.Open();
            string commText16 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
            MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
            try
            {
                comm16.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection17 = new MySqlConnection(connectionString17);
            connection17.Open();
            string commText17 = "DELETE FROM   основные_и_вспомогательные_материалы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
            MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
            try
            {
                comm17.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString11 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds11 = new DataSet();
            MySqlConnection connection11 = new MySqlConnection(connectionString11);
            MySqlDataAdapter dataAdapter11 = new MySqlDataAdapter("Select  name, Характеристика, норма_расход, колво_матер, price, sum From  основные_и_вспомогательные_материалы WHERE 	ID_project=\"" + project + "\"", connection11);
            dataAdapter11.Fill(ds11, "Name");
            dataGridView4.DataSource = ds11.Tables["Name"];
            dataGridView4.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView4.Columns[0].HeaderText = "Наименование";
            dataGridView4.Columns[1].HeaderText = "Характеристика";
            dataGridView4.Columns[2].HeaderText = "Норма расходов";
            dataGridView4.Columns[3].HeaderText = "Количество материала в упаковке";
            dataGridView4.Columns[4].HeaderText = "Цена";
            dataGridView4.Columns[5].HeaderText = "Стоимость";
            stoimingod = 0;
            stoiminmes = 0;
            stoimin = 0;
            foreach (DataGridViewRow row in dataGridView4.Rows)
            {
                double incom;
                double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                stoimin += incom;
                stoimin = Math.Round(stoimin, 2);
            }
            stoiminmes = stoimin*qym;
            stoiminmes = Math.Round(stoiminmes, 2);
            stoimingod = stoimin* qyg;
            stoimingod = Math.Round(stoimingod, 2);
            string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection12 = new MySqlConnection(connectionString12))
            {
                string commText12 = "Insert Into основные_и_вспомогательные_материалы  ( ID_project, name, Характеристика, норма_расход, колво_матер, price, sum) VALUES(  @ID_project, @name, @Характеристика, @норма_расход, @колво_матер, @price, @sum)";
                MySqlCommand comm12 = new MySqlCommand(commText12, connection12);
                comm12.Parameters.AddWithValue("@ID_project", project);
                comm12.Parameters.AddWithValue("@name", itogistoimgod);
                comm12.Parameters.AddWithValue("@Характеристика", none);
                comm12.Parameters.AddWithValue("@норма_расход", none);
                comm12.Parameters.AddWithValue("@колво_матер", none);
                comm12.Parameters.AddWithValue("@price", none);
                comm12.Parameters.AddWithValue("@sum", stoimingod);
                connection12.Open();
                try
                {
                    comm12.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
                string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection13 = new MySqlConnection(connectionString13))
                {
                    string commText13 = "Insert Into основные_и_вспомогательные_материалы  ( ID_project, name, Характеристика, норма_расход, колво_матер, price, sum) VALUES(  @ID_project, @name, @Характеристика, @норма_расход, @колво_матер, @price, @sum)";
                    MySqlCommand comm13 = new MySqlCommand(commText13, connection13);
                    comm13.Parameters.AddWithValue("@ID_project", project);
                    comm13.Parameters.AddWithValue("@name", itogistoimmes);
                    comm13.Parameters.AddWithValue("@Характеристика", none);
                    comm13.Parameters.AddWithValue("@норма_расход", none);
                    comm13.Parameters.AddWithValue("@колво_матер", none);
                    comm13.Parameters.AddWithValue("@price", none);
                    comm13.Parameters.AddWithValue("@sum", stoiminmes);
                    connection13.Open();
                    try
                    {
                        comm13.ExecuteNonQuery();
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                }
                string connectionString18 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection18 = new MySqlConnection(connectionString18))
                {
                    string commText18 = "Insert Into основные_и_вспомогательные_материалы  ( ID_project, name, Характеристика, норма_расход, колво_матер, price, sum) VALUES(  @ID_project, @name, @Характеристика, @норма_расход, @колво_матер, @price, @sum)";
                    MySqlCommand comm18 = new MySqlCommand(commText18, connection18);
                    comm18.Parameters.AddWithValue("@ID_project", project);
                    comm18.Parameters.AddWithValue("@name", itogistoim);
                    comm18.Parameters.AddWithValue("@Характеристика", none);
                    comm18.Parameters.AddWithValue("@норма_расход", none);
                    comm18.Parameters.AddWithValue("@колво_матер", none);
                    comm18.Parameters.AddWithValue("@price", none);
                    comm18.Parameters.AddWithValue("@sum", stoimin);
                    connection18.Open();
                    try
                    {
                        comm18.ExecuteNonQuery();
                        string connectionString21 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds21 = new DataSet();
                        MySqlConnection connection21 = new MySqlConnection(connectionString21);
                        MySqlDataAdapter dataAdapter21 = new MySqlDataAdapter("Select  name, Характеристика, норма_расход, колво_матер, price, sum From  основные_и_вспомогательные_материалы WHERE 	ID_project=\"" + project + "\"", connection21);
                        dataAdapter21.Fill(ds21, "Name");
                        dataGridView4.DataSource = ds21.Tables["Name"];
                        dataGridView4.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView4.Columns[0].HeaderText = "Наименование";
                        dataGridView4.Columns[1].HeaderText = "Характеристика";
                        dataGridView4.Columns[2].HeaderText = "Норма расходов";
                        dataGridView4.Columns[3].HeaderText = "Количество материала в упаковке";
                        dataGridView4.Columns[4].HeaderText = "Цена";
                        dataGridView4.Columns[5].HeaderText = "Стоимость";
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                }
            }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            if (project != "none")
            {
                double p1 = 365;
                double p2 = dv + dp;
                double p3 = p1 - p2;
                double p41 = 28;
                double p42 = p3 / 100;
                p42 = Math.Round(p42, 2);
                double p43 = p3 / 100;
                p43 = Math.Round(p43, 2);
                double p4 = p41 + p42 + p43;
                p4 = Math.Round(p4, 2);
                double p5 = p3 - p4;
                p5 = Math.Round(p5, 2);
                double p6 = t ;
                double p7 = 0.2;
                double p8 = p6 - p7;
                p8 = Math.Round(p8, 2);
                double p9 = p5 * p8;
                p9 = Math.Round(p9, 2);
                double p10 = p9 / 12;
                p10 = Math.Round(p10, 2);
                double Fmg = p10 * 12;
                Fmg = Math.Round(Fmg, 2);
                double cr = qyg * nvrsum / fmg;
                cr = Math.Round(cr, 0);
                label18.Text = "Численность персонала= " + cr;
            string connectionStrin = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connectio = new MySqlConnection(connectionStrin);
            connectio.Open();
            string comText = "DELETE FROM   оплата_труда WHERE ID_project=\"" + project + "\"";
            MySqlCommand com = new MySqlCommand(comText, connectio);
            try
            {
                com.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            string commText = "DELETE FROM   annual_time_balance  WHERE ID_project=\"" + project + "\"";
            MySqlCommand comm = new MySqlCommand(commText, connection);
            try
            {
                comm.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionStri = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connecti = new MySqlConnection(connectionStri);
            connecti.Open();
            string comTex = "DELETE FROM   conditional_variable_costs WHERE ID_project=\"" + project + "\"";
            MySqlCommand co = new MySqlCommand(comTex, connecti);
            try
            {
                co.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString1 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection1 = new MySqlConnection(connectionString1))
            {
                string commText1 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection1);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "1. Календарный фонд рабочего времени");
                comm1.Parameters.AddWithValue("@колво", p1);
                connection1.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection2 = new MySqlConnection(connectionString2))
            {
                string commText2 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm2 = new MySqlCommand(commText2, connection2);
                comm2.Parameters.AddWithValue("@ID_project", project);
                comm2.Parameters.AddWithValue("@name", "2. Выходные и праздничные дни");
                comm2.Parameters.AddWithValue("@колво", p2);
                connection2.Open();
                try
                {
                    comm2.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }

            string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection3 = new MySqlConnection(connectionString3))
            {
                string commText3 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm3 = new MySqlCommand(commText3, connection3);
                comm3.Parameters.AddWithValue("@ID_project", project);
                comm3.Parameters.AddWithValue("@name", "3. Рабочие дни");
                comm3.Parameters.AddWithValue("@колво", p3);
                connection3.Open();
                try
                {
                    comm3.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString4 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection4 = new MySqlConnection(connectionString4))
            {
                string commText4 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm4 = new MySqlCommand(commText4, connection4);
                comm4.Parameters.AddWithValue("@ID_project", project);
                comm4.Parameters.AddWithValue("@name", "4.  Плановые невыходы на работу");
                comm4.Parameters.AddWithValue("@колво", p4);
                connection4.Open();
                try
                {
                    comm4.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString41 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection41 = new MySqlConnection(connectionString41))
            {
                string commText41 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm41 = new MySqlCommand(commText41, connection41);
                comm41.Parameters.AddWithValue("@ID_project", project);
                comm41.Parameters.AddWithValue("@name", "4.1. Отпуска");
                comm41.Parameters.AddWithValue("@колво", p41);
                connection41.Open();
                try
                {
                    comm41.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString42 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection42 = new MySqlConnection(connectionString42))
            {
                string commText42 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm42 = new MySqlCommand(commText42, connection42);
                comm42.Parameters.AddWithValue("@ID_project", project);
                comm42.Parameters.AddWithValue("@name", "4.2. По болезни");
                comm42.Parameters.AddWithValue("@колво", p42);
                connection42.Open();
                try
                {
                    comm42.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString43 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection43 = new MySqlConnection(connectionString43))
            {
                string commText43 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm43 = new MySqlCommand(commText43, connection43);
                comm43.Parameters.AddWithValue("@ID_project", project);
                comm43.Parameters.AddWithValue("@name", "4.3. Прочие невыходы (учебные отпуска, выполнение административных обязанностей)");
                comm43.Parameters.AddWithValue("@колво", p43);
                connection43.Open();
                try
                {
                    comm43.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
            {
                string commText5 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                comm5.Parameters.AddWithValue("@ID_project", project);
                comm5.Parameters.AddWithValue("@name", "5. Плановый фонд рабочего времени");
                comm5.Parameters.AddWithValue("@колво", p5);
                connection5.Open();
                try
                {
                    comm5.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString6 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString6))
            {
                string commText6 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm6 = new MySqlCommand(commText6, connection6);
                comm6.Parameters.AddWithValue("@ID_project", project);
                comm6.Parameters.AddWithValue("@name", "6. Номинальная продолжительность рабочего дня");
                comm6.Parameters.AddWithValue("@колво", p6);
                connection6.Open();
                try
                {
                    comm6.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString7 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection7 = new MySqlConnection(connectionString7))
            {
                string commText7 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm7 = new MySqlCommand(commText7, connection7);
                comm7.Parameters.AddWithValue("@ID_project", project);
                comm7.Parameters.AddWithValue("@name", "7. Плановые сокращения рабочего дня ");
                comm7.Parameters.AddWithValue("@колво", p7);
                connection7.Open();
                try
                {
                    comm7.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString8 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection8 = new MySqlConnection(connectionString8))
            {
                string commText8 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm8 = new MySqlCommand(commText8, connection8);
                comm8.Parameters.AddWithValue("@ID_project", project);
                comm8.Parameters.AddWithValue("@name", "8. Плановая продолжительность рабочего времени");
                comm8.Parameters.AddWithValue("@колво", p8);
                connection8.Open();
                try
                {
                    comm8.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString9 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection9 = new MySqlConnection(connectionString9))
            {
                string commText9 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm9 = new MySqlCommand(commText9, connection9);
                comm9.Parameters.AddWithValue("@ID_project", project);
                comm9.Parameters.AddWithValue("@name", "9. Плановая продолжительность рабочего времени за год");
                comm9.Parameters.AddWithValue("@колво", p9);
                connection9.Open();
                try
                {
                    comm9.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString10 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection10 = new MySqlConnection(connectionString10))
            {
                string commText10 = "Insert Into annual_time_balance  ( ID_project, name,amount) VALUES(  @ID_project, @name, @колво)";
                MySqlCommand comm10 = new MySqlCommand(commText10, connection10);
                comm10.Parameters.AddWithValue("@ID_project", project);
                comm10.Parameters.AddWithValue("@name", "10. Месячный фонд времени");
                comm10.Parameters.AddWithValue("@колво", p10);
                connection10.Open();
                try
                {
                    comm10.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds13 = new DataSet();
            MySqlConnection connection13 = new MySqlConnection(connectionString13);
            MySqlDataAdapter dataAdapter13 = new MySqlDataAdapter("Select  name,amount From  annual_time_balance WHERE 	ID_project=\"" + project + "\"", connection13);
            dataAdapter13.Fill(ds13, "Name");
            dataGridView5.DataSource = ds13.Tables["Name"];
            dataGridView5.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView5.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView5.Columns[0].HeaderText = "Наименование";
            dataGridView5.Columns[1].HeaderText = "Количество";
            st = 0;
            if (comboBox1.SelectedIndex == 0)
            {
                st = 63.43;
            }
            if (comboBox1.SelectedIndex == 1)
            {
                st = 82.46;
            }
            if (comboBox1.SelectedIndex == 2)
            {
                st = 107.2;
            }
            if (comboBox1.SelectedIndex == 3)
            {
                st = 121.15;
            }

            if (comboBox1.SelectedIndex == 4)
            {
                st = 137;
            }
            if (comboBox1.SelectedIndex == 5)
            {
                st = 154.77;
            }
            if (comboBox1.SelectedIndex == 6)
            {
                st = 175.07;
            }
            if (comboBox1.SelectedIndex == 7)
            {
                st = 197.9;
            }

            if (comboBox1.SelectedIndex == 8)
            {
                st = 222;
            }

            if (comboBox1.SelectedIndex == 9)
            {
                st = 253.09;
            }
            if (comboBox1.SelectedIndex == 10)
            {
                st = 286.07;
            }
            if (comboBox1.SelectedIndex == 11)
            {
                st = 323.5;
            }
            if (comboBox1.SelectedIndex == 12)
            {
                st = 363.36;
            }
            if (comboBox1.SelectedIndex == 13)
            {
                st = 412.93;
            }
            if (comboBox1.SelectedIndex == 14)
            {
                st = 466.84;
            }
            if (comboBox1.SelectedIndex == 15)
            {
                st = 518.22;
            }
            if (comboBox1.SelectedIndex == 16)
            {
                st = 575.31;
            }
            if (comboBox1.SelectedIndex == 16)
            {
                st = 638.74;
            }
            double r1 = 15;
            double d1 = Convert.ToDouble(textBox25.Text);
            double p11 = 30;
            double r = st * nvrsum;
            r = Math.Round(r, 2);
            double p = (r * p11)/100;
            p = Math.Round(p, 2);
            double zm = r + p;
            double zdm = (d1 * zm )/ 100;
            zdm = Math.Round(zdm, 2);
            double zcm = zm + zdm;
            double rk = (zcm * r1) / 100;
            rk = Math.Round(rk, 2);
            double zcm1 = zcm + rk;
            double sv = (30 * zcm1) / 100;
            sv = Math.Round(sv, 2);
            ot1 = sv+ zcm1;
            ot1 = Math.Round(ot1, 2);
            otg= ot1 * qyg;
            otg = Math.Round(otg, 2);
            otm = ot1 * qym;
            otm = Math.Round(otm, 2);
            permes = otm + stoiminmes;
            permes = Math.Round(permes, 2);
            pergod = otg + stoimingod;
            pergod = Math.Round(pergod, 2);
            per = ot1 + stoimin;
            per = Math.Round(per, 2);
            stoimin = Math.Round(stoimin, 2);
            stoiminmes = Math.Round(stoiminmes, 2);
            stoimingod = Math.Round(stoimingod, 2);
            textBox24.Text += "Расчет оплаты труда:" + Environment.NewLine + Environment.NewLine;
            textBox24.Text += "R=" + st + "*" + nvrsum + "=" + r + " – расценка сдельной формы оплаты труда;" + Environment.NewLine;
            textBox24.Text += "П=" + p11 + "*" + r + "/100%=" + p + " – премия за качественное и досрочное выполнение заказа;" + Environment.NewLine;
            textBox24.Text += "З=" + r + "+" + p + "=" + zm + " – месячная заработная плата;" + Environment.NewLine;
            textBox24.Text += "Зд=" + d1 + "*" + zm + "/100%=" + zdm + " – дополнительная заработная плата;" + Environment.NewLine;
            textBox24.Text += "Зc=" + zm + "+" + zdm + "=" + zcm + " – сдельная заработная плата;" + Environment.NewLine;
            textBox24.Text += "Rk=" + zcm + "*15%/100%=" + rk + " – районный коэффициент;" + Environment.NewLine;
            textBox24.Text += "Зc1=" + zcm + "+" + rk + "=" + zcm1 + " – месячная заработная плата с учётом районного коэффициента;" + Environment.NewLine;
            textBox24.Text += "СВ=" + p11 + "%*" + zcm1 + "/100%=" + sv + " – страховые взносы;" + Environment.NewLine;
            textBox24.Text += "От=" + zcm1 + "+" + sv + "=" + ot1 + " – оплата труда персонала за месяц ;" + Environment.NewLine + Environment.NewLine;
            string connectionString22 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection22 = new MySqlConnection(connectionString22))
            {
                string commText22 = "Insert Into  оплата_труда ( ID_project, оплата_труда_мес, оплата_труда_год, оплата_труда_продукт) VALUES(  @ID_project,  @оплата_труда_мес, @оплата_труда_год, @оплата_труда_продукт)";
                MySqlCommand comm22 = new MySqlCommand(commText22, connection22);
                comm22.Parameters.AddWithValue("@ID_project", project);
                comm22.Parameters.AddWithValue("@оплата_труда_мес", otm);
                comm22.Parameters.AddWithValue("@оплата_труда_год", otg);
                comm22.Parameters.AddWithValue("@оплата_труда_продукт", ot1);
                connection22.Open();
                try
                {
                    comm22.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString20 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection20 = new MySqlConnection(connectionString20))
            {
                string commText20 = "Insert Into  conditional_variable_costs  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                MySqlCommand comm20 = new MySqlCommand(commText20, connection20);
                comm20.Parameters.AddWithValue("@ID_project", project);
                comm20.Parameters.AddWithValue("@name", "1. Основные и вспомогательные материалы");
                comm20.Parameters.AddWithValue("@single_product", stoimin);
                comm20.Parameters.AddWithValue("@month", stoiminmes);
                comm20.Parameters.AddWithValue("@year", stoimingod);
                connection20.Open();
                try
                {
                    comm20.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }

            }
            string connectionString21 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection21 = new MySqlConnection(connectionString21))
            {
                string commText21 = "Insert Into  conditional_variable_costs  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                MySqlCommand comm21 = new MySqlCommand(commText21, connection21);
                comm21.Parameters.AddWithValue("@ID_project", project);
                comm21.Parameters.AddWithValue("@name", "2. Оплата труда");
                comm21.Parameters.AddWithValue("@single_product", ot1);
                comm21.Parameters.AddWithValue("@month", otm);
                comm21.Parameters.AddWithValue("@year", otg);
                connection21.Open();
                try
                {
                    comm21.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
            }
            string connectionString23 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection23 = new MySqlConnection(connectionString23))
            {
                string commText23 = "Insert Into  conditional_variable_costs  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                MySqlCommand comm23 = new MySqlCommand(commText23, connection23);
                comm23.Parameters.AddWithValue("@ID_project", project);
                comm23.Parameters.AddWithValue("@name", "Итого");
                comm23.Parameters.AddWithValue("@single_product", per);
                comm23.Parameters.AddWithValue("@month", permes);
                comm23.Parameters.AddWithValue("@year", pergod);
                connection23.Open();
                try
                {
                    comm23.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
                string connectionString24 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                DataSet ds24 = new DataSet();
                MySqlConnection connection24 = new MySqlConnection(connectionString24);
                MySqlDataAdapter dataAdapter24 = new MySqlDataAdapter("Select name,single_product, month, year From conditional_variable_costs WHERE 	ID_project=\"" + project + "\"", connection24);
                dataAdapter24.Fill(ds24, "Name");
                dataGridView6.DataSource = ds24.Tables["Name"];
                dataGridView6.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView6.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView6.Columns[0].HeaderText = "Наименование";
                dataGridView6.Columns[1].HeaderText = "Один программный продукт, руб";
                dataGridView6.Columns[2].HeaderText = "Месяц, руб";
                dataGridView6.Columns[3].HeaderText = "Год, руб";
            }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            if (textBox16.Text != "" && textBox17.Text != "" && textBox18.Text != "" )
            {
                string itogistoim = "Итого на один программный продукт:";
                string itogistoimmes = "Итого за месяц:";
                string itogistoimgod = "Итого за год:";
                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection15 = new MySqlConnection(connectionString15);
                connection15.Open();
                string commText15 = "DELETE FROM  other_expenses WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
                MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
                try
                {
                  comm15.ExecuteNonQuery();
                }
                catch
                {
                  MessageBox.Show("Удалить проект не удалось");
                }

                string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection16 = new MySqlConnection(connectionString16);
                connection16.Open();
                string commText16 = "DELETE FROM  other_expenses WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
                MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
                try
                {
                  comm16.ExecuteNonQuery();
                }
                catch
                {
                  MessageBox.Show("Удалить проект не удалось");
                }
                string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                MySqlConnection connection17 = new MySqlConnection(connectionString17);
                connection17.Open();
                string commText17 = "DELETE FROM  other_expenses  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
                MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
                try
                {
                  comm17.ExecuteNonQuery();
                }
                catch
                {
                  MessageBox.Show("Удалить проект не удалось");
                }
                double cena = Convert.ToDouble(textBox16.Text);
                double kolvo = Convert.ToDouble(textBox17.Text);
                double stoim = cena * kolvo;
                stoim = Math.Round(stoim, 2);
                cena.ToString().Replace(",", ".");
                string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
                {
                  string commText5 = "Insert Into other_expenses   ( ID_project, name,amount, price, сумма) VALUES(  @ID_project, @name, @колво, @price, @сумма )";
                  MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                  comm5.Parameters.AddWithValue("@ID_project", project);
                  comm5.Parameters.AddWithValue("@name", textBox18.Text);
                  comm5.Parameters.AddWithValue("@колво", kolvo);
                  comm5.Parameters.AddWithValue("@price", cena);
                  comm5.Parameters.AddWithValue("@сумма", stoim);
                  connection5.Open();
                  try
                  {
                    comm5.ExecuteNonQuery();
                    string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    DataSet ds = new DataSet();
                    MySqlConnection connection3 = new MySqlConnection(connectionString3);
                    MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name,amount, price, сумма From other_expenses  WHERE 	ID_project=\"" + project + "\"", connection3);

                    dataAdapter.Fill(ds, "Name");
                    dataGridView7.DataSource = ds.Tables["Name"];
                    dataGridView7.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    dataGridView7.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    dataGridView7.Columns[0].HeaderText = "Наименование";
                    dataGridView7.Columns[1].HeaderText = "Количество (Размер помещения)";
                    dataGridView7.Columns[2].HeaderText = "Цена за единицу, руб";
                    dataGridView7.Columns[3].HeaderText = "Сумма, руб";
                  }
                  catch
                  {
                    MessageBox.Show("Вы не выбрали проект");
                  }
                  textBox16.Text = "";
                  textBox17.Text = "";
                  textBox18.Text = "";
              }
              }
              else
              {
                  MessageBox.Show("Вы заполнили не все данные");
             }
        }

        private void button8_Click_1(object sender, EventArgs e)
        {

            if (project != "none")
            {
                none = "";
            string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM  other_expenses WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection16 = new MySqlConnection(connectionString16);
            connection16.Open();
            string commText16 = "DELETE FROM  other_expenses WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
            MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
            try
            {
                comm16.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection17 = new MySqlConnection(connectionString17);
            connection17.Open();
            string commText17 = "DELETE FROM  other_expenses WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
            MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
            try
            {
                comm17.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            dataGridView7.DataSource = new object();
            proc = 0;
            procmes = 0;
            procgod = 0;
            string connectionString03 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds03 = new DataSet();
            MySqlConnection connection03 = new MySqlConnection(connectionString03);
            MySqlDataAdapter dataAdapter03 = new MySqlDataAdapter("Select  name,amount, price, сумма From other_expenses WHERE 	ID_project=\"" + project + "\"", connection03);
            dataAdapter03.Fill(ds03, "Name");
            dataGridView7.DataSource = ds03.Tables["Name"];
            dataGridView7.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView7.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView7.Columns[0].HeaderText = "Наименование";
            dataGridView7.Columns[1].HeaderText = "Количество (Размер помещения)";
            dataGridView7.Columns[2].HeaderText = "Цена за единицу, руб";
            dataGridView7.Columns[3].HeaderText = "Сумма, руб";
            foreach (DataGridViewRow row in dataGridView7.Rows)
            {
                double incom;
                double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                procmes += incom;
                procmes = Math.Round(procmes, 2);
            }
            procgod = procmes*12;
            procgod = Math.Round(procgod, 2);
            proc = procgod / qyg;
            proc = Math.Round(proc, 2);
            string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection12 = new MySqlConnection(connectionString12))
            {
                string commText12 = "Insert Into other_expenses  ( ID_project, name,amount, price, сумма) VALUES(  @ID_project, @name, @колво, @price, @сумма )";
                MySqlCommand comm12 = new MySqlCommand(commText12, connection12);
                comm12.Parameters.AddWithValue("@ID_project", project);
                comm12.Parameters.AddWithValue("@name", itogistoimgod);
                comm12.Parameters.AddWithValue("@колво", none);
                comm12.Parameters.AddWithValue("@price", none);
                comm12.Parameters.AddWithValue("@сумма", procgod);
                connection12.Open();
                try
                {
                    comm12.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }

                string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection13 = new MySqlConnection(connectionString13))
                {
                    string commText13 = "Insert Into other_expenses  ( ID_project, name,amount, price, сумма) VALUES(  @ID_project, @name, @колво, @price, @сумма )";
                    MySqlCommand comm13 = new MySqlCommand(commText13, connection13);
                    comm13.Parameters.AddWithValue("@ID_project", project);
                    comm13.Parameters.AddWithValue("@name", itogistoimmes);
                    comm13.Parameters.AddWithValue("@колво", none);
                    comm13.Parameters.AddWithValue("@price", none);
                    comm13.Parameters.AddWithValue("@сумма", procmes);
                    connection13.Open();
                    try
                    {
                        comm13.ExecuteNonQuery();
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                }
                string connectionString18 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection18 = new MySqlConnection(connectionString18))
                {
                    string commText18 = "Insert Into other_expenses  ( ID_project, name,amount, price, сумма) VALUES(  @ID_project, @name, @колво, @price, @сумма)";
                    MySqlCommand comm18 = new MySqlCommand(commText18, connection18);
                    comm18.Parameters.AddWithValue("@ID_project", project);
                    comm18.Parameters.AddWithValue("@name", itogistoim);
                    comm18.Parameters.AddWithValue("@колво", none);
                    comm18.Parameters.AddWithValue("@price", none);
                    comm18.Parameters.AddWithValue("@сумма", proc);
                    connection18.Open();
                    try
                    {
                        comm18.ExecuteNonQuery();
                        string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds = new DataSet();
                        MySqlConnection connection3 = new MySqlConnection(connectionString3);
                        MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name,amount, price, сумма From other_expenses WHERE 	ID_project=\"" + project + "\"", connection3);
                        dataAdapter.Fill(ds, "Name");
                        dataGridView7.DataSource = ds.Tables["Name"];
                        dataGridView7.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView7.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView7.Columns[0].HeaderText = "Наименование";
                        dataGridView7.Columns[1].HeaderText = "Количество (Размер помещения)";
                        dataGridView7.Columns[2].HeaderText = "Цена за единицу, руб";
                        dataGridView7.Columns[3].HeaderText = "Сумма, руб";
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проектёмаё");
                    }
                }
            }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox19.Text != "" && textBox20.Text != "" && textBox21.Text != "")
            {
                string itogistoim = "Итого на один программный продукт:";
            string itogistoimmes = "Итого за месяц:";
            string itogistoimgod = "Итого за год:";
            string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM  бюджет_рекламы WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection16 = new MySqlConnection(connectionString16);
            connection16.Open();
            string commText16 = "DELETE FROM  бюджет_рекламы WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
            MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
            try
            {
                comm16.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection17 = new MySqlConnection(connectionString17);
            connection17.Open();
            string commText17 = "DELETE FROM  бюджет_рекламы  WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
            MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
            try
            {
                comm17.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            double cena = Convert.ToDouble(textBox19.Text);
            double kolvo = Convert.ToDouble(textBox20.Text);
            double stoim = cena * kolvo;
            stoim = Math.Round(stoim, 2);
            double stoim1 = stoim * 12;
            stoim1 = Math.Round(stoim1, 2);
            cena.ToString().Replace(",", ".");
            string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
            {
                string commText5 = "Insert Into бюджет_рекламы   ( ID_project, name,amount, price, стоимость_месяц, стоимость_год) VALUES(  @ID_project, @name, @колво, @price, @стоимость_месяц, @стоимость_год )";
                MySqlCommand comm5 = new MySqlCommand(commText5, connection5);
                comm5.Parameters.AddWithValue("@ID_project", project);
                comm5.Parameters.AddWithValue("@name", textBox21.Text);
                comm5.Parameters.AddWithValue("@колво", textBox20.Text);
                comm5.Parameters.AddWithValue("@price", cena);
                comm5.Parameters.AddWithValue("@стоимость_месяц", stoim);
                comm5.Parameters.AddWithValue("@стоимость_год", stoim1);
                connection5.Open();
                try
                {
                    comm5.ExecuteNonQuery();
                    string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    DataSet ds = new DataSet();
                    MySqlConnection connection3 = new MySqlConnection(connectionString3);
                    MySqlDataAdapter dataAdapter = new MySqlDataAdapter("Select  name,amount, price, стоимость_месяц, стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"", connection3);

                    dataAdapter.Fill(ds, "Name");
                    dataGridView8.DataSource = ds.Tables["Name"];
                    dataGridView8.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    dataGridView8.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    dataGridView8.Columns[0].HeaderText = "Наименование";
                    dataGridView8.Columns[1].HeaderText = "Количество показов, шт";
                    dataGridView8.Columns[2].HeaderText = "Цена за один день (услугу), руб";
                    dataGridView8.Columns[3].HeaderText = "Стоимость в месяц, руб";
                    dataGridView8.Columns[4].HeaderText = "Стоимость в год, руб";
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
                    textBox19.Text = "";
                    textBox20.Text = "";
                    textBox21.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Вы заполнили не все данные");
            }
        }
        private void button11_Click(object sender, EventArgs e)
        {

            if (project != "none")
            {
                none = "";
            string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM  бюджет_рекламы WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoim + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString16 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection16 = new MySqlConnection(connectionString16);
            connection16.Open();
            string commText16 = "DELETE FROM  бюджет_рекламы WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimmes + "\"";
            MySqlCommand comm16 = new MySqlCommand(commText16, connection16);
            try
            {
                comm16.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString17 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection17 = new MySqlConnection(connectionString17);
            connection17.Open();
            string commText17 = "DELETE FROM  бюджет_рекламы WHERE ID_project=\"" + project + "\" AND name=\"" + itogistoimgod + "\"";
            MySqlCommand comm17 = new MySqlCommand(commText17, connection17);
            try
            {
                comm17.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionStri = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connecti = new MySqlConnection(connectionStri);
            connecti.Open();
            string comTex = "DELETE FROM   conditionally_const_expenses WHERE ID_project=\"" + project + "\"";
            MySqlCommand co = new MySqlCommand(comTex, connecti);
            try
            {
                co.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionStr = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connect = new MySqlConnection(connectionStr);
            connect.Open();
            string comTe = "DELETE FROM   costs WHERE ID_project=\"" + project + "\"";
            MySqlCommand coqw = new MySqlCommand(comTe, connect);
            try
            {
                coqw.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            dataGridView8.DataSource = new object();
            recgod = 0;
            recmes = 0;
            rec = 0;
            string connectionString26 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds26 = new DataSet();
            MySqlConnection connection26 = new MySqlConnection(connectionString26);
            MySqlDataAdapter dataAdapter26 = new MySqlDataAdapter("Select  name,amount, price, стоимость_месяц, стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"", connection26);
            dataAdapter26.Fill(ds26, "Name");
            dataGridView8.DataSource = ds26.Tables["Name"];
            dataGridView8.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView8.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView8.Columns[0].HeaderText = "Наименование";
            dataGridView8.Columns[1].HeaderText = "Количество показов, шт";
            dataGridView8.Columns[2].HeaderText = "Цена за один день (услугу), руб";
            dataGridView8.Columns[3].HeaderText = "Стоимость в месяц, руб";
            dataGridView8.Columns[4].HeaderText = "Стоимость в год, руб";
            foreach (DataGridViewRow row in dataGridView8.Rows)
            {
                double incom;
                double.TryParse((row.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom);
                recmes += incom;
                recmes = Math.Round(recmes, 2);
            }
            recgod = recmes* 12;
            recmes = Math.Round(recmes, 2);
            rec = recgod / qyg;
            rec = Math.Round(rec, 2);
            string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection12 = new MySqlConnection(connectionString12))
            {
                string commText12 = "Insert Into бюджет_рекламы  ( ID_project, name,amount, price, стоимость_месяц,  стоимость_год ) VALUES(  @ID_project, @name, @колво, @price, @стоимость_месяц,  @стоимость_год )";
                MySqlCommand comm12 = new MySqlCommand(commText12, connection12);
                comm12.Parameters.AddWithValue("@ID_project", project);
                comm12.Parameters.AddWithValue("@name", itogistoimgod);
                comm12.Parameters.AddWithValue("@колво", none);
                comm12.Parameters.AddWithValue("@price", recgod );
                comm12.Parameters.AddWithValue("@стоимость_месяц", none);
                comm12.Parameters.AddWithValue("@стоимость_год", none);
                connection12.Open();
                try
                {
                    comm12.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
                string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection13 = new MySqlConnection(connectionString13))
                {
                    string commText13 = "Insert Into бюджет_рекламы   ( ID_project, name,amount, price, стоимость_месяц,  стоимость_год ) VALUES(  @ID_project, @name, @колво, @price, @стоимость_месяц,  @стоимость_год )";
                    MySqlCommand comm13 = new MySqlCommand(commText13, connection13);
                    comm13.Parameters.AddWithValue("@ID_project", project);
                    comm13.Parameters.AddWithValue("@name", itogistoimmes);
                    comm13.Parameters.AddWithValue("@колво", none);
                    comm13.Parameters.AddWithValue("@price", recmes);
                    comm13.Parameters.AddWithValue("@стоимость_месяц", none);
                    comm13.Parameters.AddWithValue("@стоимость_год", none);
                    connection13.Open();
                    try
                    {
                        comm13.ExecuteNonQuery();
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проект");
                    }
                }
                string connectionString18 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                using (MySqlConnection connection18 = new MySqlConnection(connectionString18))
                {
                    string commText18 = "Insert Into бюджет_рекламы  ( ID_project, name,amount, price, стоимость_месяц,  стоимость_год ) VALUES(  @ID_project, @name, @колво, @price, @стоимость_месяц,  @стоимость_год )";
                    MySqlCommand comm18 = new MySqlCommand(commText18, connection18);
                    comm18.Parameters.AddWithValue("@ID_project", project);
                    comm18.Parameters.AddWithValue("@name", itogistoim);
                    comm18.Parameters.AddWithValue("@колво", none);
                    comm18.Parameters.AddWithValue("@price", rec );
                    comm18.Parameters.AddWithValue("@стоимость_месяц", none);
                    comm18.Parameters.AddWithValue("@стоимость_год", none);
                    connection18.Open();
                    try
                    {
                        comm18.ExecuteNonQuery();
                        string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds2 = new DataSet();
                        MySqlConnection connection2 = new MySqlConnection(connectionString2);
                        MySqlDataAdapter dataAdapter2 = new MySqlDataAdapter("Select  name,amount, price, стоимость_месяц, стоимость_год From бюджет_рекламы  WHERE 	ID_project=\"" + project + "\"", connection2);
                        dataAdapter2.Fill(ds2, "Name");
                        dataGridView8.DataSource = ds2.Tables["Name"];
                        dataGridView8.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView8.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView8.Columns[0].HeaderText = "name";
                        dataGridView8.Columns[1].HeaderText = "Количество показов, шт";
                        dataGridView8.Columns[2].HeaderText = "Цена за один день (услугу), руб";
                        dataGridView8.Columns[3].HeaderText = "Стоимость в месяц, руб";
                        dataGridView8.Columns[4].HeaderText = "Стоимость в год, руб";
                    }
                    catch
                    {
                        MessageBox.Show("Вы не выбрали проектёмаё");
                    }
                    postmes = agmes + procmes + recmes;
                    postmes = Math.Round(postmes, 2);
                    postgod = aggod + procgod + recgod;
                    postgod = Math.Round(postgod, 2);
                    post =  agin + proc + rec;
                    post = Math.Round(post, 2);
                    string connectionString20 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection20 = new MySqlConnection(connectionString20))
                    {
                        string commText20 = "Insert Into   conditionally_const_expenses  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                        MySqlCommand comm20 = new MySqlCommand(commText20, connection20);
                        comm20.Parameters.AddWithValue("@ID_project", project);
                        comm20.Parameters.AddWithValue("@name", "1. Амортизационные отчисления");
                        comm20.Parameters.AddWithValue("@single_product", agin);
                        comm20.Parameters.AddWithValue("@month", agmes);
                        comm20.Parameters.AddWithValue("@year", aggod);
                        connection20.Open();
                        try
                        {
                            comm20.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                    }
                    string connectionString21 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection21 = new MySqlConnection(connectionString21))
                    {
                        string commText21 = "Insert Into  conditionally_const_expenses  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                        MySqlCommand comm21 = new MySqlCommand(commText21, connection21);
                        comm21.Parameters.AddWithValue("@ID_project", project);
                        comm21.Parameters.AddWithValue("@name", "2.	Прочие расходы");
                        comm21.Parameters.AddWithValue("@single_product", proc);
                        comm21.Parameters.AddWithValue("@month", procmes);
                        comm21.Parameters.AddWithValue("@year", procgod);
                        connection21.Open();
                        try
                        {
                            comm21.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                    }
                    string connectionString25 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection25 = new MySqlConnection(connectionString25))
                    {
                        string commText25 = "Insert Into  conditionally_const_expenses  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                        MySqlCommand comm25 = new MySqlCommand(commText25, connection25);
                        comm25.Parameters.AddWithValue("@ID_project", project);
                        comm25.Parameters.AddWithValue("@name", "3.	Коммерческие расходы");
                        comm25.Parameters.AddWithValue("@single_product", rec);
                        comm25.Parameters.AddWithValue("@month", recmes);
                        comm25.Parameters.AddWithValue("@year", recgod);
                        connection25.Open();
                        try
                        {
                            comm25.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                    }
                    string connectionString23 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection23 = new MySqlConnection(connectionString23))
                    {
                        string commText23 = "Insert Into  conditionally_const_expenses  ( ID_project, name,single_product, month, year) VALUES(  @ID_project, @name, @single_product, @month, @year)";
                        MySqlCommand comm23 = new MySqlCommand(commText23, connection23);
                        comm23.Parameters.AddWithValue("@ID_project", project);
                        comm23.Parameters.AddWithValue("@name", "Итого");
                        comm23.Parameters.AddWithValue("@single_product", post);
                        comm23.Parameters.AddWithValue("@month", postmes);
                        comm23.Parameters.AddWithValue("@year", postgod);
                        connection23.Open();
                        try
                        {
                            comm23.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                        string connectionString24 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        DataSet ds24 = new DataSet();
                        MySqlConnection connection24 = new MySqlConnection(connectionString24);
                        MySqlDataAdapter dataAdapter24 = new MySqlDataAdapter("Select name,single_product, month, year From conditionally_const_expenses WHERE 	ID_project=\"" + project + "\"", connection24);
                        dataAdapter24.Fill(ds24, "Name");
                        dataGridView9.DataSource = ds24.Tables["Name"];
                        dataGridView9.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView9.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                        dataGridView9.Columns[0].HeaderText = "name";
                        dataGridView9.Columns[1].HeaderText = "Один программный продукт, руб";
                        dataGridView9.Columns[2].HeaderText = "Месяц, руб";
                        dataGridView9.Columns[3].HeaderText = "Год, руб";
                    }
                    yslyga = per + post;
                    yslygames = permes + postmes;
                    yslygagod = pergod + postgod;
                    string connectionString27 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                    using (MySqlConnection connection27 = new MySqlConnection(connectionString27))
                    {
                        string commText27 = "Insert Into  costs  ( ID_project, name, service, month, year) VALUES(  @ID_project, @name, @service, @month, @year)";
                        MySqlCommand comm27 = new MySqlCommand(commText27, connection27);
                        comm27.Parameters.AddWithValue("@ID_project", project);
                        comm27.Parameters.AddWithValue("@name", "1. Условно-переменные");
                        comm27.Parameters.AddWithValue("@service", per);
                        comm27.Parameters.AddWithValue("@month", permes);
                        comm27.Parameters.AddWithValue("@year", pergod);
                        connection27.Open();
                        try
                        {
                            comm27.ExecuteNonQuery();
                        }
                        catch
                        {
                            MessageBox.Show("Вы не выбрали проект");
                        }
                  // Добавлению данных в таблицу издержек по выполнению
                        string connectionString28 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        using (MySqlConnection connection28 = new MySqlConnection(connectionString28))
                        {
                            string commText28 = "Insert Into  costs  ( ID_project, name, service, month, year) VALUES(  @ID_project, @name, @service, @month, @year)";
                            MySqlCommand comm28 = new MySqlCommand(commText28, connection28);
                            comm28.Parameters.AddWithValue("@ID_project", project);
                            comm28.Parameters.AddWithValue("@name", "2. Условно-постоянные");
                            comm28.Parameters.AddWithValue("@service", post);
                            comm28.Parameters.AddWithValue("@month", postmes);
                            comm28.Parameters.AddWithValue("@year", postgod);
                            connection28.Open();
                            try
                            {
                                comm28.ExecuteNonQuery();
                            }
                            catch
                            {
                                MessageBox.Show("Вы не выбрали проект");
                            }
                        }
                        string connectionString29 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                        using (MySqlConnection connection29 = new MySqlConnection(connectionString29))
                        {
                            string commText29 = "Insert Into  costs  ( ID_project, name, service, month, year) VALUES(  @ID_project, @name, @service, @month, @year)";
                            MySqlCommand comm29 = new MySqlCommand(commText29, connection29);
                            comm29.Parameters.AddWithValue("@ID_project", project);
                            comm29.Parameters.AddWithValue("@name", "Итого по услуге:");
                            comm29.Parameters.AddWithValue("@service", yslyga);
                            comm29.Parameters.AddWithValue("@month", yslygames);
                            comm29.Parameters.AddWithValue("@year", yslygagod);
                            connection29.Open();
                            try
                            {
                                comm29.ExecuteNonQuery();
                                string connectionString24 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                                DataSet ds24 = new DataSet();
                                MySqlConnection connection24 = new MySqlConnection(connectionString24);
                                MySqlDataAdapter dataAdapter24 = new MySqlDataAdapter("Select name, service, month, year From costs WHERE 	ID_project=\"" + project + "\"", connection24);

                                dataAdapter24.Fill(ds24, "Name");
                                dataGridView10.DataSource = ds24.Tables["Name"];
                                dataGridView10.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                                dataGridView10.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                                dataGridView10.Columns[0].HeaderText = "name расходов";
                                dataGridView10.Columns[1].HeaderText = "Услуга, руб";
                                dataGridView10.Columns[2].HeaderText = "Месяц, руб";
                                dataGridView10.Columns[3].HeaderText = "Год, руб";
                            }
                            catch
                            {
                                MessageBox.Show("Вы не выбрали проект");
                            }
                        }

                    }
                }

            }

            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
             if (project != "none")
            { 
            if (textBox23.Text != "" && textBox22.Text != "" && textBox26.Text != "")
            {

                string connectionString15 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connection15 = new MySqlConnection(connectionString15);
            connection15.Open();
            string commText15 = "DELETE FROM  цена_по WHERE ID_project=\"" + project + "\"";
            MySqlCommand comm15 = new MySqlCommand(commText15, connection15);
            try
            {
                comm15.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
//Расчёт цены
            double nds;
                    if (checkBox1.Checked == true)
                    {
                        nds = 20;
                    }
                    else
                    {
                        nds = 0;
                    }

             procent = Convert.ToDouble(textBox22.Text);
            double pribil = (yslyga / 100) * procent;
            pribil = Math.Round(pribil, 2);
            double cena = pribil + yslyga;
            cena = Math.Round(cena, 2);
           
            double nds1 = (cena / 100) * nds;
            nds1 = Math.Round(nds1, 2);
           
            double cenands = cena +nds1;
            cenands = Math.Round(cenands, 2);
            
            double kolvo = qyg;
          
            vir = kolvo * cena;
            vir = Math.Round(vir, 2);
            nameras = textBox23.Text;

            string connectionString12 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection12 = new MySqlConnection(connectionString12))
            {
                string commText12 = "Insert Into цена_по  ( ID_project, name, полная_себ, прибыль,	цена_предприятия, НДС, Цена_НДС,amount, Выручка, процент ) VALUES(  @ID_project, @name,  @полная_себ, @прибыль, @цена_предприятия,  @НДС, @Цена_НДС, @Колво, @Выручка, @процент)";
                MySqlCommand comm12 = new MySqlCommand(commText12, connection12);
                comm12.Parameters.AddWithValue("@ID_project", project);
                comm12.Parameters.AddWithValue("@name", nameras);
                comm12.Parameters.AddWithValue("@полная_себ", yslyga);
                comm12.Parameters.AddWithValue("@прибыль", pribil);
                comm12.Parameters.AddWithValue("@цена_предприятия", cena);
                comm12.Parameters.AddWithValue("@НДС", nds1);
                comm12.Parameters.AddWithValue("@Цена_НДС", cenands);
                comm12.Parameters.AddWithValue("@Колво", kolvo);
                comm12.Parameters.AddWithValue("@Выручка", vir);
                comm12.Parameters.AddWithValue("@процент", procent);
                connection12.Open();
                try
                {
                    comm12.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект5");
                }
            }
            string connectionString24 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
                DataSet ds24 = new DataSet();
                MySqlConnection connection24 = new MySqlConnection(connectionString24);
                MySqlDataAdapter dataAdapter24 = new MySqlDataAdapter("Select name, полная_себ, прибыль, цена_предприятия,  НДС, Цена_НДС,amount, Выручка From  цена_по WHERE 	ID_project=\"" + project + "\"", connection24);
                    double p6 = 0;
                dataAdapter24.Fill(ds24, "Name");
                dataGridView11.DataSource = ds24.Tables["Name"];
                dataGridView11.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView11.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                dataGridView11.Columns[0].HeaderText = "Услуга";
                dataGridView11.Columns[1].HeaderText = "Полная себестоимость";
                dataGridView11.Columns[2].HeaderText = "Прибыль";
                dataGridView11.Columns[3].HeaderText = "Цена предприятия";
                dataGridView11.Columns[4].HeaderText = "НДС";
                dataGridView11.Columns[5].HeaderText = "Цена с НДС";
                dataGridView11.Columns[6].HeaderText = "Кол-во услуг";
                dataGridView11.Columns[7].HeaderText = "Выручка без НДС";
                double arenda = procmes;
                double reg =Convert.ToDouble(textBox26.Text);
                double p1 = vir;
                 p1 = Math.Round(p1, 2);
                double p2 = qyg*per;
                p2 = Math.Round(p2, 2);
                double p3 = p1 - p2;
            p3 = Math.Round(p3, 2);
            double p4 = qyg*post;
                double p5 = p3 - p4;
            p5 = Math.Round(p5, 2);
                    if (checkBox3.Checked == true)
                    {
                         p6 = (stoimingod1 * 32.385) / 100;
                        p6 = Math.Round(p6, 2);
                    }
                    else
                    {
                         p6 = (stoimingod1 * 2.2) / 100;
                        p6 = Math.Round(p6, 2);
                    }
                    double p8 = 0;
            double p7 = p5 - p6;
            p7 = Math.Round(p7, 2);
                    if (checkBox3.Checked == true)
                    {
                        p8 = (p7 * 15) / 100;
                        p8 = Math.Round(p8, 2);
                    }
                    else
                    {
                        p8 = (p7 * 20) / 100;
                        p8 = Math.Round(p8, 2);
                    }
            double p9 = p7 - p8;
            p9 = Math.Round(p9, 2);
            double p10 = p9 / p1 * 100;
            p10 = Math.Round(p10, 2);
            double p11 = cena;
                double p12 = per;
                double p13 = p4 / (p11 - p12);
            p13 = Math.Round(p13, 2);
            double p14 = p1 * p4 / p3;
                    rentab = p14;
                    p14 = Math.Round(p14, 2);
            double p15 = stoimingod1;
                double p16 = arenda*12;
                double p17 = reg;
            p17 = Math.Round(p17, 2);
            double p18 = stoimingod1+ p16 + reg;
            p18 = Math.Round(p18, 2);
            double p19 = p9 / p18;
            p19 = Math.Round(p19, 2);
            double p23 = 1 / p19;
                    p23 = Math.Round(p23, 2);
                    tkb =p13;
                    textBox24.Text += "Технико-экономические показатели:" + Environment.NewLine + Environment.NewLine;
                    textBox24.Text += "Вр=" + cena + "*" + kolvo + "=" + vir + " – выручка за год без НДС;" + Environment.NewLine;
                    textBox24.Text += "Упер=" + qyg + "*" + per + "=" + p2 + " – переменные расходы;" + Environment.NewLine;
                    textBox24.Text += "Мд=" + p2 + "-" + p1 + "=" + p3 + " – маржинальный доход;" + Environment.NewLine;
                    textBox24.Text += "Упост =" + qyg * post + "=" + p4 + " – постоянные расходы;" + Environment.NewLine;
                    textBox24.Text += "Поп=" + p3 + "-" + p4 + "=" + p5 + " – прибыль от продаж;" + Environment.NewLine;
                    textBox24.Text += "Нни=2,2 % *" + stoimingod1 + "/100=" + p6 + " – налог на имущество ;" + Environment.NewLine;
                    textBox24.Text += "Пдн=" + p5 + "-" + p6 + "=" + p7 + " – прибыль до налогообложения;" + Environment.NewLine;
                    textBox24.Text += "Ннп=" + p7 + "*20%/100%=" + p8 + " – налог на прибыль  ;" + Environment.NewLine;
                    textBox24.Text += "Чп =" + p7 + "-" + p8 + "=" + p9 + " – чистая прибыль;" + Environment.NewLine;
                    textBox24.Text += "Ренп=" + p9 + "/" + p1 + "100%=" + p10 + " – рентабельность продукции;" + Environment.NewLine;
                    textBox24.Text += "Тб =" + postgod + "/" + cena + "-" + per + "=" + p13 + " – точка безубыточности;" + Environment.NewLine;
                    textBox24.Text += "Прен =" + p1 + "*" + postgod + "/" + p3 + "=" + p14 + " – порог рентабельности;" + Environment.NewLine;
                    textBox24.Text += "Ар=" + procmes + "*12=" + p16 + " – аренда помещения;" + Environment.NewLine;
                    textBox24.Text += "Знсп =" + stoimingod1 + "+" + p16 + "+" + reg + "=" + p18 + " – затраты на создание проекта;" + Environment.NewLine;
                    textBox24.Text += "Эфпр=" + p9 + "/" + p18 + "=" + p19 + " – эффективность проекта;" + Environment.NewLine;
                    textBox24.Text += "Тф =1/" + p19 + "=" + p23+ " – фактический срок окупаемости;" + Environment.NewLine + Environment.NewLine;
                    string connectionStri = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            MySqlConnection connecti = new MySqlConnection(connectionStri);
            connecti.Open();
            string comTex = "DELETE FROM   techniс_and_econom_indicators WHERE ID_project=\"" + project + "\"";
            MySqlCommand co = new MySqlCommand(comTex, connecti);
            try
            {
                co.ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Удалить проект не удалось");
            }
            string connectionString1 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection1 = new MySqlConnection(connectionString1))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection1);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "1.	Выручка за год без НДС");
                comm1.Parameters.AddWithValue("@computation", p1);
                connection1.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString2 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection2= new MySqlConnection(connectionString2))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection2);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "2.	Переменные расходы (за год)");
                comm1.Parameters.AddWithValue("@computation", p2);
                connection2.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString3 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection3 = new MySqlConnection(connectionString3))
            {

                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection3);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "3.	Маржинальный доход");
                comm1.Parameters.AddWithValue("@computation", p3);
                connection3.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString4 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection4 = new MySqlConnection(connectionString4))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection4);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "4.	Постоянные расходы (за год)");
                comm1.Parameters.AddWithValue("@computation", p4);
                connection4.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString5 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection5 = new MySqlConnection(connectionString5))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection5);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "5.	Прибыль от продаж");
                comm1.Parameters.AddWithValue("@computation", p5);
                connection5.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString6 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString6))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "6.	Налог на имущество");
                comm1.Parameters.AddWithValue("@computation", p6);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString7 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString7))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "7.	Прибыль до налогообложения");
                comm1.Parameters.AddWithValue("@computation", p7);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString8 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString8))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "8.	Налог на прибыль");
                comm1.Parameters.AddWithValue("@computation", p8);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString9 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString9))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "9.	Чистая прибыль");
                comm1.Parameters.AddWithValue("@computation", p9);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString10 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString10))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "10. Рентабельность продукции");
                comm1.Parameters.AddWithValue("@computation", p10);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString11 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString11))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "11. Цена единицы продукции (без НДС)");
                comm1.Parameters.AddWithValue("@computation", p11);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString112 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString112))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "12.Условно-переменные затраты на единицу продукции");
                comm1.Parameters.AddWithValue("@computation", p12);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString13 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString13))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "13. Точка безубыточности");
                comm1.Parameters.AddWithValue("@computation", p13);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString14 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString14))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "14. Порог рентабельности");
                comm1.Parameters.AddWithValue("@computation", p14);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString115 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString115))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "15. Стоимость основных средств");
                comm1.Parameters.AddWithValue("@computation", p15);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString116 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString116))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "16.  Аренда помещения");
                comm1.Parameters.AddWithValue("@computation", p16);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString117 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString117))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "17. Регистрация деятельности");
                comm1.Parameters.AddWithValue("@computation", p17);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString118 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString118))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "18. Затраты на создание проекта");
                comm1.Parameters.AddWithValue("@computation", p18);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проектёмаё");
                }
            }
            string connectionString119 = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            using (MySqlConnection connection6 = new MySqlConnection(connectionString119))
            {
                string commText1 = "Insert Into techniс_and_econom_indicators  ( ID_project, name, computation) VALUES(  @ID_project, @name, @computation)";
                MySqlCommand comm1 = new MySqlCommand(commText1, connection6);
                comm1.Parameters.AddWithValue("@ID_project", project);
                comm1.Parameters.AddWithValue("@name", "19. Фактический срок окупаемости");
                comm1.Parameters.AddWithValue("@computation", p19);
                connection6.Open();
                try
                {
                    comm1.ExecuteNonQuery();
                }
                catch
                {
                    MessageBox.Show("Вы не выбрали проект");
                }
            }
            string connectionStrin = "SERVER=localhost;" + "DATABASE=computation_startup;" + "UID=root;" + "PASSWORD= ;";
            DataSet ds13 = new DataSet();
            MySqlConnection connection13 = new MySqlConnection(connectionStrin);
            MySqlDataAdapter dataAdapter13 = new MySqlDataAdapter("Select  name, computation From  techniс_and_econom_indicators WHERE 	ID_project=\"" + project + "\"", connection13);

            dataAdapter13.Fill(ds13, "Name");
            dataGridView12.DataSource = ds13.Tables["Name"];
            dataGridView12.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView12.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView12.Columns[0].HeaderText = "Наименование";
            dataGridView12.Columns[1].HeaderText = "Расчёт";
            }
            else
            {
                MessageBox.Show("Вы не заполнили не все данные");
            }
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
//Нужен свитч
      private void button12_Click(object sender, EventArgs e)
        {
            st = 0;
            razr = 0;
            if (comboBox1.SelectedIndex == 0)
            {
                st = 63.43;
                razr =1 ;
            }
            if (comboBox1.SelectedIndex == 1)
            {
                st = 82.46;
                razr = 2;
            }
            if (comboBox1.SelectedIndex == 2)
            {
                st = 107.2;
                razr =3 ;
            }
            if (comboBox1.SelectedIndex == 3)
            {
                st = 121.15;
                razr = 4;
            }
            if (comboBox1.SelectedIndex == 4)
            {
                st = 137;
                razr = 5;
            }
            if (comboBox1.SelectedIndex == 5)
            {
                st = 154.77;
                razr = 6;
            }
            if (comboBox1.SelectedIndex == 6)
            {
                st = 175.07;
                razr =7;
            }
            if (comboBox1.SelectedIndex == 7)
            {
                st = 197.9;
                razr =8 ;
            }
            if (comboBox1.SelectedIndex == 8)
            {
                st = 222;
                razr = 9;
            }
            if (comboBox1.SelectedIndex == 9)
            {
                st = 253.09;
                razr =10;
            }
            if (comboBox1.SelectedIndex == 10)
            {
                st = 286.07;
                razr =11 ;
            }

            if (comboBox1.SelectedIndex == 11)
            {
                st = 323.5;
                razr =12 ;
            }
            if (comboBox1.SelectedIndex == 12)
            {
                st = 363.36;
                razr =13 ;
            }
            if (comboBox1.SelectedIndex == 13)
            {
                st = 412.93;
                razr = 14;
            }
            if (comboBox1.SelectedIndex == 14)
            {
                st = 466.84;
                razr = 15;
            }
            if (comboBox1.SelectedIndex == 15)
            {
                st = 518.22;
                razr =16 ;
            }
            if (comboBox1.SelectedIndex == 16)
            {
                st = 575.31;
                razr =17 ;
            }
            if (comboBox1.SelectedIndex == 16)
            {
                st = 638.74;
                razr = 18;
            }
            zp = st * fmes;
        }
        private void button12_Click_1(object sender, EventArgs e)
        {
            if (project != "none")
            {
                double postgod1 = postgod *Math.Pow(10, -1);
            GraphPane pane = z1.GraphPane;
            pane.CurveList.Clear();
            z1.IsShowPointValues = true;
            double[] x1 = new double[2];
            double[] y1 = new double[2];
            y1[0] = postgod1;
            x1[0] = 0;
            y1[1] = postgod1;
            x1[1] = 30;
            z1.GraphPane.AddCurve("Условно-постоянные расходы", x1, y1, Color.Black, SymbolType.XCross);
            z1.AxisChange();
            z1.Invalidate();
                double rentab1 =rentab * Math.Pow(10, -1);
                double[] x4 = new double[3];
                double[] y4 = new double[3];
                y4[0] = rentab1;
                x4[0] = 0;
                y4[1] = rentab1;
                x4[1] = tkb;
                y4[2] = 0;
                x4[2] = tkb;
                LineItem myCurve = z1.GraphPane.AddCurve("Точка безубыточности", x4, y4, Color.Red);
                myCurve.Line.Style = DashStyle.Dot;
                z1.AxisChange();
                z1.Invalidate();
                double[] x = new double[3];
            double[] y = new double[3];
            double pergod1 = pergod * Math.Pow(10, -1);
            double proc1 = (pergod / 100) * 10;
            x[0] = 0;
            y[0] = postgod1;
            y[1] = rentab1+1;
            x[1] = tkb;
                y[2] = pergod1 + postgod1;
                x[2] = qyg;
                z1.GraphPane.AddCurve("Условно-переменные расходы", x, y, Color.Green);
            z1.AxisChange();
            z1.Invalidate();
            double[] x2 = new double[3];
            double[] y2 = new double[3];
                double vir1 = vir * Math.Pow(10, -1);
                double proc2 = (vir / 100) * 10;
            y2[0] = 0;
            x2[0] = 0;
                y2[1] = rentab1;
            x2[1] = tkb;
                y2[2] = vir1;
                z1.GraphPane.AddCurve("Выручка", x2, y2, Color.Blue, SymbolType.Default);
            z1.AxisChange();
            z1.Invalidate();
            z1.GraphPane.Title = "График точки безубыточности";
            pane.YAxis.IsOmitMag = true;
            pane.XAxis.Title = "";
            pane.YAxis.Title = "";
            double qkr = postgod /(vir- per);
            }
            else
            {
                MessageBox.Show("Вы не выбрали проект");
            }
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }
        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }
        private void textBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }
        private void textBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }
        private void textBox15_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }

        private void textBox21_KeyPress(object sender, KeyPressEventArgs e)
        {
            lockword(e);
        }

        private void textBox23_KeyPress(object sender, KeyPressEventArgs e)
        {
           lockword(e);
        }
   
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender,e);
        }
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox7_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void textBox17_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }

        private void textBox19_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }

        private void textBox22_KeyPress(object sender, KeyPressEventArgs e)
        {
            locknumb(sender, e);
        }
        private void обАвтореToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Расчёт IT-стартапа- программа, осуществляющая процесс выполнения просчёта себестоимости разработки программного продукта, вычисление цены программного продукта.");
        }  
        private void AllBorders(Excel.Borders _borders)
        {
            _borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            _borders.Color = Color.Black;
        }
        private void сохранитьДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet2;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet3;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet4;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet5;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet7;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet8;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet11;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet12;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet6;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet9;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet10;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet10 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet10.Name = "Изд. по выполнению ПО";
            ExcelApp.Columns.EntireColumn.AutoFit();
            ExcelApp.Cells[1, 1] = "Наименование расходов";
            ExcelApp.Cells[1, 2] = "Услуга, руб";
            ExcelApp.Cells[1, 3] = "Месяц, руб";
            ExcelApp.Cells[1, 4] = "Год, руб";
            ExcelWorkSheet10.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet10.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet10.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet10.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView10.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView10.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView10.Rows[i].Cells[j].Value;

                }
            }
            ExcelWorkSheet9 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet9.Name = "Условно-постоянные расходы";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Один пргораммный продукт, руб";
            ExcelApp.Cells[1, 3] = "Месяц, руб";
            ExcelApp.Cells[1, 4] = "Год, руб";
            ExcelWorkSheet9.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet9.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet9.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet9.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView9.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView9.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView9.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet6 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet6.Name = "Условно-переменные расходы";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Один пргораммный продукт, руб";
            ExcelApp.Cells[1, 3] = "Месяц, руб";
            ExcelApp.Cells[1, 4] = "Год, руб";
            ExcelWorkSheet6.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet6.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet6.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet6.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView6.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView6.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView6.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet12 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet12.Name = "Техн.-экономические показатели";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Расчёт";
            ExcelWorkSheet12.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet12.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView12.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView12.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView12.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet11 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet11.Name = "Цена ПО";
            ExcelApp.Cells[1, 1] = "Услуга";
            ExcelApp.Cells[1, 2] = "Полная себестоимость";
            ExcelApp.Cells[1, 3] = "Прибыль";
            ExcelApp.Cells[1, 4] = "Цена предприятия";
            ExcelApp.Cells[1, 5] = "НДС";
            ExcelApp.Cells[1, 6] = "Цена с НДС";
            ExcelApp.Cells[1, 7] = "Кол-во услуг";
            ExcelApp.Cells[1, 8] = "Выручка без НДС";
            ExcelWorkSheet11.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 5].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 6].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 7].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet11.Cells[1, 8].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView11.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView11.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView11.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet8 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet8.Name = "Прочие расходы";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "кол.показов";
            ExcelApp.Cells[1, 3] = "Цена за один день (услугу), руб";
            ExcelApp.Cells[1, 4] = "Стоимость в месяц, руб";
            ExcelApp.Cells[1, 5] = "Стоимость в год, руб";
            ExcelWorkSheet8.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet8.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet8.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet8.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet8.Cells[1, 5].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView8.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView8.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView8.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet7 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet7.Name = "Аренда";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Количество (Размер помещения)";
            ExcelApp.Cells[1, 3] = "Цена за единицу, руб";
            ExcelApp.Cells[1, 4] = "Сумма, руб";
            ExcelWorkSheet7.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet7.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet7.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet7.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView7.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView7.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView7.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet5.Name = "Годовой баланс рабочего времени";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Количество";
            ExcelWorkSheet5.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet5.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView5.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView5.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet4.Name = "Основ. и вспом. материалы";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Характеристика";
            ExcelApp.Cells[1, 3] = "кол. материала в упаковке";
            ExcelApp.Cells[1, 4] = "Цена, руб";
            ExcelApp.Cells[1, 5] = "Стоимость, руб";
            ExcelWorkSheet4.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet4.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet4.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet4.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet4.Cells[1, 5].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView4.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView4.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet3.Name = "Основные средства и амортизация";
            ExcelApp.Cells[1, 1] = "Наименование";
            ExcelApp.Cells[1, 2] = "Количество, шт";
            ExcelApp.Cells[1, 3] = "Цена,руб";
            ExcelApp.Cells[1, 4] = "Стоимость, руб";
            ExcelApp.Cells[1, 5] = "Срок службы, лет";
            ExcelApp.Cells[1, 6] = "Годовые нормы аморт. ,%";
            ExcelApp.Cells[1, 7] = "Годовые аморт. отчисления, руб";
            ExcelWorkSheet3.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 5].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 6].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet3.Cells[1, 7].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet2.Name = "Условный пок. объёма услуг";
            ExcelApp.Cells[1, 1] = "Комплексная работа";
            ExcelApp.Cells[1, 2] = "Объём услуг в день";
            ExcelApp.Cells[1, 3] = "Объём услуг в месяц";
            ExcelApp.Cells[1, 4] = "Объём услуг в год";
            ExcelWorkSheet2.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet2.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet2.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet2.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.Add();
            ExcelWorkSheet.Name = "Норма времени";
            ExcelApp.Cells[1, 1] = "Наименование работ";
            ExcelApp.Cells[1, 2] = "Топ, час";
            ExcelApp.Cells[1, 3] = "Тпз, час";
            ExcelApp.Cells[1, 4] = "Тотл, час";
            ExcelApp.Cells[1, 5] = "Тобс, час";
            ExcelApp.Cells[1, 6] = "Нвр, час";
            ExcelWorkSheet.Cells[1, 1].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet.Cells[1, 2].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet.Cells[1, 3].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet.Cells[1, 4].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet.Cells[1, 5].HorizontalAlignment = Excel.Constants.xlCenter;
            ExcelWorkSheet.Cells[1, 6].HorizontalAlignment = Excel.Constants.xlCenter;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            ExcelWorkSheet.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet2.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet3.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet4.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet5.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet7.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet8.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet11.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet12.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet6.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet9.Columns.EntireColumn.AutoFit();
            ExcelWorkSheet10.Columns.EntireColumn.AutoFit();
            Microsoft.Office.Interop.Excel.Range tRange = ExcelWorkSheet.UsedRange;
            tRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange2 = ExcelWorkSheet2.UsedRange;
            tRange2.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange2.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange3 = ExcelWorkSheet3.UsedRange;
            tRange3.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange3.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange4 = ExcelWorkSheet4.UsedRange;
            tRange4.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange4.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange5 = ExcelWorkSheet5.UsedRange;
            tRange5.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange5.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange6 = ExcelWorkSheet6.UsedRange;
            tRange6.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange6.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange7 = ExcelWorkSheet7.UsedRange;
            tRange7.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange7.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange8 = ExcelWorkSheet8.UsedRange;
            tRange8.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange8.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange9 = ExcelWorkSheet9.UsedRange;
            tRange9.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange9.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange10 = ExcelWorkSheet10.UsedRange;
            tRange10.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange10.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange11 = ExcelWorkSheet11.UsedRange;
            tRange11.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange11.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            Microsoft.Office.Interop.Excel.Range tRange12 = ExcelWorkSheet12.UsedRange;
            tRange12.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            tRange12.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            ExcelApp.AlertBeforeOverwriting = false;
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
        private void button14_Click_3(object sender, EventArgs e)
        {
            textBox24.Text = "";
        }
        private void button15_Click_1(object sender, EventArgs e)
        {
            string str = textBox24.Text;
            using (SaveFileDialog dlgSave = new SaveFileDialog())
                try
                {
                    dlgSave.Filter = "TXT Файлы (*.txt)|*.txt |DOC Файлы (*.doc)|*.doc";
                    dlgSave.Title = "Сохранить";
                    if (dlgSave.ShowDialog() == DialogResult.OK && dlgSave.FileName.Length > 0)
                    {
                        UTF7Encoding utf8 = new UTF7Encoding();
                        StreamWriter sw = new StreamWriter(dlgSave.FileName, false, Encoding.GetEncoding(1251));
                        sw.Write(str);
                        sw.Close();
                    }
                }
                catch (Exception errorMsg)
                {
                    MessageBox.Show(errorMsg.Message);
                }
        }
    }

    }
    
    


