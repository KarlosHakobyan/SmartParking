using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace Login_Project
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        void BindData()
        {
            SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
            SqlCommand command = new SqlCommand("select * from Infos_tab", con);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "")
            {
                SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
                con.Open();
                SqlCommand command = new SqlCommand("Insert into Infos_tab(Car_Model,Car_Num,First_Name,Last_Name,Phone_Number,Insert_Date) values(@Car_Model, @Car_Num, @First_Name, @Last_Name, @Phone_Number, getdate())", con);
                command.Parameters.AddWithValue("@Car_Model", textBox2.Text);
                command.Parameters.AddWithValue("@Car_Num", textBox3.Text);
                command.Parameters.AddWithValue("@First_Name", textBox4.Text);
                command.Parameters.AddWithValue("@Last_Name", textBox5.Text);
                command.Parameters.AddWithValue("@Phone_Number", textBox6.Text);
                command.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Successfully Inserted.");
                BindData();
            }
            else
            { MessageBox.Show("Insert full data without Car ID"); }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && textBox5.Text != "" && textBox6.Text != "")
            {
                SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
                con.Open();
                SqlCommand command = new SqlCommand("update Infos_tab set Car_Model = '" + textBox2.Text + "',Car_Num = '" + textBox3.Text + "',First_Name= '" + textBox4.Text + "',Last_Name= '" + textBox5.Text + "',Phone_Number= '" + textBox6.Text + "' where CarID= '" + int.Parse(textBox1.Text) + "'", con);
                command.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Successfully Updated.");
                BindData();
            }
            else
            { MessageBox.Show("Insert full data"); }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {
                if (MessageBox.Show("Are you sure to delete?", "Delete Record", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
                    con.Open();
                    SqlCommand command = new SqlCommand("delete Infos_tab where CarID= '" + int.Parse(textBox1.Text) + "'", con);
                    command.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Successfully Deleted.");
                    BindData();
                }
            }
            else
            {
                MessageBox.Show("Insert Car ID");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox3.Text != "" || textBox4.Text != "")
            {
                SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
                SqlCommand command = new SqlCommand("select * from Infos_tab where Car_Num = '" + textBox3.Text + "' or First_Name = '" + textBox4.Text + "'", con);
                SqlDataAdapter sd = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                sd.Fill(dt);
                dataGridView1.DataSource = dt;
            }
            else
            {
                MessageBox.Show("Insert First Name or Car Number");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SqlConnection con = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True");
            SqlCommand command = new SqlCommand("select * from Infos_tab", con);
            SqlDataAdapter sd = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            sd.Fill(dt);
            dataGridView1.DataSource = dt;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            dataGridView1.SelectAll();
            DataObject copydata = dataGridView1.GetClipboardContent();
            if (copydata != null)
            {
                Clipboard.SetDataObject(copydata);
                Microsoft.Office.Interop.Excel.Application xlapp = new Microsoft.Office.Interop.Excel.Application();
                xlapp.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook xlwbook;
                Microsoft.Office.Interop.Excel.Worksheet xlsheet;
                object miseddata = System.Reflection.Missing.Value;
                xlwbook = xlapp.Workbooks.Add(miseddata);
                xlsheet = (Microsoft.Office.Interop.Excel.Worksheet)xlwbook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range xlr = (Microsoft.Office.Interop.Excel.Range)xlsheet.Cells[2, 1];
                xlr.Select();
                xlsheet.PasteSpecial(xlr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            else
            {
                MessageBox.Show("No data found.");
            }
        }
    }
}
