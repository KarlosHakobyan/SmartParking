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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (SqlConnection conn = new SqlConnection("Data Source=localhost;Initial Catalog=SmartParkingDB;Integrated Security=True"))
            {
                SqlDataAdapter ssda = new SqlDataAdapter("Select Count (*) From _Login where Us_Login= '" + textBox1.Text + "' and Us_Password= '" + textBox2.Text + "'", conn);
                DataTable dtt = new DataTable();
                ssda.Fill(dtt);
                if (dtt.Rows[0][0].ToString() == "1")
                {
                    Main ss = new Main();
                    ss.Show();
                    Hide();
                }
                else
                {
                    MessageBox.Show("Please Check Login and Password");
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            { 
                textBox2.UseSystemPasswordChar = false;
            }

            else
            {
                textBox2.UseSystemPasswordChar = true;
            }

        }
    }
}
