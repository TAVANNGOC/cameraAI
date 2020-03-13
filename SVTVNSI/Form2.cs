using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SVTVNSI
{
    public partial class Form2 : Form
    {
       
        public Form2()
        {
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
          //  SqlConnection cnn = new SqlConnection("Data Source=SV-CAS-65\\SQLEXPRESS;Initial Catalog=SQLDBCAMERAAI;User ID=sa;Password=123456");
            Form1 f = new Form1();
            string line;
            SqlConnection cnn = null;

            if (txtuser.Text != "" && txtpass.Text != "")
            {

                StreamReader sr = new StreamReader("Data\\Config.txt");
                line = sr.ReadToEnd();
                cnn = new SqlConnection(line);
                try
                {
                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT COUNT (*) AS 'STCOUNT'  FROM Account WHERE Username = N'" + txtuser.Text + "' AND Password = N'" + txtpass.Text + "'";
                    cmd.ExecuteNonQuery();
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        string sbg = reader["STCOUNT"].ToString();
                        if (Convert.ToInt32(sbg) > 0)
                        {
                            
                            f.Sender(txtuser.Text, txtpass.Text);
                            f.ShowDialog();
                            this.Close();
                        }
                        else
                        {
                            MessageBox.Show("Sai tên tài khoản hoặc mật khẩu !","Lỗi");
                        }
                    }
                    cnn.Close();
                }

                catch { MessageBox.Show("Không kết nối được với server !","Lỗi"); }
            }

            else
            {
                MessageBox.Show("Tên đăng nhập hoặc mật khẩu trống", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
              
        }
        

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Bạn muốn thoát chương trình", "Thông báo", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Application.Exit();
               
            }
            else if (dialogResult == DialogResult.No)
            {
                
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                txtpass.UseSystemPasswordChar = false;
            }
            else
            {
                txtpass.UseSystemPasswordChar = true;
            }
        }

        private void btncauhinh_Click(object sender, EventArgs e)
        {
            this.Size = new System.Drawing.Size(597, 209);
            StreamReader sr = new StreamReader("Data\\Config.txt");
            textBox1.Text = sr.ReadToEnd();
            sr.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Lưu cấu hình kết nối server !", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                string chuoiketnoi = textBox1.Text.ToString();
                using (StreamWriter sw = new StreamWriter("Data\\Config.txt"))
                {

                    sw.WriteLine(chuoiketnoi);
                    sw.Close();
                    this.Size = new System.Drawing.Size(291, 209);
                }

                MessageBox.Show("Lưu cấu hình thành công !");
            }
            else if (dialogResult == DialogResult.No)
            {
                this.Size = new System.Drawing.Size(291, 209);
            }

           
        }

        private void txtpass_KeyDown(object sender, KeyEventArgs e)
        {
            
        }
    }
}
