using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Data.SqlClient;

namespace SVTVNSI
{
    public partial class Form1 : Form
    {
      
        public delegate void SendMessage(string user,string pass);
        public SendMessage Sender;
        public Form1()
        {

            InitializeComponent();
            Sender = new SendMessage(GetMessage);

        }
        string LNV = "";
        string LNV2 = "";
        string line;
        string userrel, passrel,typeac;
        SqlConnection cnn = null;
        private void GetMessage(string user,string pass)
        {
            userrel = user;
            passrel = pass;
           
        }
        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        public static string RemoveUnicode(string text)
        {
            string[] arr1 = new string[] { "á", "à", "ả", "ã", "ạ", "â", "ấ", "ầ", "ẩ", "ẫ", "ậ", "ă", "ắ", "ằ", "ẳ", "ẵ", "ặ",
                    "đ",
                    "é","è","ẻ","ẽ","ẹ","ê","ế","ề","ể","ễ","ệ",
                    "í","ì","ỉ","ĩ","ị",
                    "ó","ò","ỏ","õ","ọ","ô","ố","ồ","ổ","ỗ","ộ","ơ","ớ","ờ","ở","ỡ","ợ",
                    "ú","ù","ủ","ũ","ụ","ư","ứ","ừ","ử","ữ","ự",
                    "ý","ỳ","ỷ","ỹ","ỵ",};
            string[] arr2 = new string[] { "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a", "a",
                    "d",
                    "e","e","e","e","e","e","e","e","e","e","e",
                    "i","i","i","i","i",
                    "o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o","o",
                    "u","u","u","u","u","u","u","u","u","u","u",
                    "y","y","y","y","y",};
            for (int i = 0; i < arr1.Length; i++)
            {
                text = text.Replace(arr1[i], arr2[i]);
                text = text.Replace(arr1[i].ToUpper(), arr2[i].ToUpper());
            }
            return text;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.cbtypeac.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            try
            {
                StreamReader sr = new StreamReader("Data\\Config.txt");
                line = sr.ReadToEnd();
                cnn = new SqlConnection(line);
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [STT],[MNV],[HOTEN],[CHUCDANH],[PHONGBAN],[LNV] FROM [dbo].[DANHSACHNHANVIEN] ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                dataGridView1.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                dataGridView1.Columns[2].HeaderText = "HỌ TÊN";
                dataGridView1.Columns[3].HeaderText = "CHỨC VỤ";
                dataGridView1.Columns[4].HeaderText = "PHÒNG BAN";
                dataGridView1.AutoResizeColumns();
                dataGridView1.AutoResizeRows();
               // cnn.Close();
                dateTimePicker1.CustomFormat = "yyyy-MM-dd";
                dateTimePicker2.CustomFormat = "yyyy-MM-dd";

                dateTimePicker3.CustomFormat = "HH:mm";
                dateTimePicker4.CustomFormat = "HH:mm";

                dateTimePicker1.Format = DateTimePickerFormat.Custom;
                dateTimePicker2.Format = DateTimePickerFormat.Custom;

                dateTimePicker3.Format = DateTimePickerFormat.Custom;
                dateTimePicker4.Format = DateTimePickerFormat.Custom;

                dateTimePicker1.Value = DateTime.Today;              
                dateTimePicker2.Value = DateTime.Today;
                dataGridView1.Rows[0].Selected = true;

                
                // cnn.Open();
                // SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [HOTEN] FROM [dbo].[DANHSACHNHANVIEN] ";
                cmd.ExecuteNonQuery();
                DataTable dt2 = new DataTable();
                SqlDataAdapter da2 = new SqlDataAdapter(cmd);
                da2.Fill(dt2);

                comboBox2.Items.Add("All");
                foreach (DataRow dr in dt2.Rows)
                {

                    comboBox2.Items.Add(dr["HOTEN"].ToString());

                }
                
                cmd.CommandText = "SELECT Format([START],'yyyy-MM-dd') AS START1, Format([END],'yyyy-MM-dd') AS END1 FROM THOIGIANCHAMCONG; ";
                cmd.ExecuteNonQuery();
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    textBox3.Text = reader["START1"].ToString();
                    textBox8.Text = reader["END1"].ToString();


                }
                reader.Close();

                cmd.CommandText = "SELECT [Permission] AS TYPEAC FROM Account WHERE [Username] = '" + userrel + "'";
                cmd.ExecuteNonQuery();
                 reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    typeac = reader["TYPEAC"].ToString();
                    if (typeac == "ADMIN") { tabControl1.TabPages.Remove(tabPage2); } else
                    if (typeac == "USER") 
                    { 
                        tabControl1.TabPages.Remove(tabPage1);
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage5);
                        groupBox2.Visible = false;
                        button6.Visible = false;
                        button5.Visible = false;
                    } else { }

                }
                reader.Close();
                dateTimePicker1.Value = new DateTime(Convert.ToInt32(textBox3.Text.Substring(0, 4)), Convert.ToInt32(textBox3.Text.Substring(5, 2)), Convert.ToInt32(textBox3.Text.Substring(8, 2)));
                dateTimePicker2.Value = new DateTime(Convert.ToInt32(textBox8.Text.Substring(0, 4)), Convert.ToInt32(textBox8.Text.Substring(5, 2)), Convert.ToInt32(textBox8.Text.Substring(8, 2)));

                cnn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã có lỗi xảy ra ! ");
                cnn.Close();
                this.Close();
            }
        }
        private void btnchonanh_Click(object sender, EventArgs e)
        {
          
        }
        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (txtMNV.Text != "" && txtHoten.Text != "" && txtchucdanh.Text != "" && txtphongban.Text != "" && LNV != "")
            {
                int hang = dataGridView1.Rows.Count - 1;
                txtCalam.Text = (hang + 1).ToString();
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;

                cmd.CommandText = "INSERT INTO DANHSACHNHANVIEN (STT,MNV,HOTEN,CHUCDANH,PHONGBAN,LNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "',N'" + txtHoten.Text.ToString() + "',N'" + txtchucdanh.Text + "',N'" + txtphongban.Text + "',N'" + LNV + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO CheckinT (STT,MNV,Name) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "',N'" + RemoveUnicode(txtHoten.Text.ToString()) + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO CheckoutT (STT,MNV,Name) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "',N'" + RemoveUnicode(txtHoten.Text.ToString()) + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO TLVS (STT,MNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO TLVC (STT,MNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO TGLV (STT,MNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO BANGCHAMCONGTHANG (STT,MNV,HOTEN,CHUCVU,PHONGBAN,LNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "',N'" + txtHoten.Text.ToString() + "',N'" + txtchucdanh.Text + "',N'" + txtphongban.Text + "',N'" + LNV + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "INSERT INTO BANGCHAMCONGCHITIET (STT,MNV,HOTEN,CHUCVU,PHONGBAN,LNV) VALUES('" + txtCalam.Text + "',N'" + txtMNV.Text + "',N'" + txtHoten.Text.ToString() + "',N'" + txtchucdanh.Text + "',N'" + txtphongban.Text + "',N'" + LNV + "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "SELECT * FROM DANHSACHNHANVIEN ";
                cmd.ExecuteNonQuery();

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
              
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                cnn.Close();
                MessageBox.Show("Thêm nhân viên mới thành công !", "Thông báo");
            }
            else { MessageBox.Show("Không được để trống thông tin nhân viên !", "Thông báo"); }
        }
        private void btnsave_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "DSNV";

            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;

            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            worksheet.Columns.AutoFit();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "DANH_SACH_NHAN_VIEN";
            saveFileDialog.DefaultExt = ".xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show("Xuất dữ liệu thành công !");
            }

            app.Quit();
            


        }
        private void btncancel_Click(object sender, EventArgs e)
        {
            Application.Exit();
            
        }
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
          //  DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
            //populate the textbox from specific value of the coordinates of column and row.

        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridView1.Rows[e.RowIndex].Selected = true;
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                txtCalam.Text = row.Cells[0].Value.ToString();
                txtHoten.Text = row.Cells[2].Value.ToString();
                txtMNV.Text = row.Cells[1].Value.ToString();
                txtchucdanh.Text = row.Cells[3].Value.ToString();
                txtphongban.Text = row.Cells[4].Value.ToString();
                string LNVLD = row.Cells[5].Value.ToString();
                if (LNVLD == "CT") { radioButton1.Checked = true; } else if(LNVLD == "TV") { radioButton2.Checked = true; }
                else if (LNVLD == "ĐTF") { radioButton3.Checked = true; } else if (LNVLD == "ĐTP") { radioButton4.Checked = true; }
                else { MessageBox.Show("Loại nhân viên không xác định !"); }
               

            }
        }
        private void button2_Click(object sender, EventArgs e)
        {


            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;

            cmd.CommandText = "UPDATE DANHSACHNHANVIEN SET HOTEN=N'" + txtHoten.Text + "',CHUCDANH=N'" + txtchucdanh.Text + "',PHONGBAN=N'" + txtphongban.Text + "',LNV=N'" + LNV + "' WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET HOTEN=N'" + txtHoten.Text + "',CHUCVU=N'" + txtchucdanh.Text + "',PHONGBAN=N'" + txtphongban.Text + "',LNV=N'" + LNV + "' WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET HOTEN=N'" + txtHoten.Text + "',CHUCVU=N'" + txtchucdanh.Text + "',PHONGBAN=N'" + txtphongban.Text + "',LNV=N'" + LNV + "' WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "UPDATE CheckinT SET Name=N'" + RemoveUnicode(txtHoten.Text) + "' WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "UPDATE CheckoutT SET Name=N'" + RemoveUnicode(txtHoten.Text) + "' WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();


            cmd.CommandText = "SELECT * FROM DANHSACHNHANVIEN ";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            dataGridView1.FirstDisplayedScrollingRowIndex = Convert.ToInt32(txtCalam.Text)-7;
            dataGridView1.Rows[Convert.ToInt32(txtCalam.Text)-1].Selected = true;
            cnn.Close();
            MessageBox.Show("Cập nhật dữ liệu thành công !", "Thông báo");

        }
        private void button3_Click(object sender, EventArgs e)
        {
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;

            cmd.CommandText = "DELETE FROM DANHSACHNHANVIEN WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM BANGCHAMCONGTHANG WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM BANGCHAMCONGCHITIET WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM CheckinT WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM CheckoutT WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM TLVS WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM TLVC WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "DELETE FROM TGLV WHERE MNV = '" + txtMNV.Text + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "SELECT * FROM DANHSACHNHANVIEN ";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
            int hang = dataGridView1.Rows.Count - 2;
            for (int i = 0; i <= hang; i++)
            {
                dataGridView1.Rows[i].Cells[0].Value = i+1;
                int STT = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                cmd.CommandType = CommandType.Text;

                cmd.CommandText = "UPDATE DANHSACHNHANVIEN SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE CheckinT SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE CheckoutT SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE TLVS SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE TLVC SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE TGLV SET STT= '" + STT + "' WHERE MNV = '" + dataGridView1.Rows[i].Cells[1].Value.ToString() + "'";
                cmd.ExecuteNonQuery();
            }
            dataGridView1.FirstDisplayedScrollingRowIndex = Convert.ToInt32(txtCalam.Text) - 1;
            cnn.Close();
            MessageBox.Show("Xoá dữ liệu thành công !", "Thông báo");
        }
        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.Columns.Clear();
            if (comboBox2.Text == "" || comboBox2.Text == "All")
            {
                try
                {
                    dataGridView2.Refresh();
                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;

                   // cmd.CommandText = "SELECT ROW_NUMBER () OVER (ORDER BY [DATE]) STT FROM Backupdata WHERE [DATE] >= '" + dateTimePicker1.Value.ToShortDateString() + " " + dateTimePicker3.Value.ToShortTimeString() + "' AND [DATE] <= '" + dateTimePicker2.Value.ToShortDateString() + " " + dateTimePicker4.Value.ToShortTimeString() + "'";
                   // cmd.ExecuteNonQuery();

                    cmd.CommandText = "SELECT [STT], [DATE],[NAME] FROM [dbo].[Backupdata] WHERE [DATE] >= '" + dateTimePicker1.Value.ToShortDateString() + " " + dateTimePicker3.Value.ToShortTimeString() + "' AND [DATE] <= '" + dateTimePicker2.Value.ToShortDateString() + " " + dateTimePicker4.Value.ToShortTimeString() + "'";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    dataGridView2.DataSource = dt;
                    dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);

                    dataGridView2.AutoResizeColumns();
                    dataGridView2.AutoResizeRows();

                    dataGridView2.Columns[1].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";

                    textBox3.Text = dateTimePicker1.Value.ToString("yyyy-MM-dd").Substring(0, 7) + "-01";
                    textBox8.Text = dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    cmd.CommandText = "UPDATE [dbo].[THOIGIANCHAMCONG] SET [START] = '" + textBox3.Text + "' ,[END] = '" + textBox8.Text + "' WHERE [ID] = '1'";
                    cmd.ExecuteNonQuery();
                    cnn.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không kết nối được server ! ");
                }

            }
            else
            {


                try
                {
                    dataGridView2.Refresh();
                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT [STT],[DATE],[NAME] FROM [dbo].[Backupdata] WHERE [DATE] >= '" + dateTimePicker1.Value.ToShortDateString() + " " + dateTimePicker3.Value.ToShortTimeString() + "' AND [DATE] <= '" + dateTimePicker2.Value.ToShortDateString() + " " + dateTimePicker4.Value.ToShortTimeString() + "' AND NAME = '" + RemoveUnicode(comboBox2.Text.ToString()) + "'";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    dataGridView2.DataSource = dt;
                    dataGridView2.Sort(dataGridView2.Columns[1], ListSortDirection.Ascending);
                    dataGridView2.AutoResizeColumns();
                    dataGridView2.AutoResizeRows();
                    dataGridView2.Columns[1].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm";
                    textBox3.Text = dateTimePicker1.Value.ToString("yyyy-MM-dd").Substring(0, 7) + "-01";
                    textBox8.Text = dateTimePicker2.Value.ToString("yyyy-MM-dd");
                    cmd.CommandText = "UPDATE [dbo].[THOIGIANCHAMCONG] SET [START] = '" + textBox3.Text + "' ,[END] = '" + textBox8.Text + "' WHERE [ID] = '1'";
                    cmd.ExecuteNonQuery();
                    cnn.Close();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Không kết nối được server ! ");
                }
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {

            if (dataGridView2.Rows.Count > 0)
            {

                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;

                for (int i = 1; i <= 31; i++)
                {
                    // Xoá dữ liệu cũ //
                    cmd.CommandText = "UPDATE CheckinT SET [" + i + "] = '00:00' ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE CheckoutT SET [" + i + "] = '00:00' ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVS SET [" + i + "] = 0 ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] = 0 ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TGLV SET [" + i + "] = 0 ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + i + ")] = NULL,[OUT(" + i + ")] = NULL ; ";
                    cmd.ExecuteNonQuery();

                }



                button8.PerformClick();

                button9.PerformClick();



                // cnn.Open();
                for (int i = 1; i <= 31; i++)
                {
                    // xử lý dữ liệu cho danh sách ưu tiên //
                    cmd.CommandText = "UPDATE CheckinT SET [" + i + "] = '07:00:00' FROM UT INNER JOIN CheckinT ON UT.MNV = CheckinT.MNV,THOIGIANCHAMCONG WHERE UT.MNV = CheckinT.MNV ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE CheckoutT SET [" + i + "] = '18:00:00' FROM UT INNER JOIN CheckoutT ON UT.MNV = CheckoutT.MNV WHERE UT.MNV = CheckoutT.MNV; ";
                    cmd.ExecuteNonQuery();
                }

                for (int i = 1; i <= 31; i++)
                {
                    // BẢNG CHẤM CÔNG CHI TIẾT //
                    cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + i + ")] = CheckinT.[" + i + "] , [OUT(" + i + ")]  = CheckoutT.[" + i + "] FROM (CheckinT INNER JOIN BANGCHAMCONGCHITIET ON BANGCHAMCONGCHITIET.MNV = CheckinT.MNV) INNER JOIN CheckoutT ON BANGCHAMCONGCHITIET.MNV = CheckoutT.MNV;";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE [dbo].[BANGCHAMCONGCHITIET] SET [IN(" + i + ")] = NULL WHERE [IN(" + i + ")] = '00:00'";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE [dbo].[BANGCHAMCONGCHITIET] SET [OUT(" + i + ")] = NULL WHERE [OUT(" + i + ")] = '00:00'";
                    cmd.ExecuteNonQuery();

                }

                for (int i = 1; i <= 31; i++)
                {
                    // Tính thời gian làm việc buổi sáng //
                    cmd.CommandText = "UPDATE TLVS SET [" + i + "] =  ((DATEPART(hh,KHUNGGIO.[ETIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[ETIME1]))-((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) FROM (TLVS INNER JOIN CheckinT ON TLVS.MNV = CheckinT.MNV),KHUNGGIO WHERE(TLVS.[MNV] =CheckinT.[MNV])";
                    cmd.ExecuteNonQuery();


                    cmd.CommandText = "UPDATE TLVS SET [" + i + "] = CASE when [" + i + "] < 0 then 0 when [" + i + "] >= 240 then 240 when ([" + i + "] < 180) then 0 ELSE [" + i + "] END " +
                                     "FROM(DANHSACHNHANVIEN INNER JOIN TLVS ON DANHSACHNHANVIEN.MNV = TLVS.MNV)" +
                                     " WHERE TLVS.MNV = DANHSACHNHANVIEN.MNV";
                    cmd.ExecuteNonQuery();



                }

                try
                {

                    for (int i = 1; i <= 31; i++)
                    {
                        // Tính thời gian làm việc buổi chiều //


                        cmd.CommandText = "UPDATE TLVC SET [" + i + "] =  ((DATEPART(hh,CheckoutT.[" + i + "])*60)+DATEPART(MINUTE,CheckoutT.[" + i + "])) - ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2]))" +
                                        "FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                        " WHERE ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2]))" +
                                        " AND DANHSACHNHANVIEN.MNV = TLVC.MNV; ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE TLVC SET [" + i + "] =  ((DATEPART(hh,CheckoutT.[" + i + "])*60)+DATEPART(MINUTE,CheckoutT.[" + i + "])) - ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "]))" +
                                          "FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                          " WHERE ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) > ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2])) " +
                                          " AND((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) <= 840" +
                                          "AND DANHSACHNHANVIEN.MNV = TLVC.MNV; ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE TLVC SET [" + i + "] = CASE WHEN [" + i + "] < 0 THEN 0 ELSE [" + i + "] END FROM (DANHSACHNHANVIEN INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV) WHERE TLVC.MNV = DANHSACHNHANVIEN.MNV ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE TLVC SET [" + i + "] = 0 " +
                                          " FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                          " WHERE((DATEPART(hh, CheckinT.[" + i + "]) * 60) + DATEPART(MINUTE, CheckinT.[" + i + "])) > ((DATEPART(hh, KHUNGGIO.[STIME2]) * 60) + DATEPART(MINUTE, KHUNGGIO.[STIME2])) + 30  AND DANHSACHNHANVIEN.MNV = TLVC.MNV";
                        cmd.ExecuteNonQuery();

                        //cmd.CommandText = "UPDATE TLVC SET [" + i + "] = '0' FROM (DANHSACHNHANVIEN INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV),THOIGIANCHAMCONG WHERE DATEPART(weekday, convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + i + ")) = 7";
                        //cmd.ExecuteNonQuery();

                    }
                }
                catch { }

                for (int i = 1; i <= 31; i++)
                {
                    // Tính tổng thời gian làm việc //
                    cmd.CommandText = "UPDATE TGLV SET [" + i + "] = convert(int,TLVS.[" + i + "]) + convert(int,TLVC.[" + i + "]) FROM (TGLV INNER JOIN TLVC ON TGLV.MNV = TLVC.MNV) INNER JOIN TLVS ON TGLV.MNV = TLVS.MNV WHERE TGLV.MNV=TLVS.MNV;";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TGLV SET [" + i + "] = 0 FROM (TGLV INNER JOIN CheckinT ON TGLV.MNV = CheckinT.MNV) INNER JOIN CheckoutT ON TGLV.MNV = CheckoutT.MNV WHERE CheckinT.[" + i + "]=CheckoutT.[" + i + "] ;";
                    cmd.ExecuteNonQuery();

                }

                for (int i = 1; i <= 31; i++)
                {
                    // UPDATE thời gian làm việc //
                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + i + "] = TGLV.[" + i + "] FROM (BANGCHAMCONGTHANG INNER JOIN TGLV ON BANGCHAMCONGTHANG.MNV = TGLV.MNV) WHERE BANGCHAMCONGTHANG.MNV = TGLV.MNV;";
                    cmd.ExecuteNonQuery();

                }


                try
                {

                    for (int a = 1; a <= 31; a++)
                    {

                        // Chấm công cho nhân viên ( TV ,CT ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE " +
                            "(DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV')" +
                            "AND  convert(int,TGLV.[" + a + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                            "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) < ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                            "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE " +
                            "(DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                            "convert(int,TGLV.[" + a + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                            "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) >= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                            "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE " +
                            "(DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                           "convert(int,TGLV.[" + a + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                           "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) + 30 )" +
                           "AND  DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV=CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                            "convert(int,TGLV.[" + a + "]) >= convert(int,KHUNGGIO.[SGLV])*30  AND convert(int,TGLV.[" + a + "]) < convert(int,KHUNGGIO.[SGLV])*60 AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE" +
                            " (convert(int,TGLV.[" + a + "]) < convert(int,KHUNGGIO.[SGLV])*30 OR convert(int,TGLV.[" + a + "]) = 0) AND (DANHSACHNHANVIEN.LNV='CT' OR DANHSACHNHANVIEN.LNV='TV') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') " +
                                          "AND  ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                                          "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'M/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') " +
                          "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +

                          "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                         "convert(int,TGLV.[" + a + "]) <= convert(int,KHUNGGIO.[SGLV])*22.5" +
                        " AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + a + ")) = 1";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + a + ")] = NULL,[OUT(" + a + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + a + ")) = 1";
                        cmd.ExecuteNonQuery();

                    }

                }
                catch {  }

                try
                {

                    for (int b = 1; b <= 31; b++)
                    {

                        // Chấm công cho nhân viên ( ĐTF ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                               "AND  convert(int,TGLV.[" + b + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                               "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) < ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                               "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                            "convert(int,TGLV.[" + b + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                            "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) >= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                            "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                           "convert(int,TGLV.[" + b + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                           "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) + 30 )" +
                           "AND  DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV=CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                            "convert(int,TGLV.[" + b + "]) >= convert(int,KHUNGGIO.[SGLV])*30  AND convert(int,TGLV.[" + b + "]) < convert(int,KHUNGGIO.[SGLV])*60 AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE" +
                            " (convert(int,TGLV.[" + b + "]) < convert(int,KHUNGGIO.[SGLV])*30 OR convert(int,TGLV.[" + b + "]) = 0) AND (DANHSACHNHANVIEN.LNV=N'ĐTF') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                                          "AND  ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                                          "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'M/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                          "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                          "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                         "convert(int,TGLV.[" + b + "]) <= convert(int,KHUNGGIO.[SGLV])*22.5" +
                        " AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + b + ")) = 1";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + b + ")] = NULL,[OUT(" + b + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + b + ")) = 1";
                        cmd.ExecuteNonQuery();




                    }
                }
                catch {  }

                try
                {

                    for (int c = 1; c <= 31; c++)
                    {

                        // Chấm công cho nhân viên ( ĐTP ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') " +
                            "AND  convert(int,TGLV.[" + c + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                            "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                            "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') " +
                            "AND  convert(int,TGLV.[" + c + "]) >= convert(int,KHUNGGIO.[SGLV])*30 AND convert(int,TGLV.[" + c + "]) < convert(int,KHUNGGIO.[SGLV])*60 " +
                           "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') AND " +
                              "convert(int,TGLV.[" + c + "]) >= convert(int,KHUNGGIO.[SGLV])*60 " +
                              "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                              "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7 ";
                        cmd.ExecuteNonQuery();



                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'KL' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE" +
                                " (convert(int,TGLV.[" + c + "]) < convert(int,KHUNGGIO.[SGLV])*30 OR convert(int,TGLV.[" + c + "]) = 0) AND (DANHSACHNHANVIEN.LNV=N'ĐTP') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7 ";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTP') " +
                            "AND  convert(int,TGLV.[" + c + "]) >= convert(int,KHUNGGIO.[SGLV])*30 " +
                            "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                            "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) = 7";
                        cmd.ExecuteNonQuery();


                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTP') AND " +
                            "convert(int,TGLV.[" + c + "]) < convert(int,KHUNGGIO.[SGLV])*30 " +
                           " AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) = 7 ";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + c + ")) = 1";
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + c + ")] = NULL,[OUT(" + c + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + c + ")) = 1";
                        cmd.ExecuteNonQuery();

                    }
                }
                catch { }

                MessageBox.Show("Chấm công hoàn tất !");
                cnn.Close();
              

            }
         
           
        }
        private void button6_Click(object sender, EventArgs e)
        {
            
            
            DialogResult dialogResult = MessageBox.Show("Bạn muốn dữ liệu chấm công cũ !", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;

                for (int i = 1; i <= 31; i++)
                {
                    // Xoá dữ liệu cũ //
                    cmd.CommandText = "UPDATE CheckinT SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE CheckoutT SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVS SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TGLV SET [" + i + "] = 0 ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + i + "] = NULL ; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + i + ")] = NULL,[OUT(" + i + ")] = NULL ; ";
                    cmd.ExecuteNonQuery();

                }
                cnn.Close();
                MessageBox.Show("Xoá dữ liệu thành công !");
            }
            else
            {


            }
          

        }    
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Bạn muốn xoá toàn bộ dữ liệu !", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;

                for (int i = 1; i <= 31; i++)
                {
                    // Xoá dữ liệu cũ //
                    cmd.CommandText = "DELETE * FROM Backupdata2 ";
                    cmd.ExecuteNonQuery();

                }
                cnn.Close();
                MessageBox.Show("Xoá dữ liệu thành công !");
            }
            else
            {


            }




        }
        private void button8_Click(object sender, EventArgs e)
        {
          //  cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dataGridView2.Rows.Count;

            for (int i = dataGridView2.Rows.Count-1; i > 1 ; i--)
            {

                progressBar1.Value = dataGridView2.Rows.Count - i;
                // string mnv = dataGridView2.Rows[i].Cells[1].Value.ToString();
                string name = dataGridView2.Rows[i].Cells[2].Value.ToString();
               // string ht = dataGridView2.Rows[i].Cells[2].Value.ToString();
               // string time = dataGridView2.Rows[i].Cells[3].Value.ToString();
                DateTime time = Convert.ToDateTime( dataGridView2.Rows[i].Cells[1].Value.ToString());
                string time3 = time.ToString("yyyy-MM-dd HH:mm:ss");
                // string status = dataGridView2.Rows[i].Cells[5].Value.ToString();

                if (time3 != "")
                {
                    string time1 = DateTime.Parse(time3.Substring(11, 5)).ToString("t");

                    if (time3.Substring(8, 2) == "01")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [1]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "02")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [2]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "03")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [3]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "04")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [4]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "05")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [5]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "06")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [6]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "07")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [7]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "08")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [8]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "09")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [9]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "10")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [10]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "11")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [11]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "12")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [12]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "13")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [13]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "14")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [14]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "15")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [15]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "16")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [16]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "17")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [17]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "18")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [18]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "19")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [19]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "20")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [20]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "21")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [21]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "22")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [22]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "23")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [23]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "24")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [24]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "25")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [25]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "26")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [26]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "27")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [27]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "28")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [28]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "29")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [29]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "30")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [30]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "31")
                    {

                        cmd.CommandText = "UPDATE CheckinT SET [31]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                }
            }
            progressBar1.Value = 0;

            //cnn.Close();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            //cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = dataGridView2.Rows.Count;
            for (int i = 2; i < dataGridView2.Rows.Count; i++)
            {
                progressBar1.Value = i;
              
                string name = dataGridView2.Rows[i].Cells[2].Value.ToString();

              
               // string time = dataGridView2.Rows[i].Cells[1].Value.ToString();
                DateTime time = Convert.ToDateTime(dataGridView2.Rows[i].Cells[1].Value.ToString());
                string time3 = time.ToString("yyyy-MM-dd HH:mm:ss");

                // string status = dataGridView2.Rows[i].Cells[5].Value.ToString();
                if (time3 != "")
                {
                    string time1 = DateTime.Parse(time3.Substring(11, 5)).ToString("t");
                    if (time3.Substring(8, 2) == "01")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [1]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "02")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [2]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "03")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [3]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "04")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [4]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "05")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [5]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "06")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [6]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "07")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [7]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "08")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [8]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "09")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [9]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "10")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [10]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "11")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [11]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "12")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [12]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "13")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [13]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "14")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [14]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "15")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [15]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "16")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [16]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }

                    if (time3.Substring(8, 2) == "17")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [17]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "18")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [18]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "19")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [19]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "20")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [20]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "21")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [21]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "22")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [22]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "23")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [23]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "24")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [24]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "25")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [25]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "26")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [26]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "27")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [27]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "28")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [28]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "29")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [29]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "30")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [30]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }
                    if (time3.Substring(8, 2) == "31")
                    {

                        cmd.CommandText = "UPDATE CheckoutT SET [31]='" + time1 + "' WHERE Name = '" + name + "'";
                        cmd.ExecuteNonQuery();

                    }



                }
            }
          
           // cnn.Close();
        }
        private void button10_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "BANGCHAMCONGTHANG";

            for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;

            }
            for (int i = 0; i < dataGridView3.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView3.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                }
            }
            worksheet.Columns.AutoFit();

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = comboBox1.Text.ToString();
            saveFileDialog.DefaultExt = ".xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(saveFileDialog.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show("Xuất dữ liệu thành công !");
            }

            app.Quit();


        }
        private void tabPage3_Click(object sender, EventArgs e)
        {
           
        }
        private void button11_Click(object sender, EventArgs e)
        {

            string bcc = comboBox1.Text.ToString();

            MessageBox.Show("Bạn muốn xem bảng chấm công ? Chọn OK để tiếp tục !", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (bcc == "BANGCHAMCONGTHANG")
            {
                DateTime thang = Convert.ToDateTime(textBox3.Text);

                string thangst = thang.ToString("MM");

                if (thangst == "01" || thangst == "03" || thangst == "05" || thangst == "07" || thangst == "08" || thangst == "10" || thangst == "12")
                {

                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT * FROM BANGCHAMCONGTHANG ";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    dataGridView3.DataSource = dt;
                    dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                    cnn.Close();
                    dataGridView3.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                    dataGridView3.Columns[2].HeaderText = "HỌ TÊN";
                    dataGridView3.Columns[3].HeaderText = "CHỨC VỤ";
                    dataGridView3.Columns[4].HeaderText = "PHÒNG BAN";
                    for (int i = 0; i <= 30; i++)
                    {
                        DateTime txt3 = Convert.ToDateTime(textBox3.Text.ToString());
                        string thu = ((txt3.AddDays(+i).DayOfWeek).ToString());
                        if (thu == "Monday") thu = "T.2";
                        if (thu == "Tuesday") thu = "T.3";
                        if (thu == "Wednesday") thu = "T.4";
                        if (thu == "Thursday") thu = "T.5";
                        if (thu == "Friday") thu = "T.6";
                        if (thu == "Saturday") thu = "T.7";
                        if (thu == "Sunday") thu = "CN";


                        dataGridView3.Columns[i + 6].HeaderText = "(" + (i + 1) + ") " + thu;



                    }
                    dataGridView3.AutoResizeColumns();
                    dataGridView3.AutoResizeRows();
                    dataGridView3.Columns[2].Frozen = true;
                    MessageBox.Show("Hoàn thành ! ");
                }
                else if (thangst == "04" || thangst == "06" || thangst == "09" || thangst == "11")
                {

                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT BANGCHAMCONGTHANG.STT, BANGCHAMCONGTHANG.MNV, BANGCHAMCONGTHANG.HOTEN, BANGCHAMCONGTHANG.CHUCVU, BANGCHAMCONGTHANG.PHONGBAN, BANGCHAMCONGTHANG.LNV, BANGCHAMCONGTHANG.[1],BANGCHAMCONGTHANG.[2],BANGCHAMCONGTHANG.[3]," +
                                            "BANGCHAMCONGTHANG.[4],BANGCHAMCONGTHANG.[5]," +
                                            "BANGCHAMCONGTHANG.[6],BANGCHAMCONGTHANG.[7]," +
                                            "BANGCHAMCONGTHANG.[8],BANGCHAMCONGTHANG.[9]," +
                                            "BANGCHAMCONGTHANG.[10],BANGCHAMCONGTHANG.[11]," +
                                            "BANGCHAMCONGTHANG.[12],BANGCHAMCONGTHANG.[13]," +
                                            "BANGCHAMCONGTHANG.[14],BANGCHAMCONGTHANG.[15]," +
                                       "BANGCHAMCONGTHANG.[16],BANGCHAMCONGTHANG.[17],BANGCHAMCONGTHANG.[18],BANGCHAMCONGTHANG.[19],BANGCHAMCONGTHANG.[20],BANGCHAMCONGTHANG.[21],BANGCHAMCONGTHANG.[22],BANGCHAMCONGTHANG.[23],BANGCHAMCONGTHANG.[24],BANGCHAMCONGTHANG.[25],BANGCHAMCONGTHANG.[26],BANGCHAMCONGTHANG.[27],BANGCHAMCONGTHANG.[28],BANGCHAMCONGTHANG.[29],BANGCHAMCONGTHANG.[30]FROM BANGCHAMCONGTHANG; ";
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(dt);
                    dataGridView3.DataSource = dt;
                    dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                    cnn.Close();
                    dataGridView3.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                    dataGridView3.Columns[2].HeaderText = "HỌ TÊN";
                    dataGridView3.Columns[3].HeaderText = "CHỨC VỤ";
                    dataGridView3.Columns[4].HeaderText = "PHÒNG BAN";
                    dataGridView3.Columns[5].HeaderText = "LNV";
                    for (int i = 0; i <= 29; i++)
                    {
                        DateTime txt3 = Convert.ToDateTime(textBox3.Text.ToString());
                        string thu = ((txt3.AddDays(+i).DayOfWeek).ToString());
                        if (thu == "Monday") thu = "T.2";
                        if (thu == "Tuesday") thu = "T.3";
                        if (thu == "Wednesday") thu = "T.4";
                        if (thu == "Thursday") thu = "T.5";
                        if (thu == "Friday") thu = "T.6";
                        if (thu == "Saturday") thu = "T.7";
                        if (thu == "Sunday") thu = "CN";


                        dataGridView3.Columns[i + 6].HeaderText = "(" + (i + 1) + ") " + thu;

                    }
                    dataGridView3.AutoResizeColumns();
                    dataGridView3.AutoResizeRows();
                    MessageBox.Show("Hoàn thành ! ");
                } 
                else if(thangst == "02")
               
                
                {
                    int namst = Convert.ToInt32((thang.ToString("yyyy")));
                    if ( (namst%400==0) || (namst%4==0 && namst%100!=0)  )
                    {
                        cnn.Open();
                        SqlCommand cmd = cnn.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT BANGCHAMCONGTHANG.STT, BANGCHAMCONGTHANG.MNV, BANGCHAMCONGTHANG.HOTEN, BANGCHAMCONGTHANG.CHUCVU, BANGCHAMCONGTHANG.PHONGBAN, BANGCHAMCONGTHANG.LNV, BANGCHAMCONGTHANG.[1],BANGCHAMCONGTHANG.[2],BANGCHAMCONGTHANG.[3]," +
                                                "BANGCHAMCONGTHANG.[4],BANGCHAMCONGTHANG.[5]," +
                                                "BANGCHAMCONGTHANG.[6],BANGCHAMCONGTHANG.[7]," +
                                                "BANGCHAMCONGTHANG.[8],BANGCHAMCONGTHANG.[9]," +
                                                "BANGCHAMCONGTHANG.[10],BANGCHAMCONGTHANG.[11]," +
                                                "BANGCHAMCONGTHANG.[12],BANGCHAMCONGTHANG.[13]," +
                                                "BANGCHAMCONGTHANG.[14],BANGCHAMCONGTHANG.[15]," +
                                                "BANGCHAMCONGTHANG.[16],BANGCHAMCONGTHANG.[17]," +
                                                "BANGCHAMCONGTHANG.[18],BANGCHAMCONGTHANG.[19]," +
                                                "BANGCHAMCONGTHANG.[20],BANGCHAMCONGTHANG.[21]," +
                                                "BANGCHAMCONGTHANG.[22],BANGCHAMCONGTHANG.[23]," +
                                                "BANGCHAMCONGTHANG.[24],BANGCHAMCONGTHANG.[25]," +
                                                "BANGCHAMCONGTHANG.[26],BANGCHAMCONGTHANG.[27]," +
                                                "BANGCHAMCONGTHANG.[28],BANGCHAMCONGTHANG.[29] FROM BANGCHAMCONGTHANG; ";
                        cmd.ExecuteNonQuery();
                        DataTable dt = new DataTable();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                        dataGridView3.DataSource = dt;
                        dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                        cnn.Close();
                        dataGridView3.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                        dataGridView3.Columns[2].HeaderText = "HỌ TÊN";
                        dataGridView3.Columns[3].HeaderText = "CHỨC VỤ";
                        dataGridView3.Columns[4].HeaderText = "PHÒNG BAN";
                        dataGridView3.Columns[5].HeaderText = "LNV";
                        for (int i = 0; i <= 28; i++)
                        {
                            DateTime txt3 = Convert.ToDateTime(textBox3.Text.ToString());
                            string thu = ((txt3.AddDays(+i).DayOfWeek).ToString());
                            if (thu == "Monday") thu = "T.2";
                            if (thu == "Tuesday") thu = "T.3";
                            if (thu == "Wednesday") thu = "T.4";
                            if (thu == "Thursday") thu = "T.5";
                            if (thu == "Friday") thu = "T.6";
                            if (thu == "Saturday") thu = "T.7";
                            if (thu == "Sunday") thu = "CN";


                            dataGridView3.Columns[i + 6].HeaderText = "(" + (i + 1) + ") " + thu;

                        }
                        dataGridView3.AutoResizeColumns();
                        dataGridView3.AutoResizeRows();
                        dataGridView3.Columns[2].Frozen = true;
                        MessageBox.Show("Hoàn thành ! ");

                    }
                    else
                    {

                        cnn.Open();
                        SqlCommand cmd = cnn.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = "SELECT BANGCHAMCONGTHANG.STT, BANGCHAMCONGTHANG.MNV, BANGCHAMCONGTHANG.HOTEN, BANGCHAMCONGTHANG.CHUCVU, BANGCHAMCONGTHANG.PHONGBAN, BANGCHAMCONGTHANG.LNV, BANGCHAMCONGTHANG.[1],BANGCHAMCONGTHANG.[2],BANGCHAMCONGTHANG.[3]," +
                                                "BANGCHAMCONGTHANG.[4],BANGCHAMCONGTHANG.[5]," +
                                                "BANGCHAMCONGTHANG.[6],BANGCHAMCONGTHANG.[7]," +
                                                "BANGCHAMCONGTHANG.[8],BANGCHAMCONGTHANG.[9]," +
                                                "BANGCHAMCONGTHANG.[10],BANGCHAMCONGTHANG.[11]," +
                                                "BANGCHAMCONGTHANG.[12],BANGCHAMCONGTHANG.[13]," +
                                                "BANGCHAMCONGTHANG.[14],BANGCHAMCONGTHANG.[15]," +
                                           "BANGCHAMCONGTHANG.[16],BANGCHAMCONGTHANG.[17],BANGCHAMCONGTHANG.[18],BANGCHAMCONGTHANG.[19],BANGCHAMCONGTHANG.[20],BANGCHAMCONGTHANG.[21],BANGCHAMCONGTHANG.[22],BANGCHAMCONGTHANG.[23],BANGCHAMCONGTHANG.[24],BANGCHAMCONGTHANG.[25],BANGCHAMCONGTHANG.[26],BANGCHAMCONGTHANG.[27],BANGCHAMCONGTHANG.[28] FROM BANGCHAMCONGTHANG; ";
                        cmd.ExecuteNonQuery();
                        DataTable dt = new DataTable();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dt);
                        dataGridView3.DataSource = dt;
                        dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                        cnn.Close();
                        dataGridView3.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                        dataGridView3.Columns[2].HeaderText = "HỌ TÊN";
                        dataGridView3.Columns[3].HeaderText = "CHỨC VỤ";
                        dataGridView3.Columns[4].HeaderText = "PHÒNG BAN";
                        for (int i = 0; i <= 27; i++)
                        {
                            DateTime txt3 = Convert.ToDateTime(textBox3.Text.ToString());
                            string thu = ((txt3.AddDays(+i).DayOfWeek).ToString());
                            if (thu == "Monday") thu = "T.2";
                            if (thu == "Tuesday") thu = "T.3";
                            if (thu == "Wednesday") thu = "T.4";
                            if (thu == "Thursday") thu = "T.5";
                            if (thu == "Friday") thu = "T.6";
                            if (thu == "Saturday") thu = "T.7";
                            if (thu == "Sunday") thu = "CN";


                            dataGridView3.Columns[i + 6].HeaderText = "(" + (i + 1) + ") " + thu;

                        }
                        dataGridView3.AutoResizeColumns();
                        dataGridView3.AutoResizeRows();
                        dataGridView3.Columns[2].Frozen = true;
                        MessageBox.Show("Hoàn thành ! ");


                    }





                }

            }
            if (bcc == "BANGCHAMCONGCHITIET")
            {
              
                cnn.Open();
                SqlCommand cmd2 = cnn.CreateCommand();
                cmd2.CommandType = CommandType.Text;
                cmd2.CommandText = "SELECT * FROM BANGCHAMCONGCHITIET  ";
                cmd2.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd2);
                da.Fill(dt);
                dataGridView3.DataSource = dt;
                cnn.Close();
                dataGridView3.AutoResizeColumns();
                dataGridView3.AutoResizeRows();
                dataGridView3.Sort(dataGridView3.Columns[0], ListSortDirection.Ascending);
                dataGridView3.Columns[2].Frozen = true;
                MessageBox.Show("Hoàn thành ! ");
            }
          
        }
        private void button12_Click(object sender, EventArgs e)
        {
            textBox2.Visible = true;
            textBox4.Visible = true;
            textBox5.Visible = true;
            textBox6.Visible = true;
            textBox7.Visible = true;
            textBox9.Visible = true;
            label5.Visible = true;
            label7.Visible = true;
            label8.Visible = true;
            label9.Visible = true;
            label10.Visible = true;
            label12.Visible = true;
           
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT KHUNGGIO.MKG, Format([STIME1],'HH:mm') AS STTIME1, Format([ETIME1],'HH:mm') AS ETTIME1, Format([STIME2],'HH:mm') AS STTIME2, Format([ETIME2],'HH:mm' ) AS ETTIME2, KHUNGGIO.LateTime,KHUNGGIO.SGLV FROM KHUNGGIO;";
            cmd.ExecuteNonQuery();
            SqlDataReader reader = cmd.ExecuteReader();
            while(reader.Read())
            {
                textBox2.Text = reader["STTIME1"].ToString();
                textBox4.Text = reader["ETTIME1"].ToString();
                textBox6.Text = reader["STTIME2"].ToString();
                textBox5.Text = reader["ETTIME2"].ToString();
                textBox7.Text = reader["LateTime"].ToString();
                textBox9.Text = reader["SGLV"].ToString();

            }

            cnn.Close();

        }    
        private void button13_Click(object sender, EventArgs e)
        {
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            for (int i = 1; i <= 31; i++)
            {
                // BẢNG CHẤM CÔNG CHI TIẾT //
                cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + i + ")] = CheckinT.[" + i + "] , [OUT(" + i + ")]  = CheckoutT.[" + i + "] FROM (CheckinT INNER JOIN BANGCHAMCONGCHITIET ON BANGCHAMCONGCHITIET.MNV = CheckinT.MNV) INNER JOIN CheckoutT ON BANGCHAMCONGCHITIET.MNV = CheckoutT.MNV;";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE [dbo].[BANGCHAMCONGCHITIET] SET [IN(" + i + ")] = NULL WHERE [IN(" + i + ")] = '00:00'";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE [dbo].[BANGCHAMCONGCHITIET] SET [OUT(" + i + ")] = NULL WHERE [OUT(" + i + ")] = '00:00'";
                cmd.ExecuteNonQuery();

            }

            for (int i = 1; i <= 31; i++)
            {
                // Tính thời gian làm việc buổi sáng //
                cmd.CommandText = "UPDATE TLVS SET [" + i + "] =  ((DATEPART(hh,KHUNGGIO.[ETIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[ETIME1]))-((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) FROM (TLVS INNER JOIN CheckinT ON TLVS.MNV = CheckinT.MNV),KHUNGGIO WHERE(TLVS.[MNV] =CheckinT.[MNV])";
                cmd.ExecuteNonQuery();


                cmd.CommandText = "UPDATE TLVS SET [" + i + "] = CASE when [" + i + "] < 0 then 0 when [" + i + "] >= 240 then 240 when ([" + i + "] < 180) then 0 ELSE [" + i + "] END " +
                                 "FROM(DANHSACHNHANVIEN INNER JOIN TLVS ON DANHSACHNHANVIEN.MNV = TLVS.MNV)" +
                                 " WHERE TLVS.MNV = DANHSACHNHANVIEN.MNV";
                cmd.ExecuteNonQuery();

              

            }
            
            try
            {

                for (int i = 1; i <= 31; i++)
                {
                    // Tính thời gian làm việc buổi chiều //
                   

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] =  ((DATEPART(hh,CheckoutT.[" + i + "])*60)+DATEPART(MINUTE,CheckoutT.[" + i + "])) - ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2]))" +
                                    "FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                    " WHERE ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2]))" +
                                    " AND DANHSACHNHANVIEN.MNV = TLVC.MNV; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] =  ((DATEPART(hh,CheckoutT.[" + i + "])*60)+DATEPART(MINUTE,CheckoutT.[" + i + "])) - ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "]))" +
                                      "FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                      " WHERE ((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) > ((DATEPART(hh,KHUNGGIO.[STIME2])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME2])) " +
                                      " AND((DATEPART(hh,CheckinT.[" + i + "])*60)+DATEPART(MINUTE,CheckinT.[" + i + "])) <= 840" +
                                      "AND DANHSACHNHANVIEN.MNV = TLVC.MNV; ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] = CASE WHEN [" + i + "] < 0 THEN 0 ELSE [" + i + "] END FROM (DANHSACHNHANVIEN INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV) WHERE TLVC.MNV = DANHSACHNHANVIEN.MNV ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE TLVC SET [" + i + "] = 0 " +
                                      " FROM(DANHSACHNHANVIEN INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV) INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,KHUNGGIO" +
                                      " WHERE((DATEPART(hh, CheckinT.[" + i + "]) * 60) + DATEPART(MINUTE, CheckinT.[" + i + "])) > ((DATEPART(hh, KHUNGGIO.[STIME2]) * 60) + DATEPART(MINUTE, KHUNGGIO.[STIME2])) + 30  AND DANHSACHNHANVIEN.MNV = TLVC.MNV";
                    cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE TLVC SET [" + i + "] = '0' FROM (DANHSACHNHANVIEN INNER JOIN TLVC ON DANHSACHNHANVIEN.MNV = TLVC.MNV),THOIGIANCHAMCONG WHERE DATEPART(weekday, convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + i + ")) = 7";
                    //cmd.ExecuteNonQuery();

                }
            }
            catch { }

            for (int i = 1; i <= 31; i++)
            {
                // Tính tổng thời gian làm việc //
                cmd.CommandText = "UPDATE TGLV SET [" + i + "] = convert(int,TLVS.[" + i + "]) + convert(int,TLVC.[" + i + "]) FROM (TGLV INNER JOIN TLVC ON TGLV.MNV = TLVC.MNV) INNER JOIN TLVS ON TGLV.MNV = TLVS.MNV WHERE TGLV.MNV=TLVS.MNV;";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "UPDATE TGLV SET [" + i + "] = 0 FROM (TGLV INNER JOIN CheckinT ON TGLV.MNV = CheckinT.MNV) INNER JOIN CheckoutT ON TGLV.MNV = CheckoutT.MNV WHERE CheckinT.[" + i + "]=CheckoutT.[" + i + "] ;";
                cmd.ExecuteNonQuery();

            }

            for (int i = 1; i <= 31; i++)
            {
                // UPDATE thời gian làm việc //
                cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + i + "] = TGLV.[" + i + "] FROM (BANGCHAMCONGTHANG INNER JOIN TGLV ON BANGCHAMCONGTHANG.MNV = TGLV.MNV) WHERE BANGCHAMCONGTHANG.MNV = TGLV.MNV;";
                cmd.ExecuteNonQuery();

            }


            try
            {

                for (int a = 1; a <= 31; a++)
                {

                    // Chấm công cho nhân viên ( TV ,CT ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') " +
                        "AND  convert(int,TGLV.[" + a + "]) >= 720 " +
                        "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) < ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                        "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7";
                    cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                    //    "convert(int,TGLV.[" + a + "]) >= 480 " +
                    //    "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) >= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                    //    "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                    //   "convert(int,TGLV.[" + a + "]) >= 480 " +
                    //   "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) + 30 )" +
                    //   "AND  DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                    //cmd.ExecuteNonQuery();


                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV=CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                    //    "convert(int,TGLV.[" + a + "]) >= 240  AND convert(int,TGLV.[" + a + "]) < 480 AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG WHERE" +
                    //    " (convert(int,TGLV.[" + a + "]) < 240 OR convert(int,TGLV.[" + a + "]) = 0) AND (DANHSACHNHANVIEN.LNV='CT' OR DANHSACHNHANVIEN.LNV='TV') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) <> 7 ";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') " +
                    //                  "AND  ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +                                   
                    //                  "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7";
                    //cmd.ExecuteNonQuery();


                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'M/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') " +                   
                    //  "AND ((DATEPART(hh,CheckinT.[" + a + "])*60)+DATEPART(MINUTE,CheckinT.[" + a + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                     
                    //  "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7 ";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = 'KP/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] ='CT' OR DANHSACHNHANVIEN.[LNV] ='TV') AND " +
                    // "convert(int,TGLV.[" + a + "]) <= 180" +
     
                    //" AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + a + "))) = 7 ";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + a + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + a + ")) = 1";
                    //cmd.ExecuteNonQuery();

                    //cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + a + ")] = NULL,[OUT(" + a + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + a + ")) = 1";
                    //cmd.ExecuteNonQuery();

                }





            }
            catch { }

            try
            {

                for (int b = 1; b <= 31; b++)
                {

                    // Chấm công cho nhân viên ( ĐTF ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                        "AND  convert(int,TGLV.[" + b + "]) >= 480 AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) < ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                           "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                        "convert(int,TGLV.[" + b + "]) >= 480 " +
                        "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) >= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                        "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                       "convert(int,TGLV.[" + b + "]) >= 480 " +
                       "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) + 30 )" +
                       "AND  DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                    cmd.ExecuteNonQuery();


                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV=CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                        "convert(int,TGLV.[" + b + "]) >= 240  AND convert(int,TGLV.[" + b + "]) < 480 AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG WHERE" +
                        " (convert(int,TGLV.[" + b + "]) < 240 OR convert(int,TGLV.[" + b + "]) = 0) AND (DANHSACHNHANVIEN.LNV=N'ĐTF') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) <> 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                                      "AND  ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                                      "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7";
                    cmd.ExecuteNonQuery();


                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'M/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV INNER JOIN CheckoutT ON DANHSACHNHANVIEN.MNV = CheckoutT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') " +
                      "AND ((DATEPART(hh,CheckinT.[" + b + "])*60)+DATEPART(MINUTE,CheckinT.[" + b + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +

                      "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTF') AND " +
                     "convert(int,TGLV.[" + b + "]) <= 180" +

                    " AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + b + "))) = 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + b + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + b + ")) = 1";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + b + ")] = NULL,[OUT(" + b + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + b + ")) = 1";
                    cmd.ExecuteNonQuery();




                }
            }
            catch { }

            try
            {

                for (int c = 1; c <= 31; c++)
                {

                    // Chấm công cho nhân viên ( ĐTP ) từ T.2 đến T.7 - BẢNG CHẤM CÔNG THÁNG //

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') " +
                        "AND  convert(int,TGLV.[" + c + "]) >= 480 " +
                        "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                        "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') " +
                        "AND  convert(int,TGLV.[" + c + "]) >= 240 AND convert(int,TGLV.[" + c + "]) < 480 " +
                       "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'M' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] = N'ĐTP') AND " +
                          "convert(int,TGLV.[" + c + "]) >= 480 " +
                          "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) > ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + convert(int,KHUNGGIO.[LateTime]) )" +
                          "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1]) + 60 ) AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7 ";
                    cmd.ExecuteNonQuery();



                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'KL' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV,THOIGIANCHAMCONG WHERE" +
                            " (convert(int,TGLV.[" + c + "]) < 240 OR convert(int,TGLV.[" + c + "]) = 0) AND (DANHSACHNHANVIEN.LNV=N'ĐTP') AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) <> 7 ";
                    cmd.ExecuteNonQuery();


                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'x/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV = CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTP') " +
                        "AND  convert(int,TGLV.[" + c + "]) >= 240 " +
                        "AND ((DATEPART(hh,CheckinT.[" + c + "])*60)+DATEPART(MINUTE,CheckinT.[" + c + "])) <= ((DATEPART(hh,KHUNGGIO.[STIME1])*60)+DATEPART(MINUTE,KHUNGGIO.[STIME1])+convert(int,KHUNGGIO.[LateTime]))  " +
                        "AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) = 7";
                    cmd.ExecuteNonQuery();


                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = 'KL/2' FROM(DANHSACHNHANVIEN INNER JOIN TGLV ON DANHSACHNHANVIEN.MNV=TGLV.MNV) INNER JOIN BANGCHAMCONGTHANG ON DANHSACHNHANVIEN.MNV = BANGCHAMCONGTHANG.MNV INNER JOIN CheckinT ON DANHSACHNHANVIEN.MNV=CheckinT.MNV,THOIGIANCHAMCONG,KHUNGGIO WHERE (DANHSACHNHANVIEN.[LNV] =N'ĐTP') AND " +
                        "convert(int,TGLV.[" + c + "]) < 240 " +
                       " AND DATEPART(weekday, (convert(nvarchar,DATEPART(yyyy,THOIGIANCHAMCONG.[START])) + '-' +convert(nvarchar,DATEPART(month,THOIGIANCHAMCONG.[START])) + '-' + convert(nvarchar," + c + "))) = 7 ";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGTHANG SET [" + c + "] = NULL FROM BANGCHAMCONGTHANG,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + c + ")) = 1";
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = "UPDATE BANGCHAMCONGCHITIET SET [IN(" + c + ")] = NULL,[OUT(" + c + ")] = NULL FROM BANGCHAMCONGCHITIET,THOIGIANCHAMCONG WHERE DATEPART(weekday,convert(nvarchar, DATEPART(yyyy,THOIGIANCHAMCONG.[START])) +'-' + convert(nvarchar, DATEPART(month,THOIGIANCHAMCONG.[START])) +'-'+ convert(nvarchar, " + c + ")) = 1";
                    cmd.ExecuteNonQuery();

                }
            }
            catch { }




            cnn.Close();
            MessageBox.Show("OK");
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }
        private void button14_Click(object sender, EventArgs e)
        {
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;

           cmd.CommandText = "UPDATE KHUNGGIO SET STIME1='" + textBox2.Text + "',ETIME1='" + textBox4.Text + "',STIME2='" + textBox6.Text + "',ETIME2='" + textBox5.Text + "',LateTime='" + textBox7.Text + "',SGLV='" + textBox9.Text + "'";

           
            cmd.ExecuteNonQuery();

            cnn.Close();
            MessageBox.Show("Lưu cấu hình thành công !","Thông báo",MessageBoxButtons.OK,MessageBoxIcon.Warning);
        }
        private void button15_Click(object sender, EventArgs e)
        {
           
           
           

           
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            LNV = "CT";
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            LNV = "TV";
        }
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            LNV = "ĐTF";
        }
        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            LNV = "ĐTP";
        }
        private void button21_Click(object sender, EventArgs e)
        {
            if (txtMNV2.Text != "" && txthoten2.Text != "" && txtchucdanh2.Text != "" && txtphongban2.Text != "" && LNV2 != "")
            {
                int hang = dataGridView4.Rows.Count -1;
                txtSTT2.Text = (hang + 1).ToString();
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO UT (STT,MNV,HOTEN,CHUCVU,PHONGBAN,LNV) VALUES(N'" + txtSTT2.Text + "',N'" + txtMNV2.Text + "',N'" + txthoten2.Text + "',N'" + txtchucdanh2.Text + "',N'" + txtphongban2.Text + "',N'" + LNV2+ "')";
                cmd.ExecuteNonQuery();

                cmd.CommandText = "SELECT * FROM UT ";
                cmd.ExecuteNonQuery();

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView4.DataSource = dt;
                cnn.Close();
                dataGridView4.Sort(dataGridView4.Columns[0], ListSortDirection.Ascending);

                MessageBox.Show("Thêm nhân viên ưu tiên thành công !", "Thông báo");
            }
            else { MessageBox.Show("Không được để trống thông tin nhân viên !", "Thông báo"); }
        }
        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            LNV2 = "CT";
        }
        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            LNV2 = "TV";
        }
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            LNV2 = "ĐTF";
        }
        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            LNV2 = "ĐTP";
        }
        private void button18_Click(object sender, EventArgs e)
        {

          
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM UT ";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView4.DataSource = dt;
            cmd.CommandText = "SELECT * FROM UT ";
            cmd.ExecuteNonQuery();
            // dataGridView1.Columns.Remove("ID");
            cnn.Close();
            dataGridView4.Sort(dataGridView4.Columns[0], ListSortDirection.Ascending);
            dataGridView4.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
            dataGridView4.Columns[2].HeaderText = "HỌ TÊN";
            dataGridView4.Columns[3].HeaderText = "CHỨC VỤ";
            dataGridView4.Columns[4].HeaderText = "PHÒNG BAN";
            

        }
        private void button19_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Bạn muốn xoá nhân viên ưu tiên !", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {

                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;

                cmd.CommandText = "DELETE FROM UT WHERE MNV = '" + txtMNV2.Text + "'";
                cmd.ExecuteNonQuery();

                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT * FROM UT ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.Sort(dataGridView4.Columns[0], ListSortDirection.Ascending);
                int hang = dataGridView4.Rows.Count - 2;
                for (int i = 0; i <= hang; i++)
                {
                    dataGridView4.Rows[i].Cells[0].Value = i + 1;
                    int STT = Convert.ToInt32(dataGridView4.Rows[i].Cells[0].Value);
                    cmd.CommandType = CommandType.Text;

                    cmd.CommandText = "UPDATE UT SET STT= '" + STT + "' WHERE MNV = '" + dataGridView4.Rows[i].Cells[1].Value.ToString() + "'";
                    cmd.ExecuteNonQuery();
                }
                    // dataGridView1.Columns.Remove("ID");
                    cnn.Close();
                dataGridView4.Sort(dataGridView4.Columns[0], ListSortDirection.Ascending);
                dataGridView4.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                dataGridView4.Columns[2].HeaderText = "HỌ TÊN";
                dataGridView4.Columns[3].HeaderText = "CHỨC VỤ";
                dataGridView4.Columns[4].HeaderText = "PHÒNG BAN";
              
            }
        }
        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }
        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }
        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
               
                DataGridViewRow row = this.dataGridView4.Rows[e.RowIndex];
                txtSTT2.Text = row.Cells[0].Value.ToString();
                txthoten2.Text = row.Cells[2].Value.ToString();
                txtMNV2.Text = row.Cells[1].Value.ToString();
                txtchucdanh2.Text = row.Cells[3].Value.ToString();
                txtphongban2.Text = row.Cells[4].Value.ToString();
                string LNVLD = row.Cells[5].Value.ToString();
                if (LNVLD == "CT") { radioButton8.Checked = true; }
                else if (LNVLD == "TV") { radioButton7.Checked = true; }
               
                else { MessageBox.Show("Loại nhân viên không xác định !"); }

            }
        }
        private void button22_Click(object sender, EventArgs e)
        {
            

        }
        private void button23_Click(object sender, EventArgs e)
        {
           
           

        }
        private void button24_Click(object sender, EventArgs e)
        {
           



            }
        private void button22_Click_1(object sender, EventArgs e)
        {

            try
            {
                dataGridView2.Columns.Clear();
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Select file";
                fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "txt File(*.txt)|*.txt|All Files(*.*)|*.*";
                fdlg.FilterIndex = 1;
                fdlg.RestoreDirectory = true;

                if (fdlg.ShowDialog() == DialogResult.OK)
                {

                    textBox1.Text = fdlg.FileName;
                    Application.DoEvents();

                }
                var lines = File.ReadAllLines(textBox1.Text);
                if (lines.Count() > 0)
                {
                    foreach (var columnName in lines.FirstOrDefault()
                        .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        dataGridView2.Columns.Add(columnName, columnName);
                    }
                    foreach (var cellValues in lines.Skip(1))
                    {
                        var cellArray = cellValues
                            .Split(new[] { '\t' }, StringSplitOptions.RemoveEmptyEntries);
                        if (cellArray.Length == dataGridView2.Columns.Count)
                            dataGridView2.Rows.Add(cellArray);
                    }
                }

                if (dataGridView2.Rows.Count > 0)
                {
                    string thang = dataGridView2.Rows[2].Cells[1].Value.ToString();
                    string thang2 = dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[1].Value.ToString();
                    cnn.Open();
                    SqlCommand cmd = cnn.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "UPDATE [dbo].[THOIGIANCHAMCONG] SET [START]  = '" + thang.Substring(0, 7) + "-01' WHERE ID = 1";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "UPDATE [dbo].[THOIGIANCHAMCONG] SET [END] = '" + thang2.Substring(0, 10) + "' WHERE ID = 1;";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = "SELECT Format([START],'yyyy-MM-dd') AS START1, Format([END],'yyyy-MM-dd') AS END1 FROM THOIGIANCHAMCONG; ";
                    cmd.ExecuteNonQuery();
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        textBox3.Text = (reader["START1"]).ToString();
                        textBox8.Text = reader["END1"].ToString();

                    }
                    cnn.Close();

                    dateTimePicker1.Value = new DateTime(Convert.ToInt32(textBox3.Text.Substring(0, 4)), Convert.ToInt32(textBox3.Text.Substring(5, 2)), Convert.ToInt32(textBox3.Text.Substring(8, 2)));
                    dateTimePicker2.Value = new DateTime(Convert.ToInt32(textBox8.Text.Substring(0, 4)), Convert.ToInt32(textBox8.Text.Substring(5, 2)), Convert.ToInt32(textBox8.Text.Substring(8, 2)));
                }
                MessageBox.Show("Tải file thành công !");
            }
            catch
            {

                MessageBox.Show("Tệp chưa được chọn !");

            }

        }
        private void button20_Click(object sender, EventArgs e)
        {

        }
        private void button17_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tabPage5);
        }
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dataGridView3.Rows[e.RowIndex].Selected = true;

            }
        }
        private void btnfind_Click(object sender, EventArgs e)
        {
            if (textfind.Text != "")
            {
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [STT],[MNV],[HOTEN],[CHUCDANH],[PHONGBAN],[LNV] FROM [dbo].[DANHSACHNHANVIEN] WHERE [MNV] = N'" + textfind.Text + "' OR [HOTEN] = N'" + textfind.Text + "' OR CHUCDANH = N'" + textfind.Text + "' OR PHONGBAN = N'" + textfind.Text + "' OR LNV = N'" + textfind.Text + "'";
                cmd.ExecuteNonQuery();
                SqlDataReader dr = cmd.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(dr);
                dataGridView1.DataSource = dt;
                cnn.Close();
            }
            else
            {

                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "SELECT [STT],[MNV],[HOTEN],[CHUCDANH],[PHONGBAN],[LNV] FROM [dbo].[DANHSACHNHANVIEN] ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Sort(dataGridView1.Columns[0], ListSortDirection.Ascending);
                dataGridView1.Columns[1].HeaderText = "MÃ NHÂN VIÊN";
                dataGridView1.Columns[2].HeaderText = "HỌ TÊN";
                dataGridView1.Columns[3].HeaderText = "CHỨC VỤ";
                dataGridView1.Columns[4].HeaderText = "PHÒNG BAN";
                dataGridView1.AutoResizeColumns();
                dataGridView1.AutoResizeRows();
                cnn.Close();

            }

        }
        private void button23_Click_1(object sender, EventArgs e)
        {

        }
        private void button25_Click(object sender, EventArgs e)
        {
            if (txtusename.Text != "" && txtpassword.Text != "" && cbtypeac.Text !="" )
            {
              
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "INSERT INTO Account (Username,Password,Permission) VALUES(N'" + txtusename.Text + "',N'" + txtpassword.Text + "',N'" + cbtypeac.Text + "')";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "SELECT * FROM Account ";
                cmd.ExecuteNonQuery();

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView5.DataSource = dt;
                cnn.Close();
                dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);

                MessageBox.Show("Thêm tài khoản thành công !", "Thông báo");
            }
            else { MessageBox.Show("Không được để trống thông tin tài khoản !", "Thông báo"); }
        }
        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {

                DataGridViewRow row = this.dataGridView5.Rows[e.RowIndex];

                txtusename.Text = row.Cells[1].Value.ToString();
                txtpassword.Text = row.Cells[2].Value.ToString();
                cbtypeac.Text = row.Cells[3].Value.ToString();
                textBox10.Text = row.Cells[0].Value.ToString();

            }
        }
        private void button27_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Bạn muốn thoát chương trình", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Application.Exit();

            }
            else if (dialogResult == DialogResult.No)
            {

            }
        }
        private void button28_Click(object sender, EventArgs e)
        {
            if (txtusename.Text != "" && txtpassword.Text != "" && cbtypeac.Text != "")
            {

                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "UPDATE Account SET [Username] = N'" + txtusename.Text + "', [Password] = N'" + txtpassword.Text + "',[Permission] = N'" + cbtypeac.Text + "' WHERE ID = N'" + textBox10.Text + "' ";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "SELECT * FROM Account ";
                cmd.ExecuteNonQuery();

                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView5.DataSource = dt;
                cnn.Close();
                dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);

                MessageBox.Show("Cập nhật tài khoản thành công !", "Thông báo");
            }
            else { MessageBox.Show("Không được để trống thông tin tài khoản !", "Thông báo"); }
        }
        private void button26_Click(object sender, EventArgs e)
        {
            if (txtusename.Text.ToString().Contains(userrel)==true && txtpassword.Text.ToString().Contains(passrel)==true)
            {

                MessageBox.Show("Tài khoản đang đăng nhập  !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            else
            {
                cnn.Open();
                SqlCommand cmd = cnn.CreateCommand();
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "DELETE FROM [dbo].[Account] WHERE Username = N'" + txtusename.Text + "'";
                cmd.ExecuteNonQuery();
                cmd.CommandText = "SELECT [ID], [Username],[Password],[Permission] FROM Account ";
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView5.DataSource = dt;
                cnn.Close();
                dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);
            }
        }
        private void button24_Click_1(object sender, EventArgs e)
        {
            cnn.Open();
            SqlCommand cmd = cnn.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT [ID], [Username],[Password],[Permission] FROM Account ";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView5.DataSource = dt;       
            cnn.Close();
           // dataGridView5.Sort(dataGridView5.Columns[0], ListSortDirection.Ascending);
         
        }
    }
}
  


    

