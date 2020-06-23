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
using COMExcel = Microsoft.Office.Interop.Excel;

namespace QuanLyCuaHangSach
{
    public partial class FrmBC7 : Form
    {
        DataTable tblBC7;
        public FrmBC7()
        {
            InitializeComponent();
        }

        private void FrmBC7_Load(object sender, EventArgs e)
        {
            DAO.OpenConnection();
            btnIn.Enabled = true;
            btnHienThi.Enabled = true;
            btnThoat.Enabled = true;
            string sql = "SELECT SoHDN,MaNV,NgayNhap,MaNCC,TongTien FROM HoaDonNhap";
            tblBC7 = DAO.GetDataToTable(sql);
            dataGridView1.DataSource = tblBC7;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
            DAO.CloseConnetion();
        }

        private void btnHienThi_Click(object sender, EventArgs e)
        {
             if (cboNam.Text == "") 
            {
                MessageBox.Show("Hãy nhập đủ điều kiện!!!", "Yêu cầu ...", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            loadGridbyngaythang();

            string sql = "SELECT SoHDN,MaNV,NgayNhap,MaNCC,TongTien FROM HoaDonNhap WHERE NgayNhap LIKE '%" + cboNam.Text + "%')";
            tblBC7 = DAO.GetDataToTable(sql);
            DAO.RunSql(sql);
            
        }
        public void loadGridbyngaythang()
        {
            DAO.OpenConnection();
            string sql;
            sql = "select SoHDN,MaNV,NgayNhap,MaNCC,TongTien FROM HoaDonNhap WHERE NgayNhap LIKE '%" + cboNam.Text + "%'";
            DAO.RunSql(sql);
            dataGridView1.DataSource = tblBC7;
            //DAO.CloseConnetion();


        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBatDau_Click(object sender, EventArgs e)
        {
            DAO.OpenConnection();
            cboNam.Enabled = true;
            btnHienThi.Enabled = true;
            btnIn.Enabled = true;
            cboNam.Text = "";
            cboNam.Focus();
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook;
            COMExcel.Worksheet exSheet;
            COMExcel.Range exRange;
            string sql;
            int hang = 0, cot = 0;
            DataTable danhsach;

            exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            exSheet = exBook.Worksheets[1];
            exRange = exSheet.Cells[1, 1];
            exRange.Range["A1:Z300"].Font.Name = "Times new roman";
            exRange.Range["A1:B3"].Font.Size = 10;
            exRange.Range["A1:B3"].Font.Bold = true;
            exRange.Range["A1:B3"].Font.ColorIndex = 5;
            exRange.Range["A1:A1"].ColumnWidth = 10;
            exRange.Range["B1:B1"].ColumnWidth = 20;
            exRange.Range["A1:B1"].MergeCells = true;
            exRange.Range["A1:B1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A1:B1"].Value = "QUẢN LÝ CỬA HÀNG SÁCH";
            exRange.Range["A2:B2"].MergeCells = true;
            exRange.Range["A2:B2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A2:B2"].Value = "Cầu Giấy - Hà Nội";
            exRange.Range["A3:B3"].MergeCells = true;
            exRange.Range["A3:B3"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A3:B3"].Value = "Điện thoại: (04)37562222";
            exRange.Range["C2:G2"].Font.Size = 16;
            exRange.Range["C2:G2"].Font.Bold = true;
            exRange.Range["C2:G2"].Font.ColorIndex = 3;
            exRange.Range["C2:G2"].MergeCells = true;
            exRange.Range["C2:G2"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["C2:G2"].Value = "Báo cáo danh sách 5 hóa đơn có tổng tiền nhập hàng lớn nhất theo năm";
            sql = "SELECT TOP 5  SoHDN,MaNV,NgayNhap,MaNCC,SUM(TongTien) FROM HoaDonNhap WHERE NgayNhap LIKE '%" + cboNam.Text + "%'" ;
            
            danhsach = DAO.GetDataToTable(sql);
            exRange.Range["B5:G5"].Font.Bold = true;
            exRange.Range["B5:G5"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B5:B5"].ColumnWidth = 13;
            exRange.Range["C5:C5"].ColumnWidth = 30;
            exRange.Range["D5:D5"].ColumnWidth = 25;
            exRange.Range["E5:E5"].Value = "STT";
            exRange.Range["F5:F5"].Value = "Số hóa đơn nhập";
            exRange.Range["G5:G5"].Value = "Mã nhân viên";
            exRange.Range["H5:H5"].Value = "Ngày nhập";
            exRange.Range["I5:I5"].Value = "Mã nhà cung cấp";
            exRange.Range["K5:K5"].Value = "Tổng tiền";

            for (hang = 0; hang < danhsach.Rows.Count; hang++)
            {
                exSheet.Cells[2][hang + 6] = hang + 1;
                for (cot = 0; cot < danhsach.Columns.Count; cot++)
                {
                    exSheet.Cells[cot + 3][hang + 6] = danhsach.Rows[hang][cot].ToString();
                }
            }
            exRange = exSheet.Cells[2][hang + 8];
            exRange.Range["D1:F1"].MergeCells = true;
            exRange.Range["D1:F1"].Font.Italic = true;
            exRange.Range["D1:F1"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["D1:F1"].Value = "Hà Nội, " + DateTime.Now.ToShortDateString();
            exSheet.Name = "Báo cáo";
            exApp.Visible = true;
        }
       
    }
}
        

        
        
    

