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
    public partial class FrmBC6 : Form
    {
        DataTable tblBC6;
        public FrmBC6()
        {
            InitializeComponent();
        }

        private void FrmBC6_Load(object sender, EventArgs e)
        {
            DAO.OpenConnection();
            btnIn.Enabled = true;
            btnHienThi.Enabled = true;
            btnThoat.Enabled = true;
            string sql = "SELECT MaSach,TenSach,SoLuong,DonGiaNhap,DonGiaBan,MaLoaiSach,MaTG,MaNXB,MaLinhVuc,MaNgonNgu,Anh,SoTrang FROM KhoSach";
            tblBC6 = DAO.GetDataToTable(sql);
            dataGridView1.DataSource = tblBC6;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
            DAO.CloseConnetion();
        }
        private void ResetValues()
        {
            cboQuy.Text = "";
        }

        private void btnHienThi_Click(object sender, EventArgs e)
        {
            DAO.OpenConnection();
            string sql = "SELECT a.MaSach,TenSach,SoLuong,DonGiaNhap,DonGiaBan,MaLoaiSach,MaTG,MaNXB,MaLinhVuc,MaNgonNgu,Anh,SoTrang" +
                "FROM KhoSach a join ChiTietHDB b on a.MaSach =b.MaSach" +
                "FROM ChiTietHDB b join HoaDonBan c on b.SoHDB =c.SoHDB" +
                "WHERE NOT EXISTS (select c.SoHDB,NgayBan AS c WHERE NgayBan LIKE '%" + cboQuy.Text + "%')";
            if (cboQuy.Text == "")
                MessageBox.Show("Bạn phải chọn quý!");
            if (cboQuy.Text == "1")
            {
                sql = sql + "AND NgayBan >= '2020-01-01' AND NgayBan <= '2020-03-31'";
            }
            if (cboQuy.Text == "2")
            {
                sql = sql + "AND NgayBan >= '2020-04-01' AND NgayBan <= '2020-06-30'";
            }
            if (cboQuy.Text == "3")
            {
                sql = sql + "AND NgayBan >= '2020-07-01' AND NgayBan <= '2020-09-30'";
            }
            if (cboQuy.Text == "4")
            {
                sql = sql + "AND NgayBan >= '2020-10-01' AND NgayBan <= '2020-12-31'";
            }
            tblBC6 = DAO.GetDataToTable(sql);
            DAO.RunSql(sql);
            DAO.CloseConnetion();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBatDau_Click(object sender, EventArgs e)
        {
            DAO.OpenConnection();
            cboQuy.Enabled = true;
            btnHienThi.Enabled = true;
            btnIn.Enabled = true;
            cboQuy.Text = "";
            cboQuy.Focus();
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
            exRange.Range["C2:G2"].Value = "Báo cáo danh sách sản phẩm không bán được theo quý";
            sql = "SELECT SELECT a.MaSach,TenSach,SoLuong,DonGiaNhap,DonGiaBan,MaLoaiSach,MaTG,MaNXB,MaLinhVuc,MaNgonNgu,Anh,SoTrang" +
                "FROM KhoSach a join ChiTietHDB b on a.MaSach =b.MaSach" +
                "FROM ChiTietHDB b join HoaDonBan c on b.SoHDB =c.SoHDB" +
                "WHERE NOT EXISTS (select c.SoHDB,NgayBan AS c WHERE NgayBan LIKE '%" + cboQuy.Text + "%')";  
            if (cboQuy.Text == "")
                MessageBox.Show("Bạn phải chọn quý!");
            if (cboQuy.Text == "1")
            {
                sql = sql + "AND NgayBan >= '2020-01-01' AND NgayBan <= '2020-03-31'";
            }
            if (cboQuy.Text == "2")
            {
                sql = sql + "AND NgayBan >= '2020-04-01' AND NgayBan <= '2020-06-30'";
            }
            if (cboQuy.Text == "3")
            {
                sql = sql + "AND NgayBan >= '2020-07-01' AND NgayBan <= '2020-09-30'";
            }
            if (cboQuy.Text == "4")
            {
                sql = sql + "AND NgayBan >= '2020-10-01' AND NgayBan <= '2020-12-31'";
            }
            danhsach = DAO.GetDataToTable(sql);
            exRange.Range["B5:G5"].Font.Bold = true;
            exRange.Range["B5:G5"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["B5:B5"].ColumnWidth = 13;
            exRange.Range["C5:C5"].ColumnWidth = 30;
            exRange.Range["D5:D5"].ColumnWidth = 25;
            exRange.Range["B5:B5"].Value = "STT";
            exRange.Range["C5:C5"].Value = "Mã Sách";
            exRange.Range["D5:D5"].Value = "Tên Sách";
            exRange.Range["E5:E5"].Value = "Số Lượng";
            exRange.Range["F5:F5"].Value = "Đơn giá nhập";
            exRange.Range["G5:G5"].Value = "Đơn giá bán";
            exRange.Range["H5:H5"].Value = "Mã loại sách";
            exRange.Range["I5:I5"].Value = "Mã tác giả";
            exRange.Range["J5:J5"].Value = "Mã nhà xuất bản";
            exRange.Range["K5:K5"].Value = "Mã lĩnh vực";
            exRange.Range["L5:L5"].Value = "Mã ngôn ngữ";
            exRange.Range["M5:M5"].Value = "Ảnh";
            exRange.Range["N5:N5"].Value = "Số trang";

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
            DAO.CloseConnetion();

        }
    }



    }

