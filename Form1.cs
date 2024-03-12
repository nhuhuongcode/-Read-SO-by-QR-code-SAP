using AddlinePromotion;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp1.Models;
using WindowsFormsApp1.Modules;
using System.Configuration;

namespace WindowsFormsApp1
{
	public partial class Form1 : System.Windows.Forms.Form
	{
		public static string query = System.Configuration.ConfigurationManager.AppSettings["query"].ToString();
		public int res { get; set; }
		SqlConnection cn = new SqlConnection(query);

		public SqlDataAdapter sda = new SqlDataAdapter();
		public DataSet ds = new DataSet();
		public Dictionary<string,int> dic;
		public BindingSource bindingSource = new BindingSource();
		public Form1()
		{
			InitializeComponent();
		}

		public void laydulieu(int para)
		{
			
			try
			{
				SqlConnection cn = new SqlConnection(query);
				cn.Open();
				SqlCommand cmd = new SqlCommand("SELECT t0.DocNum,t0.CardCode,t0.CardName,t1.ItemCode,t1.Dscription,t1.Quantity,t1.Price " +
					"FROM ORDR t0 INNER JOIN  RDR1 t1 ON t0.DocEntry = t1.DocEntry WHERE t0.DocNum = "+para, cn);
				sda.SelectCommand = cmd;
				sda.Fill(ds, "ChiTietDonHang");
				
				cn.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		public void Form1_Load(object sender, EventArgs e)
		{
			laydulieu(res);
			if (ds.Tables["ChiTietDonHang"].Rows.Count > 0)
			{
				
				string cardCode = ds.Tables["ChiTietDonHang"].Rows[0]["CardCode"].ToString();
				string cardName = ds.Tables["ChiTietDonHang"].Rows[0]["CardName"].ToString();
				tb_cardcode.Text = cardCode;
				tb_cardname.Text = cardName;

				var recordItem = new Dictionary<string, int>();
				for (int i = 0; i < ds.Tables["ChiTietDonHang"].Rows.Count; i++)
				{
					string codeItem = ds.Tables["ChiTietDonHang"].Rows[i]["ItemCode"].ToString();
					recordItem.Add(codeItem, i);
				}
				dic = recordItem;
			}
			
			
			ds.Tables["ChiTietDonHang"].Columns.Remove("DocNum");
			ds.Tables["ChiTietDonHang"].Columns.Remove("CardCode");
			ds.Tables["ChiTietDonHang"].Columns.Remove("CardName");
			bindingSource.DataSource = ds.Tables["ChiTietDonHang"];
			gv_rdr1.DataSource = bindingSource;
			gv_rdr1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
			gv_rdr1.AllowUserToDeleteRows = true;
			gv_rdr1.ReadOnly = true;
			tb_docnum.Text = res.ToString();
		}

		private void bt_Add_Click(object sender, EventArgs e)
		{
			if (bindingSource.Current != null)
			{
				DataRowView currentRowView = (DataRowView)bindingSource.Current;
				DataRow currentRow = currentRowView.Row;

				// Kiểm tra xem dòng có trạng thái Deleted hay không
				if (currentRow.RowState == DataRowState.Deleted)
				{
					bindingSource.RemoveCurrent(); // Xóa dòng khỏi BindingSource
					currentRow.AcceptChanges(); // Chấp nhận việc xóa dòng khỏi DataTable
				}
			}

			bindingSource.EndEdit();
			ds.Tables["ChiTietDonHang"].AcceptChanges();

			Delivery_model deliVM = new Delivery_model();
			Connect connect = new Connect();
			Delivery delivery = new Delivery(Connect.SapApplication, Connect.SapCompany);
			deliVM.DocNum = int.Parse(tb_docnum.Text);
			deliVM.DocDate = dt_docdate.Value;
			deliVM.DocDueDate = dt_docduedate.Value;
			deliVM.CardCode = tb_cardcode.Text;
			deliVM.CardName = tb_cardname.Text;
			deliVM.Address = "Testing";
			deliVM.Comments = "Testing";
			for (int i = 0; ds.Tables["ChiTietDonHang"].Rows.Count > i; i++)
			{
				DetailItem_model detailVM = new DetailItem_model
				{
					ItemCode = ds.Tables["ChiTietDonHang"].Rows[i]["ItemCode"].ToString(),
					Dscription = ds.Tables["ChiTietDonHang"].Rows[i]["Dscription"].ToString(),
					Quantity = double.Parse(ds.Tables["ChiTietDonHang"].Rows[i]["Quantity"].ToString()),
					Price = double.Parse(ds.Tables["ChiTietDonHang"].Rows[i]["Price"].ToString()),
					BaseType = 17,
					BaseRef = int.Parse(res.ToString()),
					BaseLine = dic[ds.Tables["ChiTietDonHang"].Rows[i]["ItemCode"].ToString()]
				};
				deliVM.detailItems.Add(detailVM);
			}
			//if (!Globals.SetApplication())
			//{
			//	if (Globals.SapApplication != null)
			//		MessageBox.Show("ERROR");
			//}

			//MessageBox.Show("CONNECTED!");

			delivery.AddDelivery(deliVM);
		}
	}
}
