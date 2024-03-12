using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
	internal class Connect
	{
		private static SAPbouiCOM.Application oApp;
		private static SAPbobsCOM.Company oCompany;
		public static SAPbouiCOM.Application SapApplication { get { return oApp; } }
		public static SAPbobsCOM.Company SapCompany { get { return oCompany; } }

		private void SetAppication()
		{
			SAPbouiCOM.SboGuiApi sboGuiApi = null;
			string sCon = null;
			sboGuiApi = new SAPbouiCOM.SboGuiApi();
			sCon = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
			sboGuiApi.Connect(sCon);
			oApp = sboGuiApi.GetApplication(-1);
		}

		public Connect()
		{
			try
			{
				SetAppication();
				oCompany = (SAPbobsCOM.Company)oApp.Company.GetDICompany();
				MessageBox.Show("Connected!");
			}catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			
		}
	}
}
