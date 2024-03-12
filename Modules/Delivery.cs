using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using WindowsFormsApp1.Models;

namespace WindowsFormsApp1.Modules
{
	internal class Delivery
	{
		private SAPbouiCOM.Application SBO_Application;
		private SAPbobsCOM.Company oCompany;
		private SAPbouiCOM.Matrix oMatrix;
		public Delivery(SAPbouiCOM.Application SBO_Application, SAPbobsCOM.Company oCompany)
		{
			this.SBO_Application = SBO_Application;
			this.oCompany = oCompany;
		}

		
		public void AddDelivery(Delivery_model model)
		{
			try
			{
				SAPbobsCOM.Documents oMarketingDocument = null;
				SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);

				string ErrorMsg = string.Empty;
				int ErrorNo = 0;
				oMarketingDocument = ((SAPbobsCOM.Documents)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)));
				

				oMarketingDocument.CardCode = model.CardCode;
				oMarketingDocument.DocDate = model.DocDate;
				oMarketingDocument.DocDueDate = model.DocDueDate;
				oMarketingDocument.Comments = model.Comments;
				oMarketingDocument.Address = model.Address;
				oMarketingDocument.EDocStatus = EDocStatusEnum.edoc_Ok;

				for (int i = 0; i < model.detailItems.Count; i++)
				{
					oMarketingDocument.Lines.ItemCode = model.detailItems[i].ItemCode;
					oMarketingDocument.Lines.ItemDescription = model.detailItems[i].Dscription;
					oMarketingDocument.Lines.Quantity = model.detailItems[i].Quantity;
					oMarketingDocument.Lines.UnitPrice = model.detailItems[i].Price;
					oMarketingDocument.Lines.BaseLine= model.detailItems[i].BaseLine;
					oMarketingDocument.Lines.BaseType= model.detailItems[i].BaseType;
					oMarketingDocument.Lines.BaseEntry = model.detailItems[i].BaseRef;
					oMarketingDocument.Lines.Add();
				}
				//oMatrix.LoadFromDataSource();
				//int returnVal = oMarketingDocument.Add();
				if (oMarketingDocument.Add() == 0)
				{
					string DocEntry = oCompany.GetNewObjectKey();
					SBO_Application.StatusBar.SetText("True", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
					MessageBox.Show("True");
					//CloseDoc(model.detailItems[0].BaseRef);
				}
				else
				{
					ErrorMsg = oCompany.GetLastErrorDescription();
					ErrorNo = oCompany.GetLastErrorCode();
					MessageBox.Show(ErrorMsg);
				}


			}
			catch (Exception ex)
			{
				MessageBox.Show("Falsebig");
			}
		}

		public void CloseDoc(int para)
		{
			try
			{
				SAPbobsCOM.Documents saleOrder = null;
				SAPbobsCOM.SBObob objBridge = (SAPbobsCOM.SBObob)oCompany.GetBusinessObject(BoObjectTypes.BoBridge);

				string ErrorMsg = string.Empty;
				int ErrorNo = 0;
				saleOrder = ((SAPbobsCOM.Documents)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)));
				if (saleOrder.GetByKey(para))
				{
					// Close the document
					if (saleOrder.Close() == 0)
					{
						MessageBox.Show($"Document  closed successfully.");
					}
					else
					{
						MessageBox.Show($"Failed to close document. Error: {oCompany.GetLastErrorDescription()}");
					}
				}
				else
				{
					MessageBox.Show($"Document not found.");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"An error occurred: {ex.Message}");
			}
		}
	}
}
