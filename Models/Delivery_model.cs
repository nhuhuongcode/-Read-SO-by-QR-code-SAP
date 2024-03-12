using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Models
{
	internal class Delivery_model
	{
		public int DocNum { get; set; }
		public DateTime DocDate { get; set; }
		public DateTime DocDueDate { get; set; }
		public string CardCode { get; set; }
		public string CardName { get; set; }
		public string Comments { get; set; }
		public string Address { get; set; }
		public string EDocStatus = "C";
		public List<DetailItem_model> detailItems { get; set;}

		public Delivery_model()
		{
			detailItems = new List<DetailItem_model>();
		}
	}
}
