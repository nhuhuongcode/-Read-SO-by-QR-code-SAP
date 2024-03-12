using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Models
{
	internal class DetailItem_model
	{
		public string ItemCode { get; set; }
		public string Dscription { get; set; }
		public double Quantity { get; set; }
		public double Price { get; set; }
		public int BaseType { get; set; }
		public int BaseRef { get; set; }
		public int BaseLine { get; set; }
	}
}
