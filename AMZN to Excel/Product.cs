using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMZN_to_Excel
{
    class Product
    {
		public String name
		{
			get;
			set;
		}

		public double price
		{
			get;
			set;
		}
		public double xprice
		{
			get;
			set;
		}
		public double dif
		{
			get;
			set;
		}
		public String category
		{
			get;
			set;
		}
		public String URL
		{
			get;
			set;
		}
		public String ID
		{
			get;
			set;
		}


		public Product(String Name, double Price, double XP, String ctgry, String url, String GUID)
		{
			name = Name;
			price = Price;
			xprice = XP;
			dif = xprice - price;
			category = ctgry;
			URL = url;
			ID = GUID;
		}
	}
}
