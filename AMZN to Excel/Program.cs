using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Net;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace AMZN_to_Excel
{
    class Program
    {
        public static String pricesXPath = "//span[@data-a-strike='true' or contains(@class,'text-strike')][.//text()]/preceding::span[@class][3]";
        public static String XpricesXPath = "//span[@data-a-strike='true' or contains(@class,'text-strike')][.//text()]";
        public static String namesXPath = "//span[@data-a-strike='true' or contains(@class,'text-strike')][.//text()]/preceding::span[@class][10]";
        public static String picURLsXPath = "//span[@data-a-strike='true' or contains(@class,'text-strike')][.//text()]/preceding::img[@src][1]";
        public static String dealURLsXPath = "//span[@data-a-strike='true' or contains(@class,'text-strike')][.//text()]/parent::a[@class]";
        
        //Initialize Lists
        public static List<Product> products = new List<Product>();
        public static List<String> prices = new List<String>();
        public static List<String> names = new List<String>();
        public static List<String> Xprices = new List<String>();
        public static List<String> picURLs = new List<String>();
        public static List<String> URLs = new List<String>();
		public static List<Image> pics = new List<Image>();

		public static List<String> skippedURLs = new List<String>();


		static void Main(string[] args)
		{
			//Initialize Browser
			IWebDriver driver = new ChromeDriver();
			WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

			//Grab Data
			getInfo(driver, wait, "Laptops");
			getInfo(driver, wait, "Desktops");
			getInfo(driver, wait, "PC Gaming");
			getInfo(driver, wait, "Monitors");
			getInfo(driver, wait, "Computer Accessories");
			getInfo(driver, wait, "Networking");
			getInfo(driver, wait, "Computer Components");
			getInfo(driver, wait, "Storage");
			getInfo(driver, wait, "TV & Video");
			getInfo(driver, wait, "Cell Phones & Accessories");
			getInfo(driver, wait, "Speakers");
			getInfo(driver, wait, "Headphones");
			getInfo(driver, wait, "Bluetooth Earbuds");
			getInfo(driver, wait, "Phones");


			products = removeDuplicates(products);

			Excel.ClearExcel();
			Excel.writeExcel(products);

			//
			//Add method to save xlsx
			//

			driver.Close();
			driver.Quit();

			Process.Start(Excel.ProductsExcelPath);


		}
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

		public static void getInfo(IWebDriver driver, WebDriverWait wait, String category)
		{
			bool searched = true;
			if(category == "Laptops")
			{
				Search.searchLaptop(driver, wait);
			}
			else if(category == "Desktops")
			{
				searched = Search.searchDesktop(driver, wait);
			}
			else if(category == "Tower")
			{
				Search.searchTower(driver, wait);
				category = "Desktop";
			}
			else if(category == "All-in-One")
			{
				Search.searchAllinOne(driver, wait);
				category = "Desktop";
			}
			else if(category == "PC Gaming")
			{
				Search.searchPCGaming(driver, wait);
			}
			else if(category == "Monitors")
			{
				searched = Search.searchMonitors(driver, wait);
			}
			else if(category == "Tablets")
			{
				Search.searchTablets(driver, wait);
			}
			else if(category == "Computer Accessories")
			{
				Search.searchComputerAccessories(driver, wait);
			}
			else if(category == "Networking")
			{
				Search.searchNetworking(driver, wait);
			}
			else if(category == "Computer Components")
			{
				Search.searchComputerComponents(driver, wait);
			}
			else if(category == "Storage")
			{
				Search.searchStorage(driver, wait);
			}
			else if(category == "TV & Video")
			{
				Search.searchTV(driver, wait);
			}
			else if(category == "Cell Phones & Accessories")
			{
				Search.searchCellAccessories(driver, wait);
			}
			else if(category == "Speakers")
			{
				Search.searchBluetoothSpeakers(driver, wait);
			}
			else if(category == "Headphones")
			{
				Search.searchHeadphones(driver, wait);
			}
			else if(category == "Bluetooth Earbuds")
			{
				Search.searchBluetoothBuds(driver, wait);
			}
			else if (category == "Phones")
			{
				Search.searchPhones(driver, wait);
			}
			else
			{
				Console.WriteLine("Category not found");
			}

			if (!searched && category == "Monitors")
			{
				return;
			}

			if(!searched && category == "Desktops")
			{
				getInfo(driver, wait, "Tower");
				getInfo(driver, wait, "All-in-One");
				return;
			}

			String listingsURL;

			//Grabs data for each deal and updates Products List directly
			listingsURL = driver.Url;
			updateURLsList(driver);
			grabData(driver, wait, URLs, category);
			URLs.Clear();

			driver.Navigate().GoToUrl(listingsURL);
			wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//a [contains(@href,'pg_2')]")));
			driver.FindElement(By.XPath("//a [contains(@href,'pg_2')]")).Click();
			listingsURL = driver.Url;
			updateURLsList(driver);
			grabData(driver, wait, URLs, category);
			URLs.Clear();

			driver.Navigate().GoToUrl(listingsURL);
			wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//a [contains(@href,'pg_3')]")));
			driver.FindElement(By.XPath("//a [contains(@href,'pg_3')]")).Click();
			listingsURL = driver.Url;
			updateURLsList(driver);
			grabData(driver, wait, URLs, category);
			URLs.Clear();
			driver.Navigate().GoToUrl(listingsURL);
		}

		public static void updateURLsList(IWebDriver driver) 
		{
			foreach(IWebElement element in driver.FindElements(By.XPath(dealURLsXPath)))
			{
				URLs.Add(element.GetAttribute("href"));
			}
		}

	public static void grabData(IWebDriver driver, WebDriverWait wait, List<String> URLs, String category) 
	{
		foreach(String url in URLs)
		{
			driver.Navigate().GoToUrl(url);
			try
			{
				wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.XPath("//span [@id = 'productTitle']")));
			}
			catch (WebDriverTimeoutException)
			{
				skippedURLs.Add(url);
				Console.WriteLine("Something missing in listing. Skipped");
				return;
			}
			//TODO check to make sure sale is on amazon page as well as listings page
			bool NoName = driver.FindElements(By.XPath("//span [@id = 'productTitle']")).Count() == 0;
			bool NoXprice = driver.FindElements(By.XPath("//span [@class = 'priceBlockStrikePriceString a-text-strike']")).Count() == 0;
			bool NoPrice1 = driver.FindElements(By.XPath("//span [@id = 'priceblock_ourprice']")).Count() == 0;
			bool NoPrice2 = driver.FindElements(By.XPath("//span [@id = 'priceblock_saleprice']")).Count() == 0;
			bool NoPrice3 = driver.FindElements(By.XPath("//span [@id = 'priceblock_dealprice']")).Count() == 0;
			if (NoName || (NoPrice1 && NoPrice2 && NoPrice3) || NoXprice)
			{
				skippedURLs.Add(url);
				Console.WriteLine("Something missing in listing. Skipped");
				return;
			}
			String name = driver.FindElement(By.XPath("//span [@id = 'productTitle']")).Text;
			String StrPrice = "-1";
			if (!NoPrice1)
			{
				StrPrice = driver.FindElement(By.XPath("//span [@id = 'priceblock_ourprice']")).Text;
			}
			else if (!NoPrice2)
			{
				StrPrice = driver.FindElement(By.XPath("//span [@id = 'priceblock_saleprice']")).Text;
			}
			else if (!NoPrice3)
			{
				StrPrice = driver.FindElement(By.XPath("//span [@id = 'priceblock_dealprice']")).Text;
			}
			String StrXprice = driver.FindElement(By.XPath("//span [@class = 'priceBlockStrikePriceString a-text-strike']")).Text;

			double price = Math.Round(Double.Parse(StrPrice.Replace("$", "").Replace(",", "")));
			double Xprice = Math.Round(Double.Parse(StrXprice.Replace("$", "").Replace(",", "")));

			if (price / Xprice < 0.9 && Math.Abs(Xprice - price) > 9)
			{
				//Saves Images
				WebClient web = new WebClient();
				String GUID = Guid.NewGuid().ToString();
				String picURL = driver.FindElement(By.XPath("//img [@data-old-hires]")).GetAttribute("src");
				if (picURL.Substring(0,4) != "http")
				{
					picURL = driver.FindElement(By.XPath("//img [@data-old-hires]")).GetAttribute("data-old-hires");
				}
				web.DownloadFile(picURL, @"C:\Users\email\Desktop\Hardware Hub\images\" + GUID + ".png");

				products.Add(new Product(
						name,
						price,
						Xprice,
						category,
						url,
						GUID));
			}

			
		}
	}

		// Function to remove duplicates from an ArrayList 
		public static List<Product> removeDuplicates(List<Product> list)
		{
			restart:
			// Traverse through the first list 
			foreach (Product element1 in list)
			{
				foreach (Product element2 in list)
				{
					if((element1 != element2) && (element1 != element2) && (element1.name == element2.name) && (element1.price == element2.price) && (element1.xprice == element2.xprice))
					{
						list.Remove(element2);
						goto restart;
					}
				}
			}
			// return the new list 
			return list;
		}
    }
}
