using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AMZN_to_Excel
{
    class Excel
    {
		public static object Open { get; private set; }
		public static String ProductsExcelPath = @"C:\Users\email\Desktop\Hardware Hub\products.xlsx";
		public static String TestExcelPath = @"C:\Users\email\Desktop\test.xlsx";


		public static void writeExcel(List<Product> products) 
		{
			IWorkbook wb = new XSSFWorkbook();

			ISheet ws = wb.CreateSheet("Products");
			ws.SetColumnWidth(0, 6000);
		
			IRow header = ws.CreateRow(0);

			header.CreateCell(0).SetCellValue("Name");
			header.CreateCell(1).SetCellValue("Price");
			header.CreateCell(2).SetCellValue("Old Price");
			header.CreateCell(3).SetCellValue("Difference");
			header.CreateCell(4).SetCellValue("Category");
			header.CreateCell(5).SetCellValue("Picture");
			header.CreateCell(6).SetCellValue("ID");
			header.CreateCell(7).SetCellValue("URL");

			int rowcount = 1;

			foreach(Product product in products)
			{


				IRow ProductRow = ws.CreateRow(rowcount);
				ProductRow.CreateCell(0).SetCellValue(product.name);
				ProductRow.CreateCell(1).SetCellValue(product.price);
				ProductRow.CreateCell(2).SetCellValue(product.xprice);
				ProductRow.CreateCell(3).SetCellValue(product.dif);
				ProductRow.CreateCell(4).SetCellValue(product.category);
				ProductRow.CreateCell(5);
				ProductRow.CreateCell(6).SetCellValue(product.ID);
				ProductRow.CreateCell(7).SetCellValue(product.URL);
				ProductRow.Height = 1500;

				if(File.Exists(@"C:\Users\email\Desktop\Hardware Hub\images\" + product.ID + ".png"))
				{
					byte[] data = File.ReadAllBytes(@"C:\Users\email\Desktop\Hardware Hub\images\" + product.ID + ".png");
					int pictureIndex = wb.AddPicture(data, PictureType.PNG);

					
					IDrawing patriarch = ws.CreateDrawingPatriarch();
					IClientAnchor anchor = wb.GetCreationHelper().CreateClientAnchor();
					anchor.Col1 = 5;
					anchor.Row1 = rowcount;
					anchor.AnchorType = AnchorType.MoveAndResize;
					IPicture picture = patriarch.CreatePicture(anchor,pictureIndex);
					picture.Resize(1);


					//byte[] data = File.ReadAllBytes(@"C:\Users\email\Desktop\Hardware Hub\images\" + product.ID + ".png");
					//int pictureIndex = wb.AddPicture(data, PictureType.PNG);
					//ICreationHelper helper = wb.GetCreationHelper();
					//IDrawing drawing = ws.CreateDrawingPatriarch();
					//IClientAnchor anchor = helper.CreateClientAnchor();
					//anchor.Col1 = 5;//0 index based column
					//anchor.Row1 = ProductRow.RowNum;//0 index based row
					//IPicture picture = drawing.CreatePicture(anchor, pictureIndex);
					//picture.Resize();
				}

				rowcount++;
			}

			IRow row = ws.CreateRow(rowcount);
			row.CreateCell(0).SetCellValue("Stop");
			rowcount++;

			row = ws.CreateRow(rowcount);
			rowcount++;
			row = ws.CreateRow(rowcount);
			row.CreateCell(0).SetCellValue("Skipped products");
			rowcount++;
		
			foreach(String url in Program.skippedURLs)
			{
				row = ws.CreateRow(rowcount);
				row.CreateCell(0).SetCellValue(url);
				rowcount++;
			}
			Stream stream = new FileStream(ProductsExcelPath, FileMode.Create);
			//Stream stream = new FileStream(@"C:\Users\email\Desktop\Hardware Hub\products.xlsx", FileMode.Open);
			wb.Write(stream);
			wb.Close();
			stream.Close();
		}
	
		//Clear Excel
		public static void ClearExcel()
		{
			try
			{
				//Stream stream = new FileStream(@"C:\Users\email\Desktop\Hardware Hub\products.xlsx", FileMode.Open);
				Stream file = new FileStream(ProductsExcelPath, FileMode.Open);
				IWorkbook workbook = new XSSFWorkbook(file);
				ISheet sheet = workbook.GetSheetAt(0);
				workbook.RemoveSheetAt(0);
				workbook.Write(file);
				file.Close();
				workbook.Close();
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
			}
		}
    }
}
