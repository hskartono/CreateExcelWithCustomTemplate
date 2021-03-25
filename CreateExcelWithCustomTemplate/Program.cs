using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CreateExcelWithCustomTemplate
{
	class Program
	{
		static void Main(string[] args)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			PurchaseOrder po = new PurchaseOrder()
			{
				PIDNo = "TEST1",
				Warehouse = "Warehouse 1",
				StorageLocation = "Sloc 1",
				PostingDate = DateTime.Now
			};
			po.AddPODetail("PART1", "Busi", "A.1.1", 10, 15);
			po.AddPODetail("PART2", "Busi", "B.1.1", 10, 5);
			po.AddPODetail("PART3", "Busi", "C.1.1", 10, 10);
			po.AddPODetail("PART4", "Busi", "D.1.1", 0, 0);

			Program p = new Program();
			string result = p.GenerateExcel(po);
			Console.WriteLine($"File tersimpan di {result}. Press any key to exit");
			Console.ReadKey();
		}

		public string GenerateExcel(PurchaseOrder po)
		{
			string excelFileName = System.IO.Path.Combine(AppContext.BaseDirectory, "template.xlsx");
			string destinationFile = System.IO.Path.Combine(AppContext.BaseDirectory, "resul.xlsx");
			if (System.IO.File.Exists(destinationFile)) System.IO.File.Delete(destinationFile);
			
			using (var excel = new ExcelPackage(new System.IO.FileInfo(excelFileName)))
			{
				var ws = excel.Workbook.Worksheets["Sheet1"];
				ws.Cells[4, 3].Value = po.PIDNo;
				ws.Cells[5, 3].Value = po.Warehouse;
				ws.Cells[6, 3].Value = po.StorageLocation;
				ws.Cells[7, 3].Value = po.PostingDate;

				int row = 10;
				int totalRow = po.PurchaseOrderDetails.Count - 2;
				if (totalRow > 0)
				{
					ws.InsertRow(11, totalRow, 10);
				}
				foreach(var item in po.PurchaseOrderDetails)
				{
					ws.Cells[row, 1].Value = item.PartNumber;
					ws.Cells[row, 2].Value = item.PartName;
					ws.Cells[row, 3].Value = item.RackNo;
					ws.Cells[row, 4].Value = item.BookQty;
					ws.Cells[row, 5].Value = item.PhysicQty;
					ws.Cells[row, 6].Value = item.DiffQtyNet;
					ws.Cells[row, 7].Value = item.DiffQtyAbs;
					row++;
				}
				excel.SaveAs(new System.IO.FileInfo(destinationFile));
			}

			return destinationFile;
		}
	}

	class PurchaseOrder
	{
		public string PIDNo { get; set; }
		public string Warehouse { get; set; }
		public string StorageLocation { get; set; }
		public DateTime PostingDate { get; set; }
		public List<PurchaseOrderDetail> PurchaseOrderDetails { get; set; } = new List<PurchaseOrderDetail>();
		public void AddPODetail(string partNumber, string partName, string rackNo, int bookQty, int physicQty)
		{
			PurchaseOrderDetail pod = new PurchaseOrderDetail(partNumber, partName, rackNo, bookQty, physicQty);
			PurchaseOrderDetails.Add(pod);
		}
	}

	class PurchaseOrderDetail
	{
		public PurchaseOrderDetail(string partNumber, string partName, string rackNo, int bookQty, int physicQty)
		{
			PartNumber = partNumber;
			PartName = partName;
			RackNo = rackNo;
			BookQty = bookQty;
			PhysicQty = physicQty;
		}

		public string PartNumber { get; set; }
		public string PartName { get; set; }
		public string RackNo { get; set; }
		public int BookQty { get; set; }
		public int PhysicQty { get; set; }
		public int DiffQtyNet { get { return PhysicQty - BookQty; } }
		public int DiffQtyAbs { get { return Math.Abs(DiffQtyNet); } }

	}
}
