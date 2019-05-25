using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace BacsToExcel
{
	public class ExcelFileCreator
	{
		public string SaveExcelFile(BacsFile file)
		{
			var app = new Application();
			Workbook workBook = null;
			Worksheet workSheet = null;

			app.Visible = false;
			workBook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
			try
			{
				workSheet = workBook.Worksheets[1];

				workSheet.Cells[1, 1] = "Account ID";
				workSheet.Cells[1, 2] = "Amount";
				workSheet.Cells[1, 3] = "Account name";
				workSheet.Cells[1, 4] = "Destination sort code";
				workSheet.Cells[1, 5] = "Destination a/c number";
				workSheet.Cells[1, 6] = "Originating sort code";
				workSheet.Cells[1, 7] = "Originating a/c number";
				workSheet.get_Range("A1", "G1").Font.Bold = true;
				(workSheet.Columns[1] as Range).ColumnWidth = "16";
				(workSheet.Columns[2] as Range).ColumnWidth = "10";
				(workSheet.Columns[3] as Range).ColumnWidth = "21";
				(workSheet.Columns[4] as Range).ColumnWidth = "19";
				(workSheet.Columns[5] as Range).ColumnWidth = "21";
				(workSheet.Columns[6] as Range).ColumnWidth = "19";
				(workSheet.Columns[7] as Range).ColumnWidth = "21";
				(workSheet.Columns[2] as Range).NumberFormat = "£0.00";

				var row = 2;
				for (var t = 0; t < file.Transactions.Count(); t++)
				{
					workSheet.Cells[row + t, 1] = file.Transactions.ElementAt(t).Beneficiary;
					workSheet.Cells[row + t, 2] = file.Transactions.ElementAt(t).Amount;
					workSheet.Cells[row + t, 3] = file.Transactions.ElementAt(t).AccountName;
					workSheet.Cells[row + t, 4] = file.Transactions.ElementAt(t).DestSortCode;
					workSheet.Cells[row + t, 5] = file.Transactions.ElementAt(t).DestAccountNumber;
					workSheet.Cells[row + t, 6] = file.Transactions.ElementAt(t).OrigSortCode;
					workSheet.Cells[row + t, 7] = file.Transactions.ElementAt(t).OrigAccountNumber;
				}

				row += 2 + file.Transactions.Count();
				for (var c = 0; c < file.ContraRecords.Count(); c++)
				{
					workSheet.Cells[row + c, 1] = file.ContraRecords.ElementAt(c).Beneficiary;
					workSheet.Cells[row + c, 2] = file.ContraRecords.ElementAt(c).Amount;
					workSheet.Cells[row + c, 3] = file.ContraRecords.ElementAt(c).AccountName;
					workSheet.Cells[row + c, 4] = file.ContraRecords.ElementAt(c).DestSortCode;
					workSheet.Cells[row + c, 5] = file.ContraRecords.ElementAt(c).DestAccountNumber;
					workSheet.Cells[row + c, 6] = file.ContraRecords.ElementAt(c).OrigSortCode;
					workSheet.Cells[row + c, 7] = file.ContraRecords.ElementAt(c).OrigAccountNumber;
				}

				row += 2 + file.ContraRecords.Count();
				workSheet.Cells[row, 1] = "Credit Value Total";
				workSheet.Cells[row, 2] = file.CreditValueTotal;
				workSheet.Cells[row, 4] = "Credit Item Count";
				workSheet.Cells[row, 5] = file.CreditItemCount;
				workSheet.Cells[row + 1, 1] = "Debit Value Total";
				workSheet.Cells[row + 1, 2] = file.DebitValueTotal;
				workSheet.Cells[row + 1, 4] = "Debit Item Count";
				workSheet.Cells[row + 1, 5] = file.DebitItemCount;
				workSheet.get_Range($"A{row}", $"A{row}").Font.Bold = true;
				workSheet.get_Range($"D{row}", $"D{row}").Font.Bold = true;
				workSheet.get_Range($"A{row + 1}", $"A{row + 2}").Font.Bold = true;
				workSheet.get_Range($"D{row + 1}", $"D{row + 2}").Font.Bold = true;

				workBook.Worksheets[1].Name = $"PF-{file.PaymentFileId}";
				workBook.SaveAs(file.FileName);
				workBook.Close();

				return "";
			}

			catch (Exception ex)
			{
				return ex.Message;
			}

			finally
			{
				app.Quit();
				Marshal.ReleaseComObject(workSheet);
				Marshal.ReleaseComObject(workBook);
				Marshal.ReleaseComObject(app);
			}
		}
	}
}
