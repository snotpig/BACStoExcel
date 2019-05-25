using System;
using System.Collections.Generic;

namespace BacsToExcel
{
	public class BacsFile
	{
		public string FileName { get; set; }
		public int PaymentFileId { get; set; }
		public DateTime CreationDate { get; set; }
		public IEnumerable<Transaction> Transactions { get; set; }
		public IEnumerable<Transaction> ContraRecords { get; set; }
		public decimal DebitValueTotal { get; set; }
		public decimal CreditValueTotal { get; set; }
		public int DebitItemCount { get; set; }
		public int CreditItemCount { get; set; }
	}
}
