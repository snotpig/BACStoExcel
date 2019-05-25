namespace BacsToExcel
{
	public class Transaction
	{
		public int DestSortCode { get; set; }
		public int DestAccountNumber { get; set; }
		public int OrigSortCode { get; set; }
		public int OrigAccountNumber { get; set; }
		public decimal Amount { get; set; }
		public string Beneficiary { get; set; }
		public string AccountName { get; set; }
	}
}
