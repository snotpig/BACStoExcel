using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;

namespace BacsToExcel
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		private BackgroundWorker _backgroundWorker;

		public MainWindow()
		{
			InitializeComponent();
			_backgroundWorker = (BackgroundWorker)FindResource("backgroundWorker");
		}

		private void DropPanel_Drop(object sender, DragEventArgs e)
		{
			var files = (string[])e.Data.GetData(DataFormats.FileDrop);
			LoadFile(files[0]);
		}

		private void BtnOpen_Click(object sender, RoutedEventArgs e)
		{
			var openFileDialog = new OpenFileDialog();

			openFileDialog.Filter = $"BACS files (*.txt)|*.txt";
			var result = openFileDialog.ShowDialog();
			if (result.Value)
				LoadFile(openFileDialog.FileName);
		}

		private void LoadFile(string filePath)
		{
			List<string> fileLines;
			try
			{
				fileLines = File.ReadAllLines(filePath).ToList();
			}
			catch(Exception ex)
			{
				ShowMessage(ex.Message);
				return;
			}
			if (!fileLines.Any() || fileLines.First().Substring(0, 4) != "VOL1") return;

			var volumeHeader = fileLines.ElementAt(0);
			var hdr1 = fileLines.ElementAt(1);
			var eof = fileLines.FindIndex(l => l.Substring(0, 3) == "EOF");
			var data = fileLines.Where((s, i) => i > 3 && i < eof);
			var utl1 = fileLines.ElementAt(eof + 2);
			var file = new BacsFile
			{
				PaymentFileId = int.Parse(volumeHeader.Substring(5, 5)),
				CreationDate = GetDate(hdr1.Substring(42, 5)),
				Transactions = data.Where(l => l.Substring(15, 2) == "99").Select(GetTransaction),
				ContraRecords = data.Where(l => l.Substring(15, 2) == "17").Select(GetTransaction),
				DebitValueTotal = decimal.Parse(utl1.Substring(4, 13)) / 100M,
				CreditValueTotal = decimal.Parse(utl1.Substring(17, 13)) / 100M,
				DebitItemCount = int.Parse(utl1.Substring(30, 7)),
				CreditItemCount = int.Parse(utl1.Substring(37, 7))
			};

			var dlg = new SaveFileDialog
			{
				AddExtension = true,
				DefaultExt = ".xlsx",
				OverwritePrompt = false,
				FileName = $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\BACS\{Path.GetFileNameWithoutExtension(filePath)}.xlsx"
			};
			var result = dlg.ShowDialog();
			if (result.Value)
			{
				file.FileName = Path.ChangeExtension(dlg.FileName, ".xlsx");
				btnOpen.Visibility = Visibility.Collapsed;
				spinner.Visibility = Visibility.Visible;
				_backgroundWorker.RunWorkerAsync(file);
			}
		}

		private DateTime GetDate(string str)
		{
			return new DateTime(2000 + int.Parse(str.Substring(0, 2)), 1, 1).AddDays(int.Parse(str.Substring(2, 3)) - 1);
		}

		private Transaction GetTransaction(string str)
		{
			return new Transaction
			{
				DestSortCode = int.Parse(str.Substring(0, 6)),
				DestAccountNumber = int.Parse(str.Substring(6, 8)),
				OrigSortCode = int.Parse(str.Substring(17, 6)),
				OrigAccountNumber = int.Parse(str.Substring(23, 8)),
				Amount = decimal.Parse(str.Substring(35, 11)) / 100M,
				Beneficiary = str.Substring(64, 18).Trim(),
				AccountName = str.Substring(82, 18).Trim()
			};
		}

		public void ShowMessage(string message)
		{
			MessageBox.Show(message, "Error!");
		}

		private void BackgroundWorker_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			var file = e.Argument as BacsFile;
			var fileCreator = new ExcelFileCreator();
			e.Result = fileCreator.SaveExcelFile(file);
		}

		private void BackgroundWorker_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			spinner.Visibility = Visibility.Collapsed;
			btnOpen.Visibility = Visibility.Visible;
			var result = e.Result as string;
			if (!string.IsNullOrEmpty(result))
				ShowMessage(result);
		}
	}
}
