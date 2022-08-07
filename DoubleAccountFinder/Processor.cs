using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace DoubleAccountFinder
{
	internal class Processor
	{
		private readonly IEnumerable<string> _sourceFiles;
		private readonly Options _configOptions;
		private readonly Regex _accountRegex;

		static Processor()
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
		}

		private ILookup<string, SourceLine> GetSourceData(string filename)
		{
			var sources = new List<SourceLine>();

			using var stream = File.OpenRead(filename);
			using var package = new ExcelPackage(stream);
			var worksheets = package.Workbook.Worksheets
				.Where(w => w.Hidden == eWorkSheetHidden.Visible)
			;

			foreach (var worksheet in worksheets)
			{
				var startRow = worksheet.Dimension.Start.Row;
				var endRow = worksheet.Dimension.End.Row;

				for (var row = startRow; row <= endRow; row++)
				{
					var account = worksheet.Cells[row, _configOptions.AccountColumn].Text;
					if (_accountRegex.IsMatch(account))
					{
						var amount = worksheet.Cells[row, _configOptions.AmountColumn].Text;
						sources.Add(new SourceLine(worksheet.Name, row, account, amount));
					}
				}
			}

			return sources
				.AsParallel()
				.ToLookup(s => s.Account)
			;
		}

		private void CompareFile(string filename, ILookup<string, SourceLine> sourceData)
		{
			Console.WriteLine("Обработка: {0}", Path.GetFileNameWithoutExtension(filename));

			using var package = new ExcelPackage(new FileInfo(filename));
			var worksheet = package.Workbook.Worksheets.First(w => w.View.TabSelected);

			var startRow = worksheet.Dimension.Start.Row;
			var endRow = worksheet.Dimension.End.Row;
			var lastCol = worksheet.Dimension.End.Column;

			for (var row = startRow; row <= endRow; row++)
			{
				var cell = worksheet.Cells[row, _configOptions.AccountColumn];
				var account = cell.Text;
				if (_accountRegex.IsMatch(account))
				{
					var line = sourceData[account];
					if (line.Any())
					{
						cell[row, lastCol + 1].Value = string.Join(Environment.NewLine, line.Select(l => l.Worksheet));
						cell[row, lastCol + 2].Value = string.Join(Environment.NewLine, line.Select(l => l.Line.ToString()));
						cell[row, lastCol + 3].Value = string.Join(Environment.NewLine, line.Select(l => l.Amount));

						cell[row, lastCol + 1, row, lastCol + 3].Style.WrapText = true;

						var errorRow = cell[row, 1, row, lastCol + 3];
						errorRow.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
						errorRow.Style.Fill.BackgroundColor.SetColor(Color.Red);
						errorRow.Style.Font.Color.SetColor(Color.White);
					}
				}
			}

			if (!string.IsNullOrWhiteSpace(_configOptions.CreateNewSuffix))
			{
				var name = Path.Combine(Path.GetDirectoryName(filename), Path.GetFileNameWithoutExtension(filename) + _configOptions.CreateNewSuffix + Path.GetExtension(filename));
				package.SaveAs(name);
			}
			else
			{
				package.Save();
			}
		}

		public Processor(IEnumerable<string> sources, Options options)
		{
			this._sourceFiles = sources;
			this._configOptions = options;

			_accountRegex = new Regex(options.AccountRegex, RegexOptions.Compiled);
		}

		public void Process()
		{
			var sourceData = GetSourceData(_configOptions.Source);
			Console.WriteLine("Исходных счетов: {0}", sourceData.Count);

			Console.WriteLine("Всего файлов: {0}", _sourceFiles.Count());
			foreach (var file in _sourceFiles)
			{
				CompareFile(file, sourceData);
			}

			Console.WriteLine("Обработка завершена");
		}

		private record SourceLine(string Worksheet, int Line, string Account, string Amount);
	}
}
