using OfficeOpenXml;

namespace DoubleAccountFinder
{
	internal class Program
	{
		static int Main(string[] args)
		{
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			return 0;
		}
	}
}