using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DoubleAccountFinder
{
	internal class Options
	{
		public int AccountColumn { get; set; }
		public int AmountColumn { get; set; }
		public string AccountRegex { get; set; }
		public string Source { get; set; }
	}
}
