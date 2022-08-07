using Microsoft.Extensions.Configuration;

namespace DoubleAccountFinder
{
	internal class Program
	{
		const string SourceExtension = ".xlsx";

		private static ISet<string> GetFiles(IEnumerable<string> sources) =>
			sources
				.Where(s => Directory.Exists(s))
				.SelectMany(s => Directory.EnumerateFiles(s, SourceExtension, SearchOption.AllDirectories))
				.Concat(sources
					.Where(s => File.Exists(s)))
				.Where(s => string.Equals(Path.GetExtension(s), SourceExtension, StringComparison.OrdinalIgnoreCase))
				.ToHashSet()
			;

		static int Main(string[] args)
		{
			if (!args.Any())
			{
				Console.Error.WriteLine("Не указаны файлы для обработки");
				return 1;
			}

			var sources = args
				.Where(a => File.Exists(a) || Directory.Exists(a))
			;

			var files = GetFiles(sources);
			if (!files.Any())
			{
				Console.Error.WriteLine("Указанные файлы недоступны или не в том формате");
				return 2;
			}

			var configuration = new ConfigurationBuilder()
				.AddIniFile("config.ini")
				.AddCommandLine(args)
				.Build()
			;

			var options = configuration.GetSection("Main").Get<Options>();
			var processor = new Processor(files, options);

			return processor.Process();
		}
	}
}