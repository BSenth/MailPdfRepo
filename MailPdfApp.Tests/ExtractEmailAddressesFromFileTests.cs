using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

namespace MailPdfApp.Tests;

public class ExtractEmailAddressesFromFileTests : IDisposable
{
	private readonly string _tempDir;

	public ExtractEmailAddressesFromFileTests()
	{
		_tempDir = Path.Combine(Path.GetTempPath(), "MailPdfApp.Tests_" + Guid.NewGuid());
		Directory.CreateDirectory(_tempDir);
	}

	public void Dispose()
	{
		if(Directory.Exists(_tempDir))
		{
			Directory.Delete(_tempDir, recursive: true);
		}
	}

	private string CreateTxt(string content)
	{
		string path = Path.Combine(_tempDir, "addresses.txt");
		File.WriteAllText(path, content);
		return path;
	}

	private string CreateXlsx(string cellValue)
	{
		string path = Path.Combine(_tempDir, "addresses.xlsx");
		IWorkbook workbook = new XSSFWorkbook();
		ISheet sheet = workbook.CreateSheet("Sheet1");
		IRow row = sheet.CreateRow(0);
		row.CreateCell(0).SetCellValue(cellValue);

		using(FileStream stream = File.Create(path))
		{
			workbook.Write(stream);
		}
		return path;
	}

	private string CreateDocx(string paragraphText)
	{
		string path = Path.Combine(_tempDir, "addresses.docx");
		using var document = new XWPFDocument();
		XWPFParagraph paragraph = document.CreateParagraph();
		XWPFRun run = paragraph.CreateRun();
		run.SetText(paragraphText);

		using(FileStream stream = File.Create(path))
		{
			document.Write(stream);
		}
		return path;
	}

	[Fact]
	public void ReadsAddressesFromTxtFile()
	{
		string path = CreateTxt("alice@example.com, bob@example.com");

		List<string> emails = MainForm.ExtractEmailAddressesFromFile(path);

		Assert.Equal(2, emails.Count);
	}

	[Fact]
	public void ReadsAddressesFromXlsxCells()
	{
		string path = CreateXlsx("carol@example.com");

		List<string> emails = MainForm.ExtractEmailAddressesFromFile(path);

		Assert.Contains(emails, e => string.Equals(e, "carol@example.com", StringComparison.OrdinalIgnoreCase));
	}

	[Fact]
	public void ReadsAddressesFromDocxParagraphs()
	{
		string path = CreateDocx("Please reach dave@example.com for questions.");

		List<string> emails = MainForm.ExtractEmailAddressesFromFile(path);

		Assert.Contains(emails, e => string.Equals(e, "dave@example.com", StringComparison.OrdinalIgnoreCase));
	}

	[Fact]
	public void ThrowsForLegacyDocFiles()
	{
		string path = Path.Combine(_tempDir, "addresses.doc");
		File.WriteAllText(path, "dummy content");

		var ex = Assert.Throws<NotSupportedException>(() => MainForm.ExtractEmailAddressesFromFile(path));
		Assert.Contains(".docx", ex.Message);
	}

	[Fact]
	public void ThrowsForUnsupportedExtension()
	{
		string path = Path.Combine(_tempDir, "addresses.pdf");
		File.WriteAllText(path, "dummy content");

		Assert.Throws<NotSupportedException>(() => MainForm.ExtractEmailAddressesFromFile(path));
	}
}
