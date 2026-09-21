using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MailPdfApp.Tests;

public class SplitPdfTests : IDisposable
{
	private readonly string _tempDir;

	public SplitPdfTests()
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

	private string CreatePdf(int pageCount)
	{
		string path = Path.Combine(_tempDir, "source.pdf");
		using var document = new PdfDocument();
		for(int i = 0; i < pageCount; i++)
		{
			document.AddPage();
		}
		document.Save(path);
		return path;
	}

	[Fact]
	public void SplitPdf_CreatesOneFilePerGroupOfPages()
	{
		string sourcePath = CreatePdf(5);

		string outputFolder = MainForm.SplitPdf(sourcePath, 2);

		string[] files = Directory.GetFiles(outputFolder, "*.pdf");
		Assert.Equal(3, files.Length); // groups of 2, 2, 1
	}

	[Fact]
	public void SplitPdf_LastFileContainsRemainderPages()
	{
		string sourcePath = CreatePdf(5);

		string outputFolder = MainForm.SplitPdf(sourcePath, 2);

		string[] files = Directory.GetFiles(outputFolder, "*.pdf").OrderBy(f => f).ToArray();
		int[] pageCounts = files
			.Select(f =>
			{
				using PdfDocument doc = PdfReader.Open(f, PdfDocumentOpenMode.Import);
				return doc.PageCount;
			})
			.ToArray();

		Assert.Equal(new[] { 2, 2, 1 }, pageCounts);
	}

	[Fact]
	public void SplitPdf_SinglePageGroupsProduceOneFilePerPage()
	{
		string sourcePath = CreatePdf(3);

		string outputFolder = MainForm.SplitPdf(sourcePath, 1);

		Assert.Equal(3, Directory.GetFiles(outputFolder, "*.pdf").Length);
	}
}
