using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace MailPdfApp.Tests;

public class ExtractEmailAddressesFromPdfTests : IDisposable
{
	private readonly string _tempDir;

	public ExtractEmailAddressesFromPdfTests()
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

	private string CreatePdfWithText(string text)
	{
		string path = Path.Combine(_tempDir, Guid.NewGuid() + ".pdf");
		using var document = new PdfDocument();
		PdfPage page = document.AddPage();
		using XGraphics gfx = XGraphics.FromPdfPage(page);
		var font = new XFont("Arial", 12);
		gfx.DrawString(text, font, XBrushes.Black,
			new XRect(20, 20, page.Width.Point - 40, page.Height.Point - 40), XStringFormats.TopLeft);
		document.Save(path);
		return path;
	}

	[Fact]
	public void FindsEmailAddressInPdfText()
	{
		string path = CreatePdfWithText("Please contact jane.doe@example.com for details.");

		List<string> emails = MainForm.ExtractEmailAddresses(path);

		Assert.Single(emails);
		Assert.Equal("jane.doe@example.com", emails[0], StringComparer.OrdinalIgnoreCase);
	}

	[Fact]
	public void ReturnsEmptyListWhenPdfHasNoEmail()
	{
		string path = CreatePdfWithText("No contact information here.");

		List<string> emails = MainForm.ExtractEmailAddresses(path);

		Assert.Empty(emails);
	}

	[Fact]
	public void ReturnsEmptyListWhenPdfIsInvalid()
	{
		string path = Path.Combine(_tempDir, "not-a-pdf.pdf");
		File.WriteAllText(path, "this is not a real pdf");

		List<string> emails = MainForm.ExtractEmailAddresses(path);

		Assert.Empty(emails);
	}
}
