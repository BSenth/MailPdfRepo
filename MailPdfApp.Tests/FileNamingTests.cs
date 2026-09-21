namespace MailPdfApp.Tests;

public class FileNamingTests : IDisposable
{
	private readonly string _tempDir;

	public FileNamingTests()
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

	[Theory]
	[InlineData("john.doe", "john.doe")]
	[InlineData("bad:name?", "bad_name_")]
	public void SanitizeFileNamePart_ReplacesInvalidCharacters(string input, string expected)
	{
		Assert.Equal(expected, MainForm.SanitizeFileNamePart(input));
	}

	[Fact]
	public void GetUniqueFilePath_ReturnsSamePathWhenNoCollision()
	{
		string path = Path.Combine(_tempDir, "file.pdf");

		Assert.Equal(path, MainForm.GetUniqueFilePath(path));
	}

	[Fact]
	public void GetUniqueFilePath_AppendsCounterWhenFileExists()
	{
		string path = Path.Combine(_tempDir, "file.pdf");
		File.WriteAllText(path, "existing");

		string result = MainForm.GetUniqueFilePath(path);

		Assert.Equal(Path.Combine(_tempDir, "file_1.pdf"), result);
	}

	[Fact]
	public void GetUniqueFilePath_SkipsPastMultipleCollisions()
	{
		string path = Path.Combine(_tempDir, "file.pdf");
		File.WriteAllText(path, "existing");
		File.WriteAllText(Path.Combine(_tempDir, "file_1.pdf"), "existing");

		string result = MainForm.GetUniqueFilePath(path);

		Assert.Equal(Path.Combine(_tempDir, "file_2.pdf"), result);
	}

	[Fact]
	public void RenameSplitFile_UsesEmailLocalPartAndTimestamp()
	{
		string original = Path.Combine(_tempDir, "part.pdf");
		File.WriteAllText(original, "content");

		string renamed = MainForm.RenameSplitFile(original, "jane.doe@example.com");

		Assert.False(File.Exists(original));
		Assert.True(File.Exists(renamed));
		Assert.StartsWith("jane.doe_", Path.GetFileName(renamed));
		Assert.EndsWith(".pdf", renamed);
	}

	[Fact]
	public void RenameSplitFile_UsesMailPrefixWhenNoEmail()
	{
		string original = Path.Combine(_tempDir, "part.pdf");
		File.WriteAllText(original, "content");

		string renamed = MainForm.RenameSplitFile(original, string.Empty);

		Assert.StartsWith("mail_", Path.GetFileName(renamed));
	}
}
