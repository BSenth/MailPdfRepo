namespace MailPdfApp.Tests;

public class ExtractEmailAddressesFromTextTests
{
	[Fact]
	public void FindsAllDistinctEmailsCaseInsensitively()
	{
		const string text = "Contact a@example.com or A@Example.com, also c@example.org.";

		List<string> result = MainForm.ExtractEmailAddressesFromText(text);

		Assert.Equal(2, result.Count);
		Assert.Contains(result, e => string.Equals(e, "a@example.com", StringComparison.OrdinalIgnoreCase));
		Assert.Contains(result, e => string.Equals(e, "c@example.org", StringComparison.OrdinalIgnoreCase));
	}

	[Fact]
	public void ReturnsEmptyListWhenNoEmailPresent()
	{
		List<string> result = MainForm.ExtractEmailAddressesFromText("nothing to see here");

		Assert.Empty(result);
	}

	[Fact]
	public void PreservesFirstSeenOrder()
	{
		const string text = "second@example.com then first@example.com then second@example.com again";

		List<string> result = MainForm.ExtractEmailAddressesFromText(text);

		Assert.Equal(new[] { "second@example.com", "first@example.com" }, result);
	}
}
