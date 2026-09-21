namespace MailPdfApp.Tests;

public class IsValidEmailTests
{
	[Theory]
	[InlineData("user@example.com")]
	[InlineData("first.last+tag@sub.example.co.uk")]
	[InlineData("a@b.co")]
	public void ReturnsTrueForValidAddresses(string email)
	{
		Assert.True(MainForm.IsValidEmail(email));
	}

	[Theory]
	[InlineData("")]
	[InlineData(" ")]
	[InlineData("not-an-email")]
	[InlineData("missing-domain@")]
	[InlineData("@missing-local.com")]
	[InlineData("has spaces@example.com")]
	public void ReturnsFalseForInvalidAddresses(string email)
	{
		Assert.False(MainForm.IsValidEmail(email));
	}
}
