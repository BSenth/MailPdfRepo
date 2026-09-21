using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using NPOI.SS.UserModel;
using NPOI.XWPF.Extractor;
using NPOI.XWPF.UserModel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace MailPdfApp;

public class MainForm : Form
{
	private Label lblEmail = new();
	private TextBox txtEmail = new();

	private Label lblPassword = new();
	private TextBox txtPassword = new();
	private CheckBox chkShowPassword = new();

	private Label lblProvider = new();
	private ComboBox cboProvider = new();

	private Label lblSmtpHost = new();
	private TextBox txtSmtpHost = new();

	private Label lblSmtpPort = new();
	private NumericUpDown numSmtpPort = new();
	private CheckBox chkUseSsl = new();

	private RadioButton rbSplitAndSend = new();
	private RadioButton rbSendToMany = new();

	private Label lblPdf = new();
	private TextBox txtPdfPath = new();
	private Button btnBrowse = new();

	private Label lblSplitCount = new();
	private NumericUpDown numSplitPageCount = new();
	private Button btnSplit = new();

	private Label lblSendFile = new();
	private TextBox txtSendFilePath = new();
	private Button btnBrowseSendFile = new();

	private Label lblAddressFile = new();
	private TextBox txtAddressFilePath = new();
	private Button btnBrowseAddressFile = new();
	private Button btnLoadAddresses = new();

	private DataGridView gridFiles = new();

	private Label lblSubject = new();
	private TextBox txtSubject = new();

	private Label lblBody = new();
	private TextBox txtBody = new();

	private ProgressBar progressBar = new();
	private Button btnSendEmails = new();
	private Button btnClear = new();

	private const string ProviderGmail = "Gmail";
	private const string ProviderOutlook = "Outlook / Office365";
	private const string ProviderCustom = "Custom SMTP";

	private const string ColEmail = "EmailAddress";
	private const string ColFileName = "FileName";
	private const string ColSend = "Send";

	// Separator used when multiple email addresses are found in a single split file.
	private const string EmailSeparator = "; ";

	// Public read-only accessors in case another part of the app wants the captured values.
	public string EmailAddress => txtEmail.Text.Trim();
	public string Password => txtPassword.Text;
	public string SelectedPdfPath => txtPdfPath.Text.Trim();
	public int SplitPageCount => (int) numSplitPageCount.Value;

	public MainForm()
	{
		InitializeComponent();
	}

	private void InitializeComponent()
	{
		Text = "Mail Credentials, PDF Splitter && Sender";
		StartPosition = FormStartPosition.CenterScreen;
		ClientSize = new Size(560, 824);
		FormBorderStyle = FormBorderStyle.FixedDialog;
		MaximizeBox = false;
		Font = new Font("Segoe UI", 9F);

		int y = 20;
		const int rowHeight = 32;
		const int labelX = 20;
		const int inputX = 160;
		const int inputWidth = 360;

		// ---- From Email ----
		lblEmail.Text = "From Email:";
		lblEmail.Location = new Point(labelX, y + 3);
		lblEmail.AutoSize = true;
		txtEmail.Location = new Point(inputX, y);
		txtEmail.Size = new Size(inputWidth, 23);
		y += rowHeight;

		// ---- Password ----
		lblPassword.Text = "Password:";
		lblPassword.Location = new Point(labelX, y + 3);
		lblPassword.AutoSize = true;
		txtPassword.Location = new Point(inputX, y);
		txtPassword.Size = new Size(inputWidth, 23);
		txtPassword.UseSystemPasswordChar = true;
		y += rowHeight - 6;

		chkShowPassword.Text = "Show password";
		chkShowPassword.Location = new Point(inputX, y);
		chkShowPassword.AutoSize = true;
		chkShowPassword.CheckedChanged += (s, e) =>
			txtPassword.UseSystemPasswordChar = !chkShowPassword.Checked;
		y += rowHeight;

		// ---- Provider ----
		lblProvider.Text = "Mail Provider:";
		lblProvider.Location = new Point(labelX, y + 3);
		lblProvider.AutoSize = true;
		cboProvider.Location = new Point(inputX, y);
		cboProvider.Size = new Size(inputWidth, 23);
		cboProvider.DropDownStyle = ComboBoxStyle.DropDownList;
		cboProvider.Items.AddRange(new object[] { ProviderGmail, ProviderOutlook, ProviderCustom });
		cboProvider.SelectedIndexChanged += CboProvider_SelectedIndexChanged;
		y += rowHeight;

		// ---- SMTP Host ----
		lblSmtpHost.Text = "SMTP Host:";
		lblSmtpHost.Location = new Point(labelX, y + 3);
		lblSmtpHost.AutoSize = true;
		txtSmtpHost.Location = new Point(inputX, y);
		txtSmtpHost.Size = new Size(inputWidth, 23);
		y += rowHeight;

		// ---- SMTP Port + SSL ----
		lblSmtpPort.Text = "SMTP Port:";
		lblSmtpPort.Location = new Point(labelX, y + 3);
		lblSmtpPort.AutoSize = true;
		numSmtpPort.Location = new Point(inputX, y);
		numSmtpPort.Size = new Size(80, 23);
		numSmtpPort.Minimum = 1;
		numSmtpPort.Maximum = 65535;

		chkUseSsl.Text = "Use SSL/STARTTLS";
		chkUseSsl.Location = new Point(inputX + 100, y + 2);
		chkUseSsl.AutoSize = true;
		chkUseSsl.Checked = true;
		y += rowHeight;

		// ---- Mode selection ----
		rbSplitAndSend.Text = "Split and Send";
		rbSplitAndSend.Location = new Point(inputX, y + 3);
		rbSplitAndSend.AutoSize = true;
		rbSplitAndSend.Checked = true;

		rbSendToMany.Text = "Send to Many";
		rbSendToMany.Location = new Point(inputX + 160, y + 3);
		rbSendToMany.AutoSize = true;
		rbSendToMany.CheckedChanged += (s, e) => ApplyModeVisibility();
		y += rowHeight;

		// ---- PDF file (Split and Send) ----
		lblPdf.Text = "PDF File:";
		lblPdf.Location = new Point(labelX, y + 3);
		lblPdf.AutoSize = true;
		txtPdfPath.Location = new Point(inputX, y);
		txtPdfPath.Size = new Size(inputWidth - 80, 23);
		txtPdfPath.ReadOnly = true;
		btnBrowse.Text = "Browse...";
		btnBrowse.Location = new Point(inputX + inputWidth - 75, y - 1);
		btnBrowse.Size = new Size(75, 25);
		btnBrowse.Click += BtnBrowse_Click;

		// ---- File to send (Send to Many) ----
		lblSendFile.Text = "File to Send:";
		lblSendFile.Location = new Point(labelX, y + 3);
		lblSendFile.AutoSize = true;
		txtSendFilePath.Location = new Point(inputX, y);
		txtSendFilePath.Size = new Size(inputWidth - 80, 23);
		txtSendFilePath.ReadOnly = true;
		btnBrowseSendFile.Text = "Browse...";
		btnBrowseSendFile.Location = new Point(inputX + inputWidth - 75, y - 1);
		btnBrowseSendFile.Size = new Size(75, 25);
		btnBrowseSendFile.Click += BtnBrowseSendFile_Click;
		y += rowHeight;

		// ---- Split page count + Split button (Split and Send) ----
		lblSplitCount.Text = "Split Page Count:";
		lblSplitCount.Location = new Point(labelX, y + 3);
		lblSplitCount.AutoSize = true;
		numSplitPageCount.Location = new Point(inputX, y);
		numSplitPageCount.Size = new Size(80, 23);
		numSplitPageCount.Minimum = 1;
		numSplitPageCount.Maximum = 10000;
		numSplitPageCount.Value = 1;

		btnSplit.Text = "Split PDF";
		btnSplit.Location = new Point(inputX + 100, y - 1);
		btnSplit.Size = new Size(110, 25);
		btnSplit.Click += BtnSplit_Click;

		// ---- Address file (Send to Many) ----
		lblAddressFile.Text = "Address File:";
		lblAddressFile.Location = new Point(labelX, y + 3);
		lblAddressFile.AutoSize = true;
		txtAddressFilePath.Location = new Point(inputX, y);
		txtAddressFilePath.Size = new Size(inputWidth - 80, 23);
		txtAddressFilePath.ReadOnly = true;
		btnBrowseAddressFile.Text = "Browse...";
		btnBrowseAddressFile.Location = new Point(inputX + inputWidth - 75, y - 1);
		btnBrowseAddressFile.Size = new Size(75, 25);
		btnBrowseAddressFile.Click += BtnBrowseAddressFile_Click;
		y += rowHeight;

		// ---- Load Addresses button (Send to Many) ----
		btnLoadAddresses.Text = "Load Addresses";
		btnLoadAddresses.Location = new Point(inputX, y - 1);
		btnLoadAddresses.Size = new Size(140, 25);
		btnLoadAddresses.Click += BtnLoadAddresses_Click;
		y += rowHeight + 4;

		// ---- Grid: Email Address | File Name | Send ----
		gridFiles.Location = new Point(labelX, y);
		gridFiles.Size = new Size(inputX + inputWidth - labelX, 190);
		gridFiles.AllowUserToAddRows = false;
		gridFiles.AllowUserToDeleteRows = false;
		gridFiles.RowHeadersVisible = false;
		gridFiles.SelectionMode = DataGridViewSelectionMode.CellSelect;
		gridFiles.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
		gridFiles.EditMode = DataGridViewEditMode.EditOnEnter;

		var colEmail = new DataGridViewTextBoxColumn
		{
			Name = ColEmail,
			HeaderText = "Email Address",
			ReadOnly = false,
			Width = 210
		};
		var colFileName = new DataGridViewTextBoxColumn
		{
			Name = ColFileName,
			HeaderText = "File Name",
			ReadOnly = true,
			Width = 190
		};
		var colSend = new DataGridViewComboBoxColumn
		{
			Name = ColSend,
			HeaderText = "Send",
			ReadOnly = false,
			Width = 80,
			FlatStyle = FlatStyle.Flat
		};
		colSend.Items.AddRange("Yes", "No");

		gridFiles.Columns.AddRange(colEmail, colFileName, colSend);
		y += 200;

		// ---- Subject ----
		lblSubject.Text = "Subject:";
		lblSubject.Location = new Point(labelX, y + 3);
		lblSubject.AutoSize = true;
		txtSubject.Location = new Point(inputX, y);
		txtSubject.Size = new Size(inputWidth, 23);
		txtSubject.Text = "Split PDF File";
		y += rowHeight;

		// ---- Body ----
		lblBody.Text = "Body:";
		lblBody.Location = new Point(labelX, y + 3);
		lblBody.AutoSize = true;
		txtBody.Location = new Point(inputX, y);
		txtBody.Size = new Size(inputWidth, 70);
		txtBody.Multiline = true;
		txtBody.Text = "Please find your PDF file attached.";
		y += 80;

		// ---- Progress bar ----
		progressBar.Location = new Point(labelX, y);
		progressBar.Size = new Size(inputX + inputWidth - labelX, 18);
		progressBar.Style = ProgressBarStyle.Marquee;
		progressBar.MarqueeAnimationSpeed = 0; // off until busy
		progressBar.Visible = false;
		y += 30;

		// ---- Action buttons ----
		btnSendEmails.Text = "Send Emails";
		btnSendEmails.Location = new Point(inputX, y);
		btnSendEmails.Size = new Size(120, 32);
		btnSendEmails.Click += BtnSendEmails_Click;

		btnClear.Text = "Clear";
		btnClear.Location = new Point(inputX + 130, y);
		btnClear.Size = new Size(100, 32);
		btnClear.Click += BtnClear_Click;

		// Default provider selection (fills host/port).
		cboProvider.SelectedItem = ProviderGmail;

		Controls.AddRange(new Control[]
		{
			lblEmail, txtEmail,
			lblPassword, txtPassword, chkShowPassword,
			lblProvider, cboProvider,
			lblSmtpHost, txtSmtpHost,
			lblSmtpPort, numSmtpPort, chkUseSsl,
			rbSplitAndSend, rbSendToMany,
			lblPdf, txtPdfPath, btnBrowse,
			lblSplitCount, numSplitPageCount, btnSplit,
			lblSendFile, txtSendFilePath, btnBrowseSendFile,
			lblAddressFile, txtAddressFilePath, btnBrowseAddressFile, btnLoadAddresses,
			gridFiles,
			lblSubject, txtSubject,
			lblBody, txtBody,
			progressBar,
			btnSendEmails, btnClear
		});

		ApplyModeVisibility();
	}

	/// <summary>
	/// Shows/hides the controls specific to each mode ("Split and Send" vs "Send to Many")
	/// and adjusts the grid's columns to match.
	/// </summary>
	private void ApplyModeVisibility()
	{
		bool sendToMany = rbSendToMany.Checked;

		lblPdf.Visible = !sendToMany;
		txtPdfPath.Visible = !sendToMany;
		btnBrowse.Visible = !sendToMany;
		lblSplitCount.Visible = !sendToMany;
		numSplitPageCount.Visible = !sendToMany;
		btnSplit.Visible = !sendToMany;

		lblSendFile.Visible = sendToMany;
		txtSendFilePath.Visible = sendToMany;
		btnBrowseSendFile.Visible = sendToMany;
		lblAddressFile.Visible = sendToMany;
		txtAddressFilePath.Visible = sendToMany;
		btnBrowseAddressFile.Visible = sendToMany;
		btnLoadAddresses.Visible = sendToMany;

		gridFiles.Columns[ColFileName].Visible = !sendToMany;

		gridFiles.Rows.Clear();
	}

	private void CboProvider_SelectedIndexChanged(object? sender, EventArgs e)
	{
		switch(cboProvider.SelectedItem as string)
		{
			case ProviderGmail:
				txtSmtpHost.Text = "smtp.gmail.com";
				numSmtpPort.Value = 587;
				chkUseSsl.Checked = true;
				txtSmtpHost.ReadOnly = true;
				numSmtpPort.Enabled = false;
				break;

			case ProviderOutlook:
				txtSmtpHost.Text = "smtp.office365.com";
				numSmtpPort.Value = 587;
				chkUseSsl.Checked = true;
				txtSmtpHost.ReadOnly = true;
				numSmtpPort.Enabled = false;
				break;

			case ProviderCustom:
			default:
				txtSmtpHost.ReadOnly = false;
				numSmtpPort.Enabled = true;
				txtSmtpHost.Text = string.Empty;
				numSmtpPort.Value = 587;
				break;
		}
	}

	private void BtnBrowse_Click(object? sender, EventArgs e)
	{
		using var dialog = new OpenFileDialog
		{
			Title = "Select a PDF file",
			Filter = "PDF Files (*.pdf)|*.pdf",
			CheckFileExists = true,
			CheckPathExists = true,
			Multiselect = false
		};

		if(dialog.ShowDialog(this) == DialogResult.OK)
		{
			txtPdfPath.Text = dialog.FileName;
			gridFiles.Rows.Clear();
		}
	}

	private void BtnBrowseSendFile_Click(object? sender, EventArgs e)
	{
		using var dialog = new OpenFileDialog
		{
			Title = "Select the file to send",
			Filter = "All Files (*.*)|*.*",
			CheckFileExists = true,
			CheckPathExists = true,
			Multiselect = false
		};

		if(dialog.ShowDialog(this) == DialogResult.OK)
		{
			txtSendFilePath.Text = dialog.FileName;
		}
	}

	private void BtnBrowseAddressFile_Click(object? sender, EventArgs e)
	{
		using var dialog = new OpenFileDialog
		{
			Title = "Select the file containing email addresses",
			Filter = "Supported Files (*.xlsx;*.xls;*.doc;*.docx;*.txt)|*.xlsx;*.xls;*.doc;*.docx;*.txt|" +
					 "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|" +
					 "Word Documents (*.doc;*.docx)|*.doc;*.docx|" +
					 "Text Files (*.txt)|*.txt",
			CheckFileExists = true,
			CheckPathExists = true,
			Multiselect = false
		};

		if(dialog.ShowDialog(this) == DialogResult.OK)
		{
			txtAddressFilePath.Text = dialog.FileName;
			gridFiles.Rows.Clear();
		}
	}

	private async void BtnLoadAddresses_Click(object? sender, EventArgs e)
	{
		string addressFilePath = txtAddressFilePath.Text.Trim();
		if(string.IsNullOrEmpty(addressFilePath) || !File.Exists(addressFilePath))
		{
			MessageBox.Show(this, "Please select a valid address file (.xlsx, .xls, .doc, .docx or .txt).",
				"Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}

		SetBusy(true);
		try
		{
			List<string> emails = await Task.Run(() => ExtractEmailAddressesFromFile(addressFilePath));

			gridFiles.Rows.Clear();
			foreach(string emailAddress in emails)
			{
				gridFiles.Rows.Add(emailAddress, string.Empty, "Yes");
			}

			if(emails.Count == 0)
			{
				MessageBox.Show(this, "No email addresses were found in the selected file.",
					"No Addresses Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			else
			{
				MessageBox.Show(this,
					$"Loaded {emails.Count} email address(es).\n\nReview the Send column below, then click 'Send Emails'.",
					"Addresses Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
		}
		catch(Exception ex)
		{
			MessageBox.Show(this, $"Failed to read the address file:\n{ex.Message}", "Error",
				MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		finally
		{
			SetBusy(false);
		}
	}

	private void BtnClear_Click(object? sender, EventArgs e)
	{
		txtEmail.Clear();
		txtPassword.Clear();
		txtPdfPath.Clear();
		txtSendFilePath.Clear();
		txtAddressFilePath.Clear();
		numSplitPageCount.Value = 1;
		chkShowPassword.Checked = false;
		txtSubject.Text = "Split PDF File";
		txtBody.Text = "Please find your PDF file attached.";
		cboProvider.SelectedItem = ProviderGmail;
		rbSplitAndSend.Checked = true;
		gridFiles.Rows.Clear();
	}

	private async void BtnSplit_Click(object? sender, EventArgs e)
	{
		string pdfPath = txtPdfPath.Text.Trim();
		if(string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
		{
			MessageBox.Show(this, "Please select a valid PDF file.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}

		int splitPageCount = (int) numSplitPageCount.Value;
		if(splitPageCount < 1)
		{
			MessageBox.Show(this, "Split page count must be at least 1.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			numSplitPageCount.Focus();
			return;
		}

		SetBusy(true);
		try
		{
			string outputFolder = await Task.Run(() => SplitPdf(pdfPath, splitPageCount));
			string[] splitFiles = Directory.GetFiles(outputFolder, "*.pdf").OrderBy(f => f).ToArray();

			gridFiles.Rows.Clear();
			foreach(string filePath in splitFiles)
			{
				List<string> emails = await Task.Run(() => ExtractEmailAddresses(filePath));
				bool foundEmail = emails.Count > 0;

				string newFilePath = RenameSplitFile(filePath, foundEmail ? emails[0] : string.Empty);

				int rowIndex = gridFiles.Rows.Add(
					foundEmail ? string.Join(EmailSeparator, emails) : "(not found)",
					Path.GetFileName(newFilePath),
					foundEmail ? "Yes" : "No");

				gridFiles.Rows[rowIndex].Tag = newFilePath;
			}

			MessageBox.Show(this,
				$"Split into {splitFiles.Length} file(s).\nOutput folder:\n{outputFolder}\n\n" +
				"Review the Email Address / Send column below, then click 'Send Emails'.",
				"Split Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		catch(Exception ex)
		{
			MessageBox.Show(this, $"Failed to split the PDF:\n{ex.Message}", "Error",
				MessageBoxButtons.OK, MessageBoxIcon.Error);
		}
		finally
		{
			SetBusy(false);
		}
	}

	private async void BtnSendEmails_Click(object? sender, EventArgs e)
	{
		string email = txtEmail.Text.Trim();
		string password = txtPassword.Text;
		string smtpHost = txtSmtpHost.Text.Trim();
		int smtpPort = (int) numSmtpPort.Value;
		string subject = string.IsNullOrWhiteSpace(txtSubject.Text) ? "Split PDF File" : txtSubject.Text.Trim();
		string body = txtBody.Text;

		if(string.IsNullOrEmpty(email) || !IsValidEmail(email))
		{
			MessageBox.Show(this, "Please enter a valid from-email address.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			txtEmail.Focus();
			return;
		}

		if(string.IsNullOrEmpty(password))
		{
			MessageBox.Show(this, "Please enter a password.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			txtPassword.Focus();
			return;
		}

		if(string.IsNullOrEmpty(smtpHost))
		{
			MessageBox.Show(this, "Please enter an SMTP host.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			txtSmtpHost.Focus();
			return;
		}

		bool sendToMany = rbSendToMany.Checked;
		string sharedAttachmentPath = txtSendFilePath.Text.Trim();

		if(sendToMany && (string.IsNullOrEmpty(sharedAttachmentPath) || !File.Exists(sharedAttachmentPath)))
		{
			MessageBox.Show(this, "Please select a valid file to send.", "Validation Error",
				MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}

		gridFiles.EndEdit();

		var rowsToSend = gridFiles.Rows.Cast<DataGridViewRow>()
			.Where(r => string.Equals(r.Cells[ColSend].Value?.ToString(), "Yes", StringComparison.OrdinalIgnoreCase))
			.ToList();

		if(rowsToSend.Count == 0)
		{
			MessageBox.Show(this, "No rows are marked 'Yes' to send. Split a PDF (or load addresses) first, or set Send to Yes.",
				"Nothing To Send", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}

		SetBusy(true);
		int successCount = 0;
		var failures = new List<string>();

		try
		{
			foreach(DataGridViewRow row in rowsToSend)
			{
				string toAddressRaw = row.Cells[ColEmail].Value?.ToString()?.Trim() ?? string.Empty;
				string filePath = sendToMany ? sharedAttachmentPath : (row.Tag as string ?? string.Empty);
				string fileName = sendToMany ? Path.GetFileName(filePath) : (row.Cells[ColFileName].Value?.ToString() ?? string.Empty);

				string[] toAddresses = toAddressRaw
					.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
					.Where(IsValidEmail)
					.Distinct(StringComparer.OrdinalIgnoreCase)
					.ToArray();

				if(toAddresses.Length == 0 || !File.Exists(filePath))
				{
					failures.Add($"{fileName}: no valid email address found for this file");
					continue;
				}

				try
				{
					await SendMailAsync(
						smtpHost, smtpPort, chkUseSsl.Checked,
						email, password, toAddresses, subject, body,
						new[] { filePath });
					successCount++;
				}
				catch(Exception ex)
				{
					failures.Add($"{fileName} -> {string.Join(", ", toAddresses)}: {ex.Message}");
				}
			}
		}
		finally
		{
			SetBusy(false);
		}

		string summary = $"Sent: {successCount} of {rowsToSend.Count}.";
		if(failures.Count > 0)
		{
			summary += "\n\nFailed:\n" + string.Join("\n", failures);
		}

		MessageBox.Show(this, summary, "Send Complete", MessageBoxButtons.OK,
			failures.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
	}

	private void SetBusy(bool busy)
	{
		Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
		progressBar.Visible = busy;
		progressBar.MarqueeAnimationSpeed = busy ? 30 : 0;
		btnSplit.Enabled = !busy;
		btnLoadAddresses.Enabled = !busy;
		btnSendEmails.Enabled = !busy;
		btnClear.Enabled = !busy;
		rbSplitAndSend.Enabled = !busy;
		rbSendToMany.Enabled = !busy;
	}

	/// <summary>
	/// Splits <paramref name="sourcePdfPath"/> into multiple PDFs, each containing up to
	/// <paramref name="pagesPerSplit"/> pages. The split files are written into a new
	/// subfolder (named with the current date/time) inside the source file's folder.
	/// </summary>
	/// <returns>The full path of the folder that was created for the split files.</returns>
	internal static string SplitPdf(string sourcePdfPath, int pagesPerSplit)
	{
		string sourceDir = Path.GetDirectoryName(Path.GetFullPath(sourcePdfPath))
							?? throw new InvalidOperationException("Could not resolve the source folder.");

		// Folder name: today's date and time, filesystem-safe (colons aren't allowed in Windows paths).
		string folderName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
		string outputFolder = Path.Combine(sourceDir, folderName);
		Directory.CreateDirectory(outputFolder);

		using PdfDocument inputDocument = PdfReader.Open(sourcePdfPath, PdfDocumentOpenMode.Import);
		int totalPages = inputDocument.PageCount;

		if(totalPages == 0)
		{
			throw new InvalidOperationException("The selected PDF has no pages.");
		}

		string baseName = Path.GetFileNameWithoutExtension(sourcePdfPath);
		int partNumber = 1;

		for(int startPage = 0; startPage < totalPages; startPage += pagesPerSplit)
		{
			int endPage = Math.Min(startPage + pagesPerSplit, totalPages);

			using(PdfDocument outputDocument = new PdfDocument())
			{
				for(int pageIndex = startPage; pageIndex < endPage; pageIndex++)
				{
					outputDocument.AddPage(inputDocument.Pages[pageIndex]);
				}

				string outputFileName = $"{baseName}_Part{partNumber}.pdf";
				outputDocument.Save(Path.Combine(outputFolder, outputFileName));
			}

			partNumber++;
		}

		return outputFolder;
	}

	/// <summary>
	/// Renames a freshly split PDF using the email address found in it: the part of the
	/// address before '@' plus a timestamp, e.g. "john.doe_20260723_143205123.pdf". If no
	/// email address was found, the file is named "mail_&lt;timestamp&gt;.pdf" instead.
	/// </summary>
	/// <returns>The full path of the file after renaming.</returns>
	internal static string RenameSplitFile(string filePath, string email)
	{
		string namePrefix = string.IsNullOrEmpty(email) ? "mail" : email.Split('@')[0];
		namePrefix = SanitizeFileNamePart(namePrefix);

		string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
		string directory = Path.GetDirectoryName(filePath) ?? string.Empty;
		string newFilePath = GetUniqueFilePath(Path.Combine(directory, $"{namePrefix}_{timestamp}.pdf"));

		File.Move(filePath, newFilePath);
		return newFilePath;
	}

	/// <summary>
	/// Replaces any characters that are not valid in a Windows file name with '_'.
	/// </summary>
	internal static string SanitizeFileNamePart(string value)
	{
		foreach(char invalidChar in Path.GetInvalidFileNameChars())
		{
			value = value.Replace(invalidChar, '_');
		}
		return value;
	}

	/// <summary>
	/// If <paramref name="path"/> already exists, appends an incrementing counter until a
	/// free file name is found.
	/// </summary>
	internal static string GetUniqueFilePath(string path)
	{
		if(!File.Exists(path))
		{
			return path;
		}

		string directory = Path.GetDirectoryName(path) ?? string.Empty;
		string baseName = Path.GetFileNameWithoutExtension(path);
		string extension = Path.GetExtension(path);

		int counter = 1;
		string candidate;
		do
		{
			candidate = Path.Combine(directory, $"{baseName}_{counter}{extension}");
			counter++;
		} while(File.Exists(candidate));

		return candidate;
	}

	/// <summary>
	/// Extracts all text from the given PDF (via PdfPig) and returns every distinct email
	/// address found in it, in the order first seen, or an empty list if none are found
	/// or extraction fails.
	/// </summary>
	internal static List<string> ExtractEmailAddresses(string pdfPath)
	{
		try
		{
			var textBuilder = new StringBuilder();

			using(var pigDocument = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
			{
				foreach(var page in pigDocument.GetPages())
				{
					textBuilder.AppendLine(page.Text);
				}
			}

			return ExtractEmailAddressesFromText(textBuilder.ToString());
		}
		catch
		{
			// If text extraction fails for any reason, treat it as "no email found"
			// rather than blocking the whole split/grid population.
			return new List<string>();
		}
	}

	/// <summary>
	/// Reads the given address file (.xlsx, .xls, .doc, .docx or .txt) and returns every
	/// distinct email address found anywhere in its content, in the order first seen.
	/// </summary>
	internal static List<string> ExtractEmailAddressesFromFile(string filePath)
	{
		string extension = Path.GetExtension(filePath).ToLowerInvariant();

		string text = extension switch
		{
			".txt" => File.ReadAllText(filePath),
			".xlsx" or ".xls" => ExtractTextFromWorkbook(filePath),
			".docx" => ExtractTextFromWordDocument(filePath),
			".doc" => throw new NotSupportedException(
				"Legacy .doc files aren't supported. Please save the file as .docx, .xlsx, or .txt."),
			_ => throw new NotSupportedException(
				$"Unsupported address file type '{extension}'. Use .xlsx, .xls, .docx or .txt.")
		};

		return ExtractEmailAddressesFromText(text);
	}

	/// <summary>
	/// Concatenates the text of every cell on every sheet of the given Excel workbook.
	/// </summary>
	internal static string ExtractTextFromWorkbook(string filePath)
	{
		using FileStream stream = File.OpenRead(filePath);
		using IWorkbook workbook = WorkbookFactory.Create(stream);

		var textBuilder = new StringBuilder();
		for(int sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
		{
			ISheet sheet = workbook.GetSheetAt(sheetIndex);
			foreach(IRow row in sheet)
			{
				foreach(NPOI.SS.UserModel.ICell cell in row)
				{
					textBuilder.Append(cell.ToString()).Append(' ');
				}
			}
		}

		return textBuilder.ToString();
	}

	/// <summary>
	/// Extracts all text from a modern .docx (OOXML) Word document.
	/// </summary>
	internal static string ExtractTextFromWordDocument(string filePath)
	{
		using FileStream stream = File.OpenRead(filePath);
		using var document = new XWPFDocument(stream);
		var extractor = new XWPFWordExtractor(document);
		return extractor.Text;
	}

	/// <summary>
	/// Returns every distinct email address found in <paramref name="text"/>, in the order
	/// first seen.
	/// </summary>
	internal static List<string> ExtractEmailAddressesFromText(string text)
	{
		return Regex.Matches(text, @"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
			.Select(m => m.Value)
			.Distinct(StringComparer.OrdinalIgnoreCase)
			.ToList();
	}

	/// <summary>
	/// Sends an email with the given attachments via SMTP using MailKit.
	/// </summary>
	private static async Task SendMailAsync(
		string smtpHost, int smtpPort, bool useSsl,
		string fromEmail, string password, IEnumerable<string> toEmails,
		string subject, string body, IEnumerable<string> attachmentPaths)
	{
		var message = new MimeMessage();
		message.From.Add(MailboxAddress.Parse(fromEmail));
		foreach(string toEmail in toEmails)
		{
			message.To.Add(MailboxAddress.Parse(toEmail));
		}
		message.Subject = subject;

		var builder = new BodyBuilder { TextBody = body };
		foreach(string filePath in attachmentPaths)
		{
			builder.Attachments.Add(filePath);
		}
		message.Body = builder.ToMessageBody();

		using var client = new SmtpClient();
		var socketOptions = useSsl ? SecureSocketOptions.StartTls : SecureSocketOptions.Auto;

		await client.ConnectAsync(smtpHost, smtpPort, socketOptions);
		await client.AuthenticateAsync(fromEmail, password);
		await client.SendAsync(message);
		await client.DisconnectAsync(true);
	}

	internal static bool IsValidEmail(string email)
	{
		// Simple, pragmatic pattern - not a full RFC 5322 validator.
		return !string.IsNullOrWhiteSpace(email) &&
			   Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$");
	}
}
