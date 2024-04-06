**Excel Email Sender**

**Description:**
This VBA macro automates the process of sending emails using Microsoft Outlook based on data from an Excel worksheet. It loops through each row in the worksheet, creates an email for each recipient, and sends it using Outlook. The email details such as recipient, subject, body, and sender's email address are fetched from specific columns in the Excel sheet.

**Features:**
- Automates sending emails to multiple recipients from an Excel worksheet.
- Customizable email details such as subject, body, and sender's email address.
- Error handling to handle exceptions during execution.

**Usage:**
1. Ensure that Microsoft Outlook is installed on your system.
2. Open the Excel workbook containing the data from which emails need to be sent.
3. Enable the "Developer" tab in Excel if not already enabled.
4. Access the Visual Basic for Applications (VBA) editor by clicking on the "Developer" tab and then "Visual Basic".
5. Insert a new module and paste the provided VBA macro code.
6. Customize the macro according to your requirements (e.g., adjust the range of cells for recipient email addresses, subject, body, etc.).
7. Save the workbook as a macro-enabled Excel file (.xlsm).
8. Run the macro from the "Developer" tab in Excel.

**Example:**
```vba
Sub SendEmails()
    ' VBA macro code goes here...
End Sub
```

**Notes:**
- This macro relies on Microsoft Outlook for sending emails. Ensure Outlook is properly configured on your system.
- Modify the macro to suit your specific Excel worksheet layout and email requirements.
- Replace the default sender's email address with your actual email address.
- Error handling is included to handle any issues during execution. Check the error message for troubleshooting.

**License:**
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

**Contributing:**
Contributions are welcome. Please fork the repository, make your changes, and submit a pull request.

**Author:**
Rakkesh R
**Contact:**
rakkesh30.mbm@gmail.com

