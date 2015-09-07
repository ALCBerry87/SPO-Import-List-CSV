# SPO-Import-List-CSV
Console application for importing records from a CSV file into a list in SharePoint Online

Requirements:
1. The following must all match exactly:
    a. The column headers of your CSV file
    b. The property names of your custom class that represents a spreadsheet record
    c. The internal names of the corresponding SharePoint site columns
2. Your spreadsheet columns cannot contain data intended to be imported to publishing-specific columns - this is not supported

Other notes:
The error handling in this project is incomplete. I recommend a more robust error handling implementation.