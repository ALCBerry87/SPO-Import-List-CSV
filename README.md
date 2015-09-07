# SPO-Import-List-CSV
Console application for importing records from a CSV file into a list in SharePoint Online

Requirements:<br/>
<ol>
<li>The following must all match exactly:
<ol>
<li>The column headers of your CSV file</li>
<li>The property names of your custom class that represents a spreadsheet record</li>
<li>The internal names of the corresponding SharePoint site columns</li>
</ol>
</li>
<li>Implementation of publishing-specific columns importing is not supported </li>
</ol>

Other notes:<br/>
The error handling in this project is incomplete. I recommend a more robust error handling implementation.