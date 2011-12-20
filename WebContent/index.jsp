<html>
<head>
<title>Jasper Report Example</title>
</head>
<body>
<form name="form_jasper" method="post" action="ReportServlet">
<select name="format">
    <option value="">Select</option>
	<option value="xls">XLS</option>
	<option value="csv">CSV</option>
	<option value="docx">DOCX</option>
	<option value="html">HTML</option>
	<option value="pdf">PDF</option>
	<option value="ods">ODS</option>
	<option value="odt">ODT</option>
	<option value="txt">TXT</option>
	<option value="rtf">RTF</option>
</select>
<br /><br />
	<input type="submit" value="Generate"/>
</form>
</body>
</html>