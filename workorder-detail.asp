<%
Option Explicit
response.buffer=true
Response.Expires = 0


Function IIf(i,j,k)
    If i Then IIf = j Else IIf = k
End Function


dim szFromDate, szToDate, szTemp, szSQL

Dim OBJdbConnection
Dim objRS
Dim rsSwitches
Dim iRecordCount, iProcYesCount, iSkipFlag, iEmailAvail, iEmailSent, iCalled, iPendingCount
Dim iDeclined, iLeftVM, iNotCalled, iRecall, iRetainedCount, iNotRetainedCount, iProcNoCount
Dim strSQL


' *************************************************************************************************
'	stats.asp - John Hilsman 2/12/2014
'	Display stats for switch away calls
'		
'		
'
' *************************************************************************************************

'Get (some optional, some required) page global vars that come through the query string
szFromDate = Request.QueryString("FromDate")
szToDate = Request.QueryString("ToDate")


'<!-- #Include virtual ="/SCRIPTS/ADOVBS.INC" -->

%>

<HTML><HEAD><TITLE>Switch-away Worksheet</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">

<script language="JavaScript">
function zoommonth(monthtozoom)
{
// javascript code
//alert (monthtozoom)
document.forms[monthtozoom-1].submit();
//return false;

}


function UpdateEmailFlag(memsep)
{

//alert ("Update " + memsep);
//window.location.href = 'http://www.google.com';
//window.location.assign("http://www.w3schools.com");
//window.location = 'http://www.google.com';

//alert ('list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&emailflag=true');
//return;

document.body.style.cursor = 'wait';

//will always have a from/to date if updating, vars are set in ASP at top of script
window.location.href = 'list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&emailflag=true';


}

function UpdateCalledStatus(memsep, controlid)
{
//alert ('list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&called='+controlid.options[controlid.selectedIndex].text);
//return;

//if status = declined, open a popup-box for notes
//if cancelled abort status change, otherwise save notes and call status

if (controlid.options[controlid.selectedIndex].text == "Declined") 
    {
	//alert ("Enter NOTES");
	var notes = prompt ("Enter Notes:");
	if (notes!=null)
	{
	    //alert ("Save notes");
	    document.body.style.cursor = 'wait';
	    window.location.href = 'list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&called='+controlid.options[controlid.selectedIndex].text + '&notes='+notes;
	}
	else
	{
	    //nevermind, do not update
	}
	return;

    }



document.body.style.cursor = 'wait';

//will always have a from/to date if updating, vars are set in ASP at top of script
//alert(memsep + ":" + controlid.options[controlid.selectedIndex].text);
window.location.href = 'list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&called='+controlid.options[controlid.selectedIndex].text;


}

function UpdateRetained(memsep, controlid)
{

document.body.style.cursor = 'wait';

//alert(memsep + ":" + controlid.options[controlid.selectedIndex].text);
window.location.href = 'list-switches-csr.asp?FromDate=<%=szFromDate%>&ToDate=<%=szToDate%>&UpdateFlag=TRUE&memsep=' + memsep + '&retained_flag='+controlid.options[controlid.selectedIndex].text;


}


</script>


</HEAD><BODY BGCOLOR=abbdAf>


<%

'workorder-detail.asp 3/7/17 by JRH

'parse WO# passed in form data, lookup in db and display detail
'   pass 'new' (or blank?) to create new



' *************************************************************************************************
'	Open db connection and get ready to update and/or query
' *************************************************************************************************
'open database connection
Set OBJdbConnection = Server.CreateObject("ADODB.Connection") 
OBJdbConnection.mode = 3 ' adModeReadWrite
OBJdbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\users\john\desktop\body\mso\mso.mdb;"

Set ObjRs = Server.CreateObject("ADODB.Recordset")
strSQL = "select * from WO17"
objRS.open strSQL, OBJdbConnection

Do While Not objRs.EOF
response.write(objRS("STATUS") & "<br>" & objRS("WO#") & "<br>")
objRS.MoveNext
Loop

OBJdbConnection.close

set OBJdbConnection = nothing



response.write "</body></html>"
response.end

%>
