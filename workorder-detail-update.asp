<%
' *************************************************************************************************
'   workorder-detail.asp - 3/7/17 by JRH
'	parse WO# passed in form data, lookup in db and display detail
'
'	pass 'new' (or blank?) to create new
'		
'		
'
'
'
'
'
'
'
'
'
' *************************************************************************************************


' *************************************************************************************************
'    First up, get (some optional, some required) page global vars that come through the query string
'    ASP functions, vars, and params passed on URL query string
' *************************************************************************************************

Option Explicit
response.buffer=true
Response.Expires = 0

Function IIf(i,j,k)
    If i Then IIf = j Else IIf = k
End Function

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adUseClient = 3

dim szSQL, szParamID, iCounter
dim szFromDate, szToDate
'dim szTemp, szSQL
Dim OBJdbConnection
Dim objRS, objRSStatus
'Dim rsSwitches
'Dim iRecordCount, iProcYesCount, iSkipFlag, iEmailAvail, iEmailSent, iCalled, iPendingCount
'Dim iDeclined, iLeftVM, iNotCalled, iRecall, iRetainedCount, iNotRetainedCount, iProcNoCount


'szFromDate = Request.QueryString("FromDate")
'szToDate = Request.QueryString("ToDate")
szParamID = Request.QueryString("ID")
szParamID = Request.Form("WO_NO_OLD")





' *************************************************************************************************
'    Make a webpage for our guest
' *************************************************************************************************
'<!-- #Include virtual ="/SCRIPTS/ADOVBS.INC" -->
%>

<HTML><HEAD><TITLE>Work Order Detail	</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<!-- *************************************************************************************************
	After HTML output page opens, put our javascript here in head
     *************************************************************************************************
-->

<script language="JavaScript">

function onUpdateWO()
{

   alert ('Ha');
   //validate data, alert error or
    document.getElementById('wo-detail').submit();

}


function zoommonth(monthtozoom)
{
// javascript code
//alert (monthtozoom)
document.forms[monthtozoom-1].submit();
//return false;

}

function onCancel()
{
window.location='workorders.asp';
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

<!-- *************************************************************************************************
	Define STYLES
     *************************************************************************************************
-->
<style>
	a:hover {
		color: #0000FF;
		text-transform:	uppercase;
		font-weight: bold;
		}
	
	.Center {
		text-align: center;
		}

	.CenterBold {
		text-align: center;
		font-weight: bold;
		}

	.CenterItalicBold {
		text-align: center;
		font-weight: bold;
		font-style: italic;
		}

	.CenterBoldLarge {
		text-align: center;
		font-weight: bold;
		font-size: 150%;
		}

	#title {
		font-size: 200%;
		}

        td {
                text-align: center;
                font-weight: bold;
                #background-color:white
           }
</style>


</HEAD>
<style>
	a:hover {
		color: #0000FF;
		text-transform:	uppercase;
		font-weight: bold;
		}
	
	body {
		background-color: #ABBDAF;
		}

	.Center {
		text-align: center;
		}

	.CenterBold {
		text-align: center;
		font-weight: bold;
		}

	.CenterItalicBold {
		text-align: center;
		font-weight: bold;
		font-style: italic;
		}

	.CenterBoldLarge {
		text-align: center;
		font-weight: bold;
		font-size: 150%;
		}

	#title {
		font-size: 200%;
		}

        td {
                text-align: center;
           }
</style>

<BODY>
<!-- *************************************************************************************************
	Start HTML BODY
     *************************************************************************************************
-->
<%
' *************************************************************************************************
'	Open db connection and get ready to update and/or query
' *************************************************************************************************
'open database connection
Set OBJdbConnection = Server.CreateObject("ADODB.Connection") 
OBJdbConnection.mode = 3 ' adModeReadWrite
OBJdbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\users\john\desktop\body\mso\mso.mdb;"

Set ObjRs = Server.CreateObject("ADODB.Recordset")
Set objRSStatus = Server.CreateObject("ADODB.Recordset")

' *************************************************************************************************
'	Prep the form - counters, query string
' *************************************************************************************************
if szParamID = "" then
   response.write ("EMPTY ID! Creating?<BR>")
   'display new form?
   'response.end
   szSQL = "select * from WO17"
   objRS.open szSQL, OBJdbConnection, adOpenStatic, adLockOptimistic
   objRS.AddNew
   szParamID = Request.Form("wo_no")
else
   response.write ("ID = '" & szParamID & "'<BR>")
   szSQL = "select * from WO17 where WO_NO ='" & szParamID & "'"
   objRS.open szSQL, OBJdbConnection, adOpenStatic, adLockOptimistic

end if




' *************************************************************************************************
'	UPDATE RECORD
' *************************************************************************************************
'lookup 'queued' 'building' etc to get value for status
szSQL = "SELECT * FROM STAGES17 WHERE NAME = '" & Request.Form("statusselection") & "'"
objRSStatus.Open szSQL, OBJdbConnection, adOpenStatic, adLockOptimistic
objRs("status") = objRSStatus("ID")


objRs("wo_no") = Request.Form("wo_no")
objRs("customer") = Request.Form("customer")
objRs("po_no") = Request.Form("po_no")
objRs("vin") = Request.Form("vin")
objRs("bodyid") = Request.Form("bodyid")
objRs("model_no") = Request.Form("model_no")
objRs("inv_no") = Request.Form("inv_no")
objRs("length") = Request.Form("length")
objRs("body_weight") = Request.Form("body_weight")
objRs("body_year") = Request.Form("body_year")
objRs("bodystyle") = Request.Form("bodystyle")

'set or clear dates
if Request.Form("order_date") <> "" then 
   objRs("order_date") = Request.Form("order_date")
else
   objRs("order_date") = Null
end if
if Request.Form("req_date") <> "" then 
   objRs("req_date") = Request.Form("req_date")
else
   objRs("req_date") = Null
end if
if Request.Form("productionstart_date") <> "" then 
   objRs("productionstart_date") = Request.Form("productionstart_date")
else
   objRs("productionstart_date") = Null
end if
if Request.Form("invoice_date") <> "" then 
   objRs("invoice_date") = Request.Form("invoice_date")
else
   objRs("invoice_date") = Null
end if

objRs.Update
objRS.Close
objRSStatus.Close
set objRS = Nothing
set objRSStatus = Nothing

Dim x

For x = 1 to Request.Form.Count 
  Response.Write x & ": " _ 
    & Request.Form.Key(x) & "=" & Request.Form.Item(x) & "<BR>" 
Next 


'wipe pkgs for this WO, loop to recreate each one from form data
szSQL = "DELETE * FROM PKGS17 WHERE WO_NO = '" & szParamID & "'"
'response.write ("<BR>" & szSQL & "<BR>")
OBJdbConnection.Execute szSQL

'for each task and option, tasks first
iCounter = 1
do while iCounter < 100
   'write record if <> ""
   if Request.Form("taskname" & iCounter) <> "" then
      if Request.Form("taskbox" & iCounter) = "on" then
         szSQL = "INSERT INTO PKGS17 (NAME, STAGE, WO_NO, COMPLETED) VALUES ('" & Request.Form("taskname" & iCounter) & "','1','" & szParamID & "', True)"
      else
         szSQL = "INSERT INTO PKGS17 (NAME, STAGE, WO_NO) VALUES ('" & Request.Form("taskname" & iCounter) & "','1','" & szParamID & "')"
      end if
      'response.write (szSQL)
      OBJdbConnection.Execute szSQL
   end if
   iCounter = iCounter + 1
Loop

'for each task and option, now options
iCounter = 1
do while iCounter < 100
   'write record if <> ""
   if Request.Form("optname" & iCounter) <> "" then
      if Request.Form("optbox" & iCounter) = "on" then
         szSQL = "INSERT INTO PKGS17 (NAME, STAGE, WO_NO, COMPLETED) VALUES ('" & Request.Form("optname" & iCounter) & "','2','" & szParamID & "', True)"
      else
         szSQL = "INSERT INTO PKGS17 (NAME, STAGE, WO_NO) VALUES ('" & Request.Form("optname" & iCounter) & "','2','" & szParamID & "')"
      end if
      'response.write (szSQL)
      OBJdbConnection.Execute szSQL
   end if

   iCounter = iCounter + 1
Loop



' *************************************************************************************************
'	Update PKGS17 if we changed the WO#
' *************************************************************************************************
if szParamID <> request.form("wo_no") then
   szSQL = "UPDATE PKGS17 SET WO_NO = '" & request.form("WO_NO") & "' WHERE WO_NO = '" & szParamID & "'"
   'update pkgs.wo to request.form("wo_no") where pkgs.wo = szParamID
   'response.write ("<BR>" & szSQL & "<BR>")
   OBJdbConnection.Execute szSQL
end if

response.write("Record Updated")

response.write("<BR><A href='workorders.asp'>Continue</A>")


OBJdbConnection.close
set OBJdbConnection = nothing


' *************************************************************************************************
'	End page
' *************************************************************************************************
response.write "</body></html>"
response.end

%>
