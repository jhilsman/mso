<%
' *************************************************************************************************
'   truck-detail-update.asp - 4/18/17 by JRH
'	update truck with VIN passed in querystring
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
szParamID = Request.QueryString("VIN")
szParamID = Request.Form("VIN_OLD")





' *************************************************************************************************
'    Make a webpage for our guest
' *************************************************************************************************
'<!-- #Include virtual ="/SCRIPTS/ADOVBS.INC" -->
%>

<HTML><HEAD><TITLE>Truck Detail Update</TITLE>
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



function onCancel()
{
window.location='trucks.asp';
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
   szSQL = "select * from TRUCKS17"
   objRS.open szSQL, OBJdbConnection, adOpenStatic, adLockOptimistic
   objRS.AddNew
   szParamID = Request.Form("VIN")
else
   response.write ("VIN = '" & szParamID & "'<BR>")
   szSQL = "select * from TRUCKS17 where VIN ='" & szParamID & "'"
   response.write szSQL
   objRS.open szSQL, OBJdbConnection, adOpenStatic, adLockOptimistic

end if




' *************************************************************************************************
'	UPDATE RECORD
' *************************************************************************************************


objRs("wo_no") = Request.Form("wo_no")
objRs("customer") = Request.Form("customer")
objRs("vin") = Request.Form("vin")
objRs("mod_no") = Request.Form("model")
objRs("truckyear") = Request.Form("year")
objRs("make") = Request.Form("make")
objRs("chassisweight") = Request.Form("chassisweight")

'set or clear dates
if Request.Form("recv_date") <> "" then 
   objRs("recv_date") = Request.Form("recv_date")
else
   objRs("recv_date") = Null
end if

objRs.Update
objRS.Close
set objRS = Nothing
set objRSStatus = Nothing

Dim x

For x = 1 to Request.Form.Count 
  Response.Write x & ": " _ 
    & Request.Form.Key(x) & "=" & Request.Form.Item(x) & "<BR>" 
Next 


response.write("Record Updated")

response.write("<BR><A href='trucks.asp'>Continue</A>")


OBJdbConnection.close
set OBJdbConnection = nothing


' *************************************************************************************************
'	End page
' *************************************************************************************************
response.write "</body></html>"
response.end

%>
