<%
' *************************************************************************************************
'   trucks.asp - 4/18/17 by JRH
'	lookup all (or filtered) WO in db and display each in a <TR>
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


dim szSQL, szParamID, szTasks, szOptions, szFilterParam
dim szFromDate, szToDate
Dim OBJdbConnection
Dim objRS, objPkgsRS, objRSStatus
Dim iWOLineCount, iTaskCompleteCount, iTaskCount, iOPtCompleteCount, iOptCount
Dim szSortParam

szSortParam = Request.QueryString("ORDER")
szFilterParam = Request.QueryString("FILTER")
if Len(szSortParam) < 2 then szSortParam = "ORDER_DATE"
If Len(szFilterParam) < 2 or szFilterParam = "NOTDELIVERED" then szFilterParam = "<> '4'"

' *************************************************************************************************
'	Open db connection and get ready to update and/or query
' *************************************************************************************************
'open database connection
Set OBJdbConnection = Server.CreateObject("ADODB.Connection") 
OBJdbConnection.mode = 3 ' adModeReadWrite
OBJdbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\users\john\desktop\body\mso\mso.mdb;"

Set ObjRs = Server.CreateObject("ADODB.Recordset")

'get filter status if needed
if szFilterParam <> "<> '4'" then
   Set ObjRSStatus = Server.CreateObject("ADODB.Recordset")
   szSQL = "SELECT * FROM STAGES17 WHERE NAME = '" & szFilterParam & "'"
   objRSStatus.Open szSQL, OBJdbConnection
   szFilterParam = " = '" & ObjRSStatus("ID") &  "'"
   ObjRSStatus.Close
   set ObjRSStatus = Nothing
end if


'get WOs
'szSQL = "select * from TRUCKS17"
szSQL = "SELECT TRUCKS17.VIN, TRUCKS17.CUSTOMER, TRUCKS17.WO_NO, WO17.ORDER_DATE, WO17.REQ_DATE, WO17.PRODUCTIONSTART_DATE, TRUCKS17.RECV_DATE FROM TRUCKS17 LEFT JOIN WO17 ON TRUCKS17.VIN = WO17.VIN ORDER BY " & szSortParam

objRS.open szSQL, OBJdbConnection

if objRS.EOF then
   objRS.Close
   OBJdbConnection.Close
   set objRS = Nothing
   set OBJdbConnection = Nothing
   response.write ("<HTML><BODY bgcolor='ABBDAF'>NO RECORDS FOUND!</BODY></HTML>")
   response.end
else
   objRS.MoveFirst
end if


' *************************************************************************************************
'    Make a webpage for our guest
' *************************************************************************************************
'<!-- #Include virtual ="/SCRIPTS/ADOVBS.INC" -->
%>

<HTML><HEAD><TITLE>Truck Summary</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<!-- *************************************************************************************************
	After HTML output page opens, put our javascript here in head
     *************************************************************************************************
-->

<script type="text/javascript">

function onFilterStatus () {
   var x = document.getElementById('statusselection').value

   window.location.href = 'trucks.asp?filter=' +x ;

}

function myFunction(myMessage) {
   alert(myMessage);

}


function onNew()
{
window.location='truck-detail.asp';
}

function onWorkOrders()
{
window.location='workorders.asp';
}


</script>
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
<BODY >

<table width='100%' border='1' id='wolist'>
<TR>
<TD><A href='trucks.asp?order=TRUCKS17.VIN'>VIN #</A> </TD>
<TD><A href='trucks.asp?order=TRUCKS17.CUSTOMER'>CUSTOMER</A> <BR><INPUT type='textbox' size='8' id='custsearch'><input type='button' value='S' id='custsearchbutton'></TD> 
<TD><A href='trucks.asp?order=TRUCKS17.WO_NO'>WORK ORDER #<A> <BR><INPUT type='textbox' size='8' id='wosearch'><input type='button' value='S' id='wosearchbutton'></TD> 
<TD><A href='trucks.asp?order=ORDER_DATE'>ORDER DATE</A> </TD> 
<TD><A href='trucks.asp?order=REQ_DATE'>REQ DATE</A> </TD> 
<TD><A href='trucks.asp?order=PRODUCTIONSTART_DATE'>PROD DATE</TD>
<TD><A href='trucks.asp?order=RECV_DATE'>RECV DATE</TD>

</TR>

<%
' *********************************************
'  DISPLAY RECORDS FROM DATABASE
' *********************************************
objRS.MoveFirst
iWOLineCount = 0
Do While Not objRs.EOF
   iWOLineCount = iWOLineCount + 1
   response.write("<tr><td> <A href='truck-detail.asp?vin=" & objRS("VIN") & "'>" & objRS("VIN") & "</td> <td> " & objRS("CUSTOMER") & " </td> <td> <A href='workorder-detail.asp?id=" & objRS("WO_NO") & "'>" & objRS("WO_NO") & "</a> </td> <td> " & objRS("ORDER_DATE") & " </td><td> " & objRS("REQ_DATE") & " </td> <td> " & objRS("PRODUCTIONSTART_DATE") & "</td> <td> " & objRS("RECV_DATE") & "</td> </tr>")
   objRS.MoveNext
  
Loop

   objRS.Close
   OBJdbConnection.Close
   set objRS = Nothing
   set OBJdbConnection = Nothing


%>


</table>
<BR>
<input type='button' value='NEW' id='newbutton' onclick='onNew()'> &nbsp; <input type='button' value='WORK ORDERS' id='truckbutton' onclick='onWorkOrders()'> &nbsp;
<BR>

</BODY></HTML>
