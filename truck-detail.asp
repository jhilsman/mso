<%
' *************************************************************************************************
'   truck-detail.asp - 4/18/17 by JRH
'	parse VIN# passed in form data, lookup in db and display detail
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


dim szSQL, szParamID, szStatus
dim szFromDate, szToDate
'dim szTemp, szSQL, 
Dim OBJdbConnection
Dim objRS, objRSStatus
Dim iTaskCount, iOptCount

'Dim rsSwitches
'Dim iRecordCount, iProcYesCount, iSkipFlag, iEmailAvail, iEmailSent, iCalled, iPendingCount
'Dim iDeclined, iLeftVM, iNotCalled, iRecall, iRetainedCount, iNotRetainedCount, iProcNoCount


'szFromDate = Request.QueryString("FromDate")
'szToDate = Request.QueryString("ToDate")
szParamID = Request.QueryString("VIN")



' *************************************************************************************************
'    Make a webpage for our guest
' *************************************************************************************************
'<!-- #Include virtual ="/SCRIPTS/ADOVBS.INC" -->
%>

<HTML><HEAD><TITLE>Truck Detail	</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<!-- *************************************************************************************************
	After HTML output page opens, put our javascript here in head
     *************************************************************************************************
-->

<script language="JavaScript">
var NumOfTask = 1;
var NumOfOpt = 1;

function onUpdateWO()
{

   var v = document.getElementById('vin').value;

   //validate data, alert error or
   if ( v == "" || v == null ) {
      alert ('VIN is required');
   } else {
      //alert ('Good VIN');
      document.getElementById('wo-detail').submit();
   }

}


function onCancel()
{
//window.location='trucks.asp';
window.history.back();
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

<BODY >
<!-- *************************************************************************************************
	Start HTML BODY
     *************************************************************************************************
-->
<%
' *************************************************************************************************
'	Prep the form - counters, query string
' *************************************************************************************************
if szParamID = "" then
   response.write ("EMPTY ID! Creating?<BR>")
   'display new form?
   'response.end
else
   response.write ("VIN = '" & szParamID & "'<BR>")
   szSQL = "SELECT * FROM TRUCKS17 where VIN ='" & szParamID & "'"

end if
'response.write ( szSQL & "<br>")
response.write ("<FORM name='wo-detail' id='wo-detail' action='truck-detail-update.asp' method='post'>")
%>

<!-- *************************************************************************************************
	Main WO Table and header
     *************************************************************************************************
-->
<table width='100%' border='1' id='wolist'>
<TR>
<TD>VIN</TD> 
<TD>CUSTOMER</TD> 
<TD>WORKORDER #</TD> 
<TD>ORDER DATE</TD>
<TD></TD> 
<TD>REQUEST DATE</TD> 
<TD></TD> 
<TD>DELIVERED DATE</TD> 
</TR>

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


'if not a new record, read it
if szParamID <> "" then 
   objRS.open szSQL, OBJdbConnection

   ' *************************************************************************************************
   '	List row or search, close DB
   ' *************************************************************************************************
   'doesn't really have to be a loop - should never be > 1 truck with this VIN!
   if objRs.EOF then
      ' *************************************************************************************************
      '	For Search, change sql to like VIN and requery
      ' *************************************************************************************************
      'szSQL = "SELECT TRUCKS17.VIN, TRUCKS17.CUSTOMER, TRUCKS17.WO_NO, WO17.ORDER_DATE, WO17.REQ_DATE, WO17.PRODUCTIONSTART_DATE, TRUCKS17.RECV_DATE, WO17.DELIVERED_DATE FROM TRUCKS17 where VIN LIKE '%" & szParamID & "%'  LEFT JOIN WO17 ON TRUCKS17.VIN = WO17.VIN " 'ORDER BY WO17.DELIVERED_DATE
      szSQL = "SELECT TRUCKS17.VIN, TRUCKS17.CUSTOMER, TRUCKS17.WO_NO, WO17.ORDER_DATE, WO17.REQ_DATE, WO17.DELIVERED_DATE FROM TRUCKS17 LEFT JOIN WO17 ON TRUCKS17.VIN = WO17.VIN WHERE TRUCKS17.VIN LIKE '%" & szParamID & "%' ORDER BY DELIVERED_DATE "
      'response.write szSQL
      objRs.Close
      objRS.open szSQL, OBJdbConnection
      Do While Not objRs.EOF
         response.write("<tr><td> <A href='truck-detail.asp?vin=" & objRS("VIN") & "'>" & objRS("VIN") & "</td>  <td> " & objRS("CUSTOMER") & " </td><td> <A href='workorder-detail.asp?id=" & objRS("WO_NO") & "'>" & objRS("WO_NO") & "</a> </td> <td> " & objRS("ORDER_DATE") & " </td><td> " & " " & " </td><td> " & objRS("REQ_DATE") & " </td> <td> " & "" & "</td> <td>" & objRS("DELIVERED_DATE") & "</td> </tr>") 
         objRS.MoveNext
      Loop
      objRS.Close
      set objRS = Nothing
      set objRSStatus = Nothing
      OBJdbConnection.close
      set OBJdbConnection = nothing
      'end page
      response.write "</table> <div align='right'><input type='button' name='cancel' id='cancel' value='Cancel' onclick='onCancel();'> </div>"
      response.write "</form> </body> </html>"
      response.end
   else
      'FOR EXACT MATCH
      response.write("<tr><td> <input type='hidden' id='vin_old' name='vin_old' value='" & objRS("VIN") & "'> <input type='input' id='vin' name='vin' value='" & objRS("VIN") & "'>  </td> <td> <input type='input' id='customer' name='customer' value='" & objRS("CUSTOMER") & "'> </td><td> <input type='input' size='12' id='wo_no' name='wo_no' value='" & objRS("WO_NO") & "'> <A href='workorder-detail.asp?id=" & objRS("WO_NO") & "'>" & objRS("WO_NO") & "</a> </td> <td> <input type='input' size='12' id='make' name='make' value='" & objRS("MAKE") & "'> </td><td> <input type='input' size='12' id='model' name='model' value='" & objRS("MOD_NO") & "'> </td><td> <input type='input' size='12' id='year' name='year' value='" & objRS("TRUCKYEAR") & "'> </td> <td><input type='input' size='12' id='chassisweight' name='chassisweight' value='" & objRS("CHASSISWEIGHT") & "'></td> <td><input type='input' size='12' id='recv_date' name='recv_date' value='" & objRS("RECV_DATE") & "'></td> </tr>") 
   end if


   objRS.Close

else
   'new record, display form
   response.write("<tr><td> <input type='input' id='vin' name='vin' >  </td> <td> <input type='input' id='customer' name='customer' > </td><td> <input type='input' size='12' id='wo_no' name='wo_no' ></td> <td> <input type='input' size='12' id='make' name='make' > </td><td> <input type='input' size='12' id='model' name='model' > </td><td> <input type='input' size='12' id='year' name='year' > </td> <td><input type='input' size='12' id='chassisweight' name='chassisweight' ></td> <td><input type='input' size='12' id='recv_date' name='recv_date' ></td> </tr>") 
end if

' *************************************************************************************************
'	End WO item detail, show tasks and options
' *************************************************************************************************
response.write "</table>"




' *************************************************************************************************
'	End MAIN WO table and form
' *************************************************************************************************
set objRS = Nothing
set objRSStatus = Nothing
OBJdbConnection.close
set OBJdbConnection = nothing


response.write "<div align='right'><input type='button' name='cancel' id='cancel' value='Cancel' onclick='onCancel();'> &nbsp; &nbsp; <input type='button' name='update' id='update' value='Update' onclick='onUpdateWO();'></div>"
response.write "</form>"


' *************************************************************************************************
'	End page
' *************************************************************************************************
response.write "</body></html>"
response.end

%>
