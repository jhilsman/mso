<%
' *************************************************************************************************
'   workorders.asp - 4/7/17 by JRH
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
'szSQL = "select * from WO17 "
szSQL = "select *, DLookUp('[NAME]','STAGES17','ID=' & [STATUS]) AS STATUS_ from WO17 WHERE STATUS " & szFilterParam & " ORDER BY " & szSortParam

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

<HTML><HEAD><TITLE>Work Order Summary</TITLE>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<!-- *************************************************************************************************
	After HTML output page opens, put our javascript here in head
     *************************************************************************************************
-->

<script type="text/javascript">

//THIS LIMITS THE ROWS OF WO THAT CAN BE DISPLAYED - JAVA WILL CRASH IF TRY TO DYNAMICALLY SET TASKS/OPTIONS BEYOND THE LIMITS OF AVAIL VARS!
<%
iWOLineCount = 0
do while iWOLineCount < 50
   response.write("var bTask" & iWOLineCount & " = ""SHORT"";"  )
   response.write("var bOption" & iWOLineCount & " = ""SHORT"";" & vbCRLF)
   iWOLineCount = iWOLineCount + 1
loop
%>

function onFilterStatus () {
   var x = document.getElementById('statusselection').value

   window.location.href = 'workorders.asp?filter=' +x ;

}

function myFunction(myMessage) {
   alert(myMessage);

}


function onNew()
{
window.location='workorder-detail.asp';
}

function onTrucks()
{
window.location='trucks.asp';
}

<%
' *************************************************************************************************
'    Write Java functions for flipping tasks and options
' *************************************************************************************************
iWOLineCount = 0
Do While Not objRs.EOF
   iWOLineCount = iWOLineCount + 1
   Set objPkgsRS = Server.CreateObject("ADODB.Recordset")
   szSQL = "SELECT * FROM PKGS17 where WO_NO = '" & objRS("WO_NO") & "'"
   objPkgsRS.open szSQL, OBJdbConnection
   szTasks = ""
   szOptions = ""
   iTaskCompleteCount = 0
   iTaskCount = 0
   iOptCompleteCount = 0
   iOptCount = 0
   do while not objPkgsRS.EOF
      'gather Tasks
      if objPkgsRS("STAGE") = 1 then
         if objPkgsRS("COMPLETED") then 
	    iTaskCompleteCount = iTaskCompleteCount + 1
            szTasks = szTasks & "<TR> <TD> <input type='checkbox' disabled='disabled' checked='checked' id='task" & iWOLineCount & "-" & iTaskCount & "'> </TD> <TD> " & objPkgsRS("NAME") & " </TD> </TR> "
         else
            szTasks = szTasks & "<TR> <TD> <input type='checkbox' disabled='disabled' id='task" & iWOLineCount & "-" & iTaskCount & "'> </TD> <TD> " & objPkgsRS("NAME") & "</TD> </TR>"
         end if
         iTaskCount = iTaskCount + 1
      end if
      'gather options
      if objPkgsRS("STAGE") = 2 then
         if objPkgsRS("COMPLETED") then 
	    iOptCompleteCount = iOptCompleteCount + 1
            szOptions = szOptions & "<TR> <TD> <input type='checkbox' disabled='disabled' checked='checked' id='option" & iWOLineCount & "-" & iOptCount & "'> </TD> <TD> " & objPkgsRS("NAME") & " </TD> </TR> " 
         else
            szOptions = szOptions & "<TR> <TD> <input type='checkbox' disabled='disabled' id='option" & iWOLineCount & "-" & iOptCount & "'> </TD> <TD> " & objPkgsRS("NAME") & " </TD> </TR> " 
         end if
         iOptCount = iOptCount + 1
      end if
      objPkgsRS.MoveNext
   Loop


   'FLIPTASK () for this WO
   response.write (vbCRLF & vbCRLF & "function FlipTask" & iWOLineCount & "() { " & _
   " if (bTask" & iWOLineCount & " == ""SHORT"") { " & _
      " bTask" & iWOLineCount & " = ""FULL""; " & _
      " document.getElementById(""tasks" & iWOLineCount & """).innerHTML = ""<A href='javascript:FlipTask" & iWOLineCount & "()'>" & iTaskCompleteCount & " of " & iTaskCount & " </A><BR><TABLE ID='tasklist" & iWOLineCount & "'>" & szTasks & "</TABLE><BR>""; " & _
   " }   else { " & _
   "  bTask" & iWOLineCount & " = ""SHORT""; " & _
   "  document.getElementById(""tasks" & iWOLineCount & """).innerHTML = ""<A href='javascript:FlipTask" & iWOLineCount & "()'>" & iTaskCompleteCount & " of " & iTaskCount & " </A>""; } } ") 

   'FLIPOPT() for this WO
   response.write (vbCRLF & vbCRLF & "function FlipOp" & iWOLineCount & "() { " & _
   " if (bOption" & iWOLineCount & " == ""SHORT"") { " & _
      " bOption" & iWOLineCount & " = ""FULL""; " & _
      " document.getElementById(""options" & iWOLineCount & """).innerHTML = ""<A href='javascript:FlipOp" & iWOLineCount & "()'>" & iOptCompleteCount & " of " & iOptCount & " </A><BR><TABLE ID='optionslist" & iWOLineCount & "'>" &  szOptions & "</TABLE><BR>""; " & _
   " }   else { " & _
   "  bOption" & iWOLineCount & " = ""SHORT""; " & _
   "  document.getElementById(""options" & iWOLineCount & """).innerHTML = ""<A href='javascript:FlipOp" & iWOLineCount & "()'>" & iOptCompleteCount & " of " & iOptCount & " </A>""; } } " ) 


   objRS.MoveNext
   objPkgsRS.Close
   set objPkgsRS = Nothing
  
Loop


%>



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
<TD><A href='workorders.asp?order=STATUS'>STATUS</A> <BR><select id="statusselection" name="statusselection" onchange="onFilterStatus()">
  <option value=""></option>
  <option value="NOTDELIVERED">NOT DELIVERED</option>
  <option value="QUEUED">QUEUED</option>
  <option value="BUILDING">BUILDING</option>
  <option value="COMPLETED">COMPLETED</option>
  <option value="DELIVERED">DELIVERED</option>
</select></TD> 
<TD><A href='workorders.asp?order=WO_NO'>WORK ORDER #<A> <BR><INPUT type='textbox' size='8' id='wosearch'><input type='button' value='S' id='wosearchbutton'></TD> 
<TD>TASKS</TD>
<TD>OPTIONS</TD> 
<TD><A href='workorders.asp?order=CUSTOMER'>CUSTOMER</A> <BR><INPUT type='textbox' size='8' id='custsearch'><input type='button' value='S' id='custsearchbutton'></TD> 
<TD><A href='workorders.asp?order=ORDER_DATE'>ORDER DATE</A> </TD> 
<TD><A href='workorders.asp?order=REQ_DATE'>REQ DATE</A> </TD> 
<TD><A href='workorders.asp?order=PRODUCTIONSTART_DATE'>PROD DATE</TD>
<TD><A href='workorders.asp?order=VIN'>VIN #</A> </TD>
</TR>

<%
' *********************************************
'  DISPLAY RECORDS FROM DATABASE
' *********************************************
objRS.MoveFirst
iWOLineCount = 0
Do While Not objRs.EOF
   iWOLineCount = iWOLineCount + 1
   ' do lookup on PKGS17 for PKGS17.WO_NO = WO17.WO_NO, display sum as link:
   '<A href='javascript:FlipTask1()'>0 of 2 </A>
   '<A href='javascript:FlipOp1()'>0 of 2 </A>
   Set objPkgsRS = Server.CreateObject("ADODB.Recordset")
   szSQL = "SELECT * FROM PKGS17 where WO_NO = '" & objRS("WO_NO") & "'"
   'response.write szSQL
   objPkgsRS.open szSQL, OBJdbConnection
   szTasks = ""
   szOptions = ""
   iTaskCompleteCount = 0
   iTaskCount = 0
   iOptCompleteCount = 0
   iOptCount = 0
   do while not objPkgsRS.EOF
      'gather Tasks
      if objPkgsRS("STAGE") = 1 then
         if objPkgsRS("COMPLETED") then iTaskCompleteCount = iTaskCompleteCount + 1
         iTaskCount = iTaskCount + 1
      end if
      'gather options
      if objPkgsRS("STAGE") = 2 then
         if objPkgsRS("COMPLETED") then iOptCompleteCount = iOptCompleteCount + 1
         iOptCount = iOptCount + 1
      end if
      objPkgsRS.MoveNext
   Loop
   szTasks = szTasks & "<A href='javascript:FlipTask" & iWOLineCount & "()'>" & iTaskCompleteCount & " of " & iTaskCount &  "</A> "
   szOPtions = szOptions & "<A href='javascript:FlipOp" & iWOLineCount & "()'>" & iOptCompleteCount & " of " & iOptCount &  "</A> "

   response.write("<tr><td> " & objRS("STATUS_") & "</td><td> <A href='workorder-detail.asp?id=" & objRS("WO_NO") & "'>" & objRS("WO_NO") & "</a> </td><td id='tasks" & iWOLineCount & "'>" & szTasks & "</td><td id='options" & iWOLineCount & "'>" & szOptions & "</td><td> " & objRS("CUSTOMER") & " </td><td> " & objRS("ORDER_DATE") & " </td><td> " & objRS("REQ_DATE") & " </td><td> " & objRS("PRODUCTIONSTART_DATE") & "</td><td> <A href='truck-detail.asp?vin=" & objRS("VIN") & "'>" & objRS("VIN") & "</td> </tr>")
   objRS.MoveNext
   objPkgsRS.Close
   set objPkgsRS = Nothing
  
Loop

   objRS.Close
   OBJdbConnection.Close
   set objRS = Nothing
   set OBJdbConnection = Nothing


%>


</table>
<BR>
<input type='button' value='NEW' id='newbutton' onclick='onNew()'> &nbsp; <input type='button' value='TRUCKS' id='newbutton' onclick='onTrucks()'> &nbsp;
<BR>

</BODY></HTML>
