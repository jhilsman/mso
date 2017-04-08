<%
' *************************************************************************************************
'   workorders.asp - 4/7/17 by JRH
'	lookup all WO in db and display detail
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


dim szSQL, szParamID, szTasks, szOptions
dim szFromDate, szToDate
Dim OBJdbConnection
Dim objRS, objPkgsRS
Dim iWOLineCount, iTaskCompleteCount, iTaskCount, iOPtCompleteCount, iOptCount


' *************************************************************************************************
'	Open db connection and get ready to update and/or query
' *************************************************************************************************
'open database connection
Set OBJdbConnection = Server.CreateObject("ADODB.Connection") 
OBJdbConnection.mode = 3 ' adModeReadWrite
OBJdbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\users\john\desktop\body\mso\mso.mdb;"

Set ObjRs = Server.CreateObject("ADODB.Recordset")
'szSQL = "select * from WO17 "
szSQL = "select *, DLookUp('[NAME]','STAGES17','ID=' & [STATUS]) AS STATUS_ from WO17 "

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

var bTask1 = "SHORT";
var bTask2 = "SHORT";
var bTask3 = "SHORT";
var bTask4 = "SHORT";
var bTask5 = "SHORT";

var bOption1 = "SHORT";
var bOption2 = "SHORT";
var bOption3 = "SHORT";
var bOption4 = "SHORT";
var bOption5 = "SHORT";

function myFunction(myMessage) {
   alert(myMessage);

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
   'szTasks = szTasks & "<A href='javascript:FlipTask" & iWOLineCount & "()'>" & iTaskCompleteCount & " of " & iTaskCount &  "</A> "
   'szOPtions = szOptions & "<A href='javascript:FlipOp" & iWOLineCount & "()'>" & iOptCompleteCount & " of " & iOptCount &  "</A> "

   'response.write("<tr><td> " & objRS("STATUS_") & "</td><td> " & objRS("WO_NO") & "</td><td id='tasks" & iWOLineCount & "'>" & szTasks & "</td><td id='Options" & iWOLineCount & "'>" & szOptions & "</td><td> " & objRS("CUSTOMER") & " </td><td> " & objRS("ORDER_DATE") & " </td><td> " & objRS("REQ_DATE") & " </td><td> " & objRS("PRODUCTIONSTART_DATE") & "</td><td> " & objRS("VIN") & "</td> </tr>")
   response.write ("function FlipTask1() { " & _
   " if (bTask1 == ""SHORT"") { " & _
      " bTask1 = ""FULL""; " & _
      " document.getElementById(""tasks1"").innerHTML = ""<A href='javascript:FlipTask1()'>0 of 2 </A><BR><TABLE ID='tasklist1'><TR><TD><input type='checkbox' id='task1-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' id='task1-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>""; " & _
   " }   else { " & _
   "  bTask1 = ""SHORT""; " & _
   "  document.getElementById(""tasks1"").innerHTML = ""<A href='javascript:FlipTask1()'>0 of 2 </A>""; } } ") 

   objRS.MoveNext
   objPkgsRS.Close
   set objPkgsRS = Nothing
  
Loop


%>

function FlipTask1() {
   if (bTask1 == "SHORT") {
      //expand it
      bTask1 = "FULL";
      document.getElementById("tasks1").innerHTML = "<A href='javascript:FlipTask1()'>0 of 2 </A><BR><TABLE ID='tasklist1'><TR><TD><input type='checkbox' id='task1-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' id='task1-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bTask1 = "SHORT";
      document.getElementById("tasks1").innerHTML = "<A href='javascript:FlipTask1()'>0 of 2 </A>";
   }

}

function FlipOp1() {
   if (bOption1 == "SHORT") {
      //expand it
      bOption1 = "FULL";
      document.getElementById("options1").innerHTML = "<A href='javascript:FlipOp1()'>0 of 2 </A><BR><TABLE ID='optionslist1'><TR><TD><input type='checkbox' id='option1-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' id='option1-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bOption1 = "SHORT";
      document.getElementById("options1").innerHTML = "<A href='javascript:FlipOp1()'>0 of 2 </A>";
   }

}

function FlipTask2() {
   if (bTask2 == "SHORT") {
      //expand it
      bTask2 = "FULL";
      document.getElementById("tasks2").innerHTML = "<A href='javascript:FlipTask2()'>1 of 2 </A><BR><TABLE ID='tasklist2'><TR><TD><input type='checkbox' checked='checked' id='task2-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' id='task2-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bTask2 = "SHORT";
      document.getElementById("tasks2").innerHTML = "<A href='javascript:FlipTask2()'>1 of 2 </A>";
   }

}

function FlipTask3() {
   if (bTask3 == "SHORT") {
      //expand it
      bTask3 = "FULL";
      document.getElementById("tasks3").innerHTML = "<A href='javascript:FlipTask3()'>2 of 2 </A><BR><TABLE ID='tasklist3'><TR><TD><input type='checkbox' checked='checked' id='task3-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' checked='checked' id='task3-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bTask3 = "SHORT";
      document.getElementById("tasks3").innerHTML = "<A href='javascript:FlipTask3()'>2 of 2 </A>";
   }

}

function FlipTask4() {
   if (bTask4 == "SHORT") {
      //expand it
      bTask4 = "FULL";
      document.getElementById("tasks4").innerHTML = "<A href='javascript:FlipTask4()'>2 of 2 </A><BR><TABLE ID='tasklist4'><TR><TD><input type='checkbox' checked='checked' id='task4-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' checked='checked' id='task4-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bTask4 = "SHORT";
      document.getElementById("tasks4").innerHTML = "<A href='javascript:FlipTask4()'>2 of 2 </A>";
   }

}

function FlipTask5() {
   if (bTask5 == "SHORT") {
      //expand it
      bTask5 = "FULL";
      document.getElementById("tasks5").innerHTML = "<A href='javascript:FlipTask5()'>2 of 2 </A><BR><TABLE ID='tasklist5'><TR><TD><input type='checkbox' checked='checked' id='task5-1'></TD> <TD>PTO System</TD> </TR><TR> <TD><input type='checkbox' checked='checked' id='task5-2' ></TD> <TD>Sunroof</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bTask5 = "SHORT";
      document.getElementById("tasks5").innerHTML = "<A href='javascript:FlipTask5()'>2 of 2 </A>";
   }

}


function FlipOp2() {
   if (bOption2 == "SHORT") {
      //expand it
      bOption2 = "FULL";
      document.getElementById("options2").innerHTML = "<A href='javascript:FlipOp2()'>1 of 2 </A><BR><TABLE ID='optionslist2'><TR><TD><input type='checkbox' id='option2-1' checked='checked'></TD> <TD>CD Player</TD></TR><TR><TD><input type='checkbox' id='option2-2'></TD><TD>Tinted Windows</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bOption2 = "SHORT";
      document.getElementById("options2").innerHTML = "<A href='javascript:FlipOp2()'>1 of 2 </A>";
   }

}

function FlipOp3() {
   if (bOption3 == "SHORT") {
      //expand it
      bOption3 = "FULL";
      document.getElementById("options3").innerHTML = "<A href='javascript:FlipOp3()'>1 of 2 </A><BR><TABLE ID='optionslist3'><TR><TD><input type='checkbox' id='option3-1' checked='checked'></TD> <TD>Power Windows</TD></TR><TR><TD><input type='checkbox' id='option3-2'></TD> <TD>Stretch 17.25 inches</TD></TR></TABLE><BR>";
   }
   else {
      //collapse it
      bOption3 = "SHORT";
      document.getElementById("options3").innerHTML = "<A href='javascript:FlipOp3()'>1 of 2 </A>";
   }

}

function FlipOp4() {
   if (bOption4 == "SHORT") {
      //expand it
      bOption4 = "FULL";
      document.getElementById("options4").innerHTML = "<A href='javascript:FlipOp4()'>2 of 2 </A><BR><TABLE ID='optionslist4'><TR><TD><input type='checkbox' id='option4-1' checked='checked'></TD> <TD>Dovetail</TD></TR><TR><TD><input type='checkbox' checked='checked' id='option4-2'></TD> <TD>Sleeper</TD> </TR></TABLE><BR>";
   }
   else {
      //collapse it
      bOption4 = "SHORT";
      document.getElementById("options4").innerHTML = "<A href='javascript:FlipOp4()'>2 of 2 </A>";
   }

}
function FlipOp5() {
   if (bOption5 == "SHORT") {
      //expand it
      bOption5 = "FULL";
      document.getElementById("options5").innerHTML = "<A href='javascript:FlipOp5()'>2 of 2 </A><BR><TABLE ID='optionslist5'><TR><TD><input type='checkbox' id='option5-1' checked='checked'></TD> <TD>CB Radio</TD></TR><TR><TD><input type='checkbox' checked='checked' id='option5-2'></TD> <TD>Television</TD> </TR></TABLE><BR>";
   }
   else {
      //collapse it
      bOption5 = "SHORT";
      document.getElementById("options5").innerHTML = "<A href='javascript:FlipOp5()'>2 of 2 </A>";
   }

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
<!-- 
<TD>BODYSTYLE</TD>
<TD>LENGTH</TD>
<TD>BODYWEIGHT</TD>
<TD>PKGS</TD>
<TD>INVOICEDATE</TD>
<TD>INVNUM</TD>
<TD>BODYYEAR</TD>
<TD>TRADENAME</TD>
<TD>MODELNO</TD>

<TD> <A href='#' onclick='myFunction("blah");'>0 of 5 </A></TD>
<TD> <A href='javascript:myFunction("blah")'>0 of 5 </A></TD>

-->
<table width='100%' border='1' id='wolist'>
<TR>
<TD>STATUS<BR><select id="STATUSSELECTION">
  <option selected="SELECTED" value="ALL">ALL</option>
  <option value="QUEUED">QUEUED</option>
  <option value="BUILDING">BUILDING</option>
  <option value="COMPLETED">COMPLETED</option>
  <option value="DELIVERED">DELIVERED</option>
</select></TD> 
<TD>WORK ORDER #<BR><INPUT type='textbox' size='8' id='wosearch'><input type='button' value='S' id='wosearchbutton'></TD> 
<TD>TASKS</TD>
<TD>OPTIONS</TD> 
<TD>CUSTOMER <BR><INPUT type='textbox' size='8' id='custsearch'><input type='button' value='S' id='custsearchbutton'></TD> 
<TD>ORDER DATE</TD> 
<TD>REQ DATE</TD> 
<TD>PROD DATE</TD>
<TD>VIN #</TD>
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

   response.write("<tr><td> " & objRS("STATUS_") & "</td><td> " & objRS("WO_NO") & "</td><td id='tasks" & iWOLineCount & "'>" & szTasks & "</td><td id='Options" & iWOLineCount & "'>" & szOptions & "</td><td> " & objRS("CUSTOMER") & " </td><td> " & objRS("ORDER_DATE") & " </td><td> " & objRS("REQ_DATE") & " </td><td> " & objRS("PRODUCTIONSTART_DATE") & "</td><td> " & objRS("VIN") & "</td> </tr>")
   objRS.MoveNext
   objPkgsRS.Close
   set objPkgsRS = Nothing
  
Loop

   objRS.Close
   OBJdbConnection.Close
   set objRS = Nothing
   set OBJdbConnection = Nothing


%>


<!-- REST OF TABLE IS DUMMY 
<TR>

<TD> QUEUED </TD>
<TD> <A href='workorder-detail.asp?id=53245'>53245 </A> </TD>
<TD id='tasks1'> <A href='javascript:FlipTask1()'>0 of 2 </A></TD>
<TD id='options1'> <A href='javascript:FlipOp1()'>0 of 2 </A></TD>
<TD> TOM NEIL </TD>
<TD> 1/1/2017 </TD>
<TD> 2/1/2017 </TD>
<TD>  </TD>
<TD>  </TD>

</TR><TR>
<TD> BUILDING </TD>
<TD> <A href='workorder-detail.asp?id=123'> 123 </A> </TD>
<TD id='tasks2'> <A href='javascript:FlipTask2()'>1 of 2 </A></TD>
<TD id='options2'> <A href='javascript:FlipOp2()'>1 of 2 </A></TD>
<TD> TOM NEIL </TD>
<TD> 1/1/2017 </TD>
<TD> 2/1/2017 </TD>
<TD> 1/14/2017 </TD>
<TD>  </TD>

</TR><TR>

<TD> BUILDING </TD>
<TD> <A href='workorder-detail.asp?id=456'>456 </A> </TD>
<TD id='tasks3'> <A href='javascript:FlipTask3()'>2 of 2 </A></TD>
<TD id='options3'> <A href='javascript:FlipOp3()'>1 of 2 </A></TD>
<TD> BOB NEIL </TD>
<TD> 1/2/2017 </TD>
<TD> 2/2/2017 </TD>
<TD> 1/15/2017 </TD>
<TD> <A HREF='trucks.asp?vin=23075207598207490'> 23075207598207490 </A></TD>

</TR><TR>

<TD> COMPLETED </TD>
<TD> <A href='workorder-detail.asp?id=2342'> 2342 </A> </TD>
<TD id='tasks4'> <A href='javascript:FlipTask4()'>2 of 2 </A></TD>
<TD id='options4'> <A href='javascript:FlipOp4()'>2 of 2 </A></TD>
<TD> RICK HENDRICK </TD>
<TD> 1/2/2017 </TD>
<TD> 2/2/2017 </TD>
<TD> 1/15/2017 </TD>
<TD> <A HREF='trucks.asp?vin=23075207598207490'> 23075207598207490 </A></TD>

</TR><TR>

<TD> DELIVERED </TD>
<TD> <A href='workorder-detail.asp?id=9081'>9081 </A> </TD>
<TD id='tasks5'> <A href='javascript:FlipTask5()'>2 of 2 </A></TD>
<TD id='options5'> <A href='javascript:FlipOp5()'>2 of 2 </A></TD>
<TD> DALE EARNHARDT </TD>
<TD> 1/2/2017 </TD>
<TD> 2/2/2017 </TD>
<TD> 1/15/2017 </TD>
<TD> <A HREF='trucks.asp?vin=23075207598207490'> 23075207598207490 </A></TD>

</TR>

END DUMMY TABLE DATA -->

</table>
<BR>
<input type='button' value='NEW' id='newbutton'>
<BR>

</BODY></HTML>
