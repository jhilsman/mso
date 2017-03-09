<HTML>
<!--
'workorders.htm 2/22/17 by JRH

-->

<HEAD>
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


</table>
<BR>
<input type='button' value='NEW' id='newbutton'>
<BR>

</BODY></HTML>
