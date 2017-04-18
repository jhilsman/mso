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


dim szSQL, szParamID
dim szFromDate, szToDate
'dim szTemp, szSQL
Dim OBJdbConnection
Dim objRS
Dim iTaskCount, iOptCount

'Dim rsSwitches
'Dim iRecordCount, iProcYesCount, iSkipFlag, iEmailAvail, iEmailSent, iCalled, iPendingCount
'Dim iDeclined, iLeftVM, iNotCalled, iRecall, iRetainedCount, iNotRetainedCount, iProcNoCount


'szFromDate = Request.QueryString("FromDate")
'szToDate = Request.QueryString("ToDate")
szParamID = Request.QueryString("ID")



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

var NumOfTask = 1;
var NumOfOpt = 1;
    
        function Button1_onclick(){
            NumOfTask++;
            // get the refference of the main Div
            var mainDiv = document.getElementById('MainDiv1');
            
            // create new div that will work as a container
            var newDiv = document.createElement('div');
            newDiv.setAttribute('id','innerDiv'+NumOfTask);
            
            //create span to contain the text
            var newSpan = document.createElement('span');
            newSpan.innerHTML = "Enter Your Mail Address ";
            
            // create new textbox for email entry
            var newTextBox = document.createElement('input');
            newTextBox.type = 'text';
            newTextBox.setAttribute('id','txtAddr'+NumOfTask);
            
            // create remove button for each email adress
            var newButton = document.createElement('input');
            newButton.type = 'button';
            newButton.value = 'Remove';
            newButton.id = 'btn'+NumOfTask;
            
            // atach event for remove button click
            newButton.onclick = function RemoveEntry() { 
                var mainDiv = document.getElementById('MainDiv1');
                mainDiv.removeChild(this.parentNode);
            }
            
            // append the span, textbox and the button
            newDiv.appendChild(newSpan);
            newDiv.appendChild(newTextBox);
            newDiv.appendChild(newButton);
            
            // finally append the new div to the main div
            mainDiv.appendChild(newDiv);
    
        }

        function Button2_onclick(){
            NumOfOpt++;
            // get the refference of the main Div
            var mainDiv = document.getElementById('MainDiv2');
            
            // create new div that will work as a container
            var newDiv = document.createElement('div');
            newDiv.setAttribute('id','innerDiv'+NumOfOpt);
            
            //create span to contain the text
            var newSpan = document.createElement('span');
            newSpan.innerHTML = "Enter Your Mail Address ";
            
            // create new textbox for email entry
            var newTextBox = document.createElement('input');
            newTextBox.type = 'text';
            newTextBox.setAttribute('id','txtAddr'+NumOfOpt);
            
            // create remove button for each email adress
            var newButton = document.createElement('input');
            newButton.type = 'button';
            newButton.value = 'Remove';
            newButton.id = 'btn'+NumOfOpt;
            
            // atach event for remove button click
            newButton.onclick = function RemoveEntry() { 
                var mainDiv = document.getElementById('MainDiv2');
                mainDiv.removeChild(this.parentNode);
            }
            
            // append the span, textbox and the button
            newDiv.appendChild(newSpan);
            newDiv.appendChild(newTextBox);
            newDiv.appendChild(newButton);
            
            // finally append the new div to the main div
            mainDiv.appendChild(newDiv);
    
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

<BODY onload='onPageLoad()'>
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
   response.end
   szSQL = "select * from WO17"
else
   response.write ("ID = '" & szParamID & "'<BR>")
   szSQL = "select *, DLookUp('[NAME]','STAGES17','ID=' & [STATUS]) AS STATUS_ from WO17 where WO_NO ='" & szParamID & "'"
end if
'response.write ( szSQL & "<br>")
response.write ("<FORM name='wo-detail' id='wo-detail' action='workorder-detail-update.asp' method='post'>")
%>

<!-- *************************************************************************************************
	Main WO Table and header
     *************************************************************************************************
-->
<table width='100%' border='1' id='wolist'>
<TR>
<TD>STATUS</TD> 
<TD>WO #</TD> 
<TD>CUSTOMER </TD> 
<TD>ORDER DATE</TD> 
<TD>REQ DATE</TD> 
<TD>PROD DATE</TD>
<TD>PO#</TD> 
<TD>VIN #</TD>
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

objRS.open szSQL, OBJdbConnection

' *************************************************************************************************
'	List row, close db
' *************************************************************************************************
Do While Not objRs.EOF

response.write("<tr><td> <BR><select id='STATUSSELECTION' name='statusselection'> <option selected='SELECTED' value='ALL'>ALL</option> <option value='QUEUED'>QUEUED</option> <option value='BUILDING'>BUILDING</option> <option value='COMPLETED'>COMPLETED</option> <option value='DELIVERED'>DELIVERED</option> </select> <BR> <input type='input' size='10' id='status' name='status' value='" & objRS("STATUS") & "'><BR>" & objRS("STATUS_") & "</td><td> <input type='hidden' id='wo_no_old' name='wo_no_old' value='" & objRS("WO_NO") & "'> <input type='input' size='12' id='wo_no' name='wo_no' value='" & objRS("WO_NO") & "'></td><td> <input type='input' id='customer' name='customer' value='" & objRS("CUSTOMER") & "'> </td><td> <input type='input' size='12' id='order_date' name='order_date' value='" & objRS("ORDER_DATE") & "'> </td><td> <input type='input' size='12' id='req_date' name='req_date' value='" & objRS("REQ_DATE") & "'> </td><td> <input type='input' size='12' id='productionstart_date' name='productionstart_date' value='" & objRS("PRODUCTIONSTART_DATE") & "'> </td> <td><input type='input' size='12' id='po_no' name='po_no' value='" & objRS("PO_NO") & "'></td> <td> <input type='input' id='vin' name='vin' value='" & objRS("VIN") & "'>  </td></tr>") 
   response.write("<TR><TD>BODY ID</TD> <TD>MODEL #</TD> <TD>INV DATE</TD> <TD>INV #</TD> <TD>LENGTH</TD> <TD>BODY WEIGHT</TD> <TD>BODYYEAR</TD> <TD>BODYSTYLE</TD> </TR> ")
   response.write("<tr><td><input type='input' size='12' id='bodyid' name='bodyid' value='" & objRS("BODYID") & "'></td> <td><input type='input' size='12' id='model_no' name='model_no' value='" & objRS("MODEL_NO") & "'></td> <td><input type='input' size='12' id='invoice_date' name='invoice_date' value='" & objRS("INVOICE_DATE") & "'></td><td><input type='input' size='12' id='inv_no' name='inv_no' value='" & objRS("INV_NO") & "'></td> <td><input type='input' size='10' id='length' name='length' value='" & objRS("LENGTH") & "'></td><td><input type='input' size='12' id='body_weight' name='body_weight' value='" & objRS("BODY_WEIGHT") & "'></td> <td><input type='input' size='12' id='body_year' name='body_year' value='" & objRS("BODY_YEAR") & "'></td> <td><input type='input' id='bodystyle' name='bodystyle' value='" & objRS("BODYSTYLE") & "'></td> </tr>")

   objRS.MoveNext
Loop

' *************************************************************************************************
'	End WO item detail, show tasks and options
' *************************************************************************************************
response.write "</table>"


objRS.Close
szSQL = "select * from pkgs17 where WO_NO = '" & szParamID & "' and STAGE = '2'"
objRS.open szSQL, OBJdbConnection
'output javascript onPageLoad() function to draw array with options
%>
<script language="JavaScript">

function onPageLoad()
{
//function to add options to maindiv2

            // get the refference of the main Div
            var mainDiv = document.getElementById('MainDiv2');

// for each record in pkgs17 where wo_no = this and stage = 2

            NumOfOpt++;
            
            // create new div that will work as a container
            var newDiv = document.createElement('div');
            newDiv.setAttribute('id','innerDiv'+NumOfOpt);
            
            //create span to contain the text
            var newSpan = document.createElement('span');
            newSpan.innerHTML = "checkbox";
            
            // create new textbox for email entry
            var newTextBox = document.createElement('input');
            newTextBox.type = 'text';
            newTextBox.setAttribute('id','txtAddr'+NumOfOpt);
            
            // create remove button for each email adress
            var newButton = document.createElement('input');
            newButton.type = 'button';
            newButton.value = 'Remove';
            newButton.id = 'btn'+NumOfOpt;
            
            // atach event for remove button click
            newButton.onclick = function RemoveEntry() { 
                var mainDiv = document.getElementById('MainDiv2');
                mainDiv.removeChild(this.parentNode);
            }
            
            // append the span, textbox and the button
            newDiv.appendChild(newSpan);
            newDiv.appendChild(newTextBox);
            newDiv.appendChild(newButton);
            
            // finally append the new div to the main div
            mainDiv.appendChild(newDiv);
// end for each record


}

<%

'Enter Your Mail Address <input id="txtAddr1" type="text" />
'<input id="Button1" type="button" value="Add More" onclick="Button1_onclick()" />


response.write ("<table width='100%' border='1' id='wolist'> <TR> <TD align='center'>TASKS</TD> <TD align='center'>OPTIONS</TD> </TR>")
response.write ("<TR><TD id='MainDiv1'>")

' REUSE objRS, find all pkgs17 for tasks, loop, do same for options
objRS.Close
szSQL = "select * from pkgs17 where WO_NO = '" & szParamID & "' and STAGE = '1'"
objRS.open szSQL, OBJdbConnection
iTaskCount=0
Do While Not objRs.EOF
   iTaskCount = iTaskCount + 1
   if objRS("COMPLETED") = True then
      response.write("<input type='checkbox' checked='checked' id='task" & iTaskCount & "'>  <input type='input' id='taskname'" & iTaskCount & " name='taskname'" & iTaskCount & " value='" & objRS("NAME") & "'> <input type='button' name='RemoveTask" & iTaskCount & "' id='RemoveTask" & iTaskCount & "' value='Remove' onclick='onRemoveTask" & iTaskCount & "();'> <BR>" )
   else
      response.write("<input type='checkbox' id='task" & iTaskCount & "'>  <input type='input' id='taskname'" & iTaskCount & " name='taskname'" & iTaskCount & " value='" & objRS("NAME") & "'> <input type='button' name='RemoveTask" & iTaskCount & "' id='RemoveTask" & iTaskCount & "' value='Remove' onclick='onRemoveTask" & iTaskCount & "();'> <BR>" )
   end if

   objRS.MoveNext
Loop
'show blank line for adding
response.write("<input type='checkbox' id='taskadd1'> <input id='txtAddr' type='text' /> <input id='Button1' type='button' value='Add More' onclick='Button1_onclick()' /> ")

objRS.Close
response.write ("</TD><TD id='MainDiv2'>")
szSQL = "select * from pkgs17 where WO_NO = '" & szParamID & "' and STAGE = '2'"
objRS.open szSQL, OBJdbConnection
iOptCount = 0

Do While Not objRs.EOF
   iOptCount = iOptCount + 1
   if objRS("COMPLETED") = True then
      response.write("<input type='checkbox' checked='checked' id='option" & iOptCount & "'>  <input type='input' id='optionname'" & iOptCount & " name='optionname'" & iOptCount & " value='" & objRS("NAME") & "'> <input type='button' name='RemoveOpt" & iOptCount & "' id='RemoveOpt" & iOptCount & "' value='Remove' onclick='onRemoveOpt" & iOptCount & "();'> <BR>" )
   else
      response.write("<input type='checkbox' id='option" & iOptCount & "'>  <input type='input' id='optionname'" & iOptCount & " name='optionname'" & iOptCount & " value='" & objRS("NAME") & "'> <BR>" )
   end if

   objRS.MoveNext
Loop
'show blank line for adding
response.write("<input type='checkbox' id='optionadd1'> <input id='txtAdd2' type='text' /> <input id='Button2' type='button' value='Add More' onclick='Button2_onclick()' /> ")


' *************************************************************************************************
'	End tasks and options
' *************************************************************************************************
response.write "</td> </tr> </table>"



' *************************************************************************************************
'	End MAIN WO table and form
' *************************************************************************************************
objRS.Close
set objRS = Nothing
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
