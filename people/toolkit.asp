<!-- #include file="..\..\connDB\conn.ini" -->

<%
set condb=Server.CreateObject("adodb.connection")     
condb.open strControl 
set rs=server.createobject("adodb.recordset")

Serial=trim(request("Serial")) : DateIn=request("DateIn") : nowDT=DT(now,"yyyy/mm/dd")

if DateIn="" then DateIn=nowDT
DateIn=DT(DateSerial(mid(DateIn,1,4),mid(DateIn,6,2),mid(DateIn,9,2)),"yyyy/mm/dd")	

dateInSerial = DateSerial(mid(DateIn,1,4),mid(DateIn,6,2),mid(DateIn,9,2))
DateInYear = year(dateInSerial) 
DateInMonth = month(dateInSerial) 
DateInDay = day(dateInSerial)

week=WeekdayName(DatePart("w",DateSerial(mid(DateIn,1,4),mid(DateIn,6,2),mid(DateIn,9,2))))

select case request("DBaction")
case "c"	'取消登入
	condb.execute "update people set 值班人員2='',離開日期='',離開時間='' where Serial=" & Serial
case "d"	'刪除記錄
	condb.execute "delete from people where Serial=" & Serial
end select 

Function DT(byval sDT,byval fDT)
	dim YY0,MM0,DD0,HH0,MI0,SS0
	DT=lcase(fDT)
	YY0=year(sDT)   : if instr(1,DT,"yy",1)>0 and instr(1,DT,"yyyy",1)=0 then YY0=mid(YY0,3)
	DT=replace(replace(DT,"yyyy","yy"),"yy",YY0)
	MM0=month(sDT)  : if MM0<10 then MM0="0" & MM0
	DT=replace(DT,"mm",MM0)
	DD0=day(sDT)    : if DD0<10 then DD0="0" & DD0
	DT=replace(DT,"dd",DD0)
	HH0=hour(sDT)   : if HH0<10 then HH0="0" & HH0
	DT=replace(DT,"hh",HH0)
	MI0=minute(sDT) : if MI0<10 then MI0="0" & MI0
	DT=replace(DT,"mi",MI0)
	SS0=second(sDT) : if SS0<10 then SS0="0" & SS0
	DT=replace(DT,"ss",SS0) 
End Function

Function DT2(byval YY0, byval MM0, byval DD0, byval fDT)
	DT2=lcase(fDT)
	if instr(1,DT2,"yy",1)>0 and instr(1,DT2,"yyyy",1)=0 then YY0=mid(YY0,3)
	DT2=replace(replace(DT2,"yyyy","yy"),"yy",YY0)
	if MM0<10 then MM0="0" & MM0
	DT2=replace(DT2,"mm",MM0)
	if DD0<10 then DD0="0" & DD0
	DT2=replace(DT2,"dd",DD0)
End Function
%>

<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>進出記錄</title>
	<link href="Lib/people.css" rel="stylesheet" type="text/css" />
</head>
<body style="background:#6699ff;" onkeydown="Cmd_KeyDown(event.keyCode);">
<form id="form" name="form" method="post">

<!-- <table width="600" border="0" align="center" cellpadding="1" cellspacing="1"> -->
<table width="" border="0" align="center" cellpadding="1" cellspacing="1">
  <tr style="font-family:標楷體">
	<td>指定日期：</td>
    <td><div align="center">
      <select size="1" name="selYYYY" id="selYYYY">      
      <% for i=cint(DT(now,"yyyy"))-9 to cint(DT(now,"yyyy"))
           if i = DateInYear then
             response.write "<option value='" & i & "' selected>" & i & "</option>"
           else
             response.write "<option value='" & i & "'>" & i & "</option>"
           end if
	     next
      %>
      </select>年
    </div></td>
    <td><div align="center">
      <select size="1" name="selMM" id="selMM">
      <% for i=1 to 12
	       if i = DateInMonth then
		     if i>9 then
			   response.write "<option value='" & i & "' selected>" & i & "</option>"
	         else
	           response.write "<option value='0" & i & "' selected>0" & i & "</option>"
	         end if	  
	       else
	         if i>9 then
			   response.write "<option value='" & i & "'>" & i & "</option>"
	         else
	           response.write "<option value='0" & i & "'>0" & i & "</option>"
	         end if	  
	       end if  
	     next
      %>
      </select>月
    </div></td>
    <td><div align="center">
      <select size="1" name="selDD" id="selDD">
      <% for i=1 to 31
		   if i = DateInDay then
		     if i>9 then
	           response.write "<option value='" & i & "' selected>" & i & "</option>"
	         else
	           response.write "<option value='0" & i & "' selected>0" & i & "</option>"
	         end if
		   else
		     if i>9 then
	           response.write "<option value='" & i & "'>" & i & "</option>"
	         else
	           response.write "<option value='0" & i & "'>0" & i & "</option>"
	         end if
		   end if		   
	     next
      %>
      </select>日
    </div></td>
   <td><div align="center">
      <input type="button" class="butn" value="前往" onClick="Form_Submit_GoToDate();"/>
	  <input type="button" class="butn1" value="今日" onClick="Form_Submit_SetDateToday();"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
	
	<td>人員查詢：</td>
	<td><div align="center">		
	<input type="text" name="txtFull" id="txtFull" size="15">
	<input type="button" class="butn" value=" < " onClick="Search(-1)">	
	<input type="button" class="butn" value=" > " onClick="Search(1)">&nbsp;&nbsp;
	</div></td>
	
	<td>最近一筆：</td>
	<td><div align="center">
	<input type="button" class="butn" value="跳轉" onClick="Search(0)">
	<input type="button" class="butn" value="代入" onClick="AddApplicant();">
	</div></td>
	
  </tr>
</table>

<input type="hidden" name="DateIn" value="<%=DateIn%>">

</form>
</body>
</html>

<script language="javascript">
<!-- #include file="..\Lib\Ajax.inc" -->

var e=document.getElementById("form")

function Form_Submit_GoToDate() {
	e.action="today.asp";
	e.target="body";
	e.submit() ;
}

function Form_Submit_SetDateToday() {
	var today = new Date();
	
	document.getElementById("selYYYY").value = today.getFullYear();
	
	if ((today.getMonth() + 1) < 10) {
		document.getElementById("selMM").value = "0" + (today.getMonth()+1).toString();
	} else {
		document.getElementById("selMM").value = (today.getMonth()+1).toString();
	}
	
	if (today.getDate() < 10) {
		document.getElementById("selDD").value = "0" + today.getDate();
	} else {
		document.getElementById("selDD").value = today.getDate();
	}	
	
	Form_Submit_GoToDate();
}

function Search(direction) {
	http_request.open("POST",
		"Lib/SearchApplicant.asp?applicant=" + document.getElementById("txtFull").value
			+ "&search=" + direction
			+ "&selYYYY=" + document.getElementById("selYYYY").value
			+ "&selMM=" + document.getElementById("selMM").value
			+ "&selDD=" + document.getElementById("selDD").value
		, false);	
	http_request.send("");	
	if(http_request.responseText != "") {
		document.getElementById("selYYYY").value = http_request.responseText.substring(0,4);
		document.getElementById("selMM").value = http_request.responseText.substring(5,7);
		document.getElementById("selDD").value = http_request.responseText.substring(8,10);
		
	}
			
	e.action="today.asp";
	e.target="body";
	e.submit() ;
}

function AddApplicant(){	
	http_request.open("POST","Lib/GetApplicantSerial.asp?applicant=" + document.getElementById("txtFull").value, false);	
	http_request.send("");
	if(http_request.responseText != "") {
		window.open("people.asp?IO=I&Serial=" + http_request.responseText,"body");
	}
}

function Cmd_KeyDown(KeyCode){    //新增刪除按鍵
            if (KeyCode == 13) AddApplicant();	//Ins
}  
</script>