<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID2.inc" -->
<%
Dim conn  : Set conn=Server.CreateObject("ADODB.Connection")  : conn.Open strControl
set rs=server.createobject("ADODB.recordset")

Function DT(byval sDT,byval fDT)
  DT=lcase(fDT)
  YY=year(sDT)   : if instr(1,DT,"yy",1)>0 and instr(1,DT,"yyyy",1)=0 then YY=mid(YY,3)
  DT=replace(replace(DT,"yyyy","yy"),"yy",YY)
  MM=month(sDT)  : if MM<10 then MM="0" & MM
  DT=replace(DT,"mm",MM)
  DD=day(sDT)    : if DD<10 then DD="0" & DD
  DT=replace(DT,"dd",DD)
  HH=hour(sDT)   : if HH<10 then HH="0" & HH
  DT=replace(DT,"hh",HH)
  MI=minute(sDT) : if MI<10 then MI="0" & MI
  DT=replace(DT,"mi",MI)
  SS=second(sDT) : if SS<10 then SS="0" & SS
  DT=replace(DT,"ss",SS) 
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>歷史查詢</title>
    <style type="text/css">
	<!--
	body {
	background-repeat: repeat;
	margin: 0px;
	line-height: 100%;
	background-color: #99CC66
		}
	-->
	</style>
	<link href="Lib/people.css" rel="stylesheet" type="text/css">

</head>
<body>
<p align="center" class="title">機房進出人員歷史紀錄查詢</p>
<form name="form" id="form" method="post">
  <p>進入日期 ： 
  	<select size="1" name="selYYYY">      
    <%  response.write "<option></option>"
		for i=cint(DT(now,"yyyy"))-4 to cint(DT(now,"yyyy"))
			response.write "<option value='" & i & "'>" & i & "</option>"
		next
	%>
    </select>年　
  	<select size="1" name="selMM">
    <%  response.write "<option></option>"
		for i=1 to 12
			if i>9 then
				response.write "<option value='" & i & "'>" & i & "</option>"
			else
				response.write "<option value='0" & i & "'>0" & i & "</option>"
			end if
		next
	%>
    </select>月
  </p>
  <p>搜尋字串 ：<input type="text" name="txtFull" size="15">　
  	<font size="2" color="#CC6600">(對所有文字欄位比對字串是否符合，空白隔開搜尋字串可做and比對，填
  	&quot;<font color="red"><u onclick="txtFull.value='其它登入'">其它登入</u></font>&quot; 
    即可列出未依規定登入紀錄)</font>
  </p> 
  
  <p>樓層/機房(AND) ：<input type="text" name="txtFullAreaAnd" size="15">　
  	<font size="2" color="#CC6600">(空白隔開搜尋字串可做 <u>AND</u> 比對)</font>
  </p>  
  <p>樓層/機房(OR) ：<input type="text" name="txtFullAreaOr" size="15">　
  	<font size="2" color="#CC6600">(空白隔開搜尋字串可做 <u>OR</u> 比對)</font>
  </p>   
  
  <p>負責人 ：       
  　<select size="1" name="selUnit">      
    <%  response.write "<option></option>"
		rs.open "select distinct 負責人 from people",conn
		while not rs.eof
			if trim(rs(0))<>"" then response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"
			rs.movenext
		wend
		rs.close
	%>
    </select>
  </p> 
  <p>隨身碟 ：       
  　<select size="1" name="selUSB">      
    <%  response.write "<option></option>"
		rs.open "select Item from Config where Kind='隨身碟' order by Content",conn
		while not rs.eof
			response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"
			rs.movenext
		wend
		rs.close
	%>
    </select>
  </p>
  <p>人數 ：       
  　<select size="1" name="selAmount">      
  		<option></option>
  		<option value="=1">等於1人</option>
  		<option value="=2">等於2人</option>
  		<option value=">1">大於1人</option>
  		<option value=">2">大於2人</option>
    </select>
  </p>
  <p>查詢結果 ： 
  　<select size="1" name="selOrder">      
  		<option value="ASC">舊 -> 新</option>
  		<option value="DESC">新 -> 舊</option>
    </select> 
	<!--
  　<select size="1" name="selNewPage">      
  		<option value="_blank">開新分頁</option>
  		<option value="_self">留在此頁</option>
    </select> 
	-->
  </p>   
  <p>
	<!--
  	<input type="button" value="　紀　錄　查　詢　" name="cmdLog"  onClick="HistoryOK('report.asp');">　
  	<input type="button" value="　時　數　統　計　" name="cmdTime" onClick="HistoryOK('time.asp');">&nbsp;  	
	<input type="hidden" name="SQL">
	-->
	<input type="button" value="　紀　錄　查　詢　" name="cmdLog"  onClick="HistoryPostProc('report.asp');">　
  	<input type="button" value="　時　數　統　計　" name="cmdTime" onClick="HistoryPostProc('time.asp');">&nbsp;  	
	<input type="hidden" name="which">	
  </p> 
</form>

<script language="javascript">
var e=document.getElementById("form");

function HistoryPostProc(which) {
	if (e.elements['selYYYY'].value == "" && e.elements['selMM'].value == "" && e.elements['txtFull'].value == "" && e.elements['selUnit'].value == "" && e.elements['selUSB'].value == "" && e.elements['selAmount'].value == "") {
		alert("查詢條件不足,請輸入查詢條件 !");	
	} else {
		e.elements['which'].value = which;
		e.target="_blank";
		//e.target=e.elements['selNewPage'].value;
		e.action="history_postproc.asp";
		e.submit();
	}
}
</script>
</body>
</html>