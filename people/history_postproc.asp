<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<%
Dim conn  : Set conn=Server.CreateObject("ADODB.Connection")  : conn.Open strControl
set rs=server.createobject("ADODB.recordset")

form_selYYYY=trim(request("selYYYY"))
form_selMM=trim(request("selMM"))
form_txtFull=trim(request("txtFull"))
form_txtFullAreaAnd=trim(request("txtFullAreaAnd"))
form_txtFullAreaOr=trim(request("txtFullAreaOr"))
form_selUnit=trim(request("selUnit"))
form_selUSB=trim(request("selUSB"))
form_selAmount=trim(request("selAmount"))
form_selOrder=trim(request("selOrder"))
form_which=trim(request("which"))
SQL=""

'*******************日期搜尋****************************************
if form_selYYYY="" and form_selMM="" then
	SQLdate=""
elseif form_selYYYY<>"" and form_selMM="" then
	SQLdate="進入日期 like '" & form_selYYYY & "/__/__'"
elseif form_selYYYY="" and form_selMM<>"" then
	SQLdate="進入日期 like '____/" & form_selMM & "/__'"
else
	SQLdate="進入日期 like '" & form_selYYYY & "/" & form_selMM & "/__'"
end if
if SQLdate<>"" then SQL=SQL_format("and",SQLdate,SQL)

'*******************字串搜尋****************************************
Fulls="申請單位 申請人 負責單位 負責人 目的 攜入物品 處理過程 攜出物品 攜出原因 附註 值班人員1 值班人員2 區域 施工區域"
if form_txtFull<>"" then SQL=SQL_format("and",strSQL_Create(Fulls,DelHT(form_txtFull)), SQL)	
'*******************區域****************************************
Fulls="區域"
if form_txtFullAreaAnd<>"" then SQL=SQL_format("and",strSQL_Create(Fulls,DelHT(form_txtFullAreaAnd)), SQL)		
if form_txtFullAreaOr<>"" then SQL=SQL_format("and",strSQL_Create_Or(Fulls,DelHT(form_txtFullAreaOr)), SQL)	
'*******************負責人****************************************
if form_selUnit<>"" then SQL=SQL_format("and","負責人='" & form_selUnit & "'",SQL)
'*******************隨身碟****************************************
if form_selUSB<>"" then SQL=SQL_format("and","隨身碟='" & form_selUSB & "'",SQL)
'*******************人數****************************************
if form_selAmount<>"" then SQL=SQL_format("and","人數" & form_selAmount,SQL)

function SQL_format(byval YN,byval SQLPart,byval SQLFull) 'SQL加括號以合併
	if SQLPart<>"" then
		select case YN
		case "and"
			if SQLFull<>"" then
				SQL_format=SQLFull & " and (" & SQLPart & ")"
			else
				SQL_format="(" & SQLPart & ")"
			end if
		case "or"
			if SQLFull<>"" then
				SQL_format=SQLFull & " or " & SQLPart
			else
				SQL_format=SQLPart
			end if
		end select
	else
		SQL_format=SQLFull
	end if
end function

function strSQL_Create(byval hosts,byval strs)	'產生字串搜尋之SQL
	strA =Split(strs, " ", -1, 1)
	hostA=Split(hosts, " ", -1, 1)
	for j=0 to UBound(strA)		
		for i=0 to UBound(hostA)
			if strSQL="" then
				strSQL=hostA(i) & " like  '$" & strA(j) & "$'"
			else
				strSQL=strSQL & " or " & hostA(i) & " like  '$" & strA(j) & "$' "
			end if
		next
		strSQL_Create=SQL_format("and",strSQL,strSQL_Create)
		strSQL=""
	next
end function

function strSQL_Create_Or(byval hosts,byval strs)	'產生字串搜尋之SQL (OR)
	strA =Split(strs, " ", -1, 1)
	hostA=Split(hosts, " ", -1, 1)
	for j=0 to UBound(strA)		
		for i=0 to UBound(hostA)
			if strSQL="" then
				strSQL=hostA(i) & " like  '$" & strA(j) & "$'"
			else
				strSQL=strSQL & " or " & hostA(i) & " like  '$" & strA(j) & "$' "
			end if
		next
		strSQL_Create_Or=SQL_format("or",strSQL,strSQL_Create_Or)
		strSQL=""
	next
end function

Function DelHT(byval HT)	'去字串頭尾分號,用於DataTogether.asp,head.asp
	if trim(HT) <>"" then DelHT=replace(replace(trim(HT),";;",";"),",,",",")
	if DelHT<>"" then
		if mid(DelHT,1,1)=";" or mid(DelHT,1,1)="," then
			if DelHT=";" or DelHT=","  then 
				DelHT=""
			else
				DelHT=trim(mid(DelHT,2))
			end if
		end if
	end if
	if DelHT<>"" then		
		if mid(DelHT,len(DelHT))=";" or mid(DelHT,len(DelHT))="," then
			if DelHT=";" or DelHT="," then 
				DelHT=""
			else
				DelHT=trim(mid(DelHT,1,len(DelHT)-1))
			end if
		end if		
	end if
End Function

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
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>Post-Processing...</title>
</head>
<body onload="myFuncGo()">
<form name="form" id="form" method="post">
	<input type="hidden" name="selYYYY" id="selYYYY">
	<input type="hidden" name="selMM" id="selMM">
	<input type="hidden" name="txtFull" id="txtFull">	
	<input type="hidden" name="txtFullAreaAnd" id="txtFullAreaAnd">
	<input type="hidden" name="txtFullAreaOr" id="txtFullAreaOr">	
	<input type="hidden" name="selUnit" id="selUnit">
	<input type="hidden" name="selUSB" id="selUSB">
	<input type="hidden" name="selAmount" id="selAmount">
	<input type="hidden" name="selOrder" id="selOrder">
	<input type="hidden" name="which" id="which">
	<input type="hidden" name="SQL" id="SQL">
</form>

<script language="javascript">
var e=document.getElementById("form");

function myFuncGo(which) {
	document.getElementById("selYYYY").value="<%=form_selYYYY%>";
	document.getElementById("selMM").value="<%=form_selMM%>";	
	document.getElementById("txtFullAreaAnd").value="<%=form_txtFullAreaAnd%>";
	document.getElementById("txtFullAreaOr").value="<%=form_txtFullAreaOr%>";	
	document.getElementById("txtFull").value="<%=form_txtFull%>";
	document.getElementById("selUnit").value="<%=form_selUnit%>";
	document.getElementById("selUSB").value="<%=form_selUSB%>";
	document.getElementById("selAmount").value="<%=form_selAmount%>";
	document.getElementById("selOrder").value="<%=form_selOrder%>";
	document.getElementById("SQL").value="<%=SQL%>";	
	//alert("form_which = <%=form_which%>\n\nSQL = <%=SQL%>");	
	
	if ("<%=SQL%>" != "") {
		e.target="_self";
		e.action="<%=form_which%>";
		e.submit();
	} else {
		//alert("查詢條件不足,請輸入查詢條件 !");		
		//window.history.go(-1);
	}	
}
</script>
</body>
</html>