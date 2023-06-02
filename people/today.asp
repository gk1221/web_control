<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<%
set condb=Server.CreateObject("adodb.connection")     
condb.open strControl 
set rs=server.createobject("adodb.recordset")

Serial=trim(request("Serial")) : DateIn=request("DateIn") : nowDT=DT(now,"yyyy/mm/dd")

if request("diff")="" then
	diff=0
else
	diff=cint(request("diff"))
end if

if DateIn="" then DateIn=nowDT
DateIn=DT(DateSerial(mid(DateIn,1,4),mid(DateIn,6,2),mid(DateIn,9,2)) + diff,"yyyy/mm/dd")	

if request("selYYYY")<>"" and request("selMM")<>"" and request("selDD")<>"" then
	DateIn = DT2(cint(request("selYYYY")), cint(request("selMM")), cint(request("selDD")), "yyyy/mm/dd")
end if 

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
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>進出記錄</title>
	<!--<link href="people.css" rel="stylesheet" type="text/css" />-->
	<link href="Lib/people.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form id="form" name="form" method="post">
<% 
orderBy=request("orderBy")
'response.write len(orderBy)
if orderBy<>""	then 
	rs.open "select * from people where 進入日期='" & DateIn & "' or 離開日期='' order by " & orderBy & " desc, 進入日期 desc,進入時間 desc",condb 
else 
	rs.open "select * from people where 進入日期='" & DateIn & "' or 離開日期='' order by 進入日期 desc,進入時間 desc",condb 
end if 
i=0
%>  
<table border="0" align="center">
  <tr>
    <td align="center">
    	<span class="head">
    	<% if DateIn=nowDT then
    		response.write "今日人員進出機房紀錄"
    	   else
    	   	response.write DateIn & "(" & week & ") 人員進出機房紀錄"
    	   end if 
    	%>
    	</span>
	</td>
  </tr>
</table>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<td><div align="center">
      <input type="button" class="butn" value="<< 前10日" onClick="Form_Submit(-10);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="後3日 >>" onClick="Form_Submit(3);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="<  前一日" onClick="Form_Submit(-1);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td ><div bgcolor="red" align="center">
      <input type="button" class="butn1" value="今       日" onClick="Form_Submit(0);"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="後一日  >" onClick="Form_Submit(1);"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>	
  </tr>
</table>
<br>
<table align="center" border="1" cellpadding="5" cellspacing="1" bordercolor="#B0C4DE"> <!-- #B0C4DE: light steel blue -->
  <tr>
    <td width="120" ><div align="center" class="item"><a href="today.asp?orderBy=申請單位" target="body">廠商</a></div></td>
    <td width="100" ><div align="center" class="item"><a href="today.asp?orderBy=申請人" target="body">人員</a></div></td>
    <td width="100"><div align="center" class="item"><a href="today.asp?orderBy=" target="body">進入時間</a></div></td>
    <td width="100"><div align="center" class="item">離開時間</div></td>
	<!--<td><div align="center" class="item"><a href="today.asp?orderBy=樓層" target="body">樓層/機房</a></div></td>-->
	<td><div align="center" class="item"><a href="today.asp?orderBy=區域" target="body">樓層/機房</a></div></td>
    <td><div align="center" class="item">執行</div></td>
    <td><div align="center" class="item">借用</div></td>		
	<td width="120"><div align="center" class="item">施工區域/機櫃</div></td>	
  </tr>
<% 
while not rs.eof
   i=i+1 %>	
  <tr>
	<td class="word" align="center">&nbsp<% =rs("申請單位") 	%>&nbsp</td>
    <td class="word" align="center">&nbsp<u style="cursor:pointer" onClick="window.open('people.asp?IO=I&Serial=<%=rs("Serial")%>','_self');"><% =rs("申請人")%></u>&nbsp</td>
    <td class="word" align="center">&nbsp
    <%	if DateIn=rs("進入日期") then
    		response.write rs("進入時間")
    	else
    		response.write rs("進入日期")    		
    	end if
    %>&nbsp</td>
    <td class="word" align="center">&nbsp
	 <%	if rs("離開時間")="" then
			response.write "　"
		else
			response.write rs("離開時間")
		end if
	%>&nbsp</td>
	<!--<td>-->
	<% '=rs("樓層") 
	%>
	<!--</td>-->																					
	<td>&nbsp<% =rs("區域") %>&nbsp</td>
    <td>&nbsp
    <%	if trim(rs("離開日期")) ="" then %>			
    		<!-- <a href="people.asp?IO=O&Serial=
				<% '=rs("Serial")
				%>
			" target="body">登出</a> -->
			<b><font face="新細明體" color="#990099" onClick="Logout('<%=rs("Serial")%>','<%=trim(rs("申請人"))%>','<%=trim(rs("施工區域"))%>');" style="cursor:pointer"><u>登出</u></font></b>			
	<%	else %>
			<font color="gray" onClick="Login_Cancel('<%=rs("Serial")%>','<%=trim(rs("申請人"))%>');" style="cursor:pointer"><u>取消</u></font>
	<%	end if %>&nbsp;&nbsp;&nbsp;
    	<a href="people.asp?Serial=<%=trim(rs("Serial"))%>" target="body">更新</a>&nbsp;&nbsp;&nbsp;
    	<font color="blue" onClick="Login_Del('<%=rs("Serial")%>','<%=trim(rs("申請人"))%>');" style="cursor:pointer"><u>刪除</u></font>
	&nbsp</td>
    <td align="center"><font size="2">&nbsp
    <%	if instr(1,rs("攜入物品"),"臨時卡")>0 then response.write "臨時卡&nbsp;"
    	if instr(1,rs("攜入物品"),"鑰匙")>0 then response.write "鑰匙&nbsp;"
    	if instr(1,rs("攜入物品"),"遙控器")>0 then response.write "遙控器"
    %>
    </font>&nbsp</td>		
	<td>&nbsp<% =rs("施工區域") %>&nbsp</td>	
  </tr>
<% 
   rs.movenext  
wend
rs.close
%>
</table>

<% if i=0 then response.write "<p align='center' class='msg'>該日無人員進入！</p>" %>
<input type="hidden" name="DateIn" value="<%=DateIn%>">
<input type="hidden" name="diff">
</form>

<%	
	'rs.open "SELECT 施工區域 FROM people WHERE 進入日期='" & DateIn & "' OR 離開日期='" & DateIn & "'",condb 		
	'rs.open "SELECT 施工區域 FROM people WHERE 進入日期='" & DateIn & "' OR 離開日期='' OR 離開日期='" & DateIn & "'",condb 		
	rs.open "SELECT 施工區域 FROM people WHERE 進入日期<='" & DateIn & "' AND (離開日期='' OR 離開日期>='" & DateIn & "')",condb 		
	str_workareas = " "
	while not rs.eof
		if rs("施工區域") <> "" then
			str_workarea_i = rs("施工區域")			
			str_workarea_i = Replace(str_workarea_i, ",", " ")
			str_workarea_i = Replace(str_workarea_i, ";", " ")
			str_workarea_i = Replace(str_workarea_i, "；", " ")
			str_workarea_i = Replace(str_workarea_i, "，", " ")
			str_workarea_i = Replace(str_workarea_i, "、", " ")
			str_workareas_i = Split(str_workarea_i," ")			
			for j = 0 to ubound(str_workareas_i)
				if InStr(str_workareas, str_workareas_i(j)) = 0 then
					str_workareas = str_workareas & str_workareas_i(j) & " "
				end if					
			next			
		end if
		rs.movenext  
	wend		
	rs.close
%>
<table align="center" border="1" cellpadding="5" cellspacing="1" bordercolor="#B0C4DE"> <!-- #B0C4DE: light steel blue -->
  <tr>
    <td width="190"><div align="center" class="item"><%=DateIn%> 施工區域/機櫃</div></td>
    <td width="670"><p id="workareas"></p></td>
  </tr>
</table>
<br>
<table border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<td><div align="center">
      <input type="button" class="butn" value="<< 前10日" onClick="Form_Submit(-10);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="後3日 >>" onClick="Form_Submit(3);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="<  前一日" onClick="Form_Submit(-1);" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn1" value="今       日" onClick="Form_Submit(0);"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="後一日  >" onClick="Form_Submit(1);"/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </div></td>
    <td><div align="center">
      <input type="button" class="butn" value="夜間加班名冊" onClick="alert('此功能僅作為協助工具之用');window.open('night.asp','_blank');"/>
    </div></td>
  </tr>
</table>
<br /><br />
</body>
</html>
<script language="javascript">
var e=document.getElementById("form")
var str_workareas = '<%=str_workareas%>';
var workareas = str_workareas.trim().split(' ');
workareas.sort(function(a, b) {return a.localeCompare(b, "zh-Hant-TW");});
document.getElementById("workareas").innerHTML = workareas.toString().replaceAll(',', '、');

function Form_Submit(diff) {
	e.elements["diff"].value=diff ;
	if(diff==0) e.elements["DateIn"].value="" ;
	e.submit() ;
}

function Login_Cancel(Serial,p) {
	if(confirm(p + "確定要取消登出?")) window.open("today.asp?DBaction=c&Serial=" + Serial,"_self") ;
}

function Login_Del(Serial,p) {
	if(confirm("確定要刪除"+ p +"的登入記錄?")) window.open("today.asp?DBaction=d&Serial=" + Serial,"_self") ;
}

function Logout(Serial,p,WorkArea) {
	if(WorkArea == "" || confirm("請落實追蹤 "+ p +" 於 施工區域/機櫃："+ WorkArea + " 之環境復原情形。")) window.open("people.asp?IO=O&Serial=" + Serial,"_self") ;
}
</script>