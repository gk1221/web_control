<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<!-- #include file="..\..\Lib\DT.inc" -->
<% 
set condb=Server.CreateObject("adodb.connection")     
condb.Open strControl
set recset=server.createobject("adodb.recordset")

YYYY=trim(request("YYYY")) : if YYYY="" then YYYY=DT(now-30,"yyyy")
MM=trim(request("MM"))  : if MM="" then MM=DT(now-30,"mm")

SEL_ORDER=trim(request("selOrder")) : if SEL_ORDER="" then SEL_ORDER="ASC"

SQL=replace(request("SQL"),"$","%")

'if SQL="" then
'	strYM=trim(YYYY) & "年" & trim(MM) & "月" 
'	SQL="select * from people where 進入日期 like '"&trim(YYYY)&"/"&trim(MM)&"/%' order by 進入日期,進入時間"
'else
'	SQL="select * from people where " & SQL & " order by 進入日期,進入時間"
'end if   

if SQL="" then
	strYM=trim(YYYY) & "年" & trim(MM) & "月" 	
	SQL="select * from people where 進入日期 like '"&trim(YYYY)&"/"&trim(MM)&"/%' order by 進入日期 " & SEL_ORDER & ", 進入時間 " & SEL_ORDER
else	
	SQL="select * from people where " & SQL & " order by 進入日期 " & SEL_ORDER & ", 進入時間 " & SEL_ORDER
end if           
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title><%=strYM%>機房人員進出紀錄</title>
</head>
<body>
<p align="center"><b>
	<font color="#990099" size="6"><%=strYM%></font>
	<font color="#330099"><b><font size="6">機房人員進出紀錄</font>
</b></p>

<p style="line-height: 150%; margin-top: 0; margin-bottom: 0">　</p>
<% 
recset.open SQL,condb 
vcnt=0
do while not recset.eof 
	vcnt=vcnt+1
	if (vcnt mod 2) = 0 then
		var_bg="#C1E0FF"
		var_bd="#003399"
		var_dg="#DFEFFF"
	else
		var_bg="#CCFFCC"
		var_bd="#006633"
		var_dg="#E6FFE6"
	end if
%>

<table border="1"  align="center" cellpadding="2" cellspacing="0" bordercolor="<%=var_bd%>" width="90%">
  <tr>
	<td width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">申請單位</font></td>
	<td width="13%" bgcolor="<%=var_dg%>"><p><font size="3"><%=trim(recset("申請單位"))%>　</font></p></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">申請人</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><p><font size="3"><%=trim(recset("申請人"))%>　</font></p></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">負責單位</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><font size="3"><%=trim(recset("負責單位"))%>　</font></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">負責人</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><p><font size="3"><%=trim(recset("負責人"))%>　</font></p></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">樓層/機房</font></td>
	<td bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("區域"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">人數</font></td>
	<td bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("人數"))%>　</font></td>
  </tr>  
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">施工區域/機櫃</font></td>
	<td bgcolor="<%=var_dg%>" colspan="7"><font size="3"><%=trim(recset("施工區域"))%>　</font></td>
  </tr>  
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">目的</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("目的"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">攜入物品</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("攜入物品"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">攜出物品</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("攜出物品"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">攜出原因</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("攜出原因"))%>　</font></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>">隨身碟</td>
	<!-- <td  bgcolor="<%=var_dg%>">未攜入</td> -->
	<td  bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("隨身碟"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">處理過程</font></td>
	<td  bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("處理過程"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">附註</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("附註"))%>　</font></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size"5">值班人員１</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("值班人員1"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">進入時間</font></td>
	<td bgcolor="<%=var_dg%>"><font size="2"><%=trim(recset("進入日期")) & " " & trim(recset("進入時間"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">值班人員２</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("值班人員2"))%>　</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">離開時間</font></td>
	<td bgcolor="<%=var_dg%>"><font size="2"><%=trim(recset("離開日期")) & " " & trim(recset("離開時間"))%>　</font></td>
  </tr>
  </table>
  <br>
<%
     recset.movenext
   loop
   if vcnt=0 then
%>   
  <div align="center">
  <table border="0" cellpadding="2" width="799">
  	<tr>
  		<td><p align="center"><b><font color="#FF0000" face="標楷體">無列印資料　！！</font></b></td>
  	</tr>
  </table>
  </div>
<% end if 
   recset.close  
%>
</font></b>
</body>
</html> 
     
<%
if vcnt=0 or no="y" then
	if vcnt=0 then 
		response.write "無列印資料　！！"
	else 
		response.write "尚有人員未登出，請登出後再執行　！！"
	end if
end if
%>