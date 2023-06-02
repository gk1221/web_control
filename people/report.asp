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
'	strYM=trim(YYYY) & "�~" & trim(MM) & "��" 
'	SQL="select * from people where �i�J��� like '"&trim(YYYY)&"/"&trim(MM)&"/%' order by �i�J���,�i�J�ɶ�"
'else
'	SQL="select * from people where " & SQL & " order by �i�J���,�i�J�ɶ�"
'end if   

if SQL="" then
	strYM=trim(YYYY) & "�~" & trim(MM) & "��" 	
	SQL="select * from people where �i�J��� like '"&trim(YYYY)&"/"&trim(MM)&"/%' order by �i�J��� " & SEL_ORDER & ", �i�J�ɶ� " & SEL_ORDER
else	
	SQL="select * from people where " & SQL & " order by �i�J��� " & SEL_ORDER & ", �i�J�ɶ� " & SEL_ORDER
end if           
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title><%=strYM%>���ФH���i�X����</title>
</head>
<body>
<p align="center"><b>
	<font color="#990099" size="6"><%=strYM%></font>
	<font color="#330099"><b><font size="6">���ФH���i�X����</font>
</b></p>

<p style="line-height: 150%; margin-top: 0; margin-bottom: 0">�@</p>
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
	<td width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">�ӽг��</font></td>
	<td width="13%" bgcolor="<%=var_dg%>"><p><font size="3"><%=trim(recset("�ӽг��"))%>�@</font></p></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">�ӽФH</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><p><font size="3"><%=trim(recset("�ӽФH"))%>�@</font></p></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">�t�d���</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><font size="3"><%=trim(recset("�t�d���"))%>�@</font></td>
	<td  width="12%" align="center" bgcolor="<%=var_bg%>"><font size="3">�t�d�H</font></td>
	<td width="13%" bgcolor="<%=var_dg%>" ><p><font size="3"><%=trim(recset("�t�d�H"))%>�@</font></p></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�Ӽh/����</font></td>
	<td bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("�ϰ�"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�H��</font></td>
	<td bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("�H��"))%>�@</font></td>
  </tr>  
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�I�u�ϰ�/���d</font></td>
	<td bgcolor="<%=var_dg%>" colspan="7"><font size="3"><%=trim(recset("�I�u�ϰ�"))%>�@</font></td>
  </tr>  
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�ت�</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("�ت�"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">��J���~</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("��J���~"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">��X���~</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("��X���~"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">��X��]</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("��X��]"))%>�@</font></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>">�H����</td>
	<!-- <td  bgcolor="<%=var_dg%>">����J</td> -->
	<td  bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("�H����"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�B�z�L�{</font></td>
	<td  bgcolor="<%=var_dg%>" colspan="3"><font size="3"><%=trim(recset("�B�z�L�{"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">����</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("����"))%>�@</font></td>
  </tr>
  <tr>
	<td  align="center" bgcolor="<%=var_bg%>"><font size"5">�ȯZ�H����</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("�ȯZ�H��1"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�i�J�ɶ�</font></td>
	<td bgcolor="<%=var_dg%>"><font size="2"><%=trim(recset("�i�J���")) & " " & trim(recset("�i�J�ɶ�"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">�ȯZ�H����</font></td>
	<td bgcolor="<%=var_dg%>"><font size="3"><%=trim(recset("�ȯZ�H��2"))%>�@</font></td>
	<td  align="center" bgcolor="<%=var_bg%>"><font size="3">���}�ɶ�</font></td>
	<td bgcolor="<%=var_dg%>"><font size="2"><%=trim(recset("���}���")) & " " & trim(recset("���}�ɶ�"))%>�@</font></td>
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
  		<td><p align="center"><b><font color="#FF0000" face="�з���">�L�C�L��ơ@�I�I</font></b></td>
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
		response.write "�L�C�L��ơ@�I�I"
	else 
		response.write "�|���H�����n�X�A�еn�X��A����@�I�I"
	end if
end if
%>