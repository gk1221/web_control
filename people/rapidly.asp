<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<!-- #include file="..\..\Lib\GetConfig.inc" -->
<% 
set conn=Server.CreateObject("adodb.connection") : conn.Open strControl
set connIDMS=Server.CreateObject("adodb.connection") : connIDMS.Open strIDMS
set connGF2000=Server.CreateObject("adodb.connection") : connGF2000.Open strGF2000
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")

authority=request("authority")
SosColor="green" : ScsColor="green" : NmsColor="green" : DcsColor="green" : ElseColor="green" : MicColor="green" : CardColor="green" : HpcColor="green" : ManyColor="green"
NpsColor="green" : ApsColor="green"
select case authority
	case "card" : CardColor="red" : memo="<font color=""green"" size=""2"">(�C�X�n�O�L���T��ƪ������d�n�J���ФH���A�Ъ`�N�ſ��<font color=""red"" size=""3"">���P�t�Ӧ��ۦP�m�W</font>�H��)</font>"
	case "sos" : kind="���v�W��-�ާ@��" : SosColor="red"
	case "scs" : kind="���v�W��-�t�ν�" : ScsColor="red"
	case "nms" : kind="���v�W��-������" : NmsColor="red"
 	case "dcs" : kind="���v�W��-��޽�" : DcsColor="red" 
    case "nps" : kind="���v�W��-�ƭȽ�" : NpsColor="red"
 	case "aps" : kind="���v�W��-�n���" : ApsColor="red"
    case "mic" : kind="���v�W��-��T����" : MicColor="red"
	case else  : ElseColor="red" : memo="<font color=""green"" size=""2"">(�C�X��@�~�̱`�i�X���ФH��)</font>"
end select

select case authority
case "sos","scs","nms","dcs","mic","nps","aps"
	memo=memo & "<br><font color=""green"" size=""2"">"
	rs.open "select * from Config where kind='���v�覡'",conn
	while not rs.eof
		memo=memo & rs(1) & "." & rs(2) & "�@"
		rs.movenext
	wend
	rs.close
	memo=memo & "</font>"
end select

if authority="sos" or authority="scs" or authority="nms" or authority="dcs" or authority="nps" or authority="aps" then
	set fs=server.createobject("scripting.filesystemobject")
	set F=fs.GetFile("d:\���Ц�F\���T���\���v�H���W��.xlsx")
	rs.open "select item from Config where kind='���v�ɶ�' order by item desc",conn
	if DT(F.DateLastModified,"yyyy/mm/dd hh:mi")>rs(0) then
		memo=memo & "<br><font color=""red"" size=""5"">(���v�H���W��.xlsx�w��s,�Ч�s���v�H���α��v�ɶ����t�ΰѼ�)</font>"
	end if
	rs.close
end if

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
%> 

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>�ֳt�n�J</title>
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
<body  bgproperties="fixed">
<div align="center"><span class="title">�i�X�H���ֳt�n�J</span></div>
<form name="form" id="form" method="post" >
<table width="90%" align="center"><tr>
	<td><font color="<%=ElseColor%>"><u style="cursor:pointer" onClick="window.open('rapidly.asp','_self');">�@�~�̱`�i�X</u></font></td>
	<!--<td><font color="<%=CardColor%>"><u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=card','_self');">�����d</u></font></td> vm16�|�X�{just in time�����A�����|�X�{SQL���~�A�G����-->
	<td><font color="<%=SosColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=sos','_self');">�ާ@��</u></font></td>
	<td><font color="<%=ScsColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=scs','_self');">�t�ν�</u></font></td>
	<td><font color="<%=NmsColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=nms','_self');">������</u></font></td>	
	<td><font color="<%=DcsColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=dcs','_self');">��޽�</u></font></td>
    <td><font color="<%=NpsColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=nps','_self');">�ƭȽ�</u></font></td>
    <td><font color="<%=ApsColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=aps','_self');">�n���</u></font></td>
	<td><font color="<%=MicColor%>"> <u style="cursor:pointer" onClick="window.open('rapidly.asp?authority=mic','_self');">��T����</u></font></td>	
    <!--
	<td><font color="<%=HpcColor%>"> <u style="cursor:pointer" onClick="window.open('Many.asp?Company=�I�h�q','_self');">�I�h�q</u></font></td>
	<td><font color="<%=ManyColor%>"> <u style="cursor:pointer" onClick="window.open('Many.asp?Company=�_��','_self');">�_��</u></font></td>
    -->
</tr></table>
<hr>
<!-- <table width="90%" border="1" align="center" cellpadding="5" cellspacing="0"> -->
<table width="90%" border="1" align="center" cellpadding="5" cellspacing="1" bordercolor="#B0C4DE">
<%	select case authority
	case "sos","scs","nms","dcs","nps","aps"
		rs.open "select * from Config where kind='" & kind & "'",conn
		i=0 : SQL=""
		while not rs.eof
			pos1=instr(1,rs(1),"-") : pos2=instr(pos1+1,rs(1),"-") : unit=mid(rs(1),1,pos1-1) : name=mid(rs(1),pos1+1,pos2-pos1-1) : AuthorityNo=mid(rs(1),pos2+1)
			if i=0 then
				SQL="(�ӽг��='" & unit & "' and �ӽФH='" & name & "'"
			else
				SQL=SQL & "or �ӽг��='" & unit & "' and �ӽФH='" & name & "'"
			end if
			i=i+1
			strNames=strNames & AuthorityNo & "." & name & vbcrlf
			rs.movenext
		wend
		rs.close
		if i>0 then SQL=SQL & ")"
    case "mic"
        rs.open "select * from Config where kind='��T����'",connIDMS
		i=0 : SQL=""
		while not rs.eof
			unit=rs(1) : AuthorityNo="A"
            rs1.open "select * from Config where kind='" & unit & "'",connIDMS
            while not rs1.eof
                name=rs1(1)
                if i=0 then
				    SQL="(�ӽг��='" & unit & "' and �ӽФH='" & name & "'"
			    else
				    SQL=SQL & "or �ӽг��='" & unit & "' and �ӽФH='" & name & "'"
			    end if
                rs1.movenext
                i=i+1
			    strNames=strNames & AuthorityNo & "." & name & vbcrlf
            wend
            rs1.close
			rs.movenext
		wend
		rs.close
		if i>0 then SQL=SQL & ")"
	case "card"
		i=0 : SQL=""
		rs.open "select distinct UserName from LogDB,UserDB where LogDB.UserID=UserDB.UserID and LogDB.LogTime >='" & DT(now,"yyyy/mm/dd 00:00") & "'",connGF2000
		while not rs.eof
        	if i=0 then
				SQL="(�ӽФH='" & rs(0) & "'"
			else
				SQL=SQL & "or �ӽФH='" & rs(0) & "'"
			end if
			strNames=strNames & "?." & rs(0) & vbcrlf
        	rs.movenext
        	i=i+1			
        wend
		rs.close
		if i>0 then SQL=SQL & ")"
	case else
		rs.open "select �ӽФH,count(*) from people where �i�J���>='" & DT(now-365,"yyyy/mm/dd") & "' group by �ӽФH order by count(*) desc",conn
		i=0 : SQL="�i�J���>='" & DT(now-365,"yyyy/mm/dd") & "'"
		while not rs.eof and i<28		
			if i=0 then
				SQL=SQL & " and (�ӽФH='" & rs(0) & "'"		
			else
				SQL=SQL & " or �ӽФH='" & rs(0) & "'"
			end if
			i=i+1
			rs.movenext
		wend
		rs.close
		if i>0 then SQL=SQL & ")"
	end select
	
	'�ư��W�椣�C�J�ֳt�n�J
	rs.open "select * from Config where kind='�ư��W��'",conn
	while not rs.eof
		Exceptions=Exceptions & rs(1) & "�G" & rs(2) & "�G"
		rs.movenext
	wend
	rs.close
	
	'response.write SQL
	'response.end
	if SQL<>"" then

	rs.open "select * from people where " & SQL & " order by �ӽг��,�ӽФH,Serial desc",conn
	i=1 : r=0
	while not rs.eof
		if tmp<>trim(rs("�ӽФH")) then
			r=i mod 4 
			if r=1 then response.write "<tr>"
			tmp=trim(rs("�ӽФH")) : Names=trim(rs("�ӽФH")) : pos1=instr(1,strNames,Names)
			if pos1>0 then
				pos2=instr(pos1,strNames,vbcrlf)
				Names=mid(strNames,pos1-2,pos2-pos1+2)
			end if			
			if instr(1,Exceptions,trim(rs("�ӽг��")) & "�G" & Names)=0 then
%>				<td align="center" class="item"><%=trim(rs("�ӽг��"))%>�@</td>
				<td align="center"><input type="button" class="butn" value="<%=Names%>" onClick="set_value ('<%=rs("Serial")%>');"></td>
<%				if r=0 then response.write "</tr>"
				i=i+1
			end if
		end if
		rs.movenext
	wend
	rs.close
	
	if r<>0 then
		response.write "<td colspan=""" & 2*(4-(i-1) mod 4) & """>�@</td>"
		response.write "</tr>"
	end if

	end if
%>  
</table>
</form>
<%="<p align=""center"">" & memo & "<br><br><font color=""green"" size=""2"">(�u��ܴ��n�J�L���H��,�H�K�פJ�̪�@�����,����������M��Ш����v�M�U )</font></p>"%>
</body>
</html>
<script language="javascript">
    function set_value(Serial) {
        var e = document.getElementById("form");
        e.action = "people.asp?IO=I&Serial=" + Serial;
        e.submit();
    }
</script>