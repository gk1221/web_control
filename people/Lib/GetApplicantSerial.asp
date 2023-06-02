<!-- #include file="..\..\..\connDB\conn.ini" -->

<%
response.Charset="big5"

set condb=Server.CreateObject("adodb.connection")     
condb.open strControl 
set rs=server.createobject("adodb.recordset")

if trim(request("applicant"))<>"" then
	rs.open "select * from people where 申請人='" & trim(request("applicant")) & "' order by 進入日期 desc, 進入時間 desc",condb 
	if not rs.eof then
		response.write rs("Serial")		
	end if 
	rs.close
end if
%>
