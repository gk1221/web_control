<!-- #include file="..\..\..\connDB\conn.ini" -->
<!-- #include file="..\..\..\Lib\GetConfig.inc" -->
<%
response.Charset="big5"

Dim conn : Set conn=Server.CreateObject("ADODB.Connection") : conn.Open strControl
set rs=server.createobject("ADODB.recordset")

Kind=request("Kind")	'onchange的選取值
which=request("which")	'onchange後選項要對應改變的物件
other=request("other")	'其他參數(根據所關聯的負責人來產生選項)

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

optHTML="<option value=""""></option>"

select case which
case "selCompany"	'申請單位
	'selHTML="<select name=""" & which & """ size=""1"" class=""option"" onchange=""Option_HTML('selStaff',this.value,'');"">"
	selHTML=""
	SQL="select distinct 申請單位 from people where 進入日期>='" & DT(now-365*3,"yyyy/mm/dd") & "'"
	select case Kind
	case "All"	'(全部)
		SQL=SQL & " order by 申請單位"
	case "Adm"	'(署單位)
		SQL="select dept from accesslist  group by dept order by COUNT(dept) asc"			
	case "Link"	'(關聯)
		SQL=SQL & " and 負責人='" & other & "' order by 申請單位"
	case "No"	'(次數)
		SQL="select 申請單位,count(*) from people where 進入日期>='" & DT(now-365*3,"yyyy/mm/dd") & "' group by 申請單位 order by count(*) desc,申請單位"
	case "A-Z"	'單位(A-Z)	
		SQL=SQL & " and (申請單位 like 'A%'"	
		for i=66 to 90
			SQL=SQL & " or 申請單位 like '" & chr(i) & "%'"
		next
		SQL=SQL & ") order by 申請單位"
	case "Else"	'(未歸類)
		SQL=SQL & " and LEFT(申請單位,1) not in (select txt from phonehead)"
		for i=65 to 90
			SQL=SQL & " and 申請單位 not like '" & chr(i) & "%'"
		next
		SQL=SQL & " order by 申請單位"
		
	case "Self"	'(自填)
		'不用coding,但不能省
	case else	'單位(ㄅ-ㄦ)
		SQL=SQL & " and ("			
		rs.open "select * from phonehead where head like '" & Kind & "%'",conn
		if not rs.eof then
			SQL=SQL & "申請單位 like '" & rs(1) & "%'"
			rs.movenext
		else
			SQL=SQL & "1=2"
		end if
		while not rs.eof
			SQL=SQL & " or 申請單位 like '" & rs(1) & "%'"
			rs.movenext
		wend
		rs.close
		SQL=SQL & ") order by 申請單位"
	end select
case "selStaff"	'申請人
	'selHTML="<select name=""" & which & """ size=""1"" class=""option"" onChange=""Staff_KeyIn(this.value);"">"
	'selHTML=""
	if other="Adm" then
		SQL = "select name from accesslist  where dept='" & Kind & "'"
	else
		SQL="select distinct 申請人 from people where 申請單位='" & Kind & "' and 進入日期>='" & DT(now-365*3,"yyyy/mm/dd") & "' order by 申請人"
	end if
end select	

if which<>"selMaintainer" then
	if SQL<>"" then
		rs.open SQL, conn
		while not rs.eof
			if Kind="No" then
				optHTML=optHTML & vbcrlf & "<option value=""" & trim(rs(0)) & """>" & trim(rs(0)) & "(" & trim(rs(1)) & ")</option>"
			else
				optHTML=optHTML & vbcrlf & "<option value=""" & trim(rs(0)) & """>" & trim(rs(0)) & "</option>"
			end if
			rs.movenext
		wend
		rs.close
	end if

	if which="selStaff" then optHTML=optHTML & vbcrlf & "<option value=""Self"">(自填)</option>"

	'response.write selHTML & optHTML & "</select>"
	response.write selHTML & optHTML
else	'selMaintainer負責人
	'response.write "<select name=""" & which & """ size=""1"" class=""option"" onchange=""selComKind.options(0).selected=1;"">"
	response.write "<option value=""""></option>"
    dim MaintainerA 
    OutFormat="<option value=""#Item#"">#Item#</option>"
    response.write GetConfig("IDMS","order by mark",Kind,"","","*",OutFormat,MaintainerA)	
    'response.write "</select>"
end if
%>