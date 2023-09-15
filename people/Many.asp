<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<!-- #include file="..\..\Lib\GetConfig.inc" -->
<% 
set conn=Server.CreateObject("adodb.connection") : conn.Open strControl
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")

Company=request("Company") : if Company="" then Company="富士通"
Names=request("Names") : if Names<>"" then Names=mid(Names,1,len(Names)-1)
OP=request("selOP")

Serial=1
rs.open "select max(Serial) from people",conn	
if not rs.eof then Serial=rs(0)+1
rs.close

Unit="主任、技正" : Maintainer="林弘倉" : StuffIn="" : if Company="富士通" then StuffIn="筆記型電腦"
USB="未攜入" : Process="HPC系統安裝建置" : Cause="" : Memo="" : Amount=1

DateIn=DT(now,"yyyy/mm/dd")	: TimeIn=DT(now,"hh:mi:ss")		        

if OP<>"" then
    dim NameA : NameA=split(Names,",")
    for i=0 to UBound(NameA)
    	rs.open "select * from people where 申請單位='" & Company & "' and 申請人='" & NameA(i) & "' and (進入日期='" & DT(now,"yyyy/mm/dd") & "' or 離開日期='')",conn,3,3
    	if rs.eof then
            Purpose="系統安裝建置, " & GetABC(Company,NameA(i)) & "類人員, 依規定登入"
	        conn.execute "insert into people values(" & Serial & ",'" & Company & "','" & NameA(i) & "','" & Unit & "','" & Maintainer & "','" & Purpose & "','" _
			    & StuffIn & "','" & USB & "','" & Process & "','" & StuffIn & "','" & Cause & "','" & Memo & "'," & Amount & ",'" _
			    & OP & "','','" & DateIn & "','" & TimeIn & "','','')" 
		else
			if trim(rs("離開日期"))="" then
				rs("值班人員2")=OP : rs("離開日期")=DT(now,"yyyy/mm/dd") : rs("離開時間")=DT(now,"hh:mi:ss")
			else
				rs("值班人員2")="" : rs("離開日期")="" : rs("離開時間")=""				
			end if	
			rs.update		
		end if
		rs.close
        Serial=Serial+1
    next
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

Function GetABC(byval Company,byval Staff)
    set rsABC=server.createobject("adodb.recordset")
    rsABC.open "select Item from Config where Kind like '授權名單%' and Item like '%" & Company & "-" & Staff & "%'",conn
	if not rsABC.eof then 
		GetABC=mid(rsABC(0),len(rsABC(0)))
	else
		GetABC="C"
	end if
	rsABC.close
End Function
%> 

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>大量登入</title>
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
<div align="center"><span class="title">進出廠商大量登入 - <%=Company%></span></div>
<form name="form" id="form" method="post">
<table width="90%" border="1" align="center" cellpadding="5" cellspacing="0">
<%	rs.open "select distinct 申請人 from people where 申請單位='" & Company & "' and (進入日期='" & DT(now,"yyyy/mm/dd") & "' or 離開日期='')",conn	'TodayNames:今日登入或尚未登出者
    while not rs.eof
        TodayNames=TodayNames & rs(0) & ","
        rs.movenext
    wend
    rs.close
    
    rs.open "select distinct 申請人 from people where 申請單位='" & Company & "' and 進入日期 between '" & DT(now-7,"yyyy/mm/dd") & "' and '" & DT(now-1,"yyyy/mm/dd") & "'",conn	'WeekNames:近一週登入者
    while not rs.eof
        WeekNames=WeekNames & rs(0) & ","
        rs.movenext
    wend
    rs.close
    
    'rs.open "select distinct 申請人,substring(申請人,1,1) as FirstName from people where 申請單位='" & Company & "' order by 申請人",conn	'全部登入者
    rs.open "select * from Config where Kind='授權名單-HPC專案(" & Company & ")' order by Content",conn	'全部登入者
	i=1 : r=0
	while not rs.eof
		r=i mod 8
		pos1=instr(1,rs(1),"-") : pos2=instr(pos1+1,rs(1),"-") : Pname=mid(rs(1),pos1+1,pos2-pos1-1)
		if r=1 then response.write "<tr>"
        if instr(1,TodayNames,Pname & ",")>0 then	'TodayNames:今日登入或尚未登出者
        	rs1.open "select * from people where 申請單位='" & Company & "' and 申請人='" & Pname & "' and 離開日期=''",conn
        	if not rs1.eof then	'尚未登出者橘底藍字
	        	response.write "<td width=""12.5%"" style=""background:orange""><input type=""checkbox"" name=""chk" & i & """ value=""" & Pname & """ />&nbsp;<font color=""blue"">" & Pname & "</font></td>"
	        else	'已登出者綠底藍字
	        	response.write "<td width=""12.5%"" style=""background:green""><input type=""checkbox"" name=""chk" & i & """ value=""" & Pname & """ />&nbsp;<font color=""blue"">" & Pname & "</font></td>"
	        end if
	        rs1.close
        elseif instr(1,WeekNames,Pname & ",")>0 then	'近一週登入者藍字
        	response.write "<td width=""12.5%""><input type=""checkbox"" name=""chk" & i & """ value=""" & Pname & """ />&nbsp;<font color=""blue"">" & Pname & "</font></td>"
        else
            response.write "<td width=""12.5%""><input type=""checkbox"" name=""chk" & i & """ value=""" & Pname & """ />&nbsp;" & Pname & "</td>"
        end if
		if r=0 then response.write "</tr>"
		i=i+1
	    rs.movenext
	wend
	rs.close
	
	if r<>0 then
		response.write "<td colspan=""" & (8-(i-1) mod 8) & """>　</td>"
		response.write "</tr>"
	end if
%>  
</table>

<p align="center">
    值班人員：
    <select name="selOP" size="1">
        <option value=""></option>
	    <%	dim opA : OutFormat="<option value=""#Item#"" #Sel#>#Item#</option>"
            response.write GetConfig("IDMS","order by mark","作業管理科","Item",OP,"*",OutFormat,opA)
	    %>  
    </select>　　&nbsp;
    <input type="button" value="大量登入 / 大量登出 / 取消登出" class="butn" onClick="set_login();">　&nbsp;
    <input type="button" value="重新整理" class="butn" onClick="window.open('many.asp','_self');">
</p>
<font color="#008000">
<input type="hidden" name="Company" value="<%=Company%>" />
<input type="hidden" name="Names" />
</font>
</form>
</body>
</html>
<script>
    var e = document.getElementById("form")
    function Eobj(str) { return document.getElementById(str); }
    function Fobj(str) { return e.elements[str]; }

    if ("<%=OP%>") alert("存檔完成！");

    function set_login() {
        if (Fobj("selOP").value) {
            for (i = 0; i < document.all.length; i++) {
                if (document.all(i).tagName == "INPUT")
                    if (document.all(i).type == "checkbox")
                        if (document.all(i).checked & document.all(i).value != "") Fobj("Names").value = Fobj("Names").value + document.all(i).value + ",";
            }
            e.target = "_self";
            e.action = "Many.asp";
            e.submit();
        }
        else
            alert("請選擇值班人員！");
    }
</script>
<p><font color="#008000" size="2">
1.門禁管制--&gt;人員進出--&gt;快速登入--&gt;富士通及寶訊的選項中，已將有簽署保密切結人員清單列在上面(參考資管課-凱琦的文件),其授權等級(均屬B類人員)由系統設定值讀取; 
若在該選單上無此人員，均屬C類人員，請當班同仁注意.<br><br>
2.若有進場名單上有的人員,但在門禁管制--&gt;人員進出--&gt;快速登入--&gt;富士通的選項中無選項,當班同仁可：<br>
　a.列印<a target="_blank" href="file://10.6.1.4/d/SSM_WEB/control/people/Help/20150522_161336_CWB-MIC-ISMS-ISR-D002-第三方人員保密切結書_20150521_v2.1.doc">保密切結書</a>請其簽署,且記錄日誌.<br>
　b.請負責人帶入,登入C類人員.作法可2擇1<br><br>
-----by 資安小組</font></p>