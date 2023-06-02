<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<%
Dim connDevice  : Set connDevice=Server.CreateObject("ADODB.Connection")  : connDevice.Open strControl
set rs=server.createobject("ADODB.recordset")
%>
<!-- #include file="Lib\inScript.ini" -->
<%  Execute inScript("Lib\FxLibs.txt") %>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>search</title>
</head>
<body text="" bgcolor="#DFEFFF">
<p align="center"><b><font size="5" color="#000066">設備異動進階查詢介面</font></b></p>
<form id="FormSearch">
  <p>建檔日期(YYMMDD)介於           
  	<input type="text" name="TextCreate1" size="20">和           
  	<input type="text" name="TextCreate2" size="20">之間 　 <font color="#CC6600">          
  	<font size="2">(後面日期不填表示只查單一日期) (西元年)</font> 
  </font>
  </p>
  <p>修改日期(YYMMDD)介於           
  	<input type="text" name="TextUpd1" size="20">和           
  	<input type="text" name="TextUpd2" size="20">之間 　 <font color="#CC6600">          
  	<font size="2">(後面日期不填表示只查單一日期) (西元年)</font> 
  </font>
  </p>
  <p>搜尋字串 ：<input type="text" name="TextFull" size="30">　<font size="2" color="#CC6600">(     
  對所有文字欄位比對字串是否符合，以分號隔開搜尋字串做and比對     
  )</font>
  </p>
  <p>設備異動 ：         
  	<input type="checkbox" name="CheckI" value="ON">移入　         
  	<input type="checkbox" name="CheckO" value="ON">移出　         
  	<input type="checkbox" name="CheckN" value="ON">更換　  
  	<input type="checkbox" name="CheckM" value="ON">其它
  </p>
  <p>設備種類 ：       
  　<select size="1" name="selHostClass">      
    <%  response.write "<option></option>"
		rs.open "select Item from Config where Kind='設備種類' order by Item",connDevice
		while not rs.eof
			response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"
			rs.movenext
		wend
		rs.close
	%>
    </select>
  </p>
  <p>負責人 ： 　<select size="1" name="selMaintainer">      
    <%  response.write "<option></option>"
		rs.open "select distinct Content from Config where Kind='負責單位'",connDevice
		while not rs.eof
			response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"
			rs.movenext
		wend
		rs.close
	%>
    </select>
  <font size="2" color="#CC6600">(含申請人、硬／軟體負責人)</font>
  </p>
  <p align="right">
  	<input type="button" value="　查　詢　" name="ButtonOK">　 
  	<input type="reset" value="重新設定" name="B2">
  </p>
</form>
<script language="vbscript">
Sub ButtonOK_onClick()
	if len(formSearch.TextCreate1.value)<>6 and len(formSearch.TextUpd1.value)<>6 _
		and formSearch.TextFull.value="" _
		and (not FormSearch.checkI.checked) and (not FormSearch.checkO.checked) _
		and (not FormSearch.checkN.checked) and (not FormSearch.checkM.checked) _
		and FormSearch.selHostClass.value="" and FormSearch.selMaintainer.value="" then exit sub
	'*******************日期搜尋****************************************
	if len(FormSearch.TextCreate1.value)=6 then	'建檔日期
		if len(FormSearch.TextCreate2.value)=6 then
			SQLdate="CreateDate between '20" & FormSearch.TextCreate1.value & "' and '20" & FormSearch.TextCreate2.value & "'"
		else
			SQLdate="CreateDate like '" & FormSearch.TextCreate1.value & "%'"
		end if
		SQL=SQL_format("and",SQLdate,SQL)
	end if	
	
	if len(FormSearch.TextUpd1.value)=6 then	'修改日期
		if len(FormSearch.TextUpd2.value)=6 then
			SQLdate="UpdateDate between '20" & FormSearch.TextUpd1.value & "' and '20" & FormSearch.TextUpd2.value & "'"
		else
			SQLdate="UpdateDate like '" & FormSearch.TextUpd1.value & "%'"
		end if
		SQL=SQL_format("and",SQLdate,SQL)
	end if	

	'*******************字串搜尋****************************************
	hosts="HostName;Repair;Functions;Memo;PS;StaffName;Hw;Sw;OP"
	SQL=SQL_format("and",strSQL_Create(hosts,DelHT(FormSearch.TextFull.value)), SQL)	
	'*******************設備異動****************************************
	if FormSearch.checkI.checked then SQLIO=SQL_format("or","IO='I'",SQLIO)
	if FormSearch.checkO.checked then SQLIO=SQL_format("or","IO='O'",SQLIO)
	if FormSearch.checkN.checked then SQLIO=SQL_format("or","IO='N'",SQLIO)
	if FormSearch.checkM.checked then SQLIO=SQL_format("or","IO='M'",SQLIO)
	SQL=SQL_format("and",SQLIO,SQL)
	'*******************設備種類****************************************
	if FormSearch.selHostClass.value<>"" then SQL=SQL_format("and","HostClass='" & FormSearch.selHostClass.value & "'",SQL)
	'*******************負責人****************************************
	if FormSearch.selMaintainer.value<>"" then 
	    SQLMT="(StaffName='" & FormSearch.selMaintainer.value & "' or Hw='" & FormSearch.selMaintainer.value & "' or Sw='" & FormSearch.selMaintainer.value & "')"
	end if
	SQL=SQL_format("and",SQLMT,SQL)
	'********************執行搜尋****************************************************
	if SQL<>"" then window.open "List.asp?SQL=SELECT DevID,convert(varchar, CreateDate, 111) as CreateDate,[HostName],[HostClass],[Repair],[IO],[Functions],[Memo],[PS],[StaffName],[Hw],[Sw],[OP],[Xall],[Yall],[UpdateDate] FROM Device2 where " & SQL & " order by CreateDate Desc&SQLsearch=y","_self"
End Sub

Sub TextDate1_onKeyPress
	select case window.event.keycode
	case 48,49,50,51,52,53,54,55,56,57			
	case else
		window.event.keycode=0
	end select
End Sub
Sub TextDate2_onKeyPress
	select case window.event.keycode
	case 48,49,50,51,52,53,54,55,56,57			
	case else
		window.event.keycode=0
	end select
End Sub
Sub TextFull1_onKeyPress
	select case window.event.keycode
	case 39,44,34,124 '['][,]["][|]
		window.event.keycode=0
	end select
End Sub
Sub TextFull1_onKeyPress
	select case window.event.keycode
	case 39,44,34,124 '['][,]["][|]
		window.event.keycode=0
	end select
End Sub
Sub TextFull1_onKeyPress
	select case window.event.keycode
	case 39,44,34,124 '['][,]["][|]
		window.event.keycode=0
	end select
End Sub
</script>
</body>

</html>