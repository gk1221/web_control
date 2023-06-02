<%if Burn="" then%>
<%if Sel="" then%>
  <p align="center"><b><font size="5" color="#000066">設備異動每月報表查詢介面</font></b></p>
<% end if %>
<b><font color="#FF0000" size="4">月份：</font></b>
<input type="text" name="textShowMonth" size="6" style="font-size: 12pt; font-weight: bold; color: #FF0000">
<input type="button" value="查　尋" name="ButtonPos">　
<input type="button" value="<<上一月 " name="ButtonLast"> 
<input type="button" value="本　月" name="ButtonNow"> 
<input type="button" value="下一月>>" name="ButtonNext">　
	<%if Sel="" then%>
		<input type='button' value='月報表製作' name='cmdBurn'>　
		<font size="2" color="CC6600">(滑鼠點標題可排序,移至OP可提示全名)</font>
	<% end if %>
<script language="vbscript">
	YM="<%=request("YM")%>" : MM=month(now) : if MM<10 then MM="0" & MM
	if YM="" then YM=year(now) & "/" & MM
	textShowMonth.value=YM
	Sel="<%=request("Sel")%>"

	Sub ButtonLast_onClick
		CreateDate=DateAdd("m",-1,cdate(textShowMonth.value))
		YY=year(CreateDate) : MM=month(CreateDate) : if MM<10 then MM="0" & MM
		YM=YY & "/" & MM
		'SQL="Select * from Device where CreateDate like '" & mid(YY,3) & MM & "$' order by CreateDate Desc"
		SQL="SELECT Devid,convert(varchar, CreateDate, 111) as CreateDate,[HostName],[HostClass],[Repair],[IO],[Functions],[Memo],[PS],[StaffName],[Hw],[Sw],[OP],[Xall],[Yall],[UpdateDate] FROM Device2 WHERE YEAR([CreateDate]) = '" & YY & "' AND MONTH([CreateDate])= '"& MM & "' "
		window.open "List.asp?" & "YM=" & YM ,"_self"
	End Sub

	Sub ButtonNow_onClick
		YY=year(now) : MM=month(now) : if MM<10 then MM="0" & MM	
		YM=YY & "/" & MM	
		'SQL="Select * from Device where CreateDate like '" & mid(YY,3) & MM & "$' order by CreateDate Desc"
		SQL="SELECT Devid,convert(varchar, CreateDate, 111) as CreateDate,[HostName],[HostClass],[Repair],[IO],[Functions],[Memo],[PS],[StaffName],[Hw],[Sw],[OP],[Xall],[Yall],[UpdateDate] FROM Device2 WHERE YEAR([CreateDate]) = '" & YY & "' AND MONTH([CreateDate])= '"& MM & "' "
		window.open "List.asp?" & "YM=" & YM ,"_self"
	End Sub

	Sub ButtonNext_onClick
		CreateDate=DateAdd("m",1,cdate(textShowMonth.value))
		YY=year(CreateDate) : MM=month(CreateDate) : if MM<10 then MM="0" & MM	
		YM=YY & "/" & MM	
		'SQL="Select * from Device where CreateDate like '" & mid(YY,3) & MM & "$' order by CreateDate Desc"
		SQL="SELECT Devid,convert(varchar, CreateDate, 111) as CreateDate,[HostName],[HostClass],[Repair],[IO],[Functions],[Memo],[PS],[StaffName],[Hw],[Sw],[OP],[Xall],[Yall],[UpdateDate] FROM Device2 WHERE YEAR([CreateDate]) = '" & YY & "' AND MONTH([CreateDate])= '"& MM & "' "
		window.open "List.asp?" & "YM=" & YM ,"_self"
	End Sub

	Sub ButtonPos_onClick
		if len(textShowMonth.value)<>7 or mid(textShowMonth.value,5,1)<>"/" or cint(mid(textShowMonth.value,6))>12 or cint(mid(textShowMonth.value,6))=0 then
			msgbox "輸入資料日期[" & textShowMonth.value & "]格式有誤!",48,"資料日期錯誤"
			exit sub
		end if
		YY=mid(textShowMonth.value,1,4) :MM=mid(textShowMonth.value,6,2) : YM=YY & "/" & MM
		SQL="Select * from Device where CreateDate like '" & mid(YY,3) & MM & "$' order by CreateDate Desc"
		window.open "List.asp?Sel=" & Sel & "&YM=" & YM & "&SQL=" & SQL,"_self"
	End Sub

	Sub textShowMonth_onKeyPress
		select case window.event.keycode
		case 13
			ButtonPos_onClick
		case 47,48,49,50,51,52,53,54,55,56,57			
		case else
			window.event.keycode=0
		end select
	End Sub

	Sub cmdBurn_onClick()
		window.open "List.asp?SQLsearch=" & "<%=SQLsearch%>"  _
			& "&YM=" & textShowMonth.value & "&Burn=Y","_blank"
	End Sub
</script>
<% end if %>
