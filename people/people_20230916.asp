<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID2.inc" -->
<!-- #include file="..\..\Lib\GetConfig.inc" -->
<% 
set conn=Server.CreateObject("adodb.connection") : conn.Open strControl
set connIDMS=Server.CreateObject("adodb.connection") : connIDMS.Open strIDMS
set rs=server.createobject("adodb.recordset")

dim UnitA : dim UserA   'GetConfig(qryDB,strOrderBy,Kind,WhichIn,Xin,OutNo,OutFormat,ByRef ConfigA)
'---------------------------------------------------------------------------------------------------------------------------------------------
'DBaction		預備			預備			預備			預備				實際			實際			實際			
'IO				登入			登出			修改			帶入				登入			登出			修改			
'---------------------------------------------------------------------------------------------------------------------------------------------
'新增																				V												
'修改																				V				V
'---------------------------------------------------------------------------------------------------------------------------------------------
'讀submit																			V				V				V
'讀db							V				V				V																
'---------------------------------------------------------------------------------------------------------------------------------------------
'標題:登入		V												V																
'標題:登出						V
'標題:修改										V
'讀值-----------------------------------------------------------------------------------------------------------------------------------------
Serial=request("Serial") : DBaction=request("DBaction") : IO=request("IO") 	'I:登入流程畫面 O:登出流程畫面 空白:異動流程畫面(預設)
if Serial<>"" and DBaction="" then	'讀db : 有序號(不是預備登入,但可能是預備帶入)且不異動,即是顯示某筆資料
	rs.open "select * from people where Serial=" & Serial,conn
	if not rs.eof then
		if IO="I" then	'預備帶入
			Serial=""	'清空序號,以便實際登入時,讀到的是submit資料,而不是db的資料		
			Company=trim(rs(1)) : Staff=trim(rs(2)) : Unit=trim(rs(3)) : Maintainer=trim(rs(4)) : Purpose=trim(rs(5)) : StuffIn=""
			USB="未攜入" : Process=trim(rs(8)) : StuffOut="" : Cause="" : Memo="" : Amount=1
			OPin="" : OPout="" : DateIn="" : TimeIn="" : DateOut="" : TimeOut=""	
			FloorArea=trim(rs(19)) : FloorArea2=trim(rs(20))			
			WorkArea=""
			
            if GetConfig("IDMS","",Unit,"Item",Maintainer,"","#Item#",UserA)="" then Unit="" '預防負責人調課
            
            Purpose=replace(replace(Purpose,"依規定登入",""),"宣導資安政策門禁規定","")	'******************************************************************************************其它登入
            OtherPurpose="預備帶入"
            pos=instr(1,rs(5),"類人員") : if pos>0 then Purpose=Purpose & "," & mid(rs(5),pos-1,4)	'******************************************其它登入
		else	'讀DB資料
			Company=trim(rs(1)) : Staff=trim(rs(2)) : Unit=trim(rs(3)) : Maintainer=trim(rs(4)) : Purpose=trim(rs(5)) : StuffIn=trim(rs(6))
			USB=trim(rs(7)) : Process=trim(rs(8)) : StuffOut=trim(rs(9)) : Cause=trim(rs(10)) : Memo=trim(rs(11)) : Amount=trim(rs(12))
			OPin=trim(rs(13)) : OPout=trim(rs(14)) : DateIn=trim(rs(15))	: TimeIn=trim(rs(16)) : DateOut=trim(rs(17)) : TimeOut=trim(rs(18))	
			FloorArea=trim(rs(19)) : FloorArea2=trim(rs(20))			
			WorkArea=trim(rs(21))
			
		end if
	end if
	rs.close
else	'讀取submit過來的資料
	Company=request("selCompany") : ComKind=request("selComKind") : if ComKind="Self" then Company=trim(request("txtCompany"))
	Staff=request("selStaff") : if Staff="Self" then Staff=trim(request("txtStaff"))
	Unit=request("selUnit") : Maintainer=request("selMaintainer")
		
	if request("chkPurpose")="" then
		Purpose=trim(request("txtPurpose"))
	else
		if trim(request("txtPurpose"))="" then
			Purpose=replace(request("chkPurpose")," ,","",1,-1,1)
		else			
			Purpose=replace(request("chkPurpose")," ,","",1,-1,1) & "," & trim(request("txtPurpose"))
		end if
	end if
	
	if request("chkStuffIn")="" then
		StuffIn=trim(request("txtStuffIn"))
	else
		if trim(request("txtStuffIn"))="" then
			StuffIn=replace(request("chkStuffIn")," ,","",1,-1,1)
		else			
			StuffIn=replace(request("chkStuffIn")," ,","",1,-1,1) & "," & trim(request("txtStuffIn"))
		end if
	end if
	
	if request("chkStuffOut")="" then
		StuffOut=trim(request("txtStuffOut"))
	else
		if trim(request("txtStuffOut"))="" then
			StuffOut=replace(request("chkStuffOut")," ,","",1,-1,1)
		else			
			StuffOut=replace(request("chkStuffOut")," ,","",1,-1,1) & "," & trim(request("txtStuffOut"))
		end if
	end if

	FloorArea2=replace(request("chkFloorArea2")," ,","",1,-1,1)
	
	USB=request("selUSB") : Process=request("txtProcess") : Cause=request("txtCause") : Memo=request("txtMemo") : Amount=request("selAmount")
	OPin=request("selOPin") : OPout=request("selOPout")
	DateIn=request("txtDateIn") : TimeIn=request("txtTimeIn") : DateOut=request("txtDateOut") : TimeOut=request("txtTimeOut")	
	'FloorArea=request("selFloorArea")	
	FloorArea=""	
	WorkArea=request("txtWorkArea")
	
end if
'預設值(讀完資料後再給預設值)-----------------------------------------------------------------------------------------------------------------
if Amount=""  then Amount="0"
if DateIn=""  then DateIn=DT(now,"yyyy/mm/dd")
if TimeIn=""  then TimeIn=DT(now,"hh:mi:ss")
if DateOut="" and IO="O" then DateOut=DT(now,"yyyy/mm/dd")
if TimeOut="" and IO="O" then TimeOut=DT(now,"hh:mi:ss")
'異動-----------------------------------------------------------------------------------------------------------------------------------------
if DBaction="a" or DBaction="u" then 
		ABC="C"
		
		rs.open "select * from config where Kind='數值資訊組' and Item='" & Company & "'",connIDMS
		if not rs.eof then ABC="A"
		rs.close
		
		rs.open "select Item from Config where Kind like '授權名單%' and Item like '%" & Company & "-" & Staff & "%'",conn
		if not rs.eof then 
			ABC=mid(rs(0),len(rs(0)))			
		end if
		rs.close
				
        AbcMsg=""
		if ABC="A" and (instr(1,Purpose,"A類人員")=0 or instr(1,Purpose,"B類人員")<>0 or instr(1,Purpose,"C類人員")<>0) _ 
			or ABC="B" and (instr(1,Purpose,"B類人員")=0 or instr(1,Purpose,"A類人員")<>0 or instr(1,Purpose,"C類人員")<>0) _ 
			or ABC="C" and (instr(1,Purpose,"C類人員")=0 or instr(1,Purpose,"A類人員")<>0 or instr(1,Purpose,"B類人員")<>0) then
			AbcMsg=Company & "-" & Staff & "-" & ABC & "，請檢查您的授權方式是否設定錯誤!"	
		end if

		if AbcMsg="" then
				select case DBaction
				case "a"	'新增
			    	Serial=1
				    rs.open "select max(Serial) from people",conn	
			    	if not rs.eof then Serial=rs(0)+1
				    rs.close

					conn.execute "insert into people values(" & Serial & ",'" & Company & "','" & Staff & "','" & Unit & "','" & Maintainer & "','" & Purpose & "','" _
						& StuffIn & "','" & USB & "','" & Process & "','','" & Cause & "','" & Memo & "'," & Amount & ",'" _
						& OPin & "','','" & DateIn & "','" & TimeIn & "','','','" & FloorArea & "','" & FloorArea2 & "','" & WorkArea & "')" 
						
					'注意StuffOut='' OPout='' DateOut='' TimeOut=''	
				case "u"	'修改
					rs.open "select * from people where Serial=" & Serial,conn,3,3
						rs(1)=Company : rs(2)=Staff : rs(3)=Unit : rs(4)=Maintainer : rs(5)=Purpose : rs(6)=StuffIn : rs(7)=USB : rs(8)=Process : rs(9)=StuffOut
						rs(10)=Cause : rs(11)=Memo : rs(12)=Amount : rs(13)=OPin : rs(14)=OPout : rs(15)=DateIn : rs(16)=TimeIn : rs(17)=DateOut : rs(18)=TimeOut						
						'rs(19)=FloorArea : 
						rs(20)=FloorArea2								
						rs(21)=WorkArea						
					rs.update
					rs.close	
				end select
        else
            response.write "<script language=""javascript"">" & vbcrlf
			response.write "	alert(""" & AbcMsg & """);" & vbcrlf	
			response.write "</script>" & vbcrlf
		end if		
end if
'注音分類-----------------------------------------------------------------------------------------------------------------------------------------
if DBaction<>"" then
	phonetxt=mid(Company,1,1)
	rs.open "select * from phonehead where txt='" & phonetxt & "'",conn	'查注音分類表有沒有
	if rs.eof then
		rs.close
		rs.open "select * from phonetab where txt like '%" & phonetxt & "%'",conn	'從注音字典找出注音
		while not rs.eof
			on error resume next	'避免重複而出錯
			conn.execute "insert into phonehead values('" & mid(rs(0),1,1) & "','" & phonetxt & "')"	'新增至注音分類表
			on error goto 0
			rs.movenext
		wend
	end if
	rs.close
end if
'異動設備提醒-----------------------------------------------------------------------------------------------------------------------------------------
if DBaction<>"" and AbcMsg="" then 
'	if instr(1,Purpose,"異動設備")>0 then
'			response.write "<script language=""vbscript"">" & vbcrlf
'			response.write "Sub window_onload()" & vbcrlf
'			response.write "	msgbox ""1. 請記得向負責人詢問設備編號及移入單，並記錄之！"" & vbcrlf & ""2. 設備進出機房或移動位置，請更改IDMS放置地點，詳見本畫面底【注意事項】。"",48" & vbcrlf			
'			response.write "    window.open ""today.asp?date=" & DateIn & """,""_self""" & vbcrlf
'			response.write "End Sub" & vbcrlf
'			response.write "</script>" & vbcrlf			
'	else
		response.redirect "today.asp?date=" & DateIn '異動後,畫面轉至今日紀錄異動畫面'
'	end if
end if
'通知等級-----------------------------------------------------------------------------------------------------------------------------------------

'標題-----------------------------------------------------------------------------------------------------------------------------------------
select case IO	'因DBaction之故,不可移到前面
case "I"
	cmdSend="登入" : DBaction="a"
case "O"
	cmdSend="登出" : DBaction="u"
case ""
	cmdSend="更新" : DBaction="u"
end select
'函數-----------------------------------------------------------------------------------------------------------------------------------------
Function DT(byval sDT,byval fDT) 'if..then..後不可再加 : ,不可多行敘述
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

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<title>人員進出登記</title>
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
<body>
<table width="100%"><tr><td align="center"><span class="title">人員進出紀錄<%=cmdSend%></span></td></tr></table>
<table border="1" align="center" cellpadding="1" cellspacing="1" bordercolor="#B0C4DE">
<form name="form" id="form" method="post">
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr>
		<td align="left" valign="middle"><div align="center" class="item">負責單位</div></td>
		<td align="left" valign="middle">
			<select name="selUnit" id="selUnit" size="1" class="option" language="javascript" onchange="Option_HTML('selMaintainer',this.value,'');">
				<option></option>
            <%  OutFormat="<option value=""#Item#"" #Sel#>#Item#</option>"
                response.write GetConfig("IDMS","order by mark","數值資訊組","Item",Unit,"*",OutFormat,UnitA) 
            %>
			</select>
		</td>
		<td align="left" valign="middle"><div align="center" class="item">負責人</div></td>
		<td align="left" valign="middle">
			<select name="selMaintainer" id="selMaintainer" size="1" class="option" onchange="selComKind.options(0).selected=1;">
                <option></option>
			<%	if Unit="" then
					response.write "<option></option>"
				else
                    dim MaintainerA : OutFormat="<option value=""#Item#"" #Sel#>#Item#</option>"
                    response.write GetConfig("IDMS","order by mark",Unit,"Item",Maintainer,"*",OutFormat,MaintainerA) 
                end if
			%>
			</select>
		</td>
		<td align="left" valign="middle"><div align="center" class="item">值班<span class="style1">(入)</span> </div></td>
		<td align="left" valign="middle">
			<select name="selOPin" size="1" class="option">
				<option value=""></option>
			<%	dim opA : OutFormat="<option value=""#Item#"" #Sel#>#Item#</option>"
                response.write GetConfig("IDMS","order by mark","作業管理科","Item",OPin,"*",OutFormat,opA)
			%>  
			</select>
		</td>
		<td><div align="center" class="item">進入時間</div></td>
		<td>
			<input name="txtDateIn" type="text"  value="<%=DateIn%>" size="8">
			<input name="txtTimeIn" type="text"  value="<%=TimeIn%>" size="6">
		</td>
	</tr>
	<tr>
		<td align="left" valign="middle"><div align="center" class="item">申請單位</div></td>
		<td align="left" valign="middle">
			<select name="selComKind" id="selComKind" size="1" class="option" onChange="Option_HTML('selCompany',this.value,selMaintainer.value);">
				<option value="All">(全部)</option>
				<option value="Self">(自填)</option>				
				<option value="Link">(關聯)</option>
				<option value="No">(次數)</option>			
				<option value="Else">(未歸類)</option>
				<option value="A-Z">單位(A-Z)</option>
				<option value="ㄅ">單位(ㄅ)</option>
				<option value="ㄆ">單位(ㄆ)</option>
				<option value="ㄇ">單位(ㄇ)</option>
				<option value="ㄈ">單位(ㄈ)</option>
				<option value="ㄉ">單位(ㄉ)</option>
				<option value="ㄊ">單位(ㄊ)</option>
				<option value="ㄋ">單位(ㄋ)</option>
				<option value="ㄌ">單位(ㄌ)</option>
				<option value="ㄍ">單位(ㄍ)</option>
				<option value="ㄎ">單位(ㄎ)</option>
				<option value="ㄏ">單位(ㄏ)</option>
				<option value="ㄐ">單位(ㄐ)</option>
				<option value="ㄑ">單位(ㄑ)</option>
				<option value="ㄒ">單位(ㄒ)</option>
				<option value="ㄓ">單位(ㄓ)</option>
				<option value="ㄔ">單位(ㄔ)</option>
				<option value="ㄕ">單位(ㄕ)</option>
				<option value="ㄖ">單位(ㄖ)</option>
				<option value="ㄗ">單位(ㄗ)</option>
				<option value="ㄘ">單位(ㄘ)</option>
				<option value="ㄙ">單位(ㄙ)</option>
				<option value="ㄧ">單位(ㄧ)</option>
				<option value="ㄨ">單位(ㄨ)</option>
				<option value="ㄩ">單位(ㄩ)</option>
				<option value="ㄚ">單位(ㄚ)</option>
				<option value="ㄛ">單位(ㄛ)</option>
				<option value="ㄜ">單位(ㄜ)</option>
				<option value="ㄝ">單位(ㄝ)</option>
				<option value="ㄞ">單位(ㄞ)</option>
				<option value="ㄟ">單位(ㄟ)</option>
				<option value="ㄠ">單位(ㄠ)</option>
				<option value="ㄡ">單位(ㄡ)</option>
				<option value="ㄢ">單位(ㄢ)</option>
				<option value="ㄣ">單位(ㄣ)</option>
				<option value="ㄤ">單位(ㄤ)</option>
				<option value="ㄥ">單位(ㄥ)</option>
				<option value="ㄦ">單位(ㄦ)</option>				
			</select>
			<select name="selCompany" id="selCompany" size="1" class="option" onchange="Option_HTML('selStaff',this.value,'');">                
				<option value=""></option>
			<%	rs.open "select distinct 申請單位 from people where 進入日期>='" & DT(now-365*3,"yyyy/mm/dd") & "' order by 申請單位",conn
				selected=""			
				while not rs.eof
					if trim(rs(0))=Company then
						selected="Y"
						response.write "<option value=""" & trim(rs(0)) & """ selected>" & trim(rs(0)) & "</option>" & vbcrlf
					else
						response.write "<option value=""" & trim(rs(0)) & """>" & trim(rs(0)) & "</option>" & vbcrlf
					end if					
					rs.movenext
				wend
				rs.close
				if Company<>"" and selected="" then response.write "<option value=""" & Company & """ selected>" & Company & "</option>" & vbcrlf
			%>
			</select>
			<input name="txtCompany" type="text" class="write" size="6" style="display:none" value="<%=Company%>">                
		</td>
		<td align="left" valign="middle"><div align="center" class="item">申請人</div></td>
		<td align="left" valign="middle">
			<select name="selStaff" size="1" class="option" onChange="Staff_KeyIn(this.value);">
			<%	rs.open "select distinct 申請人 from people where 申請單位='" & Company & "' and 進入日期>='" & DT(now-365*3,"yyyy/mm/dd") & "' order by 申請人",conn				                
				selected=""
				while not rs.eof
					if trim(rs(0))=Staff then
						selected="Y"
						response.write "<option value=""" & trim(rs(0)) & """ selected>" & trim(rs(0)) & "</option>" & vbcrlf
					else
						response.write "<option value=""" & trim(rs(0)) & """>" & trim(rs(0)) & "</option>" & vbcrlf
					end if					
					rs.movenext
				wend
				rs.close
				if Staff<>"" and selected="" then response.write "<option value=""" & Staff & """ selected>" & Staff & "</option>" & vbcrlf
			%>	<option value="Self">(自填)</option>
			</select>
			<input name="txtStaff" type="text" size="6" style="display:none" value="<%=Staff%>">                
		</td>
		<td height="28" class="item"><div align="center">人數</div></td>
		<td>
			<select name="selAmount" size="1" class="option">
			<%	for i=1 to 30
					if cstr(i)=Amount then
						response.write "<option value=""" & i & """ selected>" & i & "</option>"
					else
						response.write "<option value=""" & i & """>" & i & "</option>"
					end if
				next
			%>
			</select>
		</td>
		<td class="item"><div align="center"><font color="blue">
			<u style="cursor:pointer" onClick="alert('若外單位參觀機房,申請人填局內對口人員,負責人填中心對口人員或課長,而附註則填外單位領隊姓名 !');">附註說明</u>
		</font></div></td>
		<td><input name="txtMemo" type="text"  size="16" value="<%=Memo%>" /></td>
	</tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr>
		<td align="left" valign="middle"><div align="center" class="item">樓層/機房</div></td> 
		<td colspan="7" align="left" valign="middle">
			<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			
			<%	if FloorArea2<>"" then
					FloorArea2A=split(FloorArea2,",")
				end if
				
				rs.open "select * from Config where kind='樓層/機房' order by Content",conn 
				while not rs.eof
					if FloorArea2Content="" then 
						response.write "<tr>"					
					elseif FloorArea2Content<>mid(rs("Content"),1,1) then
						response.write "</tr>"
						response.write "<tr>"
					end if
					
					FloorArea2Content=mid(rs("Content"),1,1)
					checked=""
					
					if FloorArea2<>"" then
						for i=0 to UBound(FloorArea2A)
							if trim(FloorArea2A(i))=rs("item") then
								checked="checked"
								FloorArea2A(i)=""
							end if
						next
					end if
					tmp=rs("item")
					
			%>		<td><span class="option"><input type="checkbox" name="chkFloorArea2" value="<%=rs("item")%>" <%=checked%>><%=tmp%>
					</span></td>
			<%		rs.movenext
				wend
				rs.close
			%>
				</tr>
			</table>
		</td>
	</tr>	
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
    <tr>
      <td><div align="center" class="item">施工區域/機櫃</div></td>
      <td colspan="7"><textarea name="txtWorkArea" id="txtWorkArea" cols="97" rows="1" class="intxt" value="<%=WorkArea%>"><%=WorkArea%></textarea></td>      
    </tr>	
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr style="display:<% if IO="O" then response.write "none" %>">
		<td align="left" valign="middle"><div align="center" class="item">目的及登入狀況</div></td> 
		<td colspan="7" align="left" valign="middle">
			<table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<%	PurposeA=split(Purpose,",")
				rs.open "select * from Config where kind='目的' order by Content",conn 
				while not rs.eof
					if PurposeContent="" then 
						response.write "<tr>"					
					elseif PurposeContent<>mid(rs("Content"),1,1) then
						response.write "</tr>"
						response.write "<tr>"
					end if
					
					PurposeContent=mid(rs("Content"),1,1)
					checked=""
					for i=0 to UBound(PurposeA)
						if trim(PurposeA(i))=rs("item") then
							checked="checked"
							PurposeA(i)=""
						end if
					next
					tmp=rs("item")
					if len(rs("Content"))>4 then	tmp=tmp & mid(rs("Content"),5)
					
					if instr(1,rs("Item"),"類人員")>0 then 	'*******************************************************************************其它登入
						select case mid(rs("Item"),1,1)
						case "A" : response.write "<td><span class=""option"" title=""不必通知，不需陪同"">"
						case "B" : response.write "<td><span class=""option"" title=""需要通知，不需陪同"">"
						case "C" : response.write "<td><span class=""option"" title=""需要通知，需要陪同"">"
						end select
						response.write "<hr>"
					else
						response.write "<td><span class=""option"">"
					end if									'*******************************************************************************其它登入	
			%>			<input type="checkbox" name="chkPurpose" value="<%=rs("item")%>" <%=checked%>><%=tmp%>
					</span></td>
			<%		rs.movenext
				wend
				rs.close
				
				if OtherPurpose="預備帶入" then	'*******************************************************************其它登入
					otherPurpose=""
				else
					for i=0 to UBound(PurposeA)
						if PurposeA(i)<>"" then otherPurpose=otherPurpose & trim(PurposeA(i)) & ","
					next
					if otherPurpose<>"" then otherPurpose=mid(otherPurpose,1,len(otherPurpose)-1)
				end if
			%>		<td><span class="option">
						<font color="red">登入說明</font>：<input name="txtPurpose" type="text" id="txtPurpose" value="<%=otherPurpose%>" size="30">
					</span></td>
				</tr>
			</table>
		</td>
	</tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr style="display:<% if IO="O" then response.write "none" %>">
		<td align="left" valign="middle"><div align="center" class="item">攜入物品</div></td>
		<td colspan="7" align="left" valign="middle">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr><td><span class="option">
			<%	StuffInA=split(StuffIn,",")
				rs.open "select * from Config where kind='物品' order by Content",conn 
				while not rs.eof
					if StuffContent="" then
						response.write mid(rs("Content"),5)	& "："
					elseif StuffContent<>mid(rs("Content"),5) then
						response.write "</span></td></tr>"
						response.write "<tr><td><span class=""option"">"
						if mid(rs("Content"),5)="行動裝置" then
							response.write mid(rs("Content"),5) & "<font color=""green"">(請先取得負責人同意)</font>："					
						else
							response.write mid(rs("Content"),5) & "："
						end if
					end if
					StuffContent=mid(rs("Content"),5)
			%>		<input type="checkbox" name="chkStuffIn" value="<%=rs("item")%>"
			<%		for i=0 to UBound(StuffInA)
						if trim(StuffInA(i))=rs("item") then
							response.write " checked"
							StuffInA(i)=""
						end if
					next
			%>		>
			<%		if rs("item")="臨時卡" then
						response.write rs("item") & " <font color=""green"">(請於處理過程註明 => 臨時卡號:### )" _
							& "　<u style=""cursor:pointer"" onclick=""txtProcess.value=txtProcess.value + ' 臨時卡號:'"";>導入</u></font>" & "&nbsp;&nbsp;&nbsp;&nbsp;"
                    else	
						response.write rs("item") & "&nbsp;&nbsp;&nbsp;&nbsp;"
					end if
					rs.movenext
				wend
				rs.close
				for i=0 to UBound(StuffInA)
					if StuffInA(i)<>"" then otherStuffIn=otherStuffIn & trim(StuffInA(i)) & ","
				next
				if otherStuffIn<>"" then otherStuffIn=mid(otherStuffIn,1,len(otherStuffIn)-1)
			%>	</span></td></tr>
				<tr><td>
					<span class="option">其它：&nbsp;&nbsp;<input name="txtStuffIn" type="text" value="<%=otherStuffIn%>" size="20"></span>&nbsp;&nbsp;
					<font color="red">隨身碟：</font>
					<select name="selUSB" size="1" class="option">
						<option></option>
					<%	rs.open "select Item from Config where Kind='隨身碟' order by Content",conn
						selected=""
						while not rs.eof
							if trim(rs(0))=USB then
								selected="Y"
								response.write "<option value=""" & trim(rs(0)) & """ selected>" & trim(rs(0)) & "</option>" & vbcrlf
							else
								response.write "<option value=""" & trim(rs(0)) & """>" & trim(rs(0)) & "</option>" & vbcrlf
							end if					
							rs.movenext
						wend
						rs.close
						if USB<>"" and selected="" then response.write "<option value=""" & USB & """ selected>" & USB & "</option>" & vbcrlf
					%>            
					</select>
				</td></tr>
			</table>
		</td>
    </tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
    <tr>
      <td><div align="center" class="item">處理過程</div></td>
      <td colspan="7"><textarea name="txtProcess" id="txtProcess" cols="100" rows="2" class="intxt" value="<%=Process%>"><%=Process%></textarea></td>      
    </tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr style="display:<% if IO="I" then response.write "none" %>">
		<td align="left" valign="middle"><div align="center" class="item">攜出物品</div></td>
		<td colspan="7" align="left" valign="middle">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr><td><span class="option">
			<%	StuffOutA=split(StuffOut,",")
				rs.open "select * from Config where kind='物品' order by Content",conn 
				while not rs.eof
					if StuffContent="" then
						response.write mid(rs("Content"),5)	& "："
					elseif StuffContent<>mid(rs("Content"),5) then
						response.write "</span></td></tr>"
						response.write "<tr><td><span class=""option"">"
						response.write mid(rs("Content"),5) & "："					
					end if
					StuffContent=mid(rs("Content"),5)
			%>
					<input type="checkbox" name="chkStuffOut" value="<%=rs("item")%>"
			<%		for i=0 to UBound(StuffOutA)
						if trim(StuffOutA(i))=rs("item") then
							response.write " checked"
							StuffOutA(i)=""
						end if
					next
			%>		><%=rs("item")%>&nbsp;				
			<%		rs.movenext
				wend
				rs.close
				for i=0 to UBound(StuffOutA)
					if StuffOutA(i)<>"" then otherStuffOut=otherStuffOut & trim(StuffOutA(i)) & ","
				next
				if otherStuffOut<>"" then otherStuffOut=mid(otherStuffOut,1,len(otherStuffOut)-1)
			%>	</span></td></tr>
				<tr><td>
					<span class="option">其它：&nbsp;&nbsp;<input name="txtStuffOut" type="text" value="<%=otherStuffOut%>" size="20"></span>
				</td></tr>
			</table>
		</td>
    </tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<tr style="display:<% if IO="I" then response.write "none" %>">    
		<td><span class="item">攜出原因</span></td>
		<td colspan="3"><input name="txtCause" type="text" value="<%=Cause%>" size="20"></td>
		<td><div align="center"><span class="item">值班人員<span class="style1">(出)</span></span></div></td>
		<td>
	  		<select name="selOPout" size="1" class="option">
				<option value=""></option>
			<%	OutFormat="<option value=""#Item#"" #Sel#>#Item#</option>"
                response.write GetConfig("IDMS","order by mark","作業管理科","Item",OPout,"*",OutFormat,opA)
			%>  
			</select>  
		<td><div align="center"><span class="item">離開時間</span></div></td>
		<td>
			<input name="txtDateOut" type="text"  value="<%=DateOut%>" size="10">
			<input name="txtTimeOut" type="text"  value="<%=TimeOut%>" size="8">
		</td>
	</tr>
<!---------------------------------------------------------------------------------------------------------------------------------------------------------------->
	<input name="DBaction" type="hidden" value="<%=DBaction%>">
	<input name="Serial" type="hidden" value="<%=Serial%>">	
		
</form></table>
<table width="100%"><tr>
	<td width="40%"></td>
	<td width="60%">
		<input name="cmdSend" type="button" class="butn" value="<%=cmdSend%>" onclick="CmdSend_Click();">
	</td>	
</td></tr></table>
<p align="center"><table><tr><td>
【<a target="_blank" href="../機房門禁管制輪值注意事項.doc">注意事項</a>】：<br>
<font color="red">
  ※ 有關廠商至B1維護時 , 請當班同仁 務必提醒關門及上鎖<br>
  ※ 資安規定，請OP由授權清冊確認申請人是否需負責人陪同進入<br>
  ※ 若有「硬體維修」、「軟體更新」或「組件安裝」時, 請主動提醒或詢問「執行者」, 是否需要"告知"可能受影響的對象或下游外單位<br>
  ※ 設備進出機房，請記得向負責人詢問設備編號及移入單，並記錄之！<br>
  ※ 設備進出機房或移動位置，請到資源管理子系統，更改放置地點欄位，詢問負責人若知地點就更改之，若未知請填寫(任意地點)(未定區域)，重點是它已不在機房內，放置地點不應寫機房內的任何位置。<br>
  ※ <a target="_blank" href="../主任指示門禁管制.doc">主任指示...</a>
</font>
<br><br>
【通行類型】：<br>
A類人員：數值資訊組人員或專案、駐點廠商等依授權清冊所列可不需通知、不需陪同之各課授權人員。<br>
B類人員：依授權清冊所列負責人需通知但得不陪同之各課授權人員。<br>
C類人員：未列於授權清冊上之上，負責人需通知且需陪同之人員。(授權清冊所列需通知且需陪同之人員亦然)<br><br>

【通知陪同注意事項】：<br>
1. 若負責人因故無事先通知或無法聯繫，則OP可主動通知負責人或其代理人或權責課長。<br>
2. 若負責人無法陪同，負責人可請其代理人或權責課長陪同。<br>
3. 若無法執行上述第1或第2點，OP需查核其身分證明後由操作課長同意始可進入，惟需勾選其它登入選項，記錄登入說明，並列入統計。<br>
4. B、C類人員非上班時間進入機房需事先由負責人通知，OP查驗其身分與負責人所通知者相同始可進入。<br>
5. 非上班時間之陪同得由OP代理，但僅以機房內之設備為限。
</td></tr></table></p>
<br><br><br><br><!----有的螢幕最後顯示不到------>

<script language="JavaScript">
<!-- #include file="..\Lib\Ajax.inc" -->
function trim(str) { return str.replace(/(^\s*)|(\s*$)/g,""); }
var e=document.getElementById("form");

if("<%=IO%>"=="O") window.open("outwarning.asp","_blank","width=220,height=340,resizable=0,scrollbars=no");	//登出提醒

//Kind	: onchange的選取值
//which	: onchange後選項要對應改變的物件
//other	: 其他參數(根據所關聯的負責人來產生選項)
//------------------------------------------------

function Option_HTML(which,Kind,other) { //選項變化
	var str ;
	
	if (window.ActiveXObject) {  // IE
		//http_request.setRequestHeader("Content-Type","application/x-www-form-urlencoded") ;	
		http_request.open("POST","Lib/OptionHTML_outer.asp?which=" + which + "&Kind=" + Kind + "&other=" + other,false) ;		
		http_request.send("") ;	
		str=http_request.responseText ;				
		e.elements[which].outerHTML=str ;
	} else { // Edge, Chrome, ...
		//http_request.setRequestHeader("Content-Type","application/x-www-form-urlencoded") ;	
		http_request.open("POST","Lib/OptionHTML_inner.asp?which=" + which + "&Kind=" + Kind + "&other=" + other,false) ;		
		http_request.send("") ;	
		str=http_request.responseText ;				
		e.elements[which].innerHTML=str ;
	}

	if(which=="selCompany") Company_KeyIn(Kind);	//申請單位自行輸入
	if(which=="selStaff") Staff_KeyIn(Kind);	
}

function Company_KeyIn(Kind) {	//申請單位自行輸入
	if(Kind=="Self") {
		e.elements["selCompany"].style.display="none" ;
		e.elements["txtCompany"].style.display="" ;
		e.elements["txtCompany"].focus() ;
		e.elements["selStaff"].style.display="none" ;
		e.elements["txtStaff"].style.display="" ;	
	} else {
		e.elements["selCompany"].style.display="" ;
		e.elements["txtCompany"].style.display="none" ;	
		e.elements["selStaff"].style.display="" ;
		e.elements["txtStaff"].style.display="none" ;	
	}	
}

function Staff_KeyIn(Kind) {	//申請人自行輸入
	if(Kind=="Self") {
		e.elements["selStaff"].style.display="none" ;
		e.elements["txtStaff"].style.display="" ;
		e.elements["txtStaff"].focus() ;
	} else {
		e.elements["selStaff"].style.display="" ;
		e.elements["txtStaff"].style.display="none" ;	
	}
}

function CmdSend_Click() {	//Submit前輸入檢查
if("<%=IO%>" != "O") {	//登入或修改
	//行動裝置(需放最前面判斷)--------------------------------------------------------------------
	for(i=0;i < e.elements["chkStuffIn"].length;i++) {
		if(e.elements["chkStuffIn"][i].checked) {
			if(e.elements["chkStuffIn"][i].value=="筆記型電腦" || e.elements["chkStuffIn"][i].value=="平版電腦" || e.elements["chkStuffIn"][i].value=="智慧型手機" || e.elements["chkStuffIn"][i].value=="行動電話") {
				alert("攜入行動裝置前請先取得負責人同意 !");
				break;
			}    
		}
	}
	 //--------------------------------------------------------------------------------------
	if(e.elements["selUnit"].value=="") {
		alert("請輸入負責單位 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(e.elements["selMaintainer"].value=="") {
		alert("請輸入負責人 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(e.elements["selCompany"].value=="" && e.elements["selComKind"].value=="Self" && trim(e.elements["txtCompany"].value)=="") {
		alert("請輸入申請單位 !");
		return("");
	} 
	if(e.elements["selComKind"].value=="Self" && trim(e.elements["txtCompany"].value)=="") {
		alert("請輸入申請單位 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(e.elements["selStaff"].value=="" && trim(e.elements["txtStaff"].value)=="") {
		alert("請輸入申請人 !");
		return("");
	}
	if(e.elements["selStaff"].value=="Self" && trim(e.elements["txtStaff"].value)=="") {
		alert("請輸入申請人 !");
		return("");
	} //--------------------------------------------------------------------------------------	
	//if(e.elements["selAmount"].value==2 && e.elements["txtMemo"].value=="") {
	//	alert("若申請人數為2人,附註需註明隨行人員姓名 !");
	//	return("");
	//} //--------------------------------------------------------------------------------------
	if(e.elements["selOPin"].value=="") {
		alert("請輸入值班人員(入) !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(trim(e.elements["txtDateIn"].value)=="") {
		alert("請輸入登入日期 !");
		return("");
	}
	if(e.elements["txtDateIn"].value.length != 10) {
		alert("登入日期格式不正確 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(trim(e.elements["txtTimeIn"].value)=="") {
		alert("請輸入登入時間 !");
		return("");
	}
	if(e.elements["txtTimeIn"].value.length != 8) {
		alert("登入時間格式不正確 !");
		return("");
	} //--------------------------------------------------------------------------------------
	var chkYN="n"; var tmpPurpose="" ;var chkSecPolicy="n";
	for(i=0;i < e.elements["chkPurpose"].length;i++) {
        //if (e.elements["chkPurpose"][i].value.indexOf("宣導資安政策門禁規定")!=-1){
        //    if(e.elements["chkPurpose"][i].checked)chkSecPolicy="y";
        //     continue;
        //}
		if(e.elements["chkPurpose"][i].checked) {
			tmpPurpose=tmpPurpose + e.elements["chkPurpose"][i].value + "," ;
            if(e.elements["chkPurpose"][i].value.indexOf("宣導資安政策門禁規定")!=-1){   
                chkSecPolicy="y";
                continue;
            }
            if(e.elements["chkPurpose"][i].value.indexOf("類人員")!=-1 || e.elements["chkPurpose"][i].value.indexOf("登入")!=-1)continue;
            chkYN="y";
			
		}
	}
	
	if(chkYN=="n") {
		alert("請選擇或輸入目的 !");
		return("");
	}
    
    if(chkSecPolicy == "n") {
		alert("請宣導資安政策門禁規定 !");
		return("");
	}

	if(tmpPurpose.indexOf("類人員") == -1) {
		alert("請勾選登入人員的通行類型是屬於ABC那一類人員 !");
		return("");
	} 
	
	if(tmpPurpose.indexOf("登入") == -1) {
		alert("請勾選登入的執行狀況是否依規定 !");
		return("");
	}
		
	if(tmpPurpose.indexOf("其它登入") > -1 & trim(e.elements["txtPurpose"].value)=="") {	//*********************************************其它登入
		alert("請填登入說明(例：聯絡不到相關權責人員，經操作課長同意後登入) !");
		return("");
	} //借用--------------------------------------------------------------------------------------
		
    for(i=0;i < e.elements["chkStuffIn"].length;i++) {
		if(e.elements["chkStuffIn"][i].checked) {
			if(e.elements["chkStuffIn"][i].value=="臨時卡" && e.elements["txtProcess"].value.indexOf("臨時卡號:") == -1 ) {
		        alert("請於處理過程註明臨時卡號 ! 格式如下：\n臨時卡號:###");
		        return("");
	        }
		}
	} //隨身碟--------------------------------------------------------------------------------------
	if(e.elements["selUSB"].value=="") {
		alert("請詢問是否攜入隨身碟，確實未使用請選擇未攜入 !");
		return("");
	} 
} else { //登出----------------------------------------------------------------------------------------------------------------------------------------
	if(e.elements["selOPout"].value=="") {
		alert("請輸入值班人員(出) !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(trim(e.elements["txtDateOut"].value)=="") {
		alert("請輸入登出日期 !");
		return("");
	}
	if(e.elements["txtDateOut"].value.length != 10) {
		alert("登出日期格式不正確 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(trim(e.elements["txtTimeOut"].value)=="") {
		alert("請輸入登出時間 !");
		return("");
	}
	if(e.elements["txtTimeOut"].value.length != 8) {
		alert("登出時間格式不正確 !");
		return("");
	} //--------------------------------------------------------------------------------------
	if(e.elements["txtDateIn"].value>e.elements["txtDateOut"].value) {
		alert("輸入登入日期大於登出日期 !");
		return("");
	}
	if(e.elements["txtDateIn"].value==e.elements["txtDateOut"].value && e.elements["txtTimeIn"].value>e.elements["txtTimeOut"].value) {
		alert("輸入登入時間大於登出時間 !");
		return("");
	} //借用--------------------------------------------------------------------------------------
    for(i=0;i < e.elements["chkStuffIn"].length;i++) {
		if(e.elements["chkStuffIn"][i].checked) {
			if(e.elements["chkStuffIn"][i].value=="臨時卡" && e.elements["txtProcess"].value.indexOf("臨時卡號:") == -1 ) {
		        alert("請於處理過程註明臨時卡號 ! 格式如下：\n臨時卡號:###");
		        return("");
	        }
            if(e.elements["chkStuffIn"][i].value == "臨時卡") {
		        if(! confirm("請確認臨時卡是否已經歸還 !")) return("");
	        }
	        if(e.elements["chkStuffIn"][i].value == "鑰匙") {
		        if(! confirm("請確認鑰匙是否已經歸還 !")) return("");
	        }
	        if(e.elements["chkStuffIn"][i].value == "遙控器") {
		        if(! confirm("請確認遙控器是否已經歸還 !")) return("");
	        }
		}
	} //-------------------------------------------------------------------------------------- 
	if(trim(e.elements["txtProcess"].value)=="") {
		alert("請輸入處理過程 !");
		return("");
	}	
	
	if(e.elements["selCompany"].value=="天霖") { //告知UPS用電量訊息
		window.open("upswarning.asp","_blank","width=150,height=240,resizable=0,scrollbars=no");		
	}
}//----------------------------------------------------------------------------------------------------------------------------------------------------	

	//*********************************************設備異動提醒
	for(i=0;i < e.elements["chkPurpose"].length;i++) { 
		if(e.elements["chkPurpose"][i].checked && e.elements["chkPurpose"][i].value.indexOf("異動設備")>-1) {
			alert("1. 請記得向負責人詢問設備編號及移入單，並記錄之！\n2. 設備進出機房或移動位置，請更改IDMS放置地點，詳見本畫面底【注意事項】。");
		}
	}
	
	//替換掉textarea輸入的換行字元及TAB鍵
	var txt1=document.getElementById("txtWorkArea").value;
	var txt2=document.getElementById("txtProcess").value;
	var regex_newline = /\n/g;
	var regex_tab = /\t/g;
	var regex_endspace = /(\s+)$/;
	document.getElementById("txtWorkArea").value = txt1.replace(regex_newline, ' ').replace(regex_tab, '').replace(regex_endspace, '');
	document.getElementById("txtProcess").value = txt2.replace(regex_tab, '').replace(regex_endspace, '');
	
	e.submit();
}			
</script>
</body>
</html>