<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<%
Dim connDiary 	 : Set connDiary=Server.CreateObject("ADODB.Connection")  : connDiary.Open strDiary
Dim connIDMS 	 : Set connIDMS=Server.CreateObject("ADODB.Connection")   : connIDMS.Open strIDMS
Dim connDevice   : Set connDevice=Server.CreateObject("ADODB.Connection") : connDevice.Open strControl
dim rs : set rs=server.createobject("ADODB.recordset")

'CreateDate="" ?,後面會用來判斷device.asp資料是否對應至DB(新增,修改,顯示皆有,無法用DBaction判斷),Reset用
DBaction=lcase(request("DBaction"))	'新增(c) 修改(s) 刪除(d) 顯示或建檔() 匯入(m)
CreateDate=trim(request("TextCreateDate"))		:	HostName  =trim(Request("TextHostName"))
HostClass =trim(Request("ComboHostClass"))		:	Repair    =trim(Request("TextRepair"))
IO		  =trim(Request("RadioIO"))				:	Functions =trim(Request("TextFunctions"))
Memo     =trim(Request("TextMemo"))             :   PS=trim(Request("CombPS"))                 
StaffName =trim(Request("TextStaffName"))		:	Hw        =trim(Request("TextHw"))
SW        =trim(Request("TextSw"))				:	OP        =trim(Request("TextOP"))
UpdateDate=DT(now,"yymmddhhmiss")
PointerNo =trim(Request("TextPointerNo"))       :   Xall=0 : Yall=0
Areas="尚未定位"

Sub GetXY(PointerNo)
    dim rs1 : set rs1=server.createobject("ADODB.recordset")
    if PointerNo<>"" then       
        rs1.open "select * from [定位設定] where [定位編號]=" & PointerNo,connIDMS
        if not rs1.eof then
            Xall=rs1("坐標X") : Yall=rs1("坐標Y")
            Areas=rs1("區域名稱") & rs1("定位名稱")
        end if
        rs1.close
    end if          
End Sub

Sub GetArea(Xall,Yall)
    dim rs1 : set rs1=server.createobject("ADODB.recordset")
    if Xall<>"" then       
        rs1.open "select * from [定位設定] where [定位方式]='坐標' and [坐標X]=" & Xall & " and [坐標Y]=" & Yall,connIDMS
        if not rs1.eof then
            PointerNo=rs1("定位編號")
            Areas=rs1("區域名稱") & rs1("定位名稱")
        end if
        rs1.close
    end if          
End Sub

select case DBaction
case "c","s"	'新增(c) 修改(s)
	if DBaction="c" then	'新增		
		CreateDate=UpdateDate
        call GetXY(PointerNo)
	    sql="Insert into Device values('" & CreateDate & "','" & HostName & "','" & HostClass & "','" & Repair & "','" _
			& IO & "','" & Functions & "','" & Memo & "','" & PS & "','"  _
			& StaffName & "','" & Hw & "','" & Sw & "','" & OP & "'," _
			& Xall & "," & Yall & ",'" & UpdateDate & "');"		
		connDevice.execute sql
	else	'修改
        call GetXY(PointerNo)
		rs.open "Select * from Device where CreateDate='" & CreateDate & "'",connDevice,3,3				
		rs("HostName")=HostName     :     rs("HostClass")=HostClass        :     rs("Repair")=Repair
		rs("IO")=IO                          :     rs("Functions")=Functions       :     rs ("Memo")=Memo
		rs("PS")=PS                          :     rs("StaffName")=StaffName
		rs("Hw")=Hw				:	rs("Sw")=Sw				 :     rs("OP")=OP
		rs("Xall")=Xall				:	rs("Yall")=Yall			       :	rs("UpdateDate")=UpdateDate
		rs.update					:	rs.close
	end if
	
	response.write "<script language=""vbscript"">" & vbcrlf
	response.write "Sub window_onload()" & vbcrlf
	response.write "	msgbox ""儲存完成!"",32" & vbcrlf
	response.write "End Sub" & vbcrlf
	response.write "</script>" & vbcrlf
case "d"	'刪除
	connDevice.execute "Delete from Device where CreateDate='" & CreateDate & "'"
	SQL="Select * from Device where CreateDate like '" & mid(CreateDate,1,4) & "%' order by CreateDate Desc"
	response.redirect "List.asp?SQL=" & Server.URLEnCode(SQL)
case else	'顯示,建檔("")或帶入("m")
	OP=ShowOP
	if CreateDate="" then CreateDate=trim(request("CreateDate"))
	rs.open "Select * from Device where CreateDate='" & CreateDate & "'",connDevice
	if not rs.eof then
		HostName=rs("HostName")     :     HostClass=rs("HostClass")      :     Repair=rs("Repair")
		IO=rs("IO")                           :     Functions=rs("Functions")    :     Memo=rs("Memo")
		PS=rs("PS")             :     StaffName=rs("StaffName")
		Hw=rs("Hw")				 :     Sw=rs("Sw")		           :     if DBaction<>"m" then OP=rs("OP")
		Xall=rs("Xall")				 :	 Yall=rs("Yall")			     :     UpdateDate=rs("UpdateDate")
	end if
	rs.close
	call GetArea(Xall,Yall)

	if DBaction="m" then CreateDate=""	'TextCreateDate是存檔時submit的,CreateDate是*.asp?CreateDate=...的
	
	if CreateDate="" then
		response.write "<script language=""vbscript"">" & vbcrlf
		response.write "Sub window_onload()" & vbcrlf
		response.write "	msgbox ""1. 請記得向負責人詢問設備編號及移入單，並記錄之！"" & vbcrlf & ""2. 設備進出機房或移動位置，請更改IDMS放置地點，詳見本畫面底【注意事項】。"",48" & vbcrlf
		response.write "End Sub" & vbcrlf
		response.write "</script>" & vbcrlf
	end if
end select

Function ShowOP()
	rs.open "Select * from SIGN where TOUR=(Select max(Tour) from SIGN)",connDiary
	if not rs.eof then ShowOP=trim(rs("OPNAME"))
	rs.close
End Function

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
%>
<html>
 <head>
  <meta http-equiv="Content-Type" content="text/html; charset=big5" />
  <title>資料建檔</title>	
  <link rel="stylesheet" type="text/css" href="JsDoMenu/jsdomenu.css" />
  <script type="text/javascript" src="JsDoMenu/jsdomenu.js"></script>
 </head>
 <body onLoad="initjsDOMenu()" bgcolor="#DFEFFF" text="#0000FF" onMouseOver="HideTimeOutMenus()" link="#FF0066">
  <p align="center"><b><font size="5" color="#000066">設備異動資料建檔或異動介面</font></b></p>
  <form name="form1" method="POST" target="_self" action="device.asp">   <!-- action 未寫,會很難debug-->         
   <input type="hidden"  name="DBaction" id="DBaction">                                                                                                                    
   <p align="center"><table border="1" cellpadding="0" cellspacing="0" width="98%">
    <tr>
      <td  height="40" width="105"><font size="3">建檔日期</font></td>
      <td  height="40" width="383">
        <input type="hidden" name="TextCreateDate" value='<% if CreateDate<>"" and DBaction<>"m" then response.write CreateDate%>'>      	
        <font size="3" color="black"><b><span id="divCreateDate">
<% if CreateDate<>empty and DBaction<>"m" then 
	   response.write "20" & mid(CreateDate,1,2) & "/" & mid(CreateDate,3,2) & "/" & mid(CreateDate,5,2) _
                           & " " & mid(CreateDate,7,2) & ":" & mid(CreateDate,9,2) & ":" & mid(CreateDate,11,2)
      else
         response.write "尚未建檔"
      end if 
%>
        </span></b></font>　                                                                                                                                                                                                                                                             
<% if CreateDate="" or DBaction="d" or DBaction="m" then %>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
    	  <font size="2" color="#800080"><u>                                                                                                                                                          
    	    <span  style="cursor:pointer;" onMouseOver="ShowMenu('menuHostName')" onMouseOut="MenusOut()" tagName="spanHost">參考選單</span>       
    	  </u></font>                                                                                                               
    	  <font color="#CC6600" size="2">(可匯入key in 過資料)</font>                                                                                                                                                 
	  <font color="#CC6600"><% end if %></font>                                                                
	 </td>        
       <td height="40" width="113"><font size="3">異動種類 (<font color="#FF0000">*</font>)</font></td>                                                                                                                                                                                                                                                 
       <td height="40" width="372">         <font size="3" color="black" title="設備移入機房">       
        	<input type="radio" value="I" name="RadioIO" <%if IO="" or IO="I" then response.write "checked"%>><b>移入</b>          
        	<font color="#FF0000">★</font>         
        </font>　                                                                                                                                                                                                                                                                                
        <font size="3" color="black" title="設備移出機房">       
        	<input type="radio" value="O" name="RadioIO" <%if IO="O" then response.write "checked"%>><b>移出</b>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 
        	<font color="#FF0000">Ｘ</font>       
        </font>　                                                                                                                                                                                                                                                                                
        <font size="3" color="black" title="更換設備(舊換新，壞換好)">       
        	<input type="radio" value="N" name="RadioIO" <%if IO="N" then response.write "checked"%>><b>更換</b>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                
        	<font color="#FF0000">★</font>       
        </font>　                                                                                                                                                                                                                                                                                
        <font size="3" color="black" title="主機更名，設備移位....(請於設備用途欄位中註明)">       
        	<input type="radio" value="M" name="RadioIO" <%if IO="M" then response.write "checked"%>><b>其它</b>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      
        	<font color="#FF0000">★</font>                                                                                                          
        </font>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               
      </td>        
    </tr>       
    <tr>              
      <td height="40" width="105">       
      	設備種類 (<font color="#FF0000">*</font>)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
      </td>       
      <td height="40" width="383">       
      	<select size="1" name="ComboHostClass" style="visibility:">             
<%       
    	response.write "<option></option>"       
     	YN=false       
	rs.open "select Item from Config where Kind='設備種類' order by Item",connDevice       
	while not rs.eof       
		if rs(0)=HostClass then       
			response.write "<option selected value='" & HostClass & "'>" & HostClass & "</option>"       
			YN=true       
		else       
			response.write "<option value='" & rs(0) & "'>" & rs(0) & "</option>"       
		end if       
		rs.movenext       
	wend       
	rs.close       
     
	if not YN and HostClass<>"" then response.write "<option selected value='" & HostClass & "'>" & HostClass & "</option>"	       
%>       
           </select>　                                                                                                                                                                                                                                                                         
         <font color="#CC6600" size="2">(若為主機，設備名稱欄位請填 hostname )</font>                                                                      
        </td>                             
        <td height="40" width="113"><font size="3">設備廠商</font></td>       
        <td height="40" width="372">       
          <input type="text" name="TextRepair" size="20" value="<%=repair%>">       
          <u><span onMouseOver="ShowMenu('menuRepair')" onMouseOut="MenusOut()">   
          <font size="2" color="#800080" style="cursor:pointer;">參考選單</font></span></u>       
        </td>       
      </tr>       
      <tr>             
       <td height="40" width="105">設備名稱 (<font color="#FF0000">*</font>) </td>                                                                    
       <td  height="40" width="868" colspan="3">       
         <input type="text" name="TextHostName" size="80" value="<%=HostName%>">　<font color="#CC0000" size="2">(<b>請註記IDMS的設備編號</b>)</font>                                                                      
       </td>                                   
    </tr> 
    <tr> 
      <td height="40" width="105"><font size="3">設備功能</font></td> 
      <td height="40" width="868" colspan="3"><input type="text" name="TextFunctions" size="80" value="<%=Functions%>">　<font color="#CC6600"><font size="2">(請填寫設備主要功用，附註請填寫於設備說明)</font></font>　</td> 
    </tr>         
    <tr> 
      <td height="40" width="105">設備說明</td>   
      <td height="40" width="868" colspan="3">   
        <input type="text" name="TextMemo" size="80" value="<% =Memo%>">　<font color="#CC6600" size="2">(備品、替代、移位、暫放走道、移出維護、移至地下室、報廢...)　</font>　
        <br>
        <font size="2">(需/不需)移入單</font>
        <select size="1" name="CombPS">                                       
        <option value="" <%if PS="" then response.write "selected"%> ></option>
        <option value="需單" <%if PS="需單" then response.write "selected"%> >需填移入單</option>
        <option value="測試" <%if PS="測試" then response.write "selected"%> >不需填移入單－測試</option>
        <option value="暫放" <%if PS="暫放" then response.write "selected"%> >不需填移入單－暫放</option>
        <option value="代管" <%if PS="代管" then response.write "selected"%> >不需填移入單－代管</option>
        </select>
        <font color="#CC6600" size="2">(首次移入請填移入申請單，測試、暫放、代管除外)</font>                                                                                                                  
      </td>   
    </tr>       
    <tr>  
      <td height="40" width="105"><font size="3">申請人員 (<font color="#FF0000">*</font>)</font></td>                                                                    
      <td height="40" width="383">       
        <input type="text" name="TextStaffName" size="20" value="<%=StaffName%>">       
        <u><span onMouseOver="ShowMenu('menuStaffName')" onMouseOut="MenusOut()">                                                                         
        <font size="2" color="#800080" style="cursor:pointer;">參考選單</font></span></u><font color="#CC6600" size="2">&nbsp;&nbsp;(亦可自行Key In)</font></td>       
      <td height="40" width="113">       
        <font size="3">硬體負責人 (<font color="#FF0000">*</font>)</font><font color="#CC6600" size="2">                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            
      </td>       
      <td height="40" width="372">       
        <input type="text" name="TextHw" size="20" value="<%=hw%>">        
        <u><span onMouseOver="ShowMenu('menuHw')" onMouseOut="MenusOut()">                                                                         
        <font size="2" color="#800080" style="cursor:pointer;">參考選單</font></span></u><font color="#CC6600" size="2">&nbsp;&nbsp;(亦可自行Key In)</font>                                                                               
      </td>         
    </tr>           
    <tr>       
      <td height="40" width="105"><font size="3">ＯＰ　 　(<font color="#FF0000">*</font>)</font></td>                                                                               
      <td height="40" width="383">   
        <input type="text" name="TextOP" size="20" value="<%=op%>"> 
          <u><span onMouseOver="ShowMenu('menuOP')" onMouseOut="MenusOut()">                                                                         
          <font size="2" color="#800080" style="cursor:pointer;">參考選單</font></span></u><font color="#CC6600" size="2">&nbsp;&nbsp;(亦可自行Key In)</font>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   
        </td>        
        <td height="40" width="113"><font size="3">軟體負責人 (<font color="#FF0000">*</font>)</font></td>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              
        <td height="40" width="372">       
          <input type="text" name="TextSw" size="20" value="<%=sw%>">       
          <u><span onMouseOver="ShowMenu('menuSw')" onMouseOut="MenusOut()">                                                                         
          <font size="2" color="#800080" style="cursor:pointer;">參考選單</font></span></u><font color="#CC6600" size="2">&nbsp;&nbsp;(亦可自行Key In)</font>                                                                               
        </td>                           
      </tr>       
      <tr>       
       <td height="40" width="105" align="center">       
         <input type="button" value="設備定位(*)" name="CmdMapLink" style="color: #FF0000; font-weight: bold">                                                    
       </td>       
       <td id="tdMap" height="40" width="383">       
        <input name="TextPointerNo" type="text" size="4" id="TextPointerNo" value="<%=PointerNo%>" OnKeyDown="event.returnValue=false;" style="background-color:Silver;" /> 
        <font size="3" color="black"><b><span id="divAreas"><%=Areas%></span></b></font>       
      </td>       
      <td height="40" width="113"><font size="3">修改日期</font></td>       
      <td height="40" width="372">       
        <font size="3" color="black"><b>       
<%  if UpdateDate<>"" and DBaction<>"m" then       
   		response.write "20" & mid(UpdateDate,1,2) & "/" & mid(UpdateDate,3,2) & "/" & mid(UpdateDate,5,2) & " "  _       
		& mid(UpdateDate,7,2) & ":"  & mid(UpdateDate,9,2) & ":"  & mid(UpdateDate,11,2)       
    	else       
    		response.write DT(now,"yyyy/mm/dd hh:mi:ss")       
    	end if       
%>       
     	  </b></font>       
      </td>       
    </tr>       
  </table> 
  </p> 
  
  <p align="center">      
  <table border="0" width="98%">       
    <tr>       
      <td width="71%" valign="top"> 
      	<font color="#CC6600">【<a target="_blank" href="../機房門禁管制輪值注意事項.doc">注意事項</a>】</font><br><br>      
        <font color="#CC6600">一、註明(</font><font color="#FF0000">*</font>   
        <font color="#CC6600">)為必填欄位，其它欄位若空白將以 N/A 取代</font><br><br>                                                                      
        <font color="#CC6600">二、離開前請將空箱子或雜物整理乾淨</font><br><br>                                                               
        <b><font color="#CC6600">三、若有「硬體維修」、「軟體更新」或「組件安裝」時,&nbsp;請主動提醒或詢問「執行者」<br>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 是否需要&quot;告知&quot;可能受影響的對象或下游外單位</font></b><br><br>
        <font color="#CC6600">四、 設備進出機房，請記得向負責人詢問設備編號及移入單，並記錄之！</font><br><br>
  		<font color="#CC6600">五、 設備進出機房或移動位置，請到資源管理子系統，更改放置地點欄位，詢問負責人若知地點就更改之，若未知請填寫(任意地點)(未定區域)，重點是它已不在機房內，放置地點不應寫機房內的任何位置。</font> <br><br>  
  		<font color="#CC6600">六、 <a target="_blank" href="../主任指示門禁管制.doc">主任指示...</a></font>  
      </td>      
      <td width="29%" valign="middle">       
         <p align="right">	      	 				       
<% if CreateDate="" or DBaction="m" then %>       
          <input type="button" value=" 建 檔 " name="cmdAdd" style="font-size: 24pt; font-weight: bold" title="新建一筆資料">       
<% else %>        
          <input type="button" value=" 修 改 " name="cmdUpd" style="font-size: 18pt; font-weight: bold" title="修改本筆資料">	&nbsp;&nbsp;&nbsp;&nbsp;		        
	    <input type="button" value=" 刪 除 " name="cmdDel" style="font-size: 18pt; font-weight: bold" title="刪除本筆資料">      
<% end if %>      
	   </p>                                                                                                                                                                            
       </td>       
     </tr>       
   </table> 
</p>         
 </form>   
     
<div align="center"><center>       
 <table border="0" width="656" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">       
  <tr>       
    <td width="652" align="center">  
     <marquee style="color: rgb(255,0,0); font-size: 15pt; font-family: 標楷體; font-style: italic; font-weight: bold" align="middle" width="650" scrollamount="5" scrolldelay="5" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 102年12月課務會議決議：機器上機房機架前，先請負責人在IDMS建立設備資訊，並繳交移入單，若無則請其當場填寫！並核對IDMS資料是否正確。</marquee>  
    </td>       
   </tr>       
 </table>       
</center></div>     
</body>
<script language="vbscript">
'假如window.open "*.asp"或window.location.reload,submit動作有時會失敗
'有時無法異動是因form1及其欄位名稱用了ID,必須用name才可以,但div需用id
'----------------------------------------------設備定位--------------------------------------------------------------------------
Sub cmdMapLink_onClick()
    PointerNo=form1.TextPointerNo.value
    if PointerNo="" then PointerNo="0"

	for each obj in Document.all.RadioIO
		if obj.checked then exit for
	next
	if obj.value="O" then
		window.open "/IDMS/Lib/map.aspx?IO=O&PointerNo=" & PointerNo,"_blank","location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no"
	else
		window.open "/IDMS/Lib/map.aspx?IO=I&PointerNo=" & PointerNo,"_blank","location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no"
	end if
	
	divAreas.innerText="定位地點將於存檔後顯示"
End Sub
'----------------------------------------------建檔/修改---------------------------------------------------------------------
Sub cmdAdd_onClick()
	msgbox "102年12月課務會議決議，機器上機房機架前：" & vbcrlf & vbcrlf _
		& "1. 先請負責人在IDMS建立設備資訊" & vbcrlf _
		& "2. 並繳交移入單(空白移入單IDMS有連結)" & vbcrlf _
		& "3. 在門禁、日誌追蹤及移入單註記設備編號，以便可彼此關聯" & vbcrlf _
		& "4. 若無則請其當場填寫！" & vbcrlf _
		& "5. 若有缺漏，務必記錄日誌追蹤，並建議設定日誌mail通知" & vbcrlf _
		& "6. 繳移入單時請核對IDMS資料是否正確，並移除日誌mail通知設定" & vbcrlf _
		& "7. 若無則請其當場填寫！並核對IDMS資料是否正確。"  & vbcrlf _
		& "8. 若機器移出機房，請即時更改IDMS的放置地點。",48

		
	if SaveYN then
		form1.DBaction.value="c"
		form1.submit()
	end if
End Sub
Sub cmdUpd_onClick()
	if SaveYN then
		form1.DBaction.value="s"
		form1.submit()
	end if
End Sub
Function SaveYN()
	SaveYN=True
	if trim(form1.textHostName.value)="" then
		msgbox "[設備名稱]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if
	if trim(form1.ComboHostClass.value)="" then
		msgbox "[設備種類]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if	
	if trim(form1.TextStaffName.value)="" then
		msgbox "[申請人員]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if
	if trim(form1.textHw.value)="" then
		msgbox "[硬體負責人]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if
	if trim(form1.textSw.value)="" then
		msgbox "[軟體負責人]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if
	if trim(form1.textOP.value)="" then
		msgbox "[OP]欄位空白!，請補填。",48
		SaveYN=False : exit Function
	end if
	if form1.TextPointerNo.value="" or form1.TextPointerNo.value="0" then
		msgbox "[設備定位]欄位空白!，請點選設備定位按鈕以標定設備位置。",48
		SaveYN=False : exit Function
	end if
	
	if trim(form1.TextRepair.value)="" then form1.TextRepair.value="N/A"
	if trim(form1.TextFunctions.value)="" then form1.TextFunctions.value="N/A"

	call MsgboxPower()
End Function

Sub MsgboxPower()
	if form1.RadioIO(0).checked then msgbox "請當班OP 登記 IDMS相關位置與電力資訊 !",48						'移入		
	if form1.RadioIO(1).checked then msgbox "請當班OP 修改 IDMS相關位置與電力資訊(移出維修則免)!",48			'移出
	if form1.RadioIO(2).checked then msgbox "請當班OP 詢問 IDMS相關位置與電力是否有變動，若有請更改 !",48		'更換
	if form1.RadioIO(3).checked then msgbox "請當班OP 更新 IDMS相關位置與電力資訊 !",48						'其它
End Sub 
'----------------------------------------------列印/刪除/重設--------------------------------------------------------------------------
Sub cmdDel_onClick()
	CreateDate=form1.TextCreateDate.value
	HostName=form1.TextHostName.value
	if CreateDate<>"" then
		YN=false
		if msgbox("確定要刪除[" & hostname & "]？",49,"刪除設備資料")=1 then
			call MsgboxPower()
			form1.DBaction.value="d"
			form1.submit()
		end if	
	end if	
End Sub
'----------------------------------------------按鍵--------------------------------------------------------------------------
Sub TextHostName_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextRepair_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextFunctions_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextStaffName_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextHw_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextSw_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
Sub TextOP_onKeyPress()
	select case window.event.keycode
	case 39,34,60,62,123,125 ''"<>{}
		window.event.keycode=0
	end select
End Sub
  </script>
<script language="javascript">
var whichselected;
function ShowMenu(which)
{	menuYN="N";
	HidingAllMenus();
	menuYN="Y";	
	whichselected=which;
	switch (which)
    { case "menuHostName":
  		//form1.ComboHostClass.style.visibility="hidden";
        setPopUpMenu(menuHostName);
        break;
      case "menuRepair":
        setPopUpMenu(menuRepair);
        break;
      case "menuFunctions":
        setPopUpMenu(menuFunctions);
        break;
      case "menuStaffName":
        setPopUpMenu(menuStaffName);
        break;
      case "menuHw":
        setPopUpMenu(menuStaffName);
        break;
      case "menuSw":
        setPopUpMenu(menuStaffName);
        break;
      case "menuOP":
        setPopUpMenu(menuStaffName);
        break;
    }	
	activatePopUpMenuBy(0,2);
	leftClickHandler();
}

//---------------------------陣列 = new jsDOMenu(長度); addMenuItem(new menuItem("項目名稱","項目變數","code:執行函數"));
//---------------------------陣列 = new jsDOMenu(長度); menuHostName.items.項目陣列.setSubMenu(menuUnit" & i & ");
function createjsDOMenu()
{//--------------------------------------設備名稱------------------------------------------------------------------------------
  menuHostName = new jsDOMenu(108);
  <% response.write "menuHostName.addMenuItem(new menuItem(""所有設備...."",""Host0" & """,""code:SelectMenu('other');""));" & vbcrlf
	rs.open "select hostname,CreateDate,staffName from Device order by CreateDate Desc",connDevice
	 for i=1 to 5
	 	response.write "menuHostName.addMenuItem(new menuItem(""最近" & 10*i & "筆設備"",""Host" & i & """,""""));" & vbcrlf
  		response.write "menuHost" & i & " = new jsDOMenu(400);" & vbcrlf
		response.write "with (menuHost" & i & ")" & vbcrlf
		response.write "{" & vbcrlf
		j=0
		while not rs.eof and j<10
			j=j+1 
			response.write "addMenuItem(new menuItem(""(" & mid(rs(1),3,2) & "/" & mid(rs(1),5,2) & "," & rs(2) & ")" _
				& rs(0) & ""","""",""code:SelectMenu('" & rs("CreateDate") & "');""))" & vbcrlf
			rs.movenext
		wend			 	
		response.write "}" & vbcrlf
		response.write "menuHostName.items.Host" & i & ".setSubMenu(menuHost" & i & ");" & vbcrlf
	 next
	 rs.close
   %>
//--------------------------------------設備廠商------------------------------------------------------------------------------
  menuRepair = new jsDOMenu(96);
  <% 'rs.open "select Item from Config where Kind='設備廠商' order by Item",connDevice
  	 rs.open "select distinct repair from device where CreateDate>='" & DT(now-1000,"yymmddhhmiss") & "'",connDevice
	 i=0
	 while not rs.eof
	 	i=i+1
	 	response.write "menuRepair.addMenuItem(new menuItem(""前" & 10*i & "筆廠商"",""Repair" & i & """,""""));" & vbcrlf
	 	
  		response.write "menuRepair" & i & " = new jsDOMenu(80);" & vbcrlf
		response.write "with (menuRepair" & i & ")" & vbcrlf
		response.write "{" & vbcrlf
		j=0
		while not rs.eof and j<10
			j=j+1
			response.write "addMenuItem(new menuItem(""" & rs(0) & ""","""",""code:form1.TextRepair.value='" & rs(0) & "';""));" & vbcrlf
			rs.movenext
		wend					 	
		response.write "}" & vbcrlf
		response.write "menuRepair.items.Repair" & i & ".setSubMenu(menuRepair" & i & ");" & vbcrlf
	 wend
	 rs.close
   %>
/*--------------------------------------設備用途------------------------------------------------------------------------------
  menuFunctions = new jsDOMenu(120);
  with (menuFunctions)  	
  <% 'response.write "{" & vbcrlf
'  	 for i=1 to 5
' 		response.write "addMenuItem(new menuItem(""最近" & 10*i & "筆用途"",""Func" & i & """,""""));" & vbcrlf	    
'	 next
'	 response.write "}" & vbcrlf
'	 rs.open "select distinct Functions,CreateDate,staffName from Device order by CreateDate Desc",connDevice
'	 for i=1 to 5
'  		response.write "menuFunc" & i & " = new jsDOMenu(400);" & vbcrlf
'		response.write "with (menuFunc" & i & ")" & vbcrlf
'		response.write "{" & vbcrlf
'		j=0
'		while not rs.eof and j<10
'			j=j+1
'			response.write "addMenuItem(new menuItem(""(" & mid(rs(1),3,2) & "/" & mid(rs(1),5,2) & "," & rs(2) & ")" _
'				& rs(0) & ""","""",""code:form1.TextFunctions.value='" & rs(0) & "';""));" & vbcrlf
'			rs.movenext
'		wend			 	
'		response.write "}" & vbcrlf
'		response.write "menuFunctions.items.Func" & i & ".setSubMenu(menuFunc" & i & ");" & vbcrlf
'	 next
'	 rs.close
   %>
----------------------------------------------------------------------------------------------------------------------------*/
//--------------------------------------申請單位------------------------------------------------------------------------------
  menuStaffName = new jsDOMenu(92);
  with (menuStaffName)
{<%
  	rs.open "select Item from Config where kind='資訊中心'",connIDMS
	i=1
	dim UnitNameA(50)
	while not rs.eof
  		response.write "addMenuItem(new menuItem(""" & rs(0) & """,""Unit" & i & """,""""))" & vbcrlf
	  	UnitNameA(i)=rs(0)
		rs.movenext
		i=i+1
  	wend
  	rs.close
%>}<%
  while i>1
  	i=i-1
  	response.write "menuUnit" & i & " = new jsDOMenu(72);" & vbcrlf
	response.write "with (menuUnit" & i & "){" & vbcrlf
    rs.open "select Item from Config where Kind='" & UnitNameA(i) & "' order by Mark",connIDMS
	while not rs.eof
		response.write "addMenuItem(new menuItem(""" & rs(0) & """," & """""" & ",""code:SelectMenu('" & rs(0) & "');""))" & vbcrlf
		rs.movenext
	wend
	rs.close
	response.write "}" & vbcrlf
	response.write "menuStaffName.items.Unit" & i & ".setSubMenu(menuUnit" & i & ");"	
  wend
%>}
//--------------------------------------欄位填值------------------------------------------------------------------------------
function SelectMenu(txt)
{	switch (whichselected)
	{ case "menuStaffName":
        form1.TextStaffName.value=txt;
        break;
      case "menuHw":
        form1.TextHw.value=txt;
        break;
      case "menuSw":
        form1.TextSw.value=txt;
        break;
      case "menuOP":
        form1.TextOP.value=form1.TextOP.value+"  "+txt;
        break;
      case "menuHostName":
        switch (txt)
        { case "":        	
            break;
          case "other":
          	window.open("List.asp?Sel=Y","_blank");
            break;
          default:
          	window.open("device.asp?CreateDate=" + txt + "&DBaction=m","_self");
            break;
        }      
        break;
    }
    //form1.ComboHostClass.style.visibility="";		
}
  </script>
</html>