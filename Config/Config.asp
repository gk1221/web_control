<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID2.inc" -->
<%
Dim connControl  : Set connControl=Server.CreateObject("ADODB.Connection")  : connControl.Open strControl
set rs=server.createobject("ADODB.recordset")

head="系統參數設定"	'標題

Which=request("which") : Kind=request("selKind")
if kind<>"" then
	Item=request("txtItem") : oldItem=request("selItem") : Content=request("txtContent")
	DBaction=request("DBaction") 'a:新增 u:修改 d:刪除 m:顯示

	select case DBaction
	case "a"	'新增
		connControl.execute "insert into Config values('" & kind & "','" & Item & "','" & Content & "')"			
	case "u"	'修改
		rs.open "select * from Config where kind='" & kind & "' and Item='" & oldItem & "'",connControl,3,3
		rs(1)=Item : rs(2)=Content
		rs.update : rs.close
	case "d"	'刪除
		connControl.execute "delete from Config where kind='" & kind & "' and Item='" & oldItem & "'"
		Item="" : Content=""
	case "m"	'顯示
	end select
end if
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>系統設定</title>
</head>
<body bgcolor="#DFEFFF"><br>
<p align="center"><font color="blue" size="5"><b>
<%	select case which
	case "people"
		response.write "人員進出系統設定"
	case "device"
		response.write "設備異動系統設定"
	end select
%>
</font></b></p>
<form name="form" id="form" method="POST" target="_self"> 
<table width="300" align="center">
<tr>
	<td>類別：</td>
	<td>
		<select size="1" name="selKind" onchange="form.txtItem.value='';form.txtContent.value='';form.DBaction.value='m';form.submit();"> 
			<option></option>
		<%	rs.open "select * from Config where kind='" & which & "' order by Content",connControl
			while not rs.eof
       			if rs(1)=Kind then
	       			response.write "<option value='" & rs(1) & "' selected>" & rs(1) & "</option>"
			   	else
			   		response.write "<option value='" & rs(1) & "'>" & rs(1) & "</option>"
	    	   	end if
	       		rs.movenext
	    	wend	
    	   	rs.close  	  	
		%>        
		</select>
	</td>
</tr>
<tr>
	<td>項目：</td>
	<td>
		<select size="18" name="selItem" onclick="selItem_onClick(this);"> 
		    <%	rs.open "select * from Config where kind='" & Kind & "' order by Content",connControl
    			while not rs.eof	'Item(Content)：order
    				if rs(2)="" then
    					if rs(1)=Item then
			   				response.write "<option value='" & rs(1) & "' selected>" & Server.HTMLEnCode(rs(1)) & "</option>"
				   		else
		   					response.write "<option value='" & rs(1) & "'>" & Server.HTMLEnCode(rs(1))  & "</option>"
	  					end if
    				else
	    				if rs(1)=Item then
			   				response.write "<option value='" & rs(1) & "' selected>" & Server.HTMLEnCode(rs(1) & "：" & rs(2)) & "</option>"
				   		else
		   					response.write "<option value='" & rs(1) & "'>" & Server.HTMLEnCode(rs(1) & "：" & rs(2)) & "</option>"
	  					end if
	  				end if
	    			rs.movenext
    			wend	
    			rs.close
			%>        
    	</select><br>
    </td>
</tr>
<tr>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;值：</td>
    <td>
		<input type="text" name="txtItem" id="txtItem" value="<%=Item%>" size="30"><br><br>
	</td>
</tr>
<tr>
	<td>內容：</td>
	<td><input type="text" name="txtContent" id="txtContent" value="<%=Content%>" size="30"><br><br></td>
</tr>
<tr>
	<td colspan="2">	
	  <p align="center">       
		  <input type="button" value="新增" onclick="form.DBaction.value='a';form.submit();">　  
		  <input type="button" value="修改" onclick="form.DBaction.value='u';form.submit();">    　  
		  <input type="button" value="刪除" onclick="cmdDel_onClick();"> 
	  </p>
	</td>
</tr>
</table>
<input type="hidden" name="DBaction">
</form>
<p align="center"><table>
    <tr><td><font size="2" color="green">1.收到負責人更新授權清冊時，請順便更新此處的授權名單及授權時間</font></td></tr>
    <tr><td><font size="2" color="green">2.授權時間：檢查授權人員名單.xls是否更新之時間比較點</font></td></tr>
    <tr><td><font size="2" color="green">3.授權等級之比對，目前只做到不同公司同人名，不同課或同公司同人名尚無</font></td></tr>
    <tr><td><font size="2" color="green">5.授權人員的內容欄位，目前僅供參考，沒有實際用途</font></td></tr>
    <tr><td><font size="2" color="green">6.常駐人員未列入時數統計，排除名單不列入快速登入</font></td></tr>    
    <tr><td><font size="2" color="green">7.某些特殊功能的目的與物品有與程式連動，若不確定，請先洽SSM小組</font></td></tr>
    
</table></p>
</body>
</html>
<script language="JavaScript"> 
function selItem_onClick(ee) {	//Item(Content)：order
	var e=document.getElementById("form") ;
	e.elements["txtItem"].value=ee.value ;
	txt=ee.options[ee.selectedIndex].text;		
	if(txt.indexOf("：")>0) e.elements["txtContent"].value=txt.substr(txt.indexOf("：")+1) ;
}

function cmdDel_onClick() {	
	with(document.getElementById("form")) {
		var e=elements["selItem"]
		if(confirm("確定刪除[" + e.options[e.selectedIndex].text + "]?")) {
			DBaction.value="d";
			submit();
		}
	}		
}
</script>