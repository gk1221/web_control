<!-- #include file="..\Lib\GetClientIO.inc" -->
<html>
<head>
	<title>門禁管制</title>
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
</head>
<body style="background:#6699ff url('images/head.png') no-repeat;">
<p align="right">
<table id="tblHead" border="0" cellpadding="0" cellspacing="0">
  <tr height="40">
	<td align="right" style="vertical-align:bottom;font-family:標楷體;font-size:22px">
	<u><div id="menu1" style="cursor:hand; margin-left:5px; margin-right:10px" onmousedown="richDropDown(0,20,100,227,menu1,oContextHTML1)">人員進出</div></u>	
	<DIV id="oContextHTML1" style="display:none;">		
  		<DIV style="LEFT: 50px; position:absolute; top:0; left:0; overflow: scroll; overflow-x: hidden; width: 100px; height:227px;  scrollbar-base-color:PINK; border-bottom:2px solid block; SCROLLBAR-HIGHLIGHT-COLOR: block; SCROLLBAR-ARROW-COLOR: block;">
    		<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='people/rapidly.asp'">快速登入</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='people/people.asp?IO=I'">人員登入</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='people/today.asp'">今日紀錄</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='people/history.asp'">歷史查詢</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='Config/Config.asp?which=people'">系統設定</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='/GF2000SYS'">刷卡系統</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='file://10.6.1.11/d/機房行政/門禁資料/授權人員名單.xlsx'">授權清冊</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='file://10.6.1.11/d/機房行政/門禁資料/進場名單_revised.xls'">進場名單</SPAN>
    		</DIV>
    		<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='file://10.6.1.11/d/機房行政/門禁資料/照片檔.xls'">照片檔</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='home_transfer.htm'">快捷功能</SPAN>
    		</DIV>
  		</DIV>
	</DIV>
	</td>

	<td align="right" style="vertical-align:bottom;font-family:標楷體;font-size:22px">
	<u><div id="menu2" style="cursor:hand; margin-left:5px; margin-right:10px" onmousedown="richDropDown(100,20,100,102,menu1,oContextHTML2)">設備異動</div></u>	
	<DIV id="oContextHTML2" style="display:none;">
  		<DIV style="LEFT: 50px; position:absolute; top:0; left:0; overflow: scroll; overflow-x: hidden; width: 100px; height:102px;  scrollbar-base-color:PINK; border-bottom:2px solid block; SCROLLBAR-HIGHLIGHT-COLOR: block; SCROLLBAR-ARROW-COLOR: block;">
    		<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='device/newdev.aspx'">資料建檔</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='device/List.asp'">每月查詢</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='device/search.asp'">進階查詢</SPAN>
    		</DIV>
			<DIV onmouseover="this.style.filter='progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr=pink, EndColorStr=#FFFFFF)';" onmouseout="this.style.filter='';" style="font-family:verdana; font-size:70%; height:25px; background:#E5FFFF; border:1px solid black; padding-left:20px;  cursor:hand; filter:; padding-right:3px; padding-top:3px; padding-bottom:3px">
      			<SPAN onclick="parent.parent.body.location.href='Config/Config.asp?which=device'">系統設定</SPAN>
    		</DIV>
  		</DIV>
	</DIV>
	</td>
  </tr>
</table>
</p>
</body>
</html>
<script>
var oPopup = window.createPopup();
function richDropDown(x1,x2,x3,x4,objMenu,objContex) {
  oPopup.document.body.innerHTML = objContex.innerHTML; 
  oPopup.show(x1,x2,x3,x4,objMenu);
  //alert(x1);
}
</script>