<!----------------------------- used in querym.asp,hostmodify.asp,search.asp,device.asp ----------------------------->
<!-- #include file="MapNo.inc" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<title>機房設備位置圖</title>
</head>
<body bgcolor="#6699FF" text="#FFFFFF">
  <img ID="bigplot" border="0" src="../images/plot/bigplot.jpg" style="position: absolute; left: 12; top: 1" width="719" height="402">
  <img ID="pos1" tagName="" border="0" src="../images/plot/In.gif"  style="z-index:2;position:absolute;left:-100" width="16" height="16" onClick='bigplot_onClick()' onMouseMove='bigplot_onMouseMove'> 
  <img ID="pos2" tagName="" border="0" src="../images/plot/Out.gif" style="z-index:2;position:absolute;left:-100" width="16" height="16" onClick='bigplot_onClick()' onMouseMove='bigplot_onMouseMove'>  
  <input type="button" id="buttonOK" style="z-index: 2; font-weight: bold; position: absolute; left: 640; top: 30; color: #FF0000; font-size: 18pt; background-color: #FFFF00" value="確定"> 
  <input type="button" id="buttonNO" style="z-index: 2; font-weight: bold; position: absolute; left: 640; top: 80; color: #FF0000; font-size: 18pt; background-color: #FFFF00" value="取消" onClick='Window.close'> 
<script language="vbscript">
bigplot.style.width=Xno*13
bigplot.style.height=Yno*14
Dialogwidth =cstr(Xno*13+10) & "px"	'顯示地圖視窗大小(長)
Dialogheight=cstr(Yno*14+30) & "px"	'顯示視窗大小(寬)
	
Action=mid(Window.DialogArguments,1,1) : XYZ=mid(Window.DialogArguments,2)
X1=mid(XYZ,1,instr(1,XYZ,",")-1)	   : XYZ=mid(XYZ,instr(1,XYZ,",")+1)
Y1=mid(XYZ,1,instr(1,XYZ,",")-1)	   : XYZ=mid(XYZ,instr(1,XYZ,",")+1)
Z1=mid(XYZ,1,instr(1,XYZ,",")-1)	   : XYZ=mid(XYZ,instr(1,XYZ,",")+1)
X2=mid(XYZ,1,instr(1,XYZ,",")-1)	   : XYZ=mid(XYZ,instr(1,XYZ,",")+1)
Y2=mid(XYZ,1,instr(1,XYZ,",")-1)
Z2=mid(XYZ,instr(1,XYZ,",")+1)
select case Action
case "O"
	X1=-100 : Y1=-100 : Z1=-100
case "M"
case else
	X2=-100 : Y2=-100 : Z2=-100
end select

pos1.tagName=X1 & "," & Y1 : pos2.tagName=X2 & "," & Y2

Sub window_onLoad()
	call mach_move(0,0)    '放在window_onload物件才能就定位(物件才產生好)
End Sub
	
Sub buttonOK_onClick() 	
	if (X1>"0" or X2>"0") and (Y1>"0" or Y2>"0") then
		Window.ReturnValue=X1 & "," & Y1 & "," & Z1 & "," & X2 & "," & Y2 & "," & Z2
		Window.close
	else
		msgbox "請定位,謝謝!　",48
	end if
End Sub
InYN=true
Sub bigplot_onClick()	
	select case Action
	case "M" 	'2點(*X)
		if inYN then								
			call plotXY("in")
		else				
			call plotXY("out")			
		end if
		inYN=not inYN
	case "O" 	'移出設備(X)
		inYN=false
		pos1.style.left=-100 : pos1.style.top =-100
		X1=0 : Y1=0 : Z1=0		
		call plotXY("out")		
	case else 	'1點(*)
		inYN=true
		pos2.style.left=-100 : pos2.style.top =-100
   		X2=0 : Y2=0 : Z2=0
		call plotXY("in")		
	end select		
End Sub
Sub plotXY(IO)
	if IO="in" then	'移入
		X1=int(window.event.x/Xw)-Xstart	'int無條件捨去
		Y1=int(window.event.y/Yh)-Ystart	'int無條件捨去
		Z1=0
		X=(X1+Xstart)*Xw+bigplot.offsetleft
		Y=(Y1+Ystart)*Yh+bigplot.offsetTop
		pos1.style.posleft=X	'window.event.x-8
		pos1.style.postop =Y	'window.event.y-8
		if instr(1,window.event.srcElement.tagName,",")>0 then pos1.Alt=window.event.srcelement.alt		
	elseif IO="out" then	'移出
		X2=int(window.event.x/Xw)-Xstart
		Y2=int(window.event.y/Yh)-Ystart
		Z2=0
		X=(X2+Xstart)*Xw+bigplot.offsetleft
		Y=(Y2+Ystart)*Yh+bigplot.offsetTop		
		pos2.style.posleft=X	'window.event.x-8
		pos2.style.postop =Y	'window.event.y-8
		if instr(1,window.event.srcElement.tagName,",")>0 then pos2.Alt=window.event.srcelement.alt	
	end if
end Sub
  </script>	
</body>

</html>