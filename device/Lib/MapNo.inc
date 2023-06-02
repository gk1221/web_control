<!----------used in bigplot.asp,pos.asp,map.asp------------------------------------------>
<DIV ID='Lx' style='z-index:1;color:red;font-weight:bold;position:absolute'>－</DIV>
<DIV ID='Rx' style='Z-index:1;color:red;font-weight:bold;position:absolute'>－</DIV>
<DIV ID='Ty' style='Z-index:1;color:red;position:absolute'>|</DIV>
<DIV ID='By' style='Z-index:1;color:red;position:absolute'>|</DIV>

<script language="vbscript">
'長(格):寬(格):每格長(pixel):每格寬(pixel):包含邊條,從(左:上)開始算
Xno=103 : Yno=32 : Xw=0 : Yh=0 : Xstart=1 : Ystart=0

Sub bigplot_onMouseMove()	'坐標指針	
	Lx.style.Left=bigplot.style.posLeft
	Lx.style.Top=window.event.y-8
	Rx.style.Left=bigplot.style.posLeft+bigplot.width-15
	Rx.style.Top=window.event.y-8
	Ty.style.Left=window.event.x-4
	Ty.style.Top=bigplot.style.posTop
	By.style.Left=window.event.x-4
	By.style.Top=bigplot.style.posTop+bigplot.height-15
End Sub

Sub Mach_Move(LeftX,TopY)	'底圖與機器就定位
	bigplot.style.left=LeftX
	bigplot.style.top=TopY	
	Xw=bigplot.width/Xno : Yh=bigplot.height/Yno
	for each obj in Document.images
		pos=instr(1,obj.tagName,",")
		if pos>0 then
			Xall=cint(mid(obj.tagName,1,pos-1))
			Yall=cint(mid(obj.tagName,pos+1))
			obj.style.left=LeftX+(Xall+Xstart)*Xw
			obj.style.top=TopY+(Yall+Ystart)*Yh
		end if	
	next
End Sub

Sub bigplot_click()	'定位或顯示照片
	call bigplot_onClick()	'定位,方式因系統而異(different in map.asp(2) and bigplot.asp(1) and pos.asp(0)
	if window.event.shiftkey or window.event.ctrlkey or window.event.altkey then
		'*******tagName在client端不能指定值********************************
		if mid(window.event.srcElement.alt,1,1)="(" then '用alt是因定位圖示tag無法有坐標
			XY=mid(window.event.srcElement.alt,2,instr(1,window.event.srcElement.alt,")")-2)
			Xall=cint(mid(XY,1,instr(1,XY,",")-1))
			Yall=cint(mid(XY,instr(1,XY,",")+1))
			photo="../script/photo.asp?Xall=" & Xall & "&Yall=" & Yall
			window.ShowModalDialog photo,"", "toolbar=no;menubar=no;location=no;help=no;status=no" _
				& ";directories=no;DialogWidth=" & screen.width & ";DialogHeight=" & screen.height
			'window.open photo,"_blank"
		end if
	end if
End Sub
</script>