<!----------used in bigplot.asp,pos.asp,map.asp------------------------------------------>
<DIV ID='Lx' style='z-index:1;color:red;font-weight:bold;position:absolute'>��</DIV>
<DIV ID='Rx' style='Z-index:1;color:red;font-weight:bold;position:absolute'>��</DIV>
<DIV ID='Ty' style='Z-index:1;color:red;position:absolute'>|</DIV>
<DIV ID='By' style='Z-index:1;color:red;position:absolute'>|</DIV>

<script language="vbscript">
'��(��):�e(��):�C���(pixel):�C��e(pixel):�]�t���,�q(��:�W)�}�l��
Xno=103 : Yno=32 : Xw=0 : Yh=0 : Xstart=1 : Ystart=0

Sub bigplot_onMouseMove()	'���Ы��w	
	Lx.style.Left=bigplot.style.posLeft
	Lx.style.Top=window.event.y-8
	Rx.style.Left=bigplot.style.posLeft+bigplot.width-15
	Rx.style.Top=window.event.y-8
	Ty.style.Left=window.event.x-4
	Ty.style.Top=bigplot.style.posTop
	By.style.Left=window.event.x-4
	By.style.Top=bigplot.style.posTop+bigplot.height-15
End Sub

Sub Mach_Move(LeftX,TopY)	'���ϻP�����N�w��
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

Sub bigplot_click()	'�w�����ܷӤ�
	call bigplot_onClick()	'�w��,�覡�]�t�ΦӲ�(different in map.asp(2) and bigplot.asp(1) and pos.asp(0)
	if window.event.shiftkey or window.event.ctrlkey or window.event.altkey then
		'*******tagName�bclient�ݤ�����w��********************************
		if mid(window.event.srcElement.alt,1,1)="(" then '��alt�O�]�w��ϥ�tag�L�k������
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