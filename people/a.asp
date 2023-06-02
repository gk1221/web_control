<!-- #include file="..\..\connDB\conn.ini" -->
<!-- #include file="..\..\Lib\GetID.inc" -->
<!-- #include file="..\..\Lib\DT.inc" -->
<%
set condb=Server.CreateObject("adodb.connection")
set connIDMS=Server.CreateObject("adodb.connection")
condb.Open strControl
connIDMS.Open strIDMS
set recset=server.createobject("adodb.recordset")
set rs=server.createobject("adodb.recordset")

rs.open "select Item from Config where Kind='資訊中心'",connIDMS
while not rs.eof
  strMIC=strMIC & rs(0) & ","
  rs.movenext
wend
rs.close

YYYY=trim(request("YYYY")) : if YYYY="" then YYYY=DT(now-30,"yyyy")
MM=trim(request("MM"))  : if MM="" then MM=DT(now-30,"mm")

SQL=replace(request("SQL"),"$","%")
if SQL="" then
  strYM=trim(YYYY) & "年" & trim(MM) & "月"
  SQL="select * from people where 進入日期 like '"&trim(YYYY)&"/"&trim(MM)&"/%' order by 申請單位,目的"
else
  strYM=trim(request("selYYYY")) & "年" & trim(request("selMM")) & "月"
  SQL="select * from people where " & SQL & " order by 申請單位,目的"
end if

Function CleanHw(strPurposes)
  CleanHw="" : dim PurposeA,i : PurposeA=split(strPurposes,",")
  for i=0 to ubound(PurposeA)
    if instr(1,PurposeA(i),"負責人")=0 then CleanHw=CleanHw & PurposeA(i) & ","
  next

  if CleanHw<>"" then CleanHw=mid(CleanHw,1,len(CleanHw)-1)
End Function

dim A(100)  '廠商
dim B(100,20)  '目的
dim C(100,20)  '時間
dim D(100)  '總計時間
dim E(100)  '總計人次
for i=0 to 99
  A(i)=""
next
x=0  '廠商個數
dim y(100)  '同廠商目的個數

vcnt=0 : no=""
'vresident="HP--阮榮梃  衛道科技--闕壯文  IBM--鄭智仁.李長華"  '2008.03
recset.open "select * from Config where Kind='常駐人員' order by Content",condb
while not recset.eof
  if Company<>recset(2) then
    vresident=vresident & "  " & recset(2) & "--" & recset(1)
    Company=recset(2)
  else
    vresident=vresident & "." & recset(1)
  end if
  recset.movenext
wend
recset.close

dim hiddenPurpose(9)  ' 2016.11.5 請將要隱藏的目的加到此array
hiddenPurpose(0)="宣導資安政策門禁規定"
hiddenPurpose(1)="A類人員"
hiddenPurpose(2)="B類人員"
hiddenPurpose(3)="C類人員"
hiddenPurpose(4)="依規定登入"
hiddenPurpose(5)="其它登入"
on error resume next
recset.open SQL,condb
do while not recset.eof
  if trim(recset("離開日期"))="" or trim(recset("離開時間"))="" then
    no="y"
    exit do
  end if

  ' ----------- 2008.03 ------------
  if InStr(1,vresident,trim(recset("申請人")),1)= 0 then
    vcnt=vcnt+1

    ' 2016.11.5 篩選要隱藏的目的
    vpurpose=recset("目的")
    'for each item in hiddenPurpose
    '  if(item<>"") then
    '    if(instr(vpurpose,item)=1) then  ' 未勾選目的
    '      vpurpose=replace(vpurpose,item,"")
    '    else
    '      vpurpose=replace(vpurpose,", "&item,"")
    '    end if
    '  end if
    'next
    
    for each item in hiddenPurpose
      itemPosition=instr(vpurpose,item)
      if(item<>"" and itemPosition<>0) then
        if(itemPosition=1) then  ' 未勾選目的
          if(item=hiddenPurpose(4) or item=hiddenPurpose(5)) then
            vpurpose=left(vpurpose,0)
          else
            vpurpose=replace(vpurpose,item,"")
          end if
        else
          if(item=hiddenPurpose(4) or item=hiddenPurpose(5)) then
            vpurpose=left(vpurpose,itemPosition-3)
          else
            vpurpose=replace(vpurpose,", "&item,"")
          end if
        end if
      end if
    next

    if(trim(vpurpose)="") then vpurpose="NULL"  ' 未勾選目的
    vpurpose=split(CleanHw(vpurpose),",")  '分解目的內容
    ' 2016.11.5 end

    vtime1=cdate(trim(recset("進入日期"))&" "&trim(recset("進入時間")) )  '-- 文字轉日期
    vtime2=cdate(trim(recset("離開日期"))&" "&trim(recset("離開時間")) )

    if err.Number>0 then 
      response.write recset("進入日期") & " " & recset("進入時間") & ","
      err.Number=0
    end if
  end if
  ' ----------------------------------------------------
  recset.movenext
loop
%>
<html>
<head>
  <meta charset="big5">
  <title><%=strYM%>機房進出時間月統計表</title>
</head>
<body>
  <div align="center">
    <table border="0" cellspacing="1" width="90%">
      <tr>
        <td width="100%">
          <p align="center" style="line-height: 200%; margin-top: 0; margin-bottom: 0">
          <b><font color="#800000" size="6"><%=strYM%></font></b>
          <b><font color="#800000" size="6">機房進出時間(人次)月統計表</font></b>
        </td>
      </tr>
    </table>
  </div>
  <p style="line-height: 100%; margin-top: 0; margin-bottom: 0">　</p>
  <div align="center">
    <center>
      <table border="1" cellspacing="0" bordercolorlight="#993300" width="100%" height="1">
        <tr>
          <td align="center" width="20%" bgcolor="#FFBFBF">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0"><b><font size="4" color="#800000">廠商</font></b>
          </td>
          <td align="center" width="50" bgcolor="#FFBFBF">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0"><b><font size="4" color="#800000">人次</font></b>
          </td>
          <td align="center" width="110" bgcolor="#FFBFBF">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0"><b><font size="4" color="#800000">時數總計</font></b>
          </td>
          <td align="center" bgcolor="#FFBFBF">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0"><b><font size="4" color="#800000">目的</font></b>
          </td>
        </tr>

        <tr>
          <td height="16" align="left" bgcolor="<%=vbgcolor%>">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0; margin-left: 10; margin-right: 10">
            <font color="#0000FF"><b><%response.write A(i)%></b></font>
          </td>
          <td height="16" align="right" bgcolor="<%=vbgcolor%>">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0; margin-left: 10; margin-right: 15">
            <font color="#0000FF"><%response.write vamount%></font>
          </td>
          <td height="16" align="right" bgcolor="<%=vbgcolor%>">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0; margin-left: 10; margin-right: 10">
            <font color="#0000FF"><b><% response.write vtotal%></b></font>
          </td>
          <td height="16" align="left" bgcolor="<%=vbgcolor%>">
            <p style="line-height: 150%; margin-top: 0; margin-bottom: 0; margin-left: 10; margin-right: 10">
            <font color="#0000FF"><%=vpurposeList%></font>
          </td>
        </tr>

      </table>
      <br />
      <font color="blue">
        總人次：<%=TotalPeople%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        總時數：<%=(TotalTime\1440) & "日" & ((TotalTime mod 1440)\60) & "小時" & (TotalTime mod 60) & "分" %>
      </font>
      <br />
      <br />
      <font color="red"> ** 未列入時數統計之常駐人員：本中心人員 <%=vresident%></font>
    </center>
  </div>
</body>
</html>
<%
if vcnt=0 or no="y" then
  if vcnt=0 then 
    response.write "無列印資料　！！"
  else 
    response.write "尚有人員未登出，請登出後再執行　！！"
  end if
end if
%>