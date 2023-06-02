<%@ Page Language="C#" Title="設備編輯" AutoEventWireup="true" CodeFile="newdev.aspx.cs" Inherits="Control_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
        <meta http-equiv="X-UA-Compatible" content="IE=11">
        <title>設備編輯</title>
        <style type="text/css">
            body {
                margin: 0 10px 0 10px;
                width: 99%;
            }
            
            p {
                margin-top: 0px;
                margin-bottom: 8px;
            }
            
            .psword {
                font-size: small;
            }
            
            .psoan {
                color: rgb(204, 102, 0);
            }
            
            .cent {
                text-align: center;
                color: #FF0000;
                font-size: 1.2rem;
                font-weight: bold;
            }
            
            .psref {
                text-decoration: underline;
                cursor: pointer;
            }
            
            .bottondiv {
                width: 400px;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
            }
            
            .psdiv {
                color: #CC6600;
                line-height: 1.2;
            }
            
            table {
                margin: 0px 2px 10px 2px;
                line-height: 35px;
            }
            
            tr td {
                border-collapse: collapse;
                border: 1px black solid;
                padding-left: 2px;
            }
            
            .title {
                text-align: center;
                color: rgb(0, 0, 102);
                font-weight: bold;
                font-size: x-large;
                margin-top: 5px;
            }
            
            .min-box {
                width: 170px;
            }
            
            .id-box {
                width: 65px;
            }
            
            .peo-box {
                min-width: 103px;
            }
            
            #lasthost {
                max-width: 383px;
            }
            
            .text-black {
                color: black;
            }
            
            .fw-bold {
                font-weight: 700;
            }
            
            .mx-2 {
                margin-left: 8px;
            }
            
            .btny {
                margin-top: 15px;
            }
        </style>


        <script type="text/javascript" src="/control/device/JsDoMenu/jsmap.js "></script>

    </head>

    <body style="background-color: #DFEFFF;color:#0000FF;">
        <p class="title">設備異動資料建檔或異動介面</p>


        <form id="form1" runat="server">
            <asp:Table class="tbl" ID="Table1" runat="server" Width="100%" GridLines="Both">

                <asp:TableRow ID="rowTitle0" runat="server">

                    <asp:TableCell runat="server" Width="105px">建檔日期</asp:TableCell>
                    <asp:TableCell runat="server" width="383px">
                        <asp:TextBox visible="False" class="text-center form-control form-control-sm border-dark id-box" ID="DevID" runat="server" BackColor="Silver" Columns="2" Font-Size="Large" ForeColor="Red" ReadOnly="True"></asp:TextBox>
                        <asp:Label ID="CreateDate" runat="server" class="mx-2 fw-bold text-black">
                            {{CreateDate}}}
                        </asp:Label>
                        <asp:DropDownList class="mx-2" width="300px" id="MenuHost" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuHost_SelectedIndexChanged"></asp:DropDownList>
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">
                        異動種類 (<span style="color:red">*</span>)
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="372">

                        <asp:RadioButton id="Radio1" Checked="True" GroupName="RadioGroup1" runat="server" /><span span style="color:black;font-weight: bold;">移入</span><span style="color:red">★</span>&nbsp;
                        <asp:RadioButton id="Radio2" GroupName="RadioGroup1" runat="server" /><span span style="color:black;font-weight: bold;">移出</span><span style="color:red">Ｘ</span>&nbsp;
                        <asp:RadioButton id="Radio3" GroupName="RadioGroup1" runat="server" /><span span style="color:black;font-weight: bold;">更換</span><span style="color:red">●</span>&nbsp;
                        <asp:RadioButton id="Radio4" GroupName="RadioGroup1" runat="server" /><span span style="color:black;font-weight: bold;">其它</span><span style="color:red">▲</span>&nbsp;
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle1" runat="server">
                    <asp:TableCell runat="server" Width="105px">設備種類 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <asp:DropDownList class="mx-2" id="HostClass" runat="server">
                            <asp:ListItem Selected="True" Value=""> (請選擇) </asp:ListItem>
                            <asp:ListItem Value="主機"> 主機 </asp:ListItem>
                            <asp:ListItem Value="耗材"> 耗材 </asp:ListItem>
                            <asp:ListItem Value="機櫃"> 機櫃 </asp:ListItem>
                            <asp:ListItem Value="零附件"> 零附件 </asp:ListItem>
                            <asp:ListItem Value="週邊設備"> 週邊設備 </asp:ListItem>
                            <asp:ListItem Value="網路設備"> 網路設備 </asp:ListItem>
                            <asp:ListItem Value="辦公設備"> 辦公設備 </asp:ListItem>
                            <asp:ListItem Value="儲存設備"> 儲存設備 </asp:ListItem>
                            <asp:ListItem Value="其它機器"> 其它機器 </asp:ListItem>

                        </asp:DropDownList>

                        <span class="psword psoan mx-2">(若為主機，設備名稱欄位請填 hostname )</span>
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">設備廠商</asp:TableCell>
                    <asp:TableCell runat="server" Width="372">
                        <div class="d-flex form-group mx-2 my-2 w-60">
                            <asp:TextBox class="form-control-sm form-control" ID="Repair" runat="server" Text="N/A" ForeColor="black"></asp:TextBox>
                            <asp:DropDownList class="mx-2" id="MenuRepair" ForeColor="green" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuRepair_SelectedIndexChanged"></asp:DropDownList>
                        </div>

                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle2" runat="server">
                    <asp:TableCell runat="server" Width="105px">設備名稱 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px" colspan="3">
                        <div class="d-flex form-group mx-2 my-2">
                            <asp:TextBox class=" form-control w-50" ID="HostName" runat="server" ForeColor="black"></asp:TextBox>
                            <span class="psword mx-2 fw-bold " style="color:red">  (請註記IDMS的設備編號)</span>
                        </div>
                    </asp:TableCell>

                </asp:TableRow>

                <asp:TableRow ID="rowTitle3" runat="server">
                    <asp:TableCell runat="server" Width="105px">設備功能</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px" colspan="3">
                        <div class="d-flex form-group mx-2 my-2 ">
                            <asp:TextBox class=" form-control w-50" ID="Functions" size=80 runat="server" ForeColor="black"></asp:TextBox>
                            <span class="psword psoan mx-2">(請填寫設備主要功用，附註請填寫於設備說明)</span>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle4" runat="server">
                    <asp:TableCell runat="server" Width="105px">設備說明</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px" colspan="3">
                        <div class="d-flex form-group mx-2 my-2">
                            <asp:TextBox class=" form-control  w-50" ID="Memo" size=80 runat="server" ForeColor="black"></asp:TextBox>
                            <span class="psword psoan mx-2">(備品、替代、移位、暫放走道、移出維護、移至地下室、報廢...)</span>
                        </div>
                        <span class="psword mx-2">(需/不需)移入單</span>
                        <asp:DropDownList id="PS" runat="server">
                            <asp:ListItem Selected="True" Value=""> </asp:ListItem>
                            <asp:ListItem Value="需填移入單"> 需填移入單 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－測試"> 不需填移入單－測試 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－暫放"> 不需填移入單－暫放 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－代管"> 不需填移入單－代管 </asp:ListItem>

                        </asp:DropDownList>

                        <span class="psword psoan mx-2">(首次移入請填移入申請單，測試、暫放、代管除外)</span>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle5" runat="server">
                    <asp:TableCell runat="server" Width="105px">申請人員 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group  my-2 w-30">
                            <asp:TextBox class="mx-2 form-control-sm form-control" ID="StaffName" size=20 runat="server" ForeColor="black"></asp:TextBox>
                            <asp:DropDownList class="peo-box mx-2" id="Menu_StaffName_1" ForeColor="green" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuSN1_SelectedIndexChanged"></asp:DropDownList>
                            <asp:DropDownList class="peo-box mx-2" ForeColor="green" id="Menu_StaffName_2" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuSN2_SelectedIndexChanged"></asp:DropDownList>
                            <span class="psword psoan">  (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113"><span size="3">硬體負責人 (<span style="color:red">*</span>)</span><span color="#CC6600" size="2"></asp:TableCell>
                    <asp:TableCell runat="server" Width="372">
                        <div class="d-flex form-group  my-2 w-30">
                            <asp:TextBox class="mx-2 form-control form-control-sm" ID="HW" size=20 runat="server"  ForeColor="black"></asp:TextBox>
                            <asp:DropDownList class="peo-box mx-2" id="Menu_HW_1" ForeColor="green" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuHW1_SelectedIndexChanged"></asp:DropDownList>
                            
                            <asp:DropDownList class="peo-box mx-2"  ForeColor="green" id="Menu_HW_2" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuHW2_SelectedIndexChanged"></asp:DropDownList>
                            
                            <span class="psword psoan">  (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle6" runat="server">
                    <asp:TableCell runat="server" Width="105px">ＯＰ　　 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group mx-2 my-2 w-30">
                            <asp:TextBox class="form-control form-control-sm" ID="OP" size=20 runat="server" ForeColor="black"></asp:TextBox>

                        </div>

                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">軟體負責人 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group  my-2 w-30">
                            <asp:TextBox class="mx-2 form-control form-control-sm" ID="SW" size=20 runat="server" ForeColor="black"></asp:TextBox>
                            <asp:DropDownList class="peo-box mx-2" id="Menu_SW_1" ForeColor="green" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuSW1_SelectedIndexChanged"></asp:DropDownList>
                            <asp:DropDownList class="peo-box mx-2" ForeColor="green" id="Menu_SW_2" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuSW2_SelectedIndexChanged"></asp:DropDownList>


                            <span class="psword psoan">  (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle7" runat="server">
                    <asp:TableCell runat="server" Width="105px" style="text-align: center;">
                        <input type="button" value=" 設備定位(*) " style="color: red;font-weight: bold;" onclick="PointerNo=form1.TextDevNo.value;if(PointerNo=='')PointerNo=form1.TextDevNo.value; window.open('../../IDMS/Lib/map.aspx?PointerNo='+PointerNo,'_blank','location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no');console.log('return');"
                        />
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group mx-2 my-2">
                            <asp:TextBox class="text-center form-control form-control-sm border-dark id-box" OnKeyDown="event.returnValue=false;" ID="TextDevNo" runat="server" BackColor="Silver" Columns="2" ForeColor="black"></asp:TextBox>
                            <asp:Label class="mx-2 fw-bold text-black" ID="Location" runat="server">尚未定位</asp:Label>


                        </div>



                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">修改日期</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <asp:Label class="mx-2" ID="UpdateDate" runat="server" style="font-weight: bold;color: black;">{{time}}</asp:Label>
                    </asp:TableCell>
                </asp:TableRow>
            </asp:Table>
            <div style="display: flex;">
                <div class="psdiv">
                    <p>【注意事項】</p>
                    <p>一、註明(* )為必填欄位，其它欄位若空白將以 N/A 取代</p>
                    <p>二、離開前請將空箱子或雜物整理乾淨</p>
                    <p>三、若有「硬體維修」、「軟體更新」或「組件安裝」時, 請主動提醒或詢問「執行者」 是否需要"告知"可能受影響的對象或下游外單位
                    </p>
                    <p>四、 設備進出機房，請記得向負責人詢問設備編號及移入單，並記錄之！</p>
                    <p>五、 設備進出機房或移動位置，請到資源管理子系統，更改放置地點欄位，詢問負責人若知地點就更改之，若未知請填寫(任意地點)(未定區域)，重點是它已不在機房內，放置地點不應寫機房內的任何位置。</p>
                    <p>六、 <a target="_blank" href="../主任指示門禁管制.doc">主任指示...</a></p>
                </div>
                <div class="bottondiv">
                    <asp:Button class="" ID="BtnAdd" runat="server" Text="　新　增　" OnClick="BtnAdd_Click" OnClientClick="return confirm('您確定要新增這筆資料嗎？')" ForeColor="Red" /><br>
                    <asp:Button class="btny" ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" OnClientClick="return confirm('您確定要修改這筆資料嗎？')" /><br>
                    <asp:Button class="btny" ID="BtnDel" runat="server" Text="　刪　除　" CausesValidation="False" OnClick="BtnDel_Click" OnClientClick="return confirm('您確定要刪除這筆資料嗎？')" />
                </div>
            </div>
        </form>

        <p class="cent">102年12月課務會議決議：機器上機房機架前，先請負責人在IDMS建立設備資訊，並繳交移入單，若無則請其當場填寫！並核對IDMS資料是否正確。</p>




    </body>

    </html>