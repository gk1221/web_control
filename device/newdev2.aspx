<%@ Page Language="C#" Title="設備編輯" AutoEventWireup="true" CodeFile="newdev2.aspx.cs" Inherits="Control_DevEdit"
    MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">

        <title>設備編輯</title>
        <style type="text/css">
            body {
                margin: 0 10px 0 10px!important;
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
            
            .div-mod {
                height: 600px;
            }
            
            .psdiv {
                color: #CC6600;
                line-height: 1.2;
            }
            
            table {
                margin: 0px 2px 10px 2px;
                line-height: 20px;
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
        </style>
        <link rel="stylesheet" href="/control/device/css/bootstrap.min.css">
        <script type="text/javascript" src="/control/device/js/jquery-3.6.0.min.js"></script>

        <script type="text/javascript" src="/control/device/js/bootstrap.bundle.min.js"></script>
        <meta http-equiv="X-UA-Compatible" content="ie=edge">
        <script type="text/javascript" src="/control/device/JsDoMenu/jsmap.js "></script>

        <script>
            //人員資料從後端傳入
            var rec = <%=pre_staff()%>]; //[class, name]
            var key = new Set(); //class
            rec.forEach(element => {
                key.add(element[0]);
            });
            $(document).ready(function() {
                //建立課清單
                key.forEach(element => {
                    $(".staffsel").append('<li><a class="dropdown-item" href="#">' + element + '</a></li>');
                });
                //建立選擇課之後的事件
                $('.dropdown-item').click(function() {
                    onselect($(this).text()); //create name list
                    $(".selebot").text($(this).text()); //class button change name
                    $('.staffopt').dblclick(function() { //after create == button click
                        var baba = $(this).parent().parent().parent().parent(); //option ->select->dropdown->body->offcanvas
                        baba[0].getElementsByClassName('btn')[1].click(); //offcanvas->buttonOK
                    });
                });
                //$('.dropdown-item')[0].click(); //pre click first item
            });
            //選擇課之後建立該課人員名單
            function onselect(clss) {
                $(".staffchos").empty();
                rec.forEach(element => {
                    if (element[0] == clss) {
                        $(".staffchos").append('<option class="staffopt" value="' + element[1] + '">' + element[1] + '</option>');
                    }
                })
            }
            //選擇人員後將名字填入空格
            function staffok(id, select, modal) {
                if (id == '#OP') { //除了OP為疊加其餘都為更改
                    $(id).val($(id).val() + ' ' + $(select).val());
                } else {
                    $(id).val($(select).val());
                }

                $(modal).offcanvas('hide');
            }
        </script>
    </head>

    <body style="background-color: #DFEFFF;color:#0000FF;">
        <p class="title">設備異動資料建檔</p>
        <form id="form1" runat="server">
            <asp:Table class="tbl" ID="Table1" runat="server" Width="100%" GridLines="Both">


                <asp:TableRow ID="rowTitle0" runat="server">

                    <asp:TableCell runat="server" Width="105px">建檔日期</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="form-group mx-2 my-2">
                            <asp:TextBox class="text-center form-control form-control-sm border-dark id-box" ID="DevID" visible="False" runat="server" BackColor="Silver" Columns="2" Font-Size="Large" ForeColor="Red" ReadOnly="True"></asp:TextBox>
                            <asp:Label ID="CreateDate" runat="server" class="mx-2 fw-bold text-black">
                                {{CreateDate}}}
                            </asp:Label>
                            <a class="btn btn-sm btn-secondary" id="hostmodal" data-bs-toggle="offcanvas" href="#offcanvasExample" role="button" data-target="#offcanvasExample" aria-controls="offcanvasExample">
                        參考選單
                        </a>
                        </div>


                        <!-- Modal -->

                        <div class="offcanvas offcanvas-start w-25" tabindex="-1" id="offcanvasExample" aria-labelledby="offcanvasExampleLabel">
                            <div class="offcanvas-header">
                                <h5 class="offcanvas-title" id="offcanvasExampleLabel">代入資料</h5>
                                <button type="button" class="btn-close text-reset" data-bs-dismiss="offcanvas" aria-label="Close"></button>
                            </div>
                            <div class="offcanvas-body">

                                <div class="div-mod mt-2">
                                    <asp:DropDownList id="MenuHost" runat="server" AutoPostBack="True" class="h-100 form-select" multiple aria-label="multiple select example" OnSelectedIndexChanged="MenuHost_SelectedIndexChanged"></asp:DropDownList>
                                </div>
                            </div>
                        </div>

                        <!-- Button trigger modal -->


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
                        <div class="d-flex">
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
                            <span class="psword psoan">(若為主機，設備名稱欄位請填 hostname )</span>
                        </div>



                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">設備廠商</asp:TableCell>
                    <asp:TableCell runat="server" Width="372">
                        <div class="d-flex form-group mx-2 my-2 w-60">
                            <asp:TextBox class="form-control-sm form-control mx-2" ID="Repair" runat="server" Text="N/A" ForeColor="black"></asp:TextBox>
                            <asp:DropDownList id="MenuRepair" runat="server" AutoPostBack="True" OnSelectedIndexChanged="MenuRepair_SelectedIndexChanged"></asp:DropDownList>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle2" runat="server">
                    <asp:TableCell runat="server" Width="105px">設備名稱 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px" colspan="3">
                        <div class="d-flex form-group mx-2 my-2">
                            <asp:TextBox class=" form-control w-50" ID="HostName" runat="server" ForeColor="black">
                            </asp:TextBox>
                            <span class="psword mx-2 fw-bold " style="color:red"> (請註記IDMS的設備編號)</span>
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
                            <asp:TextBox class=" form-control  w-50" ID="Memo" size=80 runat="server" ForeColor="black">
                            </asp:TextBox>
                            <span class="psword psoan mx-2">(備品、替代、移位、暫放走道、移出維護、移至地下室、報廢...)</span>
                        </div>
                        <span class="psword mx-2">(需/不需)移入單</span>
                        <asp:DropDownList id="PS" runat="server">
                            <asp:ListItem Selected="True" Value=""> </asp:ListItem>
                            <asp:ListItem Value="需填移入單"> 需填移入單 </asp:ListItem>
                            <asp:ListItem Value="已填移入單"> 已填移入單 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－測試"> 不需填移入單－測試 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－暫放"> 不需填移入單－暫放 </asp:ListItem>
                            <asp:ListItem Value="不需填移入單－代管"> 不需填移入單－代管 </asp:ListItem>
                        </asp:DropDownList>
                        <span class="psword psoan">(首次移入請填移入申請單，測試、暫放、代管除外)</span>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle5" runat="server">
                    <asp:TableCell runat="server" Width="105px">申請人員 (<span style="color:red">*</span>)
                    </asp:TableCell>

                    <asp:TableCell runat="server" Width="372px">
                        <div class="d-flex form-group  my-2 w-30 align-items-center">
                            <asp:TextBox class="mx-2 form-control-sm form-control" ID="StaffName" size=20 runat="server" ForeColor="black"></asp:TextBox>
                            <a class="btn btn-sm btn-secondary " data-bs-toggle="offcanvas" href="#offcanvasExample2" role="button" aria-controls="offcanvasExample">
                                參考選單
                            </a>
                            <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample2" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">代入人員</h5>
                                </div>
                                <div class="offcanvas-body">
                                    <div class="dropdown mt-3">
                                        <button class="btn btn-secondary dropdown-toggle selebot" type="button" data-bs-toggle="dropdown">選擇課別</button>
                                        <ul class="dropdown-menu staffsel" aria-labelledby="dropdownMenuButton"></ul>
                                        <select size="20" class="mt-3 form-select staffchos" id="staffchos2"></select>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-primary" onclick="staffok('#StaffName', '#staffchos2' , '#offcanvasExample2')">確認</button>
                                </div>
                            </div>
                            <span class="psword psoan mx-2"> (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>

                    <asp:TableCell runat="server" Width="113"><span size="3">硬體負責人 (<span
                                style="color:red">*</span>)</span><span color="#CC6600" size="2">
                    </asp:TableCell>

                    <asp:TableCell runat="server" Width="372">
                        <div class="d-flex form-group mx-2 my-2 w-30 align-items-center">
                            <asp:TextBox class="mx-2 form-control form-control-sm" ID="HW" size=20 runat="server" ForeColor="black"></asp:TextBox>                            
                            <a class="btn btn-sm btn-secondary " data-bs-toggle="offcanvas" href="#offcanvasExample3" role="button" aria-controls="offcanvasExample">
                                參考選單
                            </a>
                            <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample3" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">代入人員</h5>
                                </div>
                                <div class="offcanvas-body">
                                    <div class="dropdown mt-3">
                                        <button class="btn btn-secondary dropdown-toggle selebot" type="button" id="dropdownMenuButton" data-bs-toggle="dropdown">選擇課別</button>
                                        <ul class="dropdown-menu staffsel" aria-labelledby="dropdownMenuButton"></ul>
                                        <select size="20" class="mt-3 form-select staffchos" id="staffchos3"></select>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-primary" onclick="staffok('#HW', '#staffchos3' , '#offcanvasExample3')">確認</button>
                                </div>
                            </div>
                            <span class="psword psoan mx-2"> (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle6" runat="server">
                    <asp:TableCell runat="server" Width="113px">ＯＰ　　 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="372px">
                        <div class="d-flex form-group my-2 w-30 align-items-center">
                            <asp:TextBox class="mx-2 form-control form-control-sm" ID="OP" size=20 runat="server" ForeColor="black"></asp:TextBox>
                            <a class="btn btn-sm btn-secondary " data-bs-toggle="offcanvas" href="#offcanvasExample5" role="button" aria-controls="offcanvasExample">
                                參考選單
                            </a>
                            <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample5" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">代入人員</h5>
                                </div>
                                <div class="offcanvas-body">
                                    <div class="dropdown mt-3">
                                        <button class="btn btn-secondary dropdown-toggle selebot" type="button" id="dropdownMenuButton" data-bs-toggle="dropdown">選擇課別</button>
                                        <ul class="dropdown-menu staffsel" aria-labelledby="dropdownMenuButton"></ul>
                                        <select size="20" class="mt-3 form-select staffchos" id="staffchos5"></select>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-primary" onclick="staffok('#OP', '#staffchos5' , '#offcanvasExample5')">確認</button>
                                </div>
                            </div>
                            <span class="psword psoan mx-2"> (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">軟體負責人 (<span style="color:red">*</span>)</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group mx-2 my-2 w-30 align-items-center">
                            <asp:TextBox class="mx-2 form-control form-control-sm" ID="SW" size=20 runat="server" ForeColor="black"></asp:TextBox>
                            <a class="btn btn-sm btn-secondary " data-bs-toggle="offcanvas" href="#offcanvasExample4" role="button" aria-controls="offcanvasExample">
                                參考選單
                            </a>
                            <div class="offcanvas offcanvas-start" tabindex="-1" id="offcanvasExample4" aria-labelledby="offcanvasExampleLabel">
                                <div class="offcanvas-header">
                                    <h5 class="offcanvas-title" id="offcanvasExampleLabel">代入人員</h5>
                                </div>
                                <div class="offcanvas-body">
                                    <div class="dropdown mt-3">
                                        <button class="btn btn-secondary dropdown-toggle selebot" type="button" id="dropdownMenuButton" data-bs-toggle="dropdown">選擇課別</button>
                                        <ul class="dropdown-menu staffsel" aria-labelledby="dropdownMenuButton"></ul>
                                        <select size="20" class="mt-3 form-select staffchos" id="staffchos4"></select>
                                    </div>
                                </div>
                                <div class="modal-footer">
                                    <button type="button" class="btn btn-primary" onclick="staffok('#SW', '#staffchos4' , '#offcanvasExample4')">確認</button>
                                </div>
                            </div>
                            <span class="psword psoan mx-2"> (亦可自行填入)</span>
                        </div>
                    </asp:TableCell>
                </asp:TableRow>

                <asp:TableRow ID="rowTitle7" runat="server">
                    <asp:TableCell runat="server" Width="105px" style="text-align: center;">
                        <input type="button" value=" 設備定位(*) " style="font-weight: bold;" class="btn btn-info btn-sm " onclick="map_click()" />
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <div class="d-flex form-group mx-2 my-2">
                            <asp:TextBox class="text-center form-control form-control-sm border-dark id-box" OnKeyDown="event.returnValue=false;" ID="TextDevNo" runat="server" BackColor="Silver" Columns="2" ForeColor="black"></asp:TextBox>
                            <asp:Label class="mx-2 fw-bold text-black" ID="Location" runat="server">尚未定位</asp:Label>
                        </div>
                    </asp:TableCell>
                    <asp:TableCell runat="server" Width="113">修改日期</asp:TableCell>
                    <asp:TableCell runat="server" Width="383px">
                        <asp:Label class="mx-2" ID="UpdateDate" runat="server" style="font-weight: bold;color: black;">
                            {{time}}</asp:Label>
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
                    <p>五、 設備進出機房或移動位置，請到資源管理子系統，更改放置地點欄位，詢問負責人若知地點就更改之，若未知請填寫(任意地點)(未定區域)，重點是它已不在機房內，放置地點不應寫機房內的任何位置。
                    </p>
                    <p>六、 <a target="_blank" href="../主任指示門禁管制.doc">主任指示...</a></p>
                </div>
                <div class="bottondiv">
                    <asp:Button ID="BtnAdd" runat="server" Text="　新　增　" OnClick="BtnAdd_Click" class="btn btn-primary" OnClientClick="return confirm('您確定要新增這筆資料嗎？')" /><br>
                    <asp:Button ID="BtnEdit" runat="server" Text="　修　改　" OnClick="BtnEdit_Click" class="btn btn-warning" OnClientClick="return confirm('您確定要修改這筆資料嗎？')" /><br>
                    <asp:Button ID="BtnDel" runat="server" Text="　刪　除　" CausesValidation="False" OnClick="BtnDel_Click" class="btn btn-danger" OnClientClick="return confirm('您確定要刪除這筆資料嗎？')" />
                </div>
            </div>
        </form>

        <p class="cent">102年12月課務會議決議：機器上機房機架前，先請負責人在IDMS建立設備資訊，並繳交移入單，若無則請其當場填寫！並核對IDMS資料是否正確。</p>




    </body>

    </html>