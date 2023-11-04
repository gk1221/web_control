<%@ Page Language="C#" Title="人員權限編輯" AutoEventWireup="true"
CodeFile="access.aspx.cs" Inherits="Control_PeoEdit"
MaintainScrollPositionOnPostback="true" Debug="false" Trace="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head runat="server">
    <title>人員權限編輯</title>
    <style type="text/css">
      body {
        margin: 0 10px 0 10px !important;
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
        color: #ff0000;
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
        color: #cc6600;
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

      .tbl {
        width: 98%;
      }

      .off4width {
        width: 300px !important;
      }
    </style>
    <link rel="stylesheet" href="/control/device/css/bootstrap.min.css" />
    <script
      type="text/javascript"
      src="/control/device/js/jquery-3.6.0.min.js"
    ></script>

    <script
      type="text/javascript"
      src="/control/device/js/bootstrap.bundle.min.js"
    ></script>
    <meta http-equiv="X-UA-Compatible" content="ie=edge" />
    <script
      type="text/javascript"
      src="/control/device/JsDoMenu/jsmap.js "
    ></script>

    <script>
      //人員資料從後端傳入
      var rec = <%=pre_staff()%>]; //[class, name]
      var com = <%=pre_company()%>]; //[class, name]
      var nam = <%=pre_name()%>]; //[{}]
      var key = new Set(); //class
      rec.forEach(element => {
          key.add(element[0]);
      });

      $(document).ready(function() {
          //放入各課至參考選單
          rec.forEach(element => {
              $("#staffchos5").append('<option class="staffopt" value="' + element[0] + '">' + element[0] + '</option>');
          })
          com.forEach(element => {
              $("#staffchos4").append('<option class="staffopt" value="' + element[0] + '">' + element[0] + '</option>');
          })
          nam.forEach(element => {
              let cons = `${element[0]}`;
              if (element[1] != "") cons += ` -- ${element[1]}`
              console.log(cons)
              $("#staffchos1").append(`<option class="staffopt" value="${element[0]}d">` + cons + '</option>');
          })

          $('.staffopt').dblclick(function() { //針對參考選單的雙擊型為
              var baba = $(this).parent().parent().parent().parent(); //option ->select->dropdown->body->offcanvas
              baba[0].getElementsByClassName('btn')[0].click(); //offcanvas->buttonOK
          });

      })

      function staffok(id, select, modal) {
          $(id).val($(select).val());
          $(modal).offcanvas('hide');
      }
    </script>
  </head>

  <body style="background-color: #dfefff; color: #0000ff">
    <p class="title">人員權限資料建檔</p>
    <form id="form1" runat="server">
      <asp:TextBox
        ID="ID"
        visible="False"
        runat="server"
        BackColor="Silver"
        ReadOnly="True"
      />

      <asp:Table class="tbl ms-3" ID="Table1" runat="server" GridLines="Both">
        <asp:TableRow ID="rowTitle0" runat="server">
          <asp:TableCell runat="server" Width="105px">人員資料</asp:TableCell>
          <asp:TableCell runat="server" Width="383px">
            <div class="d-flex form-group mx-2 my-2 align-content-center">
              <asp:TextBox
                class="form-control"
                ID="Name"
                size="15"
                runat="server"
                ForeColor="black"
                placeholder="人員名稱"
              ></asp:TextBox>
              <asp:TextBox
                class="form-control mx-2"
                ID="NameID"
                size="15"
                runat="server"
                ForeColor="black"
                placeholder="人員編號 (可空白)"
              ></asp:TextBox>
              <a
                class="btn btn-sm btn-secondary"
                data-bs-toggle="offcanvas"
                href="#offcanvasExample1"
                role="button"
                aria-controls="offcanvasExample"
              >
                參考選單
              </a>
              <div
                class="offcanvas offcanvas-end off4width"
                tabindex="-1"
                id="offcanvasExample1"
                aria-labelledby="offcanvasExampleLabel"
              >
                <div class="offcanvas-header">
                  <h5 class="offcanvas-title" id="offcanvasExampleLabel">
                    代入人員
                  </h5>
                </div>
                <div class="offcanvas-body">
                  <div class="div-mod mt-2">
                    <asp:DropDownList
                      id="MenuName"
                      runat="server"
                      AutoPostBack="True"
                      class="h-100 form-select"
                      multiple
                      aria-label="multiple select example"
                      OnSelectedIndexChanged="MenuName_SelectedIndexChanged"
                    ></asp:DropDownList>
                  </div>
                </div>
              </div>
            </div>
          </asp:TableCell>

          <asp:TableCell runat="server" Width="105px">電話/分機</asp:TableCell>
          <asp:TableCell runat="server" Width="383px">
            <div class="d-flex form-group mx-2 my-2">
              <asp:TextBox
                class="form-control w-50"
                ID="Telephone"
                size="40"
                runat="server"
                ForeColor="black"
              ></asp:TextBox>
            </div>
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle1" runat="server">
          <asp:TableCell runat="server" Width="105px"> 單位課別 </asp:TableCell>
          <asp:TableCell runat="server" Width="372px">
            <div class="d-flex form-group my-2 w-30 align-items-center">
              <asp:TextBox
                class="mx-2 form-control form-control-sm"
                ID="Dept"
                size="20"
                runat="server"
                ForeColor="black"
              ></asp:TextBox>
              <a
                class="btn btn-sm btn-secondary"
                data-bs-toggle="offcanvas"
                href="#offcanvasExample5"
                role="button"
                aria-controls="offcanvasExample"
              >
                參考選單
              </a>
              <div
                class="offcanvas offcanvas-start"
                tabindex="-1"
                id="offcanvasExample5"
                aria-labelledby="offcanvasExampleLabel"
              >
                <div class="offcanvas-header">
                  <h5 class="offcanvas-title" id="offcanvasExampleLabel">
                    單位課別
                  </h5>
                </div>
                <div class="offcanvas-body">
                  <div class="dropdown mt-3">
                    <select
                      size="20"
                      class="mt-3 form-select staffchos"
                      id="staffchos5"
                    ></select>
                  </div>
                </div>
                <div class="modal-footer">
                  <button
                    type="button"
                    class="btn btn-primary"
                    onclick="staffok('#Dept', '#staffchos5' , '#offcanvasExample5')"
                  >
                    確認
                  </button>
                </div>
              </div>
              <span class="psword psoan mx-2"> (亦可自行填入)</span>
            </div>
          </asp:TableCell>

          <asp:TableCell runat="server" Width="113">設備廠商</asp:TableCell>
          <asp:TableCell runat="server" Width="372">
            <div class="d-flex form-group my-2 w-30 align-items-center">
              <asp:TextBox
                class="mx-2 form-control form-control-sm"
                ID="Company"
                size="20"
                runat="server"
                ForeColor="black"
                placeholder="EX：富士通 (可空白)"
              />
              <a
                class="btn btn-sm btn-secondary"
                data-bs-toggle="offcanvas"
                href="#offcanvasExample4"
                role="button"
                aria-controls="offcanvasExample"
              >
                參考選單
              </a>
              <div
                class="offcanvas offcanvas-end off4width"
                tabindex="-1"
                id="offcanvasExample4"
                aria-labelledby="offcanvasExampleLabel"
              >
                <div class="offcanvas-header">
                  <h5 class="offcanvas-title" id="offcanvasExampleLabel">
                    設備廠商
                  </h5>
                </div>
                <div class="offcanvas-body">
                  <div class="dropdown mt-3">
                    <select
                      size="20"
                      class="mt-3 form-select staffchos"
                      id="staffchos4"
                    ></select>
                  </div>
                </div>
                <div class="modal-footer">
                  <button
                    type="button"
                    class="btn btn-primary"
                    onclick="staffok('#Company', '#staffchos4' , '#offcanvasExample4')"
                  >
                    確認
                  </button>
                </div>
              </div>
              <span class="psword psoan mx-2"> (亦可自行填入)</span>
            </div>
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle2" runat="server">
          <asp:TableCell runat="server" Width="113"> 1F權限 </asp:TableCell>
          <asp:TableCell runat="server" Width="372">
            <div class="mx-2 my-2">
              <asp:RadioButton
                id="Radio1"
                GroupName="RadioGroup1"
                runat="server"
              /><span style="color: red">●</span>&nbsp;
              <asp:RadioButton
                id="Radio2"
                GroupName="RadioGroup1"
                runat="server"
              /><span style="color: red">Ｘ</span>&nbsp;
            </div>
          </asp:TableCell>
          <asp:TableCell runat="server" Width="113"> 2F權限 </asp:TableCell>
          <asp:TableCell runat="server" Width="372">
            <div class="mx-2 my-2">
              <asp:RadioButton
                id="Radio3"
                GroupName="RadioGroup2"
                runat="server"
              /><span style="color: red">●</span>&nbsp;
              <asp:RadioButton
                id="Radio4"
                GroupName="RadioGroup2"
                runat="server"
              /><span style="color: red">Ｘ</span>&nbsp;
            </div>
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle3" runat="server">
          <asp:TableCell runat="server" Width="105px">業務名稱</asp:TableCell>
          <asp:TableCell runat="server" Width="383px" colspan="4">
            <div class="d-flex form-group mx-2 my-2">
              <asp:TextBox
                class="form-control w-50"
                ID="Job"
                size="80"
                runat="server"
                ForeColor="black"
              ></asp:TextBox>
            </div>
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle4" runat="server">
          <asp:TableCell runat="server" Width="105px">授權方式</asp:TableCell>

          <asp:TableCell runat="server" Width="383px" colspan="4">
            <div class="d-flex form-group mx-2 my-2 flex-column">
              <div>
                <asp:RadioButton
                  id="Radio5"
                  GroupName="RadioGroup3"
                  runat="server"
                />
                <span style="color: rgb(132, 2, 2)">A類</span>
                <span class="psword psoan mx-2">不必通知、不需陪同</span>
              </div>
              <div>
                <asp:RadioButton
                  id="Radio6"
                  GroupName="RadioGroup3"
                  runat="server"
                />
                <span style="color: rgb(132, 2, 2)">B類</span>
                <span class="psword psoan mx-2">需要通知、不需陪同</span>
              </div>
              <div>
                <asp:RadioButton
                  id="Radio7"
                  GroupName="RadioGroup3"
                  runat="server"
                />
                <span style="color: rgb(132, 2, 2)">C類</span>
                <span class="psword psoan mx-2">需要通知且需陪同進入</span>
              </div>
            </div>
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle6" runat="server">
          <asp:TableCell runat="server" Width="113px">表單日期 </asp:TableCell>
          <asp:TableCell runat="server" Width="372px">
            <div class="d-flex form-group my-2 w-30 align-items-center">
              <asp:TextBox
                class="mx-2 form-control form-control-sm w-50"
                ID="Apdate"
                size="40"
                runat="server"
                ForeColor="black"
                placeholder="EX：112/2/20 (可空白)"
              />
            </div>
          </asp:TableCell>

          <asp:TableCell runat="server" Width="113">表單號碼 </asp:TableCell>
          <asp:TableCell runat="server" Width="383px">
            <asp:TextBox
              class="mx-2 form-control form-control-sm w-50"
              ID="Apnumber"
              size="40"
              runat="server"
              ForeColor="black"
              placeholder="EX：SOS05-2023-0007 (可空白)"
            />
          </asp:TableCell>
        </asp:TableRow>

        <asp:TableRow ID="rowTitle7" runat="server">
          <asp:TableCell runat="server" Width="113">修改日期</asp:TableCell>

          <asp:TableCell runat="server" Width="383px" colspan="3">
            <asp:Label
              class="mx-2"
              ID="UpdateDate"
              runat="server"
              style="font-weight: bold; color: black"
            >
            </asp:Label>
          </asp:TableCell>
        </asp:TableRow>
      </asp:Table>

      <div class="d-flex justify-content-between w-100">
        <div class="psdiv mx-5"></div>
        <div class="d-flex">
          <asp:Button
            ID="BtnAdd"
            runat="server"
            Text="　新　增　"
            onclick="BtnAdd_Click"
            class="btn btn-primary mx-2"
            OnClientClick="return confirm('您確定要新增這筆資料嗎？')"
          /><br />
          <asp:Button
            ID="BtnEdit"
            runat="server"
            Text="　修　改　"
            onclick="BtnEdit_Click"
            class="btn btn-warning mx-2"
            OnClientClick="return confirm('您確定要修改這筆資料嗎？\n\n請確認此筆修改與權限申請單內容是否相符!')"
          /><br />
          <asp:Button
            ID="BtnDel"
            runat="server"
            Text="　刪　除　"
            onclick="BtnDel_Click"
            CausesValidation="False"
            class="btn btn-danger mx-2"
            OnClientClick="return confirm('您確定要刪除這筆資料嗎？')"
          />
        </div>
      </div>
    </form>
  </body>
</html>
