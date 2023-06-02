<%@ Page Language="C#" Title="待完成清單" AutoEventWireup="true"
CodeFile="check.aspx.cs" Inherits="Check_list" Trace="false"
MaintainScrollPositionOnPostback="true" Debug="true" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head runat="server">
    <link rel="stylesheet" href="/control/device/css/bootstrap.min.css" />

    <script src="/control/device/js/DataTables/jQuery-3.6.0/jquery-3.6.0.min.js"></script>
    <script src="/control/device/js/bootstrap.bundle.min.js "></script>
    <style type="text/css">
        .form-h input[type="text"], .form-h span {
            min-height: 40px!important;
        }
        .w-60{
            width: 60%;
        }
        .w-10{
            width: 1%!important;
        }
        .mydisable{
            color: #9d9d9d;
        }
    </style>
    <script>
      var data = '<%=Get_Json()%>';
      data = Array(JSON.parse(data))[0];
 
      var c = data[2];

      function DateDiffNow(tar) {
        const now = new Date();
        const tar_time = new Date(tar);
        var M = tar_time.getTime() - now.getTime();
        var m = Math.floor(M / 1000 / 60);
        var H = Math.floor(M / 1000 / 60 / 60);
        var D = Math.floor(H / 24);
        return [
          { day: D ,
           hour: H % 24 ,
           minute: m % 60 ,
           timestamps: M },
        ];
      }
     

      function FormateDate(deadline) {
        const datelist = DateDiffNow(deadline)[0];
        //console.log(datelist);
        var ans = "";
        if (datelist.timestamps < 0) {
          ans = `已經過了  ${-datelist.day}天  ${-datelist.hour}小時`;
        } else {
          ans = `剩下  ${datelist.day}天  ${datelist.hour}小時`;
        }
        return ans;
      }

      function FormateTF(check){
        if( check === 'True'){
            return ["已完成", "mydisable"]

        }
        else{
            return ["尚未完成", ""]
        }
      }
      var d = FormateTF(c.check)
      var ax = ""
      $(document).ready(function () {
        const checklist = $("#checklist");
        
       
        data.forEach((c,index) => {
            var TF = FormateTF(c.check)
            //console.log(TF)
            checklist.append(`<a href="#" class="list-group-item list-group-item-action ${TF[1]} ${index==0?'active':''}" >
                    <div class="d-flex w-100 justify-content-between">
                        <h3 class="mb-1" aria-label="title">${c.title}</h3>
                        <small class="visually-hidden" >${c.ID}</small>
                        <small>${FormateDate(c.deadline)}</small>
                    </div>
                    
                    <h5 class="mb-1 d-inline " aria-label="context">  ${c.context}</h5>
                    <div class="d-flex justify-content-between">
                    <small>${TF[0]} </small>
                    <small aria-label="deadline">${c.deadline}</small>
                    </div>
                    
                </a>`);
        });
         ax = $(".list-group-item-action")
        $(".list-group-item-action").click(function() {
            ax = $(this)
            
            $(":not(this)").removeClass("active")
            ax.addClass("active")
            var id = ax.children()[0].children[1].innerText
            var tit = ax.children()[0].children[0].innerText
            var cont = ax.children()[1].innerText
            var dline = ax.children()[2].children[1].innerText
            var tf = ax.children()[2].children[0].innerText
                console.log(tf.length==3)
            $('[aria-label="ID"]').val(id)
            $('[aria-label="titlebox"]').val(tit)
            $('[aria-label="contextbox"]').val(cont)
            $('[aria-label="deadlinebox"]').val(dline)
            $('[aria-label="tf"]').val(String( tf.length==3))
            
            console.log(tit, cont, dline)

        })

        
      });
        
    </script>
  </head>

  <body>
    <form runat="server">
        <div class="d-flex flex-column align-items-center ">
            <div class="list-group w-60 mt-lg-3" id="checklist" role="tablist">
              
            </div>
            <asp:TextBox  ID="checkID" aria-label="ID" runat="server" class="invisible"></asp:TextBox>
            <div class="w-100 mt-lg-5 d-flex align-content-center justify-content-evenly mx-lg-5">
              <div class="input-group form-h w-75">
                  <span class="input-group-text ">標題</span>
                  <asp:TextBox  class="form-control " aria-label="titlebox"  ID="Titlea" runat="server"  />
                 <span class="input-group-text">內容</span>
                  <asp:TextBox runat="server" aria-label="contextbox" ID="context" class="form-control w-20" />
                  <label class="input-group-text" for="inputGroupSelect01">完成</label>
                  <asp:DropDownList runat="server" id="TF" class="  form-select  w-10" aria-label="tf" >
                
                      <asp:ListItem Value="true" selected> 是 </asp:ListItem>
                      <asp:ListItem Value="false"> 否 </asp:ListItem>
                      
                  </asp:DropDownList>
                  <span class="input-group-text">截止日期</span>
                  <asp:TextBox runat="server" aria-label="deadlinebox" ID="deadline" class="form-control " />
       
              </div>
              <div>
                <asp:button runat="server" id="Btn_edit" class="btn btn-warning " onclientclick="return confirm('Sure?')" onclick="Btn_addclick" text="送出"></asp:button>
        
                <asp:button runat="server" id="Btn_add" class="btn btn-success " onclientclick="return confirm('Sure?')" onclick="Btn_editclick" text="編輯"></asp:button>
                <asp:button runat="server" id="Btn_delete" class="btn btn-danger " onclientclick="return confirm('Sure?')" onclick="Btn_deleteclick" text="刪除"></asp:button>
                
              </div>
            </div>
      
            
          </div>
         
    </form>
    

   
  </body>
</html>
