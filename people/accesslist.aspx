<%@ Page Language="C#" Title="權限查詢" AutoEventWireup="true" CodeFile="accesslist.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
        <meta http-equiv="X-UA-Compatible" content="IE=11">
        <title>每月查詢</title>
        <style type="text/css">
            .title {
                text-align: center;
                color: rgb(0, 0, 102);
                font-weight: bold;
                font-size: x-large;
                margin-top: 5px;
                font-family: '微軟正黑體';
            }
            
            .main {
                margin: 5px 10px 0 10px;
                max-width: 98%!important;
            }
            
            .head-div {
                display: flex;
                justify-content: space-between;
            }
            
            .funp {
                overflow: hidden;
                margin-bottom: 0px;
                text-overflow: ellipsis;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
            }
            
            #table_id {
                max-width: 100%!important;
            }
            
            #table_id_filter input {
                border-radius: 5px;
            }
        </style>

        <script>
            function IE_Check() {
                if ("ActiveXObject" in window) {
                    alert("本頁不支援IE，請使用其他瀏覽器");
                    //window.history.go(-1); //返回上一頁
                }
            }
            IE_Check();
        </script>

        <link rel="stylesheet" type="text/css" href="/control/device/js/DataTables/DataTables-1.11.4/css/dataTables.bootstrap5.min.css" />
        <link rel="stylesheet" href="/control/device/css/bootstrap.min.css">

        <script src="/control/device/js/DataTables/jQuery-3.6.0/jquery-3.6.0.min.js"></script>
        <script src="/control/device/js/bootstrap.bundle.min.js "></script>
        <script src="/control/device/js/DataTables/DataTables-1.11.4/js/jquery.dataTables.min.js"></script>
        <script src="/control/device/js/DataTables/DataTables-1.11.4/js/dataTables.bootstrap5.min.js"></script>
        <!--    <link rel="stylesheet" href="/control/device/js/DataTables/DataTables-1.11.4/css/jquery.dataTables.min.css">!-->

        <script src="/control/device/JsDoMenu/jsmap.js"></script>

    </head>

    <body style="background-color: #DFEFFF; height: 644px;">
        <script>
            //吃後端給的資料，轉成json處理
            var data = '<%=Get_Json() %>]';
            data = data.replace("	", " ").replace("　", " ");
            var b = JSON.parse(data);


            //月報表上的說明標籤
            $(function() {
                $("[data-toggle='tooltip']").tooltip();
            });
            //initial dataTable 
            $(document).ready(function() {
                $('#table_id').dataTable({
                    "order": [ //order by what
                        
                    ],
                    "search": { //搜尋欄enter鍵確認 false=不用按
                        return: false,
                    },
                    "fixedHeader": {
                        footer: true,
                    },

                    "info": true, //下方資訊欄

                    "lengthMenu": [20, 50], //選一頁幾個
                    stateSave: true, //狀態儲存 瀏覽完返回

                    "language": { //文字設定
                        "search": "搜尋:",
                        "sProcessing": "處理中...",
                        "sLengthMenu": "顯示 _MENU_ 項搜尋結果",
                        "sZeroRecords": "無資料",
                        "infoEmpty": "沒有搜尋結果",
                        "sInfo": "顯示第 _START_ 至 _END_ 項結果，搜尋到 _TOTAL_ 筆",
                        "infoFiltered": "(共_MAX_ 筆記錄)",
                        "oPaginate": {
                            "sFirst": "首頁",
                            "sPrevious": "<<",
                            "sNext": ">>",
                            "sLast": "尾頁",
                        },
                    },



                    data: b, //資料來源
                    "columns": [{ //各欄位名稱、設定
                        data: "name",
                        width: "20px"
                    }, {
                        data: "cwb_id",
                        width: "20px",

                    }, {
                        data: "dept",
                        width: "50px"
                    }, {
                        data: "company",
                        width: "30px",
                        className: "func",
                        render: function(data) {
                            return '<p  class="funp">' + data + '</p>';
                        }
                    }, {
                        data: "access",
                        width: "20px"
                    }, {
                        data: "job",
                        width: "250px",
                        render: function(data) {
                            return '<p  class="funp">' + data + '</p>';
                        }
                    }, {
                        data: "access_1f",
                        width: "20px"
                    }, {
                        data: "access_2f",
                        width: "20px"
                    }, {
                        data: "ID",
                        width: "40px",
                        title: "操作", // 這邊是欄位
                        render: function(data, type, row) {
                            var link = 'access.aspx?ID=' + data;
                            return '<a class="btn btn-warning btn-sm"  href="' + link + '" >編輯</a> ';
                        }
                    }, ]

                });
            });
        </script>

        <div class="main">
            <div class="position-relative">

                <p class="title">
                    人員權限列表
                </p>


      
            </div>

            <table id="table_id" class=" table table-striped table-hover">
                <thead>
                    <tr>
                        <th>姓名</th>
                        <th>員工編號</th>
                        <th>部門</th>
                        <th>廠商</th>
                        <th>權限</th>
                        <th>業務名稱</th>
                        <th>1F</th>
                        <th>2F</th>
                        <th>操作</th>


                    </tr>
                </thead>
                <tbody>

                </tbody>

            </table>

        </div>

    </body>

    </html>