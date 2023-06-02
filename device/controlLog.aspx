<%@ Page Language="C#" Title="異動紀錄" AutoEventWireup="true" CodeFile="controlLog.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">
        <meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
        <meta http-equiv="X-UA-Compatible" content="IE=11">
        <title>異動紀錄</title>
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
                vertical-align: middle;
            }
            
            #table_id_filter input {
                border-radius: 5px;
            }
            .IDlink{
                text-decoration: none;
                color: black;
            }
        </style>
        <!--
      
    <script src="/control/device/js/jquery-3.6.0.min.js"></script>
    <script src="/control/device/js/jquery.dataTables.min.js"></script>
    <link rel="stylesheet" href="/control/device/js/jquery.dataTables.min.css">    
    -->


        <link rel="stylesheet" type="text/css" href="/control/device/js/DataTables/DataTables-1.11.4/css/dataTables.bootstrap5.min.css" />
        <link rel="stylesheet" href="/control/device/css/bootstrap.min.css">

        <script src="/control/device/js/DataTables/jQuery-3.6.0/jquery-3.6.0.min.js"></script>
        <script src="/control/device/js/bootstrap.bundle.min.js "></script>
        <script src="./assets/papaparse.min.js"></script>
        <script src="/control/device/js/DataTables/DataTables-1.11.4/js/jquery.dataTables.min.js"></script>
        <script src="/control/device/js/DataTables/DataTables-1.11.4/js/dataTables.bootstrap5.min.js"></script>
        <!--    <link rel="stylesheet" href="/control/device/js/DataTables/DataTables-1.11.4/css/jquery.dataTables.min.css">!-->

        <script src="/control/device/JsDoMenu/jsmap.js"></script>

    </head>

    <body style="background-color: #DFEFFF; height: 644px;">
        <script>
            var data = '<%=Get_Json() %>]';

            //data = data.replace("	", " ").replace("　", " ");
            var b = JSON.parse(data);
            //var b = Papa.parse(data);
            //console.log((b));
            //console.log((typeof(b)));
            var otb;

            $(function() {
                $("[data-toggle='tooltip']").tooltip();
            });

            $(document).ready(function() {

                $('#table_id').dataTable({
                    "order": [
                        [0, "desc"]
                    ],
                    "search": {
                        return: false,
                    },
                    "fixedHeader": {

                        footer: true,
                    },

                    "info": true,

                    "lengthMenu": [20, 50],
                    stateSave: true,

                    "language": {
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



                    data: b,
                    "columns": [{
                        data: "LifeID",
                        width: "10px"
                    }, {
                        data: "Type",
                        width: "30px",
                        render: function(data) {
                            return data;
                        }

                    }, {
                        data: "controlID",
                        width: "30px",
                        render: function(data, type, row, meta) {
                            if(row["Type"]==="人員")
                                return `<a class="IDlink" href="/control/people/access.aspx?ID=${data}">${data} </a>`;
                            else return `<a class="IDlink" href="newdev2.aspx?DevID=${data}">${data} </a>`;
                           // return '<span class="IDlink">' + data + '</span>';
                        }
                    }, {
                        data: "Action",
                        width: "700px",
                        className: "func",
                        render: function(data) {
                            return data;
                        }
                    }, {
                        data: "OP",
                        width: "58px"
                    }, {
                        data: "Time",
                        width: "60px"
                    }, {
                        data: "IP",
                        width: "40px"
                    }, ]

                });
                $('.IDlink').click(function() {
                    window.open('newdev2.aspx?DevID=' + this.textContent, '_self');
                    //alert(this.textContent);
                });

                $('tr').mouseover(function() {
                    var tar = this.getElementsByClassName('IDlink')[0];
                    tar.style = (
                        "color:#3377f9;font-weight:bolder;font-size:18px;text-decoration: underline;"
                    );
                })
                $('tr').mouseout(function() {
                    var tar = this.getElementsByClassName('IDlink')[0];
                    tar.style = (
                        "color:black"
                    );
                })
            });
        </script>

        <div class="main">
            <div class="position-relative">

                <p class="title">
                    門禁異動紀錄
                </p>




            </div>

            <table id="table_id" class=" table border-secondary table-bordered order-column table-hover">
                <thead class=" table-success  ">
                    <tr>
                        <th>NO.</th>
                        <th>門禁</th>
                        <th>編號</th>
                        <th>生命履歷</th>
                        <th>操作人員</th>
                        <th>異動時間</th>
                        <th>IP</th>



                    </tr>
                </thead>
                <tbody>

                </tbody>

            </table>

        </div>

    </body>

    </html>