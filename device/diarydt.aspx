<%@ Page Language="C#" Title="每月查詢" AutoEventWireup="true" CodeFile="diarydt.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html>
    <html>

    <head>
        <style type="text/css">
            body {
                margin: 20px 10px 20px 10px !important;
            }
            
            pre {
                border-top: 2px rgba(199, 199, 199, 0.647);
            }
            
            .minbox {
                min-width: 120px;
                max-width: 120px;
            }
            
            .funp {
                overflow: hidden;
                margin: 3px 10px 0px 5px;
                text-overflow: ellipsis;
                display: -webkit-box;
                -webkit-line-clamp: 2;
                -webkit-box-orient: vertical;
            }
            
            .mybut {
                border-radius: 0.5rem !important;
                margin: 0.25rem 0.5rem !important;
            }
            
            #table_detail {
                margin-top: 10px;
            }
            
            #table_detail td {
                padding: 5px 10px 15px 5px;
                margin-bottom: 5px;
                min-width: 130px;
                border-left: 2px rgba(116, 116, 116, 0.66) solid;
            }
            
            #table_detail tr {
                padding-bottom: 5px;
            }
            
            .detail_tit {
                font-size: 1.6rem;
                margin: 0px 10px 30px 0px;
                padding-bottom: 20px;
            }
        </style>

        <meta charset="utf-16" />
        <title>日誌查詢</title>
        <script src="./assets/jquery-3.6.0.min.js"></script>
        <script src="./assets/jquery.csv.js"></script>
        <script src="./assets/papaparse.min.js"></script>
        <script src="./assets/highlight.min.js"></script>
        <script src="./assets/helpers.js"></script>
        <link rel="stylesheet" type="text/css" href="./assets/datatables.min.css" />

        <script type="text/javascript" src="./assets/datatables.min.js"></script>
        <script>
            var datacsv = <%=Get_Json() %>];
            var tbl = {};
            var tt = "";
            var usedata = [];
            $(document).ready(function() {
                tbl = $("#example").DataTable({
                    search: {
                        return: false,
                    },
                    info: true,
                    lengthMenu: [10, 20, 50],
                    stateSave: false,
                    language: {
                        search: "搜尋:",
                        sProcessing: "處理中...",
                        sLengthMenu: "顯示 _MENU_ 項搜尋結果",
                        sZeroRecords: "無資料",
                        infoEmpty: "沒有搜尋結果",
                        sInfo: "顯示第 _START_ 至 _END_ 項結果，搜尋到 _TOTAL_ 筆",
                        infoFiltered: "(共_MAX_ 筆記錄)",
                        oPaginate: {
                            sFirst: "首頁",
                            sPrevious: "<<",
                            sNext: ">>",
                            sLast: "尾頁",
                        },
                    },
                    order: [
                        [0, ""],
                        [1, "desc"],
                    ],
                    data: datacsv,
                    columns: [{
                        data: "日誌班別",
                        title: "日誌",
                        class: "minbox",
                        render: function(data, type, row) {
                            return `${data.slice(0, 4)}.${data.slice(4, 6)}.${data.slice(
                              6,8)}(${data.slice(8, 9)})-${row["日誌編號"]}`;
                        },
                    }, {
                        data: "處理時間",
                        title: "處理時間",
                        class: "minbox",
                    }, {
                        data: "處理過程",
                        title: "處理過程",
                        render: function(data) {
                            return '<div  class="funp">' + data + "</div>";
                        },
                    }, ],
                });
                $("#example tbody").on("click", "tr", function() {
                    var dlist = this.getElementsByTagName("td"); //class time process
                    var look_class = dlist[0].innerText;
                    document.getElementById("class").innerText = look_class;
                    var look_dia = look_class.slice(14, 15);
                    look_class =
                        look_class.slice(0, 4) +
                        look_class.slice(5, 7) +
                        look_class.slice(8, 10) +
                        look_class.slice(11, 12);
                    $("#table_detail tbody").empty();
                    var temp = [];
                    datacsv.forEach(function(data) {
                        if (
                            data["日誌班別"] === look_class &&
                            data["日誌編號"] === look_dia
                        ) {
                            console.log(data)
                            temp.push([data["處理時間"], data["處理過程"]]);
                        }
                    });
                    temp.sort((a, b) => a[0] - b[0])
                    temp.forEach(function(data) {
                        $(`<tr>
                            <td>${data[0]}</td>
                            <td>${data[1]}</td>
                          </tr>`).appendTo("#table_detail>tbody");
                    })

                });

            });
        </script>
    </head>

    <body>

        <h2 class="title is-3">日誌內容</h2>
        <table id="example" class="display" style="width: 100%">
            <thead>
                <tr>
                    <th>Diary</th>
                    <th>Time</th>
                    <th>Process</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>日誌</td>
                    <td>時間</td>
                    <td>過程</td>
                </tr>
            </tbody>
        </table>
        <hr />
        <div id="diary-div">
            <h3 class="title is-3">日誌詳情</h3>

            <span class="detail_tit">班別時間:</span><span id="class">(點擊日誌查看詳細資訊)</span>

            <br />

            <span class="detail_tit">發生訊息:</span><span id="message"></span>

            <br />

            <span class="detail_tit">處理過程:</span><span id="process"></span>

            <table id="table_detail">
                <tr>
                    <td>時間</td>
                    <td>過程</td>
                </tr>
            </table>
        </div>
        <script>
            // enable syntax highlighting
            hljs.initHighlightingOnLoad();
        </script>
    </body>

    </html>