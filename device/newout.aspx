<%@ Page Language="C#" Title="月報表" AutoEventWireup="true" CodeFile="newout.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml">

    <head runat="server">

        <title>月報表</title>
        <style type="text/css">
            .title {
                text-align: center;
                color: #000066;
                font-size: xx-large;
                font-weight: bolder;
            }
            
            body {
                margin: 0px 5px 0 5px;
            }
            
            #table_id {
                max-width: 100%!important;
            }
            
            table {
                border: 1px solid black;
                border-collapse: collapse;
                width: 100%;
                padding: 0;
            }
            
            tr,
            td {
                font-size: 13px;
                line-height: 14px;
            }
            
            th {
                color: blue;
                text-decoration: underline;
            }
            
            .head10 {
                width: 44px;
            }
            
            .head9 {
                width: 60px;
            }
            
            .head8 {
                width: 56px;
            }
            
            .head7 {
                width: 56px;
            }
            
            .head6 {
                width: 200px;
            }
            
            .head5 {
                width: 40px;
            }
            
            .head4 {
                width: 40px;
            }
            
            .head3 {
                width: 28px;
            }
            
            .head2 {
                width: 50px;
            }
            
            .head1 {
                width: 40px;
            }
        </style>
        <script>
            var data = '<%=Get_Json()%>]';

            var b = JSON.parse(data);
            //console.log((b));
            //console.log((typeof(b)));
            if (b.length == 0) {
                alert("查無資料!");
                window.close();
            } else {
                alert("查詢結果:共有 " + b.length + " 筆記錄");
            }
            var otb;
            var table = document.getElementById("myTable");


            function buildHtmlTable(selector) {
                var columns = addAllColumnHeaders(b, selector);

                for (var i = 0; i < b.length; i++) {
                    var row$ = $('<tr/>');
                    for (var colIndex = 0; colIndex < columns.length; colIndex++) {
                        var cellValue = b[i][columns[colIndex]];
                        if (cellValue == null) cellValue = "";
                        row$.append($('<td/>').html(cellValue));
                    }
                    $(selector).append(row$);
                }
            }

            // Adds a header row to the table and returns the set of columns.
            // Need to do union of keys from all records as some records may not contain
            // all records.
            function addAllColumnHeaders(myList, selector) {
                var columnSet = [];
                var headerTr$ = $('<tr/>');

                for (var i = 0; i < myList.length; i++) {
                    var rowHash = myList[i];
                    /*
                                        for(var j=0; j<rowHash.length; j++) {
                                            if($.inArray(rowHash[j], columnSet) == -1){
                                                columnSet.push(rowHash[j]);
                                                headerTr$.append($('<th />').html(rowHash[j]));
                                            }
                                        }

                     */
                    var counter = 10;
                    for (var key in rowHash) {
                        console.log(rowHash);
                        if ($.inArray(key, columnSet) == -1) {
                            columnSet.push(key);
                            headerTr$.append($('<th class="head' + counter + '"/>').html(key));
                            counter -= 1;
                        }

                    }

                }

                $(selector).append(headerTr$);
                return columnSet;
            }
        </script>
        <script src="/control/device/js/jquery-3.6.0.min.js"></script>



    </head>

    <body onLoad="buildHtmlTable('#excelDataTable')" style="background-color: #DFEFFF; margin: 8px;">





        <h1 class="title">機房設備異動紀錄</h1>
        <script>
            var url = new URL(location.href);
            var times = url.searchParams.get('Time');
            var ot = document.getElementsByClassName('title')[0];
            if (times) {
                document.getElementsByClassName('title')[0].innerHTML = times.substr(0, 4) + '年' + times.substr(5, 2) + '月' + ot.innerHTML;
            } else {
                document.getElementsByClassName('title')[0].innerHTML = ot.innerHTML;
            }
        </script>
        <table id="excelDataTable" border="1">



        </table>


    </body>

    </html>