<%@ Page Language="C#" Title="進階查詢" AutoEventWireup="true" CodeFile="newsearch.aspx.cs" Inherits="Device_DevEdit" MaintainScrollPositionOnPostback="true" Debug="true" %>

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
            
            .head-div {
                display: flex;
                justify-content: space-between;
            }
            .psword {
                font-size: small;
            }
            
            .psoan {
                color: rgb(204, 102, 0);
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
            
            table {
                border: 1px solid black;
                border-collapse: collapse;
                width: 100%;
                padding: 0;
            }
            
            tr,
            td {
                font-size: 14px;
            }
            
            th {
                color: blue;
                text-decoration: underline;
            }
            
            .d1 {
                max-width: 45px;
            }
        </style>
        <script>
            function Sendout(){


            }
        </script>
        <script src="/control/device/js/jquery-3.6.0.min.js"></script>



    </head>

    <body  style="background-color: #DFEFFF; margin: 8px;">
        <h1 class="title">進階查詢</h1>

        <form id="FormSearch" action="./newout.aspx" method="get" target="_blank">
            <p>建檔日期(YYYYMMDD)介於           
                <input type="text" name="From" size="20" value="20210202">和           
                <input type="text" name="TO" size="20" value="20211231">之間 　 <font color="#CC6600">          
                <font size="2">(查單一日期請用新版搜尋介面) (西元年)</font> 
            </font>
            </p>
            
            <p>
                搜尋字串 ：
                <input type="text" name="Find" size="30">
                <input  type="radio" name="Type" value="A" checked="true">AND
                <input  type="radio" name="Type" value="O">OR
                <font size="2" color="#CC6600">　  (對所有文字欄位比對字串是否符合，以空白隔開搜尋字串做AND或OR比對)</font>
            </p>
            <p>設備異動 ：        
                <input type="checkbox" name="Check" value="I">移入      
                <input type="checkbox" name="Check" value="O">移出       
                <input type="checkbox" name="Check" value="N">更換
                <input type="checkbox" name="Check" value="M">其它
            </p>
            <p>設備種類 ：
                <input type="checkbox" name="Class" value="主機">主機         
                <input type="checkbox" name="Class" value="其它機器">其它機器  
                <input type="checkbox" name="Class" value="耗材">耗材
                <input type="checkbox" name="Class" value="週邊設備">週邊設備   
                <input type="checkbox" name="Class" value="零附件">零附件  
                <input type="checkbox" name="Class" value="網路設備">網路設備  
                <input type="checkbox" name="Class" value="辦公設備">辦公設備  
                <input type="checkbox" name="Class" value="儲存設備">儲存設備  
            
            </p>
            <span class="psword psoan"> (若有其他查詢需求，請聯絡SSM小組) </span>
           
            <p align="right">
                <input type="submit" value="　查　詢　" >　 
              
            </p>
          </form>


    </body>

    </html>