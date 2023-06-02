function map_click() {
    var PointerNo = form1.TextDevNo.value;
    if (PointerNo == '') PointerNo = form1.TextDevNo.value;
    var IO = document.getElementById('Radio2').checked == true ? "O" : "I";
    window.open('../../IDMS/Lib/map.aspx?PointerNo=' + PointerNo + '&IO=' + IO, '_blank', 'location=no;menubar=no;resizable=yes;scrollbars=no;status=no;toolbar=no');
    console.log('return');
}

function IE_Check() {
    if ("ActiveXObject" in window) {
        alert("本頁不支援IE，請使用其他瀏覽器");
        //window.history.go(-1); //返回上一頁
    } else {
        //alert("DD");
        Device_notice();
    }
}

function Get_month() { //舊網頁修改
    var f1 = document.getElementsByClassName('form-control')[0].value; // 2022/01 2~4 5~7
    //var f2 = "List.asp?SQLsearch=&SQL=Select * from Device where CreateDate like '2112$' order by CreateDate Desc&YM=2021/12&Burn=Y";
    //window.open(f2, "_blank");
    var f2 = [f1.substring(2, 4) + f1.substring(5, 7), f1];
    var f3 = "List.asp?SQLsearch=&SQL=Select * from Device where CreateDate like '" + f2[0] + "$' order by CreateDate Desc&YM=" + f2[1] + "&Burn=Y";
    //alert(f3);
    window.open(f3, "_blank");
    return (f2);
}

function Get_month2() {
    var f1 = document.getElementsByClassName('form-control')[0].value; // 2022/01 
    if (f1) {
        var f3 = "newout.aspx?Time=" + f1;

        window.open(f3, "_blank");
    } else {
        window.alert('格式錯誤 請重新輸入!')
    }
}

function Device_notice() {
    alert('1.請記得向負責人詢問設備編號及移入單，並紀錄之! \n2.設備進出機房或移動位置，請更改IDMS放置地點，詳見本頁底【注意事項】 ')

}

function Add_notice() {
    alert("102年12月課務會議決議，機器上機房機架前：\n" +
        "1. 先請負責人在IDMS建立設備資訊\n" +
        "2. 並繳交移入單(空白移入單IDMS有連結)\n" +
        "3. 在門禁、日誌追蹤及移入單註記設備編號，以便可彼此關聯\n" +
        "4. 若無則請其當場填寫！\n" +
        "5. 若有缺漏，務必記錄日誌追蹤，並建議設定日誌mail通知\n" +
        "6. 繳移入單時請核對IDMS資料是否正確，並移除日誌mail通知設定\n" +
        "7. 若無則請其當場填寫！並核對IDMS資料是否正確。\n" +
        "8. 若機器移出機房，請即時更改IDMS的放置地點。")
}