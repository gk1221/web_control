using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Control_DevEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {   
        if (!IsPostBack)
        {            
            if(Request.QueryString["DevID"] == null){
                Page.RegisterStartupScript("function","<script>IE_Check();</script>");                
            }
            DevID.Text = (Request["DevID"]);  
            GetOP();
            UpdateDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            CreateDate.Text = "尚未建檔";
            TextDevNo.Text = "0";
            Last_Host(); //產生最近新增清單
            Last_Repair(); //產生前幾筆廠商清單
            //Last_Staff(); //產生申請人
            //Last_Staff_Name(Menu_SW_1, Menu_SW_2);
            //Last_Staff_Name(Menu_HW_1, Menu_HW_2);
            //Last_Staff_Name(Menu_StaffName_1, Menu_StaffName_2);

            
                  
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
        if (Request["DevNo"] != null) DevID.Text = Request["DevNo"];

        if (DevID.Text=="")    //新增狀態
        {            
            BtnEdit.Enabled = false;
            BtnDel.Enabled = false;
        }
        else //編輯狀態
        {
            BtnEdit.Enabled = true;
            BtnDel.Enabled = true;
        }
        
        if (!IsPostBack & DevID.Text != "") { //讀取資料當輸入完成及連結進來時
            try{ 
                //有些會讀取異常做外面寫入
                List<string> re = ReadDev(int.Parse(DevID.Text));
                DevID.Text = HttpUtility.HtmlEncode(re[0]);
                CreateDate.Text = HttpUtility.HtmlEncode(re[1]);
                OP.Text = HttpUtility.HtmlEncode(re[2]);
                UpdateDate.Text = HttpUtility.HtmlEncode(re[3]);
                Location.Text = HttpUtility.HtmlEncode(re[4]);
            }catch(System.Exception){
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('讀取失敗');window.location.replace('newdevlist.aspx')</script>";
                Page.Controls.Add(Msg);
            }
            
        }
    }

    protected List<string> ReadDev(int Devid)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand("SELECT * FROM [Device2] WHERE [DevID]=@dev", Conn);
		cmd.Parameters.AddWithValue("@dev", Devid);
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        List<string> re = new List<string>();
        if (dr.Read())
        {
            //DevID.Text = HttpUtility.HtmlEncode(dr["DevID"].ToString());
            //CreateDate.Text = HttpUtility.HtmlEncode(dr["CreateDate"].ToString());
            HostName.Text = HttpUtility.HtmlEncode(dr["HostName"].ToString());
            //HostClass.SelectedValue = dr["HostClass"].ToString();
            for (int i = 0; i < HostClass.Items.Count; i++) { //比對至設備種類與選單相同
                if (HostClass.Items[i].Value == dr["HostClass"].ToString().Replace(" ", "")) HostClass.SelectedIndex = i;
            }            
            IO_Button(dr["IO"].ToString());
            Repair.Text = HttpUtility.HtmlEncode(dr["Repair"].ToString());
            Functions.Text = HttpUtility.HtmlEncode(dr["Functions"].ToString());
            Memo.Text = HttpUtility.HtmlEncode(dr["Memo"].ToString());
            PS.Text = HttpUtility.HtmlEncode(dr["PS"].ToString());
            StaffName.Text = HttpUtility.HtmlEncode(dr["StaffName"].ToString());
            HW.Text = HttpUtility.HtmlEncode(dr["HW"].ToString());
            SW.Text = HttpUtility.HtmlEncode(dr["SW"].ToString());
            TextDevNo.Text = HttpUtility.HtmlEncode(Get_Area(dr["Xall"].ToString(), dr["Yall"].ToString(), "定位編號")); //-(X-2)*31-Y
            //Location.Text = HttpUtility.HtmlEncode();
            //OP.Text = HttpUtility.HtmlEncode(dr["OP"].ToString());
            //UpdateDate.Text = HttpUtility.HtmlEncode(dr["UpdateDate"].ToString());
            re.Add(dr["DevID"].ToString());
            re.Add(dr["CreateDate"].ToString());
            re.Add(dr["OP"].ToString());
            re.Add(dr["UpdateDate"].ToString());
            re.Add(Get_Area(dr["Xall"].ToString(), dr["Yall"].ToString(), "區域名稱"));
        }
        if(Location.Text=="")Location.Text="設備定位錯誤";
        return re;
    }
    public static string pre_staff(){ //dev2--建立員工清單給Js用
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"SELECT COUNT(*) FROM Config a1, Config a2
                       WHERE a1.Kind = a2.Item AND a2.Kind='資訊中心' AND a1.Kind<>' ' ";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }

        cmd.CommandText = @"SELECT a1.kind, a1.Item
                            FROM Config a1, Config a2
                            WHERE a1.Kind = a2.Item AND a2.Kind='資訊中心' AND a1.Kind<>' '
                            ORDER BY a1.kind, a1.Memo DESC";
        dr.Close();
        dr = cmd.ExecuteReader();
        //[["系統課","001"], ["網管課","002"],.....]
        string out_s = "[";
        while(dr.Read() & count>0){
            out_s += "[\"" + dr[0].ToString() + "\",\"" + dr[1].ToString() ;
            if(count==1){
                out_s += "\"]";
            }
            else{
                out_s += "\"],";
            }
            count -= 1;
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return out_s;
    }

    protected void MenuHost_SelectedIndexChanged(object sender, EventArgs e)//最近一筆異動選取時帶入
    {
        ReadDev(int.Parse(MenuHost.SelectedValue));
        DevID.Text = "";
        CreateDate.Text = "尚未建檔";
        //HostName.Text = MenuHost.SelectedValue.ToString();
    }

    protected void Last_Host(){ //建立最近40筆異動清單
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(@"SELECT  TOP(100) DevID
                                            ,concat(Month(CreateDate),'/',DAY([CreateDate]))      
                                            ,[StaffName]
                                            ,[HOSTNAME]
                                            FROM [dbo].[Device2]
                                            ORDER BY CreateDate DESC", Conn);
        DataSet ds = RunQuery("Control", cmd);
        MenuHost.Items.Add(new ListItem("(請選擇)","0"));
        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows) // 丟進列表裡:(5/10, 王小名)SSD*1
            {
                //MenuHost.Items.Add(new ListItem(" ("+row[1].ToString() + "," + row[2].ToString()+")"+row[3].ToString(), row[0].ToString()));
                MenuHost.Items.Add(new ListItem( String.Format("({0},{1}) {2}", row[1].ToString(), row[2].ToString(), row[3].ToString()) , row[0].ToString()));
            }
        }
    }
    protected void MenuRepair_SelectedIndexChanged(object sender, EventArgs e)//廠商選項選取後填入廠商空格
    {
        Repair.Text = MenuRepair.SelectedValue;
    }
    protected void Last_Repair(){ //最近10筆廠商
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        
		SqlCommand cmd = new SqlCommand(@"SELECT TOP(10) Repair, COUNT(*) as c FROM [dbo].[Device2]
                                            WHERE Repair<>'N/A'
                                            GROUP BY Repair
                                            ORDER BY c DESC", Conn);
        DataSet ds = RunQuery("Control", cmd);
        MenuRepair.Items.Add(new ListItem("(無)","N/A"));
        if (ds.Tables.Count > 0)
        {
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                MenuRepair.Items.Add(new ListItem(row[0].ToString(), row[0].ToString()));
            }
        }
    }


    protected void BtnAdd_Click(object sender, EventArgs e) //新增按鈕
    {        
        Literal Msg = new Literal();
        List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        if(Null_Check(Msg)){} //空白確認
        else if(xy.Count()==0){ //錯誤確認
            Msg.Text = "<script>alert('設備定位錯誤，請重新定位!');</script>";
        }
        else{
            int id = GetPKNo("DevID", "Device2");
            DevID.Text = HttpUtility.HtmlEncode(id.ToString());
            CreateDate.Text = HttpUtility.HtmlEncode(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")); 
            List<SqlParameter> pars = new List<SqlParameter>();            
            string SQL=GetInsConDevSQL(id, pars);            //獲取新增至設備資料表語法
            string SQL2=String.Format("新增[設備2]：： ({0}, {1}, {2}, {3}, {4}, {5})", //生命履歷語法
                                HostName.Text.Trim(),
                                HostClass.SelectedValue.Replace(" ", ""),
                                Functions.Text.Trim(),
                                StaffName.Text.Trim(),
                                IO_Change_Button(Button_check()),
                                Get_Area(xy[0], xy[1], "區域名稱"));  
            try
            {
                ExecDbSQL(SQL, pars);
                InsLifeSQL(SQL2);         
                Location.Text = Get_Area(xy[0], xy[1], "區域名稱");
                Msg.Text = "<script>alert('新增完成! ');</script>";
            }
            catch (System.Exception)
            {                
                Msg.Text = "<script>alert('新增失敗!請重新嘗試 ');</script>";
            }
        }
        Page.Controls.Add(Msg);
    }

    protected void BtnEdit_Click(object sender, EventArgs e) //編輯按鈕
    {   
        Literal Msg = new Literal();
        List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        if(Null_Check(Msg)){} //空白確認
        else if(xy.Count()==0){ //錯誤確認
            Msg.Text = "<script>alert('設備定位錯誤，請重新定位!');</script>";
        }
        else{        
            List<SqlParameter> pars = new List<SqlParameter>();				
            pars.Add(new SqlParameter("@devid", SqlDbType.Int)); 
            pars.Last().Value = int.Parse(DevID.Text);
            
            try
            {
                InsLifeSQL("修改2 [" + HostName.Text + "] ：： " + GetUpdate(int.Parse(DevID.Text), "Life"));
                ExecDbSQL("UPDATE [device2] SET " + GetUpdate(int.Parse(DevID.Text),"SQL", pars) + " WHERE [DevID]= @devid", pars); 
                Msg.Text = "<script>alert('修改完成');window.location.replace('newdev2.aspx?DevID="+DevID.Text+" ')</script>";
            }
            catch (System.Exception)
            {                
                Msg.Text = "<script>alert('修改失敗，請再試一次');window.location.replace('newdev2.aspx?DevID="+DevID.Text+" ')</script>";
            }
        }
        Page.Controls.Add(Msg);
    }
    protected void BtnDel_Click(object sender, EventArgs e) //刪除按鈕
    {
        Literal Msg = new Literal();
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("DELETE FROM [device2] WHERE [DevID]=@devid", Conn);
        cmd.Parameters.AddWithValue("devid", DevID.Text);

        List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        string SQL2=String.Format("新增[設備] ({0}, {1}, {2}, {3}, {4}, {5})",
                                HostName.Text.Trim(),
                                HostClass.SelectedValue.Replace(" ", ""),
                                Functions.Text.Trim(),
                                StaffName.Text.Trim(),
                                IO_Change_Button(Button_check()),
                                Get_Area(xy[0], xy[1], "區域名稱"));
        InsLifeSQL("刪除[設備2]．[" + HostName.Text.Trim() + "]，原始資料：" + SQL2);
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        Msg.Text = "<script>alert('已刪除!');window.close();window.location.replace('newdevlist.aspx');</script>";
        Page.Controls.Add(Msg);
    }
    protected void BtnTest_Click(object sender, EventArgs e)
    {
        Literal Msg = new Literal();
        List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        string SQL2=String.Format("新增[設備] ({0}, {1}, {2}, {3}, {4}, {5})",
                                HostName.Text.Trim(),
                                HostClass.SelectedValue.Replace(" ", ""),
                                Functions.Text.Trim(),
                                StaffName.Text.Trim(),
                                IO_Change_Button(Button_check()),
                                Get_Area(xy[0], xy[1], "區域名稱"));
        string out2 = "刪除[設備]．[" + HostName.Text.Trim() + "]，原始資料：" + SQL2;
        Msg.Text = "<script>alert('"+out2+"');</script>";
        Page.Controls.Add(Msg);
    }

    protected string GetUserName(){ //將使用者ID -> Name
        string UserID = Request.Cookies["UserID"].Value.ToString();    //自cookie取得
        if (UserID == "") return "無";
        else{
            SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
            Conn.Open();
            SqlCommand cmd = new SqlCommand("Select Kind,Item from Config where Mark=@UserID", Conn);
            cmd.Parameters.AddWithValue("UserID", UserID);
            SqlDataReader dr = null;
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {                
                return dr[1].ToString();//UserName                 
            }
            else{
                return "無";
            }
        }
    }


    protected bool Null_Check(Literal back){ //空白確認
        if(HostClass.SelectedValue==""){
            back.Text = "<script>alert('尚未選擇設備種類!');</script>";return true;
        }else if(HostName.Text==""){
            back.Text = "<script>alert('尚未填寫設備名稱!');</script>";return true;
        }else if(StaffName.Text==""){
            back.Text = "<script>alert('尚未填寫申請人員!');</script>";return true;
        }else if(HW.Text==""){
            back.Text = "<script>alert('尚未填寫硬體負責人!');</script>";return true;
        }else if(OP.Text==""){
            back.Text = "<script>alert('尚未填寫OP!');</script>";return true;
        }else if(SW.Text==""){
            back.Text = "<script>alert('尚未填寫軟體負責人!');</script>";return true;
        }else if(TextDevNo.Text=="0"){
            back.Text = "<script>alert('尚未定位設備位置!');</script>";return true;
        }else return false;
    }
    protected string Button_check() //異動種類:名稱->存資料庫
    {
        if (Radio1.Checked) {
            return("I");//in
        }
        else if (Radio2.Checked) {
            return("O");//out
        }
        else if (Radio3.Checked) {
            return("N");//change
        }
        else {
            return("M");//other
        }
    }
    protected void IO_Button(string IO) //異動種類:存資料庫->名稱
    {
        if (IO == "I") {
            Radio1.Checked = true;
        }
        else if (IO == "O") {
            Radio2.Checked = true;
        }
        else if (IO == "N") {
            Radio3.Checked = true;
        }
        else {
            Radio4.Checked = true;
        }
    }

    protected static string IO_Change_Button(string IO) //For life
    {
        if (IO == "I") {
                return("移入");//in
        }
        else if (IO == "O") {
            return("移出");//out
        }
        else if (IO == "N") {
            return("更換");//change
        }
        else {
            return("其它");//other
        }
    }
    protected List<string> Get_XY(int location){ //IDMS地圖回傳PointerNumber後 -> XY座標
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //SqlCommand cmd = new SqlCommand("SELECT * FROM [View_作業主機] WHERE [作業編號]=" + TextApNo.Text, Conn);
		SqlCommand cmd = new SqlCommand("SELECT 坐標X,坐標Y,區域名稱 FROM [定位設定] WHERE [定位編號]=@loc", Conn);
		cmd.Parameters.AddWithValue("@loc", location);		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        List<string> re = new List<string>();
        if (dr.Read()){
            re.Add((dr["坐標X"].ToString()));
            re.Add((dr["坐標Y"].ToString()));
            re.Add((dr["區域名稱"].ToString()));
        }        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return re;
    }
    protected string Get_Area(string xall, string yall, string ColName){ //輸入XY可得位置名稱
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        //SqlCommand cmd = new SqlCommand("SELECT * FROM [View_作業主機] WHERE [作業編號]=" + TextApNo.Text, Conn);
		SqlCommand cmd = new SqlCommand("SELECT * FROM [定位設定] WHERE [坐標X]= @Xall AND [坐標Y]=@Yall", Conn);
		cmd.Parameters.AddWithValue("@Xall", xall);
        cmd.Parameters.AddWithValue("@yall", yall);		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();        
        if (dr.Read()){
            if(ColName=="區域名稱"){
                return dr[ColName].ToString() + dr["定位名稱"].ToString();
            }
            else{ //定位編號
                return dr[ColName].ToString();
            }            
        }else{
            return "定位錯誤";
        }        
        return null;
    }
    protected string GetInsConDevSQL(int devid ,List<SqlParameter> pars) //取得新增資料的語法
    { 
        //int PointerNo = TextDevNo.Text != ""?int.Parse(TextDevNo.Text):-88;
        List<string> XY =  Get_XY(int.Parse(TextDevNo.Text));
        pars.Add(new SqlParameter("@DevID", devid));
        pars.Add(new SqlParameter("@Time", SqlDbType.DateTime2));
        pars.Last().Value = DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");		
        
		pars.Add(new SqlParameter("@HostName", HostName.Text.Trim()));
        pars.Add(new SqlParameter("@HostClass", HostClass.SelectedValue.Replace(" ", "")));
        pars.Add(new SqlParameter("@Repair", Repair.Text.Trim()));
        pars.Add(new SqlParameter("@IO", Button_check()));
        if(Functions.Text=="")Functions.Text = "N/A";
        pars.Add(new SqlParameter("@Functions", Functions.Text.Trim().Replace("	", " ").Replace("　"," ")));
        pars.Add(new SqlParameter("@Memo", Memo.Text.Trim()));
        pars.Add(new SqlParameter("@PS", PS.SelectedValue));
        pars.Add(new SqlParameter("@StaffName", StaffName.Text.Trim()));
        pars.Add(new SqlParameter("@HW", HW.Text.Trim()));
        pars.Add(new SqlParameter("@SW", SW.Text.Trim()));
        pars.Add(new SqlParameter("@OP", OP.Text.Trim()));
        pars.Add(new SqlParameter("@Xall", XY[0]));
        pars.Add(new SqlParameter("@Yall", XY[1]));
		//15 row
        return ("INSERT INTO [Device2] values("
            + "@DevID" + ","
            + "@Time" + ","
            + "@HostName" + ","
            + "@HostClass" + ","
            + "@Repair" + ","
            + "@IO" + ","
            + "@Functions" + ","
            + "@Memo" + ","
            + "@PS" + ","
            + "@StaffName" + ","
            + "@HW" + ","
            + "@SW" + ","
            + "@OP" + ","
            + "@Xall" + ","
            + "@Yall" + ","
            + "@Time )");
    }

    protected void InsLifeSQL(string action) //lifelog
    {   
        
        string who = GetUserName();
        List<SqlParameter> pars = new List<SqlParameter>();                     
        pars.Add(new SqlParameter("@ID", GetPKNo("LifeID", "ChangeLog")));
        pars.Add(new SqlParameter("@DevID", DevID.Text));
        pars.Add(new SqlParameter("@Time", SqlDbType.DateTime2));
        pars.Last().Value = DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
		pars.Add(new SqlParameter("@type", "設備"));
        pars.Add(new SqlParameter("@action", action));
        pars.Add(new SqlParameter("@OP", who));
        pars.Add(new SqlParameter("@IP", Request.ServerVariables["REMOTE_ADDR"].ToString()));		
		//7 row
        ExecDbSQL ("INSERT INTO [ChangeLog] values("
            + "@ID" + ","
            + "@type" + ","
            + "@DevID" + ","
            + "@Action" + ","
            + "@OP" + ","
            + "@Time" + ","
            + "@IP )", pars  );
    }


    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select ISNULL(MAX(" + PKfield + "), 0) FROM " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1; 
        if (dr.Read()) PkNo=int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (PkNo);
    }
    

    protected string GetInsConDevSQL() //取得新增資料的語法-一般版
    {
		
		//15 row
        return (@"INSERT INTO [Device2] values('"
            + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"','"
            + HostName.Text +"','"
            + HostClass.SelectedValue +"','"
            + Repair.Text + "','"
            + Button_check() + "','"
            + Functions.Text + "','"
            + Memo.Text + "','"
            + PS.SelectedValue + "','"
            + StaffName.Text + "','"
            + HW.Text + "','"
            + SW.Text + "','"
            + "d1 and d2" + "',"
            + 79 + ","
            + 82 + ",'"
            + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+"'");
            
    }



    

    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind, string SQLorLife, List<SqlParameter> pars) //取得單一欄位修改資料的語法
    {
        string SQL = "";
        

        if (Source != Target)
        {
            if (SQLorLife == "SQL")
            {
                switch (Kind)
                {
					case "string": case "date": case "datetime": 						
						SQL = SQL + ",[" + ColName + "]=" + "@" + Convert.ToString(pars.Count); 
						pars.Add(new SqlParameter("@" + Convert.ToString(pars.Count), Target));				
						break;
                    case "integer": case "money": 						
						SQL = SQL + ",[" + ColName + "]=" + "@" + Convert.ToString(pars.Count);
						pars.Add(new SqlParameter("@" + Convert.ToString(pars.Count), SqlDbType.Int));
						pars.Last().Value = int.Parse(Target);
						break;
                    case "null": 
						SQL = SQL + ",[" + ColName + "]=" + null; 
						break;
                    default: 						
						SQL = SQL + ",[" + ColName + "]=" + "@" + Convert.ToString(pars.Count); 
						pars.Add(new SqlParameter("@" + Convert.ToString(pars.Count), Target));				
						break;
                }
            }
            
        }
        return (SQL);
    }
    
    protected string GetUpdate(int DevID, string SQLorLife, List<SqlParameter> pars) //取得修改資料的SQL語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand("select * from [Device2] where [DevID]=@Dev", Conn);
		cmd.Parameters.AddWithValue("@Dev", DevID);
		List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {                 
            SQL =     GetUpdateCol("HostName", dr["HostName"].ToString(), HostName.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("HostClass", dr["HostClass"].ToString(), HostClass.SelectedValue, "string", SQLorLife, pars);
            SQL += GetUpdateCol("Repair", dr["Repair"].ToString(), Repair.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("IO", dr["IO"].ToString(), Button_check(), "string", SQLorLife, pars);
            SQL += GetUpdateCol("Functions", dr["Functions"].ToString(), Functions.Text.Trim().Replace("	", " ").Replace("　"," "), "string", SQLorLife, pars);
            SQL += GetUpdateCol("Memo", dr["Memo"].ToString(), Memo.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("PS", dr["PS"].ToString(), PS.SelectedValue, "string", SQLorLife, pars);
            SQL += GetUpdateCol("StaffName", dr["StaffName"].ToString(), StaffName.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("HW", dr["HW"].ToString(), HW.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("SW", dr["SW"].ToString(), SW.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("Xall", dr["Xall"].ToString(), xy[0].ToString(), "integer", SQLorLife, pars);
            SQL += GetUpdateCol("Yall", dr["Yall"].ToString(), xy[1].ToString(), "integer", SQLorLife, pars);
            SQL += GetUpdateCol("OP", dr["OP"].ToString(), OP.Text, "string", SQLorLife, pars);
            SQL += GetUpdateCol("UpdateDate", dr["UpdateDate"].ToString(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "datetime", SQLorLife, pars);;

            if (SQL != "") SQL = SQL.Substring(1);  
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind, string SQLorLife) //取得單一欄位修改資料的語法
    {
        string SQL = "";

        if (Source != Target)
        {
            if (SQLorLife == "SQL")
            {
                switch (Kind)
                {
					case "string":
                    case "date":
                    case "datetime": SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                    case "integer":
                    case "money": SQL = SQL + ",[" + ColName + "]=" + Target; break;
                    case "null": SQL = SQL + ",[" + ColName + "]=" + null; break;
                    default: SQL = SQL + ",[" + ColName + "]='" + Target + "'"; break;
                }
            }
            else if (SQLorLife == "Life")
            {
                if (ColName != "定位編號")
                {
                    if (Source == "") Source = "(空白)";
                    if (Target == "") Target = "(空白)";
                    SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
                }
                else
                {
                    SQL = SQL + ",[" + ColName + "]：" + Source + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=" + Source)
                        + ") -> " + Target + "(" + GetConfig("select [區域名稱]+[定位名稱] as [放置地點] from [定位設定] where [定位編號]=" + Target) + ")";
                }
            }
        }
        
        return (SQL);
    }
    
    protected string GetUpdate(int DevID, string SQLorLife) //取得修改資料的SQL語法 -- LIFELOG
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand("select * from [Device2] where [DevID]=@Dev", Conn);
		cmd.Parameters.AddWithValue("@Dev", DevID);
		List<string> xy = Get_XY(int.Parse(TextDevNo.Text));
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {    
            SQL =  GetUpdateCol("HostName", dr["HostName"].ToString(), HostName.Text, "string", SQLorLife);
            SQL += GetUpdateCol("HostClass", dr["HostClass"].ToString(), HostClass.SelectedValue, "string", SQLorLife);
            SQL += GetUpdateCol("Repair", dr["Repair"].ToString(), Repair.Text, "string", SQLorLife);
            SQL += GetUpdateCol("IO", dr["IO"].ToString(), Button_check(), "string", SQLorLife);
            SQL += GetUpdateCol("Functions", dr["Functions"].ToString(), Functions.Text.Trim().Replace("	", " ").Replace("　"," "), "string", SQLorLife);
            SQL += GetUpdateCol("Memo", dr["Memo"].ToString(), Memo.Text, "string", SQLorLife);
            SQL += GetUpdateCol("PS", dr["PS"].ToString(), PS.SelectedValue, "string", SQLorLife);
            SQL += GetUpdateCol("StaffName", dr["StaffName"].ToString(), StaffName.Text, "string", SQLorLife);
            SQL += GetUpdateCol("HW", dr["HW"].ToString(), HW.Text, "string", SQLorLife);
            SQL += GetUpdateCol("SW", dr["SW"].ToString(), SW.Text, "string", SQLorLife);
            SQL += GetUpdateCol("Xall", dr["Xall"].ToString(), xy[0].ToString(), "integer", SQLorLife);
            SQL += GetUpdateCol("Yall", dr["Yall"].ToString(), xy[1].ToString(), "integer", SQLorLife);
            SQL += GetUpdateCol("OP", dr["OP"].ToString(), OP.Text, "string", SQLorLife);
            SQL += GetUpdateCol("UpdateDate", dr["UpdateDate"].ToString(), DateTime.Now.ToString("yyyy/MM/dd tt hh:mm:ss"), "datetime", SQLorLife);

            if (SQL != "") SQL = SQL.Substring(1);  
        }
        //return (dr["Xall"].ToString() +" "+ xy[0].ToString());
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (SQL);
    }

    protected string GetConfig(string SQL) //讀取某系統設定值
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        string cfg = ""; if (dr.Read()) cfg = HttpUtility.HtmlEncode(dr[0].ToString());
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();

        return (cfg);
    }
    protected void GetOP() //從工作日誌抓目前當班OP
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["DiaryConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("SELECT TOP(1) * FROM SIGN ORDER BY TOUR DESC", Conn);
		SqlDataReader dr = cmd.ExecuteReader();
        if(dr.Read())
        {
            OP.Text = dr["OPname"].ToString();
        }
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }

    protected DataSet RunQuery(string db, SqlCommand sqlQuery) //讀取DB資訊
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings[db+ "ConnectionString"].ConnectionString);
        SqlDataAdapter dbAdapter = new SqlDataAdapter();
        dbAdapter.SelectCommand = sqlQuery;
        sqlQuery.Connection = Conn;
        DataSet QueryDataSet = new DataSet();
        dbAdapter.Fill(QueryDataSet);
        dbAdapter.Dispose(); Conn.Close(); Conn.Dispose();
        return (QueryDataSet);
    }

    protected void ExecDbSQL(string SQL, List<SqlParameter> pars=null) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);		
		cmd.Parameters.AddRange(pars.ToArray());		
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }
    
}
