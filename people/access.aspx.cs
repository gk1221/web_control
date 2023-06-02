using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;

public partial class Control_PeoEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            ID.Text = (Request["ID"]);
            Trace.Write(ID.Text);
            Last_Name();
            UpdateDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {

          if (!IsPostBack & ID.Text != "") { //讀取資料當輸入完成及連結進來時
            try{ 
                //有些會讀取異常做外面寫入
                Trace.Write(ID.Text);
                ReadName(Int32.Parse(ID.Text));
                Trace.Write("HERE");
          
            }catch(System.Exception){
                Literal Msg = new Literal();
                Msg.Text = "<script>alert('讀取失敗');window.location.replace('accesslist.aspx')</script>";
                Page.Controls.Add(Msg);
            }
            
        }


    }

    public static string pre_staff()
    { //
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        string sql = @"SELECT [dept]
					  FROM [dbo].[accesslist]
					  group by dept
					  order by LEN(dept) ASC
                   ";
        int count_rows = 0;
        DataTable Devtbl = new DataTable();
        using (SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString))
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            conn.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
            dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
            count_rows = Devtbl.Rows.Count;
        }
        string out_s = "[";
        foreach (DataRow dr in Devtbl.Rows)
        {
            out_s += "[" + string.Format("  \"{0}\", \"{1}\" ", dr[0].ToString(), count_rows);
            if (count_rows == 1)
            {
                out_s += "]";
            }
            else
            {
                out_s += "],";
            }
            count_rows -= 1;
        }
        return out_s;
    }
    public static string pre_company()
    { //dev2--建立廠商清單給Js用
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        string sql = @"SELECT [company]
                        FROM [dbo].[accesslist]
                        WHERE company <> ''
                        group by company ";
        int count_rows = 0;
        DataTable Devtbl = new DataTable();
        using (SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString))
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            conn.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
            dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
            count_rows = Devtbl.Rows.Count;
        }
        string out_s = "[";
        foreach (DataRow dr in Devtbl.Rows)
        {
            out_s += "[\"" + dr[0].ToString();
            if (count_rows == 1)
            {
                out_s += "\"]";
            }
            else
            {
                out_s += "\"],";
            }
            count_rows -= 1;
        }
        return out_s;
    }
    public static string pre_name()
    { //dev2--建立廠商清單給Js用
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["IDMSConnectionString"].ConnectionString);
        string sql = @"SELECT [name], [cwb_id]
                        FROM [dbo].[accesslist]
                   ";
        int count_rows = 0;
        DataTable Devtbl = new DataTable();
        using (SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString))
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            conn.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
            dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
            count_rows = Devtbl.Rows.Count;
        }
        string out_s = "[";
        foreach (DataRow dr in Devtbl.Rows)
        {
            out_s += "[" + string.Format("  \"{0}\", \"{1}\" ", dr[0].ToString(), dr[1].ToString());
            if (count_rows == 1)
            {
                out_s += "]";
            }
            else
            {
                out_s += "],";
            }
            count_rows -= 1;
        }
        return out_s;
    }

    protected void Last_Name(){ //建立最近異動清單
    	string sql = @"SELECT [name], [cwb_id], [ID]
                      FROM [dbo].[accesslist]";
        DataTable Devtbl = new DataTable();
        using (SqlConnection conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString))
        using (SqlCommand cmd = new SqlCommand(sql, conn))
        {
            conn.Open();
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//從Command取得資料存入dataAdapter
            dataAdapter.Fill(Devtbl);//將dataAdapter資料存入dataset
            //count_rows = Devtbl.Rows.Count;
        }


   
            foreach (DataRow row in Devtbl.Rows) // 丟進列表裡:(5/10, 王小名)SSD*1
            {
                string cons = row[1].ToString()!=""? String.Format(" <{0}>", row[1].ToString()):"";
                MenuName.Items.Add(new ListItem( String.Format("{0} {1} ", row[0].ToString(), cons) , row[2].ToString()));
            }
        
    }
    protected void MenuName_SelectedIndexChanged(object sender, EventArgs e)//最近一筆異動選取時帶入
    {
        
        ReadName(Int32.Parse(MenuName.SelectedValue));
    
    }

    protected void ReadName(int id)
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand(@"SELECT *
                                        FROM [dbo].[accesslist]
                                        WHERE id=@id   ", Conn);
        
        cmd.Parameters.AddWithValue("@id", id);
        Trace.Write(MenuName.SelectedValue);
        Trace.Write(id.ToString());
		
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        if (dr.Read())
        {   Trace.Write("ID=" + dr["ID"].ToString());
            ID.Text = HttpUtility.HtmlEncode(dr["ID"].ToString());
            Name.Text = HttpUtility.HtmlEncode(dr["Name"].ToString());
            NameID.Text = HttpUtility.HtmlEncode(dr["cwb_id"].ToString());
            Dept.Text = HttpUtility.HtmlEncode(dr["dept"].ToString());
            Company.Text = HttpUtility.HtmlEncode(dr["company"].ToString());
            Job.Text = HttpUtility.HtmlEncode(dr["job"].ToString());
            Telephone.Text = HttpUtility.HtmlEncode(dr["tel"].ToString());
            Check_Access(HttpUtility.HtmlEncode( dr["access"].ToString()));
            Check_Access(HttpUtility.HtmlEncode(dr["access_1f"].ToString()), "1f");
            Check_Access(HttpUtility.HtmlEncode(dr["access_2f"].ToString()), "2f");
            Apdate.Text = HttpUtility.HtmlEncode(dr["ap_date"].ToString());
            Apnumber.Text = HttpUtility.HtmlEncode(dr["ap_number"].ToString());
            UpdateDate.Text = HttpUtility.HtmlEncode(dr["date_modified"].ToString());
        }
        else{
            throw new InvalidCastException("ERROR");
        }
        
        
    }
    protected void BtnAdd_Click(object sender, EventArgs e) //新增按鈕
    {
        Trace.Write("btn on");
        Literal Msg = new Literal();
        if(Null_Check(Msg)){
          
        }
        else{
            List<SqlParameter> pars = new List<SqlParameter>();
            string SQL=GetInsConDevSQL(pars);            //獲取新增至設備資料表語法
            
            string SQL2 = String.Format("新增[{0}]：： ({0}, {1}, {2}, {3}類人員,1F={4}, 2F={5})",//生命履歷語法
                    Name.Text,
                    Dept.Text,
                    Job.Text,
                    Access_Check(),
                    Access_Check("1f"),
                    Access_Check("2f")
                    );
            Trace.Warn(SQL);
            try
            {
                ID.Text = GetPKNo("ID", "accesslist").ToString();
                ExecDbSQL(SQL, pars);
                
                InsLifeSQL(SQL2);
                UpdateDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                Msg.Text = "<script>alert('新增完成! ');</script>";
            }
            catch (System.Exception exx)
            {
                Trace.Warn(exx.ToString());
                Msg.Text = "<script>alert('新增失敗!請重新嘗試\n請注意!姓名+課別名稱 請勿重複');</script>";
            }
        }
        

        Page.Controls.Add(Msg);
    }
    protected void BtnEdit_Click(object sender, EventArgs e) //編輯按鈕
    {   
        Literal Msg = new Literal();
        Trace.Write("EDIT");
        if(Null_Check(Msg)){} //空白確認
        else{        
            List<SqlParameter> pars = new List<SqlParameter>();	
            pars.Add(new SqlParameter("@ID", ID.Text));
            try
            {
                InsLifeSQL("修改 [" + Name.Text + "] ：： " + GetUpdate(ID.Text));
                Trace.Write("修改 [" + Name.Text + "] ：： " + GetUpdate(ID.Text));
                ExecDbSQL("UPDATE [accesslist] SET " + GetUpdate(ID.Text, pars) + " WHERE [ID]= @ID", pars); 
                UpdateDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                Msg.Text = "<script>alert('修改完成')</script>";
            }
            catch (System.Exception ex)
            {     
                Trace.Write(ex.ToString());        
                Msg.Text = "<script>alert('修改失敗，請再試一次');</script>";
            }
        }
        Page.Controls.Add(Msg);
    }
    protected void BtnDel_Click(object sender, EventArgs e) //刪除按鈕
    {
        Literal Msg = new Literal();
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("DELETE FROM [accesslist] WHERE [ID]=@ID", Conn);
        cmd.Parameters.AddWithValue("ID", ID.Text);

        InsLifeSQL(string.Format("刪除[{0}] ： ({1}, {2}類人員, 1F={3}, 2F={4})", 
                                        Name.Text,
                                        Dept.Text,
                                        Access_Check(),
                                        Access_Check("1f"),
                                        Access_Check("2f")
                                        ) );
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
        Msg.Text = "<script>alert('已刪除!');window.close();window.location.replace('access.aspx');</script>";
        Page.Controls.Add(Msg);
    }
    protected string GetInsConDevSQL( List<SqlParameter> pars) //取得新增資料的語法
    {
        //int PointerNo = TextDevNo.Text != ""?int.Parse(TextDevNo.Text):-88;
        pars.Add(new SqlParameter("@name", Name.Text));
        pars.Add(new SqlParameter("@cwbid", NameID.Text));
        pars.Add(new SqlParameter("@tel", Telephone.Text));
        pars.Add(new SqlParameter("@dept", Dept.Text));
        pars.Add(new SqlParameter("@company", Company.Text));
        pars.Add(new SqlParameter("@job", Job.Text));
        pars.Add(new SqlParameter("@access", Access_Check()));
        pars.Add(new SqlParameter("@access1f", Access_Check("1f")));
        pars.Add(new SqlParameter("@access2f", Access_Check("2f")));
        pars.Add(new SqlParameter("@ap_date", Apdate.Text));
        pars.Add(new SqlParameter("@ap_number", Apnumber.Text));
        pars.Add(new SqlParameter("@Time", SqlDbType.DateTime2));
        pars.Last().Value = DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
        return ("INSERT INTO [accesslist] values("
            + "@name" + ","
            + "@cwbid" + ","
            + "@dept" + ","
            + "@company" + ","
            + "@job" + ","
            + "@tel" + ","
            + "@access" + ","
            + "@access1f" + ","
            + "@access2f" + ","
            + "@ap_date" + ","
            + "@ap_number" + ","
            + "@Time )");
    }
    protected string GetUpdate(string ID) //取得修改資料的SQL語法 -- LIFELOG
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand("select * from [accesslist] where [ID]=@ID", Conn);
		cmd.Parameters.AddWithValue("@ID", Int16.Parse(ID));
		
        SqlDataReader dr = cmd.ExecuteReader();
        if (dr.Read())
        {    
            SQL =  GetUpdateCol("name", dr["name"].ToString(), Name.Text, "string");
            SQL += GetUpdateCol("cwb_id", dr["cwb_id"].ToString(), NameID.Text, "string");
            SQL += GetUpdateCol("dept", dr["dept"].ToString(), Dept.Text, "string");
            SQL += GetUpdateCol("company", dr["company"].ToString(), Company.Text, "string");
            SQL += GetUpdateCol("job", dr["job"].ToString(), Job.Text.Trim().Replace("	", " ").Replace("　"," "), "string");
            SQL += GetUpdateCol("tel", dr["tel"].ToString(), Telephone.Text, "string");
            SQL += GetUpdateCol("access", dr["access"].ToString(), Access_Check(), "string");
            SQL += GetUpdateCol("access_1f", dr["access_1f"].ToString(), Access_Check("1f"), "string");
            SQL += GetUpdateCol("access_2f", dr["access_2f"].ToString(), Access_Check("2f"), "string");
            SQL += GetUpdateCol("ap_date", dr["ap_date"].ToString(), Apdate.Text, "string");
            SQL += GetUpdateCol("ap_number", dr["ap_number"].ToString(), Apnumber.Text, "string");
            SQL += GetUpdateCol("date_modified", dr["date_modified"].ToString(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "datetime");

            if (SQL != "") SQL = SQL.Substring(1);  
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        Trace.Write(SQL);
        return (SQL);
    }
    protected string GetUpdate(string ID, List<SqlParameter> pars) //取得修改資料的SQL語法
    {
        string SQL = "";
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();        
		SqlCommand cmd = new SqlCommand("select * from [accesslist] where [ID]=@ID", Conn);
		cmd.Parameters.AddWithValue("@ID", Int16.Parse(ID));
        SqlDataReader dr = cmd.ExecuteReader();
        if (dr.Read())
        {                 
            SQL =  GetUpdateCol("name", dr["name"].ToString(), Name.Text, "string", pars);
            SQL += GetUpdateCol("cwb_id", dr["cwb_id"].ToString(), NameID.Text, "string", pars);
            SQL += GetUpdateCol("dept", dr["dept"].ToString(), Dept.Text, "string", pars);
            SQL += GetUpdateCol("company", dr["company"].ToString(), Company.Text, "string", pars);
            SQL += GetUpdateCol("job", dr["job"].ToString(), Job.Text.Trim().Replace("	", " ").Replace("　"," "), "string", pars);
            SQL += GetUpdateCol("tel", dr["tel"].ToString(), Telephone.Text, "string", pars);
            SQL += GetUpdateCol("access", dr["access"].ToString(), Access_Check(), "string", pars);
            SQL += GetUpdateCol("access_1f", dr["access_1f"].ToString(), Access_Check("1f"), "string", pars);
            SQL += GetUpdateCol("access_2f", dr["access_2f"].ToString(), Access_Check("2f"), "string", pars);
            SQL += GetUpdateCol("ap_date", dr["ap_date"].ToString(), Apdate.Text, "string", pars);
            SQL += GetUpdateCol("ap_number", dr["ap_number"].ToString(), Apnumber.Text, "string", pars);
            SQL += GetUpdateCol("date_modified", dr["date_modified"].ToString(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "datetime", pars);;

            if (SQL != "") SQL = SQL.Substring(1);  
        }
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        Trace.Write("GETUPDATE2" + SQL);
        return (SQL);
    }

    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind, List<SqlParameter> pars) //取得單一欄位修改資料的語法
    {
        string SQL = "";
        if (Source != Target)
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
        return (SQL);
    }
    protected string GetUpdateCol(string ColName, string Source, string Target, string Kind) //取得單一欄位修改資料的語法
    {
        string SQL = "";
        if (Source != Target)
        {
            if (Source == "") Source = "(空白)";
            if (Target == "") Target = "(空白)";
            SQL = SQL + ",[" + ColName + "]：" + Source + " -> " + Target;
        }
        return (SQL);
    }
    

    protected void ExecDbSQL(string SQL, List<SqlParameter> pars = null) //執行資料庫異動
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand(SQL, Conn);
        cmd.Parameters.AddRange(pars.ToArray());
        cmd.ExecuteNonQuery();
        cmd.Cancel(); cmd.Dispose(); Conn.Close(); Conn.Dispose();
    }
    protected void InsLifeSQL(string action) //lifelog
    {

        string who = GetUserName();
        List<SqlParameter> pars = new List<SqlParameter>();               
        pars.Add(new SqlParameter("@ID", GetPKNo("LifeID", "ChangeLog")));
        pars.Add(new SqlParameter("@DevID", ID.Text));
        pars.Add(new SqlParameter("@Time", SqlDbType.DateTime2));
        pars.Last().Value = DateTime.Now.ToString("yyyy'-'MM'-'dd HH':'mm':'ss");
		pars.Add(new SqlParameter("@type", "人員"));
        pars.Add(new SqlParameter("@action", action));
        pars.Add(new SqlParameter("@OP", who));
        pars.Add(new SqlParameter("@IP", Request.ServerVariables["REMOTE_ADDR"].ToString()));	

        ExecDbSQL("INSERT INTO [ChangeLog] values("
            + "@ID" + ","
            + "@type" + ","
            + "@DevID" + ","
            + "@Action" + ","
            + "@OP" + ","
            + "@Time" + ","
            + "@IP )", pars);
    }
    protected bool Null_Check(Literal back){ //空白確認
        if(Name.Text==""){
            back.Text = "<script>alert('尚未填寫員工名稱!');</script>";return true;
        }else if(Dept.Text==""){
            back.Text = "<script>alert('尚未填寫單位課別!');</script>";return true;
        }else if(Job.Text==""){
            back.Text = "<script>alert('尚未填寫業務名稱!');</script>";return true;
        }else if(Access_Check()=="0"){
            back.Text = "<script>alert('尚未選擇授權方式!');</script>";return true;
        }else if(Apnumber.Text.Length > 15 ){
            Trace.Write(Apnumber.Text.Length.ToString());
            back.Text = "<script>alert('表單號碼格式錯誤!請重新填寫!');</script>";return true;
        }else if(Apdate.Text.Length > 12 ){
            back.Text = "<script>alert('表單日期格式錯誤!請重新填寫!');</script>";return true;
        }
        else return false;
    }
    protected int GetPKNo(string PKfield, string PKtbl) //取得主鍵編號
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        SqlCommand cmd = new SqlCommand("select ISNULL(MAX(" + PKfield + "), 0) FROM " + PKtbl, Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int PkNo = 1;
        if (dr.Read()) PkNo = int.Parse(dr[0].ToString()) + 1;
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return (PkNo);
    }
    protected string GetUserName()
    { //將使用者ID -> Name
        string UserID = Request.Cookies["UserID"].Value.ToString();    //自cookie取得
        if (UserID == "") return "無";
        else
        {
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
            else
            {
                return "無";
            }
        }
    }
    protected string Access_Check(string loc = null)
    {
        if(loc == "1f"){
            if (Radio1.Checked)
            {
                return "True";
            }
            else if (Radio2.Checked)
            {
                return "False";
            }
        }
        else if(loc == "2f"){
            if (Radio3.Checked)
            {
                return "True";
            }
            else if (Radio4.Checked)
            {
                return "False";
            }
        }else{
            if (Radio5.Checked)
            {
                return "A";
            }
            else if (Radio6.Checked)
            {
                return "B";
            }
            else if (Radio7.Checked)
            {
                return "C";
            }
        }
        return "0";
        
    }
    protected void Check_Access(string TF, string loc = null)
    {
        
        if(loc == "1f"){
            Trace.Write(TF);
      
            if ( Convert.ToBoolean(TF))
            {
                Radio1.Checked = true;
                Radio2.Checked = false;
            }
            else
            {
                Radio2.Checked = true;
                Radio1.Checked = false;
            }
        }
        else if(loc == "2f"){

            if ( Convert.ToBoolean(TF))
            {
                Radio3.Checked = true;
                Radio4.Checked = false;
            }
            else
            {
                Radio4.Checked = true;
                Radio3.Checked = false;
              
            }
        }else{
    
            if (TF=="A")
            {
                Radio6.Checked = false;
                Radio7.Checked = false;
                Radio5.Checked = true;
                
            }
            else if (TF=="B")
            {
                Radio5.Checked = false;
                Radio7.Checked = false;
                Radio6.Checked = true;
            }
            else if (TF=="C")
            {
                Radio5.Checked = false;
                Radio6.Checked = false;
                Radio7.Checked = true;
            }
        }
    }


}


