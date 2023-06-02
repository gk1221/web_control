using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.Services;
using System.Text;
public partial class Device_DevEdit : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
        
        if (!IsPostBack)
        {
            
        
        }
    }

    protected void Page_PreRenderComplete(object sender, EventArgs e)   //放Page_Load無法顯示
    {
       
    
    }
    public  string f1 (){
        List<SqlParameter> pars = new List<SqlParameter>();
        string outt = "";
        string[] lis = Request["Class"].Split(','); //分割class字串
        foreach (var item in lis)
        {
            outt += " AND HostClass=@" + pars.Count();
            pars.Add(new SqlParameter(pars.Count().ToString(), item));
        }
        
        return outt;
    }
//由此進入後端
    public string Get_Json(){
        if(Request.QueryString["Time"] != null) return One_Date();
        else {            
            //List<SqlParameter> pars = new List<SqlParameter>();return Con_Add(pars);
            return Multi_Sel();
        }
    }
    protected  string Con_Add (List<SqlParameter> pars){
        string Addition_s = "";
        if(Request.QueryString["Class"]!=null){
            Addition_s +=  " AND (";
            string[] lis1 = Request["Class"].Split(','); //分割種類字串
            for(int i=0;i<lis1.Length;i++){
                Addition_s += " HostClass=@" + pars.Count();
                pars.Add(new SqlParameter(pars.Count().ToString(), lis1[i]));
                if(i!=lis1.Length-1)Addition_s += " OR "; //最後一項不加OR
            }
            Addition_s += " ) ";
        }
        
        if(Request.QueryString["Check"]!=null){
            string[] lis2 = Request["Check"].Split(','); //分割異動字串
            Addition_s += " AND (";
            for(int i=0;i<lis2.Length;i++){
                Addition_s += " IO=@" + pars.Count();
                pars.Add(new SqlParameter(pars.Count().ToString(), lis2[i]));
                if(i!=lis2.Length-1)Addition_s += " OR "; //最後一項不加OR
            }
            Addition_s += " ) ";
        }
        
        if(Request.QueryString["Find"]!=null){
            string[] lis3 = Request["Find"].Split(' ');//分割搜尋字串
            string anor = Request["Type"]=="A"?"AND":"OR";//聯集交集
            
            Addition_s += " AND (";
            for(int i=0;i<lis3.Length;i++){ 

                Addition_s += " (";
                Addition_s += " HostName LIKE @" + pars.Count().ToString() + " OR ";
                Addition_s += " Functions LIKE @" + pars.Count().ToString() + " OR ";
                Addition_s += " HW LIKE @" + pars.Count().ToString() + " OR " ;
                Addition_s += " StaffName LIKE @" + pars.Count().ToString()  ;
                Addition_s += ") ";

                pars.Add(new SqlParameter(pars.Count().ToString(), "%" + lis3[i] + "%"));
                
                if(i!=lis3.Length-1)Addition_s += anor; //最後一項不加OR/AND
            }
            Addition_s += ") ";            
        }        
        return Addition_s ;        
    }
//資料庫的進出轉成文字
    protected static string Button_check(string IO)
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
//取得OP的姓
    protected static string OP_split(string op){
        string[] sub = op.Split();
        string re = "";
        foreach (var item in sub) 
        {
            try{
                re += " " + item[0] ;
            }
            catch (System.Exception){}                
        }
        return re;
    }

    
//單月報表方法
    protected string One_Date()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT COUNT(*) FROM [control].[dbo].[Device2] WHERE Year(CreateDate)=@YYY AND MONTH(CreateDate)=@MMM";
        SqlCommand cmd = new SqlCommand(sql, Conn);
        string[] YYMM = new string[3];
        YYMM = Request["Time"].Split('/');
        //若輸入錯誤則回傳空白陣列        
        if(YYMM.Length!=2)
        {              
            return "[";
        }
        cmd.Parameters.AddWithValue("YYY", YYMM[0]);
        cmd.Parameters.AddWithValue("MMM", YYMM[1]);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        
        if(dr.Read()) count = int.Parse(dr[0].ToString());
        
        cmd.CommandText = @"SELECT
                            convert(varchar, CreateDate, 111) as CreateDate
                            ,[HostName]
                            ,[HostClass]
                            ,[Repair]  
                            ,[Functions]
                            ,[Hw]  
                            ,[StaffName]                                  
                            ,[IO]                                                                               
                            ,CONCAT(w.區域名稱,'<br>', w.定位名稱) as area                                                                                 
                            ,[OP]                                                                              
                            FROM [control].[dbo].[Device2] d left join [IDMS].dbo.定位設定 w 
                            ON d.Xall=w.坐標X2 AND d.Yall=w.坐標Y2 AND 定位方式='坐標'
                            WHERE Year(CreateDate)=@YYY AND MONTH(CreateDate)=@MMM
                            ORDER BY DevID DESC";
        dr.Close();
        dr = cmd.ExecuteReader();
        StringBuilder out_s=new StringBuilder("[");
        //string out_s="\"data\": [";
        Make_JSON(dr, count, out_s);
        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return out_s.ToString();
    }
    protected string Make_JSON(SqlDataReader dr, int count ,StringBuilder myStringBuilder){
        while(dr.Read() & count>0){
            myStringBuilder.Append( "{\"建檔\":\"" + dr[0].ToString().Substring(2,8)+                    
                    "\",\"設備名稱\":\"" +  dr[1].ToString() +
                    "\",\"設備種類\":\"" + dr[2].ToString() +
                    "\",\"設備廠商\":\"" + dr[3].ToString() +
                    "\",\"設備用途\":\"" + dr[4].ToString() +
                    "\",\"硬體\":\"" + dr[5].ToString() +
                    "\",\"申請\":\"" + dr[6].ToString() +
                    "\",\"異動\":\"" + Button_check(dr[7].ToString()) +
                    "\",\"位置\":\"" + dr[8].ToString() +
                    "\",\"OP\":\"" + OP_split(dr[9].ToString()));              
            if(count==1){
                myStringBuilder.Append( "\"}" );
                
            }
            else{
                myStringBuilder.Append( "\"}," );
            }
            count -= 1;        
        }
        return null;
    }

    protected string Multi_Sel(){
        string from = Request["From"];
        string to = Request["To"];
        string sql = "SELECT COUNT(*) FROM [control].[dbo].[Device2] WHERE CreateDate between @from and @to";   
        List<SqlParameter> pars = new List<SqlParameter>();
        pars.Add(new SqlParameter("from", from));
        pars.Add(new SqlParameter("to", to));
        string moreCon = Con_Add(pars);
        sql += moreCon;
        
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        
        SqlCommand cmd = new SqlCommand(sql, Conn);
        cmd.Parameters.AddRange(pars.ToArray());
        SqlDataReader dr = cmd.ExecuteReader();
        int count = 0;        
        if(dr.Read()) count = int.Parse(dr[0].ToString());
        
        cmd.CommandText = @"SELECT
                            convert(varchar, CreateDate, 111) as CreateDate
                            ,[HostName]
                            ,[HostClass]
                            ,[Repair]  
                            ,[Functions]
                            ,[Hw]  
                            ,[StaffName]                                  
                            ,[IO]                                                                               
                            ,CONCAT(w.區域名稱,'<br>', w.定位名稱) as area                                                                                 
                            ,[OP]                                                                              
                            FROM [control].[dbo].[Device2] d left join [IDMS].dbo.定位設定 w 
                            ON d.Xall=w.坐標X2 AND d.Yall=w.坐標Y2 AND 定位方式='坐標'
                            WHERE CreateDate BETWEEN @from AND @to ";
        cmd.CommandText += moreCon;
        dr.Close();
        dr = cmd.ExecuteReader();
        StringBuilder myStringBuilder = new StringBuilder("[");

        Make_JSON(dr, count, myStringBuilder);

        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
        
    }
    
}
