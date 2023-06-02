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

    protected static string f1 (){
        return null;
    }

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

    protected static string OP_split(string op){ //抓OP的第一個字
        string[] sub = op.Split();//用空白分割成array
        string re = "";
        foreach (var item in sub) 
        {
            try{ //避免抓到空白
                re += " " + item[0] ;
            }
            catch (System.Exception)
            {
            }                
        }
        
        return re;
    }
    //產生前端Json資料
    public static string Get_Json()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT COUNT(*) FROM Device2";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }
        
       
        cmd.CommandText = @"SELECT  [DevID]
                ,convert(varchar, CreateDate, 111) as CreateDate
                ,[HostName]
                ,[HostClass]                                        
                ,[IO]
                ,[Functions]
                ,[Memo]
                ,[PS]                                      
                ,[StaffName]
                ,[Hw]                                        
                ,[OP]
                ,ISNULL(w.區域名稱, '定位錯誤') as ar
				,ISNULL(w.定位名稱, '') as lo               
                FROM [control].[dbo].[Device2] d left join [IDMS].dbo.定位設定 w 
                ON d.Xall=w.坐標X AND d.Yall=w.坐標Y
                where 定位方式='坐標' or 定位方式 is null
                ";
        dr.Close();
        dr = cmd.ExecuteReader();
        
        StringBuilder myStringBuilder = new StringBuilder("["); //string加入很慢
        while(dr.Read() & count>0)
        {          
            string ps="", memo="";
            if(dr[6].ToString()!="") ps = "(" + dr[6].ToString() + ")";
            if(dr[7].ToString()!="") memo = "[" + dr[7].ToString() + "]";
            myStringBuilder.Append( "{\"CreateDate\":\"" + dr[1].ToString()+
                "\",\"ID\":\"" +  dr[0].ToString() +
                "\",\"HostName\":\"" +  dr[2].ToString() +
                "\",\"HostClass\":\"" + dr[3].ToString() +
                "\",\"Functions\":\"" + dr[5].ToString() + ps + memo +
                "\",\"HW\":\"" + dr[9].ToString() +
                "\",\"StaffName\":\"" + dr[8].ToString() +
                "\",\"IO\":\"" + Button_check(dr[4].ToString()) +
                "\",\"Xall\":\"" + dr[11].ToString() + "<br>" + dr[12].ToString() +
                "\",\"OP\":\"" + OP_split(dr[10].ToString())  );              
            
            if(count==1){
                myStringBuilder.Append( "\"}" );                
            }
            else{
                myStringBuilder.Append( "\"}," );
            }
            count -= 1;        
        }         
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
    }

    
}
