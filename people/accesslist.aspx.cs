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

    protected static string TF_check(string TF)
    {
        if (TF == "True") {
                return("V");//in
		}
		 else {
			return(" ");//other
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
        string sql = "SELECT COUNT(*) FROM accesslist";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }
        
       
        cmd.CommandText = @"SELECT  [ID]
                            ,[name]
                            ,[cwb_id]
                            ,[dept]
                            ,[company]
                            ,[job]
                    
                            ,[access]
                            ,[access_1f]
                            ,[access_2f]
                      
                        FROM [dbo].[accesslist]
                        ORDER BY date_modified desc
                ";
        dr.Close();
        dr = cmd.ExecuteReader();
        
        StringBuilder myStringBuilder = new StringBuilder("["); //string加入很慢
        while(dr.Read() & count>0)
        {          
            myStringBuilder.Append( "{" +
                "\"ID\":\"" +  dr[0].ToString() +
                "\",\"name\":\"" +  dr[1].ToString() +
                "\",\"cwb_id\":\"" + dr[2].ToString() +
                "\",\"dept\":\"" + dr[3].ToString() +
                "\",\"company\":\"" + dr[4].ToString() +
                "\",\"job\":\"" + dr[5].ToString() +
                "\",\"access\":\"" + dr[6].ToString() +
                "\",\"access_1f\":\"" + TF_check(dr[7].ToString()) +
                "\",\"access_2f\":\"" + TF_check(dr[8].ToString()) 
    
                 );              
            
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
