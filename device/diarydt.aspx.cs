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

    protected static string OP_split(string op){
        string[] sub = op.Split();
        string re = "";
        foreach (var item in sub) 
        {
            try{
                re += " " + item[0] ;
            }
            catch (System.Exception)
            {
            }                
        }
        
        return re;
    }

    public static string Get_Json()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["DiaryConnectionString"].ConnectionString);
        Conn.Open();
        string sql = @"SELECT                     [日誌班別]                            
                            ,[日誌編號]
                            ,[訊息編號]
                            ,[日誌時間]
                            ,[處理時間]
                            ,[狀態代碼]
                            ,[根因系統]
                            ,[資產類型]
                            ,[異常等級]
                            ,[發生訊息]
                            ,[訊息代碼]
                            ,[處理編號]
                            ,[處理過程]
                            ,[單位名稱]
                            ,[當班OP]
                            ,[員工姓名]
                        FROM [diary].[dbo].[View_工作日誌]
                        where Year(convert(datetime, 日誌時間 ) ) BETWEEN YEAR(GETDATE())-3 AND YEAR(GETDATE())
                        AND 員工姓名<>'機房OP'
                        AND 訊息代碼 NOT IN('1992d20', '1992d00', '1202d00')                        
                        order by 日誌班別 desc ,日誌編號 ";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        
  
        StringBuilder myStringBuilder = new StringBuilder("[");
        int count = 1, max = 18000;
        while(dr.Read() && count<max)
        {
            
                myStringBuilder.Append( "{\"日誌班別\":\"" + dr[0].ToString()+
                "\",\"日誌編號\":\"" +  dr[1].ToString() +
                "\",\"訊息編號\":\"" +  dr[2].ToString() +
                "\",\"日誌時間\":\"" + dr[3].ToString() +
                "\",\"處理時間\":\"" + dr[4].ToString() +
                "\",\"狀態代碼\":\"" + dr[5].ToString() +
                "\",\"根因系統\":\"" + dr[6].ToString() +
                "\",\"資產類型\":\"" + dr[7].ToString() +
                "\",\"異常等級\":\"" + dr[8].ToString() + 
                "\",\"發生訊息\":\"" + dr[9].ToString().Replace("\\", "\\\\").Replace("\"", " \\\"") + 
                "\",\"訊息代碼\":\"" + dr[10].ToString() + 
                "\",\"處理編號\":\"" + dr[11].ToString() +
                "\",\"處理過程\":\"" + dr[12].ToString().Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r\n", "\\r\\n").Replace("\n","\\n").Replace("\r","\\r") +
                "\",\"單位名稱\":\"" + dr[13].ToString() +
                "\",\"當班OP\":\"" + dr[14].ToString() +
                "\",\"員工姓名\":\"" + dr[15].ToString() 
                ); 
                count += 1;

                if(count >= max){
                    myStringBuilder.Append( "\"}" );
                }
                else{
                    myStringBuilder.Append( "\"}," );
                }
          
              
        } 
                  
        
             
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
    }

    
}
