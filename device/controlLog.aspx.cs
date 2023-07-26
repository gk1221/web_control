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

    


    public static string Get_Json()
    {
        SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["ControlConnectionString"].ConnectionString);
        Conn.Open();
        string sql = "SELECT COUNT(*) FROM ChangeLog";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }
        
       
        cmd.CommandText = @"SELECT [LifeID]
                            ,[Type]
                            ,[controlID]
                            ,[Action]
                            ,[OP]
                            ,convert(varchar, [Time], 120)
                            ,[IP]
                        FROM [dbo].[ChangeLog]
                ";
        dr.Close();
        dr = cmd.ExecuteReader();
      
        //string out_s="[";
        StringBuilder myStringBuilder = new StringBuilder("[");
        //string out_s="\"data\": [";
        //while(dr.Read())
        while(dr.Read() & count>0)
        {          
            myStringBuilder.Append( "{\"LifeID\":\"" + dr[0].ToString()+
                "\",\"Type\":\"" +  dr[1].ToString() +
                "\",\"controlID\":\"" +  dr[2].ToString() +
                "\",\"Action\":\"" + dr[3].ToString().Replace("\""," ") +
                "\",\"OP\":\"" + dr[4].ToString() +
                "\",\"Time\":\"" + dr[5].ToString() +
                "\",\"IP\":\"" + dr[6].ToString()
                  );              
            
            //myStringBuilder.Append("{\"CreateDate\":\"2004/01/16\",\"ID\":\"1\",\"HostName\":\"cwbftl\",\"HostClass\":\"主機\",\"Functions\":\"VPP5000 console\",\"HW\":\"詹國華\",\"StaffName\":\"詹國華\",\"IO\":\"更換\",\"Xall\":\"HPC專區<br>(24,2)\",\"OP\":\"鄭 柯 " );
            
            if(count==1){
                myStringBuilder.Append( "\"}" );
                
            }
            else{
                myStringBuilder.Append( "\"}," );
            }
            count -= 1;        
        }        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString().Replace("	", " ").Replace("　"," ");
    }

    
}
