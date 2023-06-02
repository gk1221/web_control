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
public partial class Check_list : System.Web.UI.Page
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
        string sql = "SELECT COUNT(*) FROM [control_dev2].[dbo].[Check]";

		SqlCommand cmd = new SqlCommand(sql,Conn);
        SqlDataReader dr = null;
        dr = cmd.ExecuteReader();
        int count = 0;
        if(dr.Read()){
            count = int.Parse(dr[0].ToString());
        }
        
       
        cmd.CommandText = @"SELECT [ID]
                            ,[title]
                            ,[check]
                            , CONVERT(VARCHAR(16), deadline, 120) AS formatted_datetime
                            ,[context]
                            FROM [dbo].[Check]
                            ORDER BY [check] asc, formatted_datetime asc
                ";
        dr.Close();
        dr = cmd.ExecuteReader();
      
        //string out_s="[";
        StringBuilder myStringBuilder = new StringBuilder("[");
        //string out_s="\"data\": [";
        //while(dr.Read())
        while(dr.Read() & count>0)
        {          
            myStringBuilder.Append( "{\"ID\":\"" + dr[0].ToString()+
                "\",\"title\":\"" +  dr[1].ToString() +
                "\",\"check\":\"" +  dr[2].ToString() +
                  "\",\"context\":\"" + dr[4].ToString() +
                "\",\"deadline\":\"" + dr[3].ToString() 
            
                  );              
          
            if(count==1){
                myStringBuilder.Append( "\"}]" );
                
            }
            else{
                myStringBuilder.Append( "\"}," );
            }
            count -= 1;        
        }        
        cmd.Cancel(); cmd.Dispose(); dr.Close(); Conn.Close(); Conn.Dispose();
        return myStringBuilder.ToString();
    }

    protected void Btn_addclick(object sender, EventArgs e){
        using (SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["controlConnectionString"].ConnectionString))
        {
            SqlCommand cmd =
                new SqlCommand(@"INSERT INTO [dbo].[check] 
                                VALUES(@ti, @co, @ch, @dl)", Conn);
            Conn.Open();
            cmd.Parameters.AddWithValue("@ti", Titlea.Text);
            cmd.Parameters.AddWithValue("@co", context.Text);
            cmd.Parameters.AddWithValue("@ch", TF.SelectedValue);
            cmd.Parameters.AddWithValue("@dl", deadline.Text);

            cmd.ExecuteNonQuery();
           
        }
    }

    protected void Btn_deleteclick(object sender, EventArgs e){
        using (SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["controlConnectionString"].ConnectionString))
        {
            SqlCommand cmd =
                new SqlCommand(@"DELETE FROM [dbo].[check] 
                                WHERE ID = @id", Conn);
            Conn.Open();
            cmd.Parameters.AddWithValue("@id", checkID.Text);
           

            cmd.ExecuteNonQuery();
           
        }
    }

    protected void Btn_editclick(object sender, EventArgs e){
        using (SqlConnection Conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["controlConnectionString"].ConnectionString))
        {
            SqlCommand cmd =
                new SqlCommand(@"UPDATE [dbo].[check]
                                SET [title]=@ti, [context]=@co, [check]=@ch, [deadline]=@dl
                                WHERE ID=@id", Conn);
            Conn.Open();
            
            cmd.Parameters.AddWithValue("@id", checkID.Text);
            cmd.Parameters.AddWithValue("@ti", Titlea.Text);
            cmd.Parameters.AddWithValue("@co", context.Text);
            cmd.Parameters.AddWithValue("@ch", TF.SelectedValue);
            cmd.Parameters.AddWithValue("@dl", deadline.Text);
            Trace.Write(checkID.Text);
            Trace.Write(TF.SelectedValue);

            cmd.ExecuteNonQuery();
           
        }
    }
}