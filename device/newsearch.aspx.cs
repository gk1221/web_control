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

    

    
}
