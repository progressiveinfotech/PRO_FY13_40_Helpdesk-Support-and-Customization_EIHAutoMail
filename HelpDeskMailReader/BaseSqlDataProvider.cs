using System;
using System.Data;
using System.Data.Common;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using Microsoft.Practices.EnterpriseLibrary.Data;
using Microsoft.Practices.EnterpriseLibrary.Data.Sql;
using Microsoft.CSharp;
using System.Collections;
using System.Globalization;
using System.Configuration;

/// <summary>
/// Summary description for BaseSqlDataProvider
/// </summary>



//public class BaseSqlDataProvider
public partial class SqlDataProvider
{
    public string _ProvideName = ConfigurationManager.ConnectionStrings["CSM_DB"].ConnectionString;
    protected delegate CollectionBase GenerateCollectionFromReader( ref IDataReader returnData);
    protected delegate object GenerateObjectFromReader(ref IDataReader returnData);

   
     
    //public BaseSqlDataProvider()
    //{
		
    //}
   
   
    public int ExecuteNonQuery(string storedProcedure, object[] @params)
    {

        Database db = default(Database);
        int i = 0;
        try
        {
            db = DatabaseFactory.CreateDatabase(_ProvideName);
            i = db.ExecuteNonQuery(storedProcedure, @params);
        }
        catch (Exception ex)
        {

            //log to db 
            i = -1;
        }
        finally
        {
            db = null;
        }
        
        return i;
      }

    public IDataReader ExecuteReader(string storedProcedure, object[] @params)
    {
        Database db = default(Database);
        try
        {
            db = DatabaseFactory.CreateDatabase(_ProvideName);
            return db.ExecuteReader(storedProcedure, @params);
        }
        catch (Exception ex)
        {
        }
        finally
        {
            //log to db 
            db = null;
        }
        return null;
    }

    public int ExecuteScalar(string storedProcedure, object[] @params)
    {
        Database db = default(Database);
         int  i = 0;
        try
        {
            db = DatabaseFactory.CreateDatabase(_ProvideName);
            i = (int)db.ExecuteScalar(storedProcedure, @params);
                
        }
         catch (Exception ex)
        {

        }
        finally
        {
            
            db = null;
        }

        return i;
    }

   protected  object ExecuteReader(string storedProcedure, object[] @params, GenerateCollectionFromReader fofr)
    {
        Database db = default(Database);
        try
        {
            db = DatabaseFactory.CreateDatabase(_ProvideName);
            IDataReader Idata = (IDataReader)(db.ExecuteReader(storedProcedure, @params));
            return fofr(ref Idata);
        }
        catch (Exception ex)
        {
        }
        finally
        {
            //log to db 
            db = null;
        }
        return null;
    }

   protected object ExecuteReader(string storedProcedure, object[] @params, GenerateObjectFromReader fofr)
   {
       Database db = default(Database);
       try
       {
           db = DatabaseFactory.CreateDatabase(_ProvideName);
           IDataReader Idata = (IDataReader)(db.ExecuteReader(storedProcedure, @params));
           return fofr(ref Idata);
       }
        catch (Exception ex)
       {
       }
       finally
       {
           //log to db 
           db = null;
       }
       return null;
   }

  



    



    


}
