using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;

/// <summary>
/// Summary description for UserToSiteMapping
/// </summary>
public class UserToSiteMapping
{
    #region Private Fields
    private int _userid;
    private int _siteid;
    #endregion

    #region Public Fields
    public int Userid
    {
        get { return _userid; }
        set { _userid = value; }
    }
    public int Siteid
    {
        get { return _siteid; }
        set { _siteid = value; }
    }
    #endregion

    #region Constructors
    public UserToSiteMapping()
    {
        
    }
   
    public UserToSiteMapping(int userid, int siteid)
    {
        _userid = userid;
        _siteid = siteid;
    }
    #endregion

    #region Public Methods
    public int Insert()
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Insert_UserToSiteMapping_mst(this);

    }

    //public int Delete(int userid,int siteid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Delete_UserToSiteMapping_mst_By_id(userid, siteid);
    //}

    //public BLLCollection<UserToSiteMapping> Get_All()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_UserToSiteMapping_mst_All();
    //}
    public int Get_By_Id(int userid, int siteid)
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Get_UserId_mst_Get_By_SiteId(userid, siteid);
    }
    public BLLCollection<UserToSiteMapping> Get_All_By_userid(int userid)
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Get_UserToSiteMapping_mst_All_By_Userid(userid);
    }
    public BLLCollection<UserToSiteMapping> Get_All_By_siteid(int siteid)
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Get_UserToSiteMapping_mst_All_By_siteid(siteid);
    }

    #endregion

	
}
