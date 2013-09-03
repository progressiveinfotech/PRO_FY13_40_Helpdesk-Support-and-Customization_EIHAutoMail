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
/// Summary description for CustomerToSiteMapping
/// </summary>
public class CustomerToSiteMapping
{
    #region Private Fields
    private int _siteid;
    private int _custid;
    #endregion

    #region Public Fields
    public int Siteid
    {
        get { return _siteid; }
        set { _siteid = value; }
    }
    public int Custid
    {
        get { return _custid; }
        set { _custid = value; }
    }
    #endregion

    #region Constructors
    public CustomerToSiteMapping()
    {
    }
    public CustomerToSiteMapping(int siteid, int custid)
    {
        _siteid = siteid;
        _custid = custid;
    }
    #endregion



    #region Public Methods
    //public int Insert()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Insert_CustomerToSiteMapping_mst(this);

    //}

    //public int Delete(int CustId, int siteid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Delete_CustomerToSiteMapping_mst_By_id(CustId, siteid);
    //}

    //public BLLCollection<CustomerToSiteMapping> Get_All()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_CustomerToSiteMapping_mst_All();
    //}
    //public int Get_By_Id(int CustId, int siteid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_CustId_mst_Get_By_SiteId(CustId, siteid);
    //}
    //public BLLCollection<CustomerToSiteMapping> Get_All_By_CustId(int CustId)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_CustomerToSiteMapping_mst_All_By_CustId(CustId);
    //}
    //public BLLCollection<CustomerToSiteMapping> Get_All_By_siteid(int siteid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_CustomerToSiteMapping_mst_All_By_siteid(siteid);
    //}




    ////////Added by vishal 18-05-2012

    //public BLLCollection<Customer_mst> Get_UserName_All_By_Likename(string username)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_UserName_All_By_Likename(username);
    //}

    #endregion


}
