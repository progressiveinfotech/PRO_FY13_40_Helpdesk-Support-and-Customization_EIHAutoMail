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
/// Summary description for Organization_mst
/// </summary>
public class Organization_mst
{

    #region Private Fields
    private string _orgname;
    private int _orgid;
    private string _description;
    private string _createdatetime;
    #endregion
   
    #region Public Fields
    public string Orgname
    {
        get { return _orgname; }
        set { _orgname = value; }
    }
    public int Orgid
    {
        get { return _orgid; }
        set { _orgid = value; }
    }
    public string Description
    {
        get { return _description; }
        set { _description = value; }
    }
    public string Createdatetime
    {
        get { return _createdatetime; }
        set { _createdatetime = value; }
    }
    #endregion

    #region Constructors
    public Organization_mst()
    {
    }
    public Organization_mst(string orgname, int orgid, string description, String createdatetime)
    {
        _orgname = orgname;
        _orgid = orgid;
        _description = description;
        _createdatetime = createdatetime;
    }
    #endregion

    #region Public Methods
    
   
    public BLLCollection<Organization_mst> Get_All()
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Get_Organization_mst_All();
    }

    public Organization_mst Get_Organization()
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Get_Organization_mst();
    } 
   
    #endregion
}
