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
/// Summary description for ContactInfo_mst
/// </summary>
public class ContactInfo_mst
{
    #region Private Fields
    private int _userid;
    private string _mobile;
    private string _lastname;
    private string _landline;
    private string _firstname;
    private string _empid;
    private string _emailid;
    private string _description;
   
    private int _siteid;
    private int _deptid;
    private int _userpriority;
    #endregion

    #region Public Fields
    public int Userid
    {
        get { return _userid; }
        set { _userid = value; }
    }
    public string Mobile
    {
        get { return _mobile; }
        set { _mobile = value; }
    }
   
    public string Lastname
    {
        get { return _lastname; }
        set { _lastname = value; }
    }
    public string Landline
    {
        get { return _landline; }
        set { _landline = value; }
    }
    public string Firstname
    {
        get { return _firstname; }
        set { _firstname = value; }
    }
    public string Empid
    {
        get { return _empid; }
        set { _empid = value; }
    }
    public string Emailid
    {
        get { return _emailid; }
        set { _emailid = value; }
    }
    public string Description
    {
        get { return _description; }
        set { _description = value; }
    }
    
    public int Siteid
    {
        get { return _siteid; }
        set { _siteid = value; }
    }
    public int Deptid
    {
        get { return _deptid; }
        set { _deptid = value; }
    }
    public int userpriority
    {
        get { return _userpriority; }
        set { _userpriority = value; }
    }
    #endregion


    #region Constructors
    public ContactInfo_mst()
    {
    }
    public ContactInfo_mst(int userid, string mobile, string lastname, string landline, string firstname, string empid, string emailid, string description, int siteid, int deptid, int userpriority)
    {
        _userid = userid;
        _mobile = mobile;
        _lastname = lastname;
        _landline = landline;
        _firstname = firstname;
        _empid = empid;
        _emailid = emailid;
        _description = description;
        _siteid = siteid;
        _deptid = deptid;
        _userpriority = userpriority;
    }
    #endregion

    #region Public Methods
    public int Insert()
    {
        SqlDataProvider db = new SqlDataProvider();
        return db.Insert_ContactInfo_mst(this);
    }
    //public int Update()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Update_ContactInfo_mst_By_id(this);
    //}
    //public int Delete(int userid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Delete_ContactInfo_mst_By_id(userid);
    //}
    //public ContactInfo_mst Get_By_id(int userid)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.ContactInfo_mst_Get_By_Userid(userid);
    //}


    //public BLLCollection<ContactInfo_mst> Get_All()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
       
    //    return db.Get_ContactInfo_All();
    //}
    ////added by lalit---------------------------------
    //public IDataReader Get_By_siteid()
    //{
    //    SqlDataProvider db = new SqlDataProvider();
    //    return db.Get_contactInfo_ms_By_siteid();
    //}
    //public BLLCollection<ContactInfo_mst> Get_By_comandname(string commandname)
    //{
    //    SqlDataProvider db = new SqlDataProvider();
      
    //    return db.Get_ContactInfo_ms_By_commandname(commandname);
    //}
    
    #endregion
}
