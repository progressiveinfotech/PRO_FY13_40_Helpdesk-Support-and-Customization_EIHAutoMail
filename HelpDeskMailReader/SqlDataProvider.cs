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
using System.Collections;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Xml.Linq;

/// <summary>
/// Summary description for SqlDataProvider
/// </summary>


public partial class SqlDataProvider
{
    private const string Sp_Get_Organization_All = "sp_Get_Organization_All_mst";
    private const string Sp_Get_Organization = "sp_Get_Organization_mst";
    private const string Sp_Get_UserId_By_SiteId = "sp_Get_UserId_By_SiteId_mst";
    private const string Sp_Get_UserToSiteMapping_All_By_userid = "sp_Get_UserToSiteMapping_All_By_userid_mst";
    private const string Sp_Get_UserToSiteMapping_All_By_siteid = "sp_Get_UserToSiteMapping_All_By_siteid_mst";
    private const string Sp_Get_UserId_By_UserName = "sp_Get_UserId_By_UserName_mst";
    private const string Sp_UserLogin_Insert = "sp_UserLogin_Insert_mst";
    private const string Sp_ContactInfo_Insert = "sp_ContactInfo_Insert_mst";
    private const string Sp_UserToSiteMapping_Insert = "sp_UserToSiteMapping_Insert_mst";

    public int Insert_UserToSiteMapping_mst(UserToSiteMapping objUserToSiteMapping)
    {
        return (int)ExecuteNonQuery(Sp_UserToSiteMapping_Insert, new object[] { objUserToSiteMapping.Userid, objUserToSiteMapping.Siteid });
    }
    public int Insert_ContactInfo_mst(ContactInfo_mst objContactInfo)
    {
        return (int)ExecuteNonQuery(Sp_ContactInfo_Insert, new object[] { objContactInfo.Userid, objContactInfo.Mobile, objContactInfo.Lastname, objContactInfo.Landline, objContactInfo.Firstname, objContactInfo.Empid, objContactInfo.Emailid, objContactInfo.Description, objContactInfo.Siteid, objContactInfo.Deptid });
    }
    public int Insert_UserLogin_mst(UserLogin_mst objUserLogin)
    {
        return (int)ExecuteNonQuery(Sp_UserLogin_Insert, new object[] { objUserLogin.Username, objUserLogin.Userid, objUserLogin.Roleid, objUserLogin.Password, objUserLogin.Orgid, objUserLogin.Enable, objUserLogin.Createdatetime, objUserLogin.ADEnable, objUserLogin.DomainName, objUserLogin.Company, objUserLogin.City });
    }

    public BLLCollection<UserToSiteMapping> Get_UserToSiteMapping_mst_All_By_Userid(int userid)
    {
        return (BLLCollection<UserToSiteMapping>)ExecuteReader(Sp_Get_UserToSiteMapping_All_By_userid, new object[] { userid }, new GenerateCollectionFromReader(GenerateUserToSiteMapping_mstCollection));
    }
    public int Get_UserId_mst_Get_By_UserName(string UserName, int orgid)
    {
        return (int)ExecuteScalar(Sp_Get_UserId_By_UserName, new object[] { UserName, orgid });
    }
    public CollectionBase GenerateUserToSiteMapping_mstCollection(ref IDataReader returnData)
    {
        BLLCollection<UserToSiteMapping> col = new BLLCollection<UserToSiteMapping>();
        while (returnData.Read())
        {


            UserToSiteMapping obj = new UserToSiteMapping();
            obj.Userid = (int)returnData["Userid"];
            obj.Siteid = (int)returnData["Siteid"];
            col.Add(obj);
        }
        returnData.Close();
        returnData.Dispose();
        return col;
    }
    public BLLCollection<UserToSiteMapping> Get_UserToSiteMapping_mst_All_By_siteid(int siteid)
    {
        return (BLLCollection<UserToSiteMapping>)ExecuteReader(Sp_Get_UserToSiteMapping_All_By_siteid, new object[] { siteid }, new GenerateCollectionFromReader(GenerateUserToSiteMapping_mstCollection));
    }
    public BLLCollection<Organization_mst> Get_Organization_mst_All()
    {
        return (BLLCollection<Organization_mst>)ExecuteReader(Sp_Get_Organization_All, new object[] { }, new GenerateCollectionFromReader(GenerateOrganization_mstCollection));
    }

    public CollectionBase GenerateOrganization_mstCollection(ref IDataReader returnData)
    {
        BLLCollection<Organization_mst> col = new BLLCollection<Organization_mst>();
        while (returnData.Read())
        {
            DateTime Mydatetime = new DateTime();

            Organization_mst obj = new Organization_mst();
            obj.Orgid = (int)returnData["Orgid"];
            obj.Orgname = (string)returnData["Orgname"];
            obj.Description = (string)returnData["Description"];
            Mydatetime = (DateTime)returnData["Createdatetime"];
            obj.Createdatetime = Mydatetime.ToString();
            col.Add(obj);
        }
        returnData.Close();
        returnData.Dispose();
        return col;
    }

    public int Get_UserId_mst_Get_By_SiteId(int userid, int siteid)
    {
        return (int)ExecuteScalar(Sp_Get_UserId_By_SiteId, new object[] { userid, siteid });
    }

    public Organization_mst Get_Organization_mst()
    {
        return (Organization_mst)ExecuteReader(Sp_Get_Organization, new object[] { }, new GenerateObjectFromReader(GenerateOrganization_mstObject));
    }

    public object GenerateOrganization_mstObject(ref IDataReader returnData)
    {
        Organization_mst obj = new Organization_mst();
        while (returnData.Read())
        {
            DateTime Mydatetime = new DateTime();
            obj.Orgid = (int)returnData["Orgid"];
            obj.Orgname = (string)returnData["Orgname"];
            obj.Description = (string)returnData["Description"];
            Mydatetime = (DateTime)returnData["Createdatetime"];
            obj.Createdatetime = Mydatetime.ToString();

        }
        returnData.Close();
        returnData.Dispose();
        return obj;
    }

}
