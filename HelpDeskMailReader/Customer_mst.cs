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
/// Summary description for Customer_mst
/// </summary>
public class Customer_mst
{
	    #region Private Fields
		private string _emailid;
		private string _customer_name;
		private int _custid;
		private DateTime _createdatetime;
		private string _contact_person;
		private string _contact_no;
		private string _address;
		#endregion

		#region Public Fields
		public string Emailid
		{
			 get{ return _emailid; }
			 set{ _emailid = value; }
		}
		public string Customer_name
		{
			 get{ return _customer_name; }
			 set{ _customer_name = value; }
		}
		public int Custid
		{
			 get{ return _custid; }
			 set{ _custid = value; }
		}
		public DateTime Createdatetime
		{
			 get{ return _createdatetime; }
			 set{ _createdatetime = value; }
		}
		public string Contact_person
		{
			 get{ return _contact_person; }
			 set{ _contact_person = value; }
		}
		public string Contact_no
		{
			 get{ return _contact_no; }
			 set{ _contact_no = value; }
		}
		public string Address
		{
			 get{ return _address; }
			 set{ _address = value; }
		}
		#endregion

		#region Constructors
		public Customer_mst()		{
		}
		public Customer_mst(string emailid,string customer_name,int custid,DateTime createdatetime,string contact_person,string contact_no,string address)		{
			_emailid = emailid;
			_customer_name = customer_name;
			_custid = custid;
			_createdatetime = createdatetime;
			_contact_person = contact_person;
			_contact_no = contact_no;
			_address = address;
		}
		#endregion

        #region Public Methods
        //public int Insert()
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Insert_Customer_mst(this);

        //}

        //public int Update()
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Update_Customer_mst_By_id(this);
        //}
        //public int Delete(int Customerid)
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Delete_Customer_mst_By_id(Customerid);
        //}
        //public BLLCollection<Customer_mst> Get_All()
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Get_Customer_mst_All();
        //}

        //public Customer_mst Get_By_id(int Customerid)
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Customer_mst_Get_By_Customerid(Customerid);
        //}

        //public Customer_mst Get_Customer_By_customername(string customername)
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Get_Customer_By_customername(customername);
        //}

        //public int Get_Customer_By_Custid(string Customername)
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Check_Customer_By_Custid(Customername);
        //}

        //public int Get_TopCustId()
        //{
        //    SqlDataProvider db = new SqlDataProvider();
        //    return db.Get_TopCustomerId();
        //}

        #endregion
}
