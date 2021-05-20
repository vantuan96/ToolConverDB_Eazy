using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ToolConvertSqlToMongoDB.Models;

namespace ToolConvertSqlToMongoDB
{
    public partial class IBMS : Form
    {
        private static bool ischeck1 = false;
        private static bool ischeck2 = false;
       
        private static MongoClient connect;
        protected static IMongoClient _client;
        protected static IMongoDatabase _db;
        public IBMS()
        {
            InitializeComponent();
        }
        private void IBMS_Load(object sender, EventArgs e)
        {
            txtServer1.Focus();
            txtServer1.Text = "125.212.226.137";
            txtLogin1.Text = "admin";
            txtPassword1.Text = "Futech12345678";
            txtDatabase1.Text = "MPARKING-TH2";
            txtPort.Text = "712";
            txtServer2.Text = "125.212.226.137";
            txtLogin2.Text = "adminLogin";
            txtPassword2.Text = "kztek123456";
            txtDatabase2.Text = "MPARKING-TH2";

            //cbType.SelectedIndex = 0;
            btnTestConnect2.Enabled = false;

        }
        private void btnTestConnect1_Click(object sender, EventArgs e)
        {
            btnTestConnect1.Enabled = false;

            #region validate
            if (string.IsNullOrEmpty(txtServer1.Text))
            {
                MessageBox.Show("Vui lòng nhập tên server");
                txtServer1.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtLogin1.Text))
            {
                MessageBox.Show("Vui lòng nhập tài khoản");
                txtLogin1.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtPassword1.Text))
            {
                MessageBox.Show("Vui lòng nhập mật khẩu");
                txtPassword1.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtDatabase1.Text))
            {
                MessageBox.Show("Vui lòng nhập tên database");
                txtDatabase1.Focus();
                return;
            }


            #endregion
            var connectionString = "mongodb://" + txtLogin1.Text.Trim() + ":" + txtPassword1.Text.Trim() + "@" + txtServer1.Text.Trim() + ":" + txtPort.Text.Trim() + "/" + txtDatabase1.Text.Trim();
            connect = new MongoClient(connectionString);

            try
            {

                MessageBox.Show("Kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ischeck1 = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + " Vui lòng kiểm tra thông tin kết nối", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ischeck1 = false;
            }
            finally
            {

            }

            btnTestConnect1.Enabled = true;

            //ẩn hiển chuyển đổi dữ liệu
            showExchangeData();

            //ẩn hiện nút test 2
            if (ischeck1)
            {
                btnTestConnect2.Enabled = true;
            }
            else
            {
                btnTestConnect2.Enabled = false;
            }
        }
        //ẩn hiện nút chuyển đổi dữ liệu
        void showExchangeData()
        {
            if (ischeck1 && ischeck2)
            {
                 btnExchangeData.Enabled = true;

            }
            else
            {
               btnExchangeData.Enabled = false;

            }
        }
        private void btnExchangeData_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtServer1.Text.Trim()) && !string.IsNullOrEmpty(txtServer2.Text.Trim()) && ischeck1 && ischeck2)
            {
                Task.Run(() =>
                {
                    btnExchangeData.Invoke(new Action(delegate { btnExchangeData.Enabled = false; }));
                    statusStrip1.Invoke(new Action(delegate
                    {
                        toolStripProgressBar1.MarqueeAnimationSpeed = 30;
                        toolStripProgressBar1.Style = ProgressBarStyle.Marquee;
                    }));


                    ExchangeData(connect);

                    statusStrip1.Invoke(new Action(delegate
                    {
                        toolStripProgressBar1.Style = ProgressBarStyle.Continuous;
                        toolStripProgressBar1.MarqueeAnimationSpeed = 0;
                    }));

                    btnExchangeData.Invoke(new Action(delegate { btnExchangeData.Enabled = true; }));
                });
            }
        }
        public void ExchangeData(MongoClient conn)
        {
            var conString = "mongodb://" + txtLogin1.Text.Trim() + ":" + txtPassword1.Text.Trim() + "@" + txtServer1.Text.Trim() + ":" + txtPort.Text.Trim() ;
            _client = new MongoClient(conString);
            _db = _client.GetDatabase(txtDatabase1.Text.Trim());     
            var path2 = "[" + txtServer2.Text.Trim() + "].[" + txtDatabase2.Text.Trim() + "]";
           // string con1 = "Data Source=125.212.226.137;Initial Catalog=MPARKING-TH2;User Id=adminLogin;Password=kztek123456;";
           string con2 = "Data Source=125.212.226.137;Initial Catalog=MPARKING-TH;User Id=adminLogin;Password=kztek123456;";
            try
            {
               
                var query = new StringBuilder();
               // Chuyển dữ liệu từ trong bảng
                // Lấy dữ liệu từ sqlDb
                #region Nhóm thẻ
                var lst = new List<tblCardGroup>();
                query.AppendLine(string.Format("select * from {0}.dbo.[tblCardGroup]", path2));
                var obj = DBConnect.GetTable(query.ToString(), con2, false);
                foreach (DataRow item in obj.Rows)
                {
                    var obj1 = new tblCardGroup();
                    obj1.CardGroupID = ObjectId.GenerateNewId().ToString();
                    obj1.Id = obj1.CardGroupID;
                    obj1.CardGroupCode = item["CardGroupCode"].ToString();
                    obj1.CardGroupName = item["CardGroupName"].ToString();
                    obj1.Description = item["Description"].ToString();
                    obj1.CardType = Convert.ToInt32(item["CardType"].ToString());
                    obj1.LaneIDs = item["LaneIDs"].ToString();
                    obj1.DayTimeFrom = item["DayTimeFrom"].ToString();
                    obj1.DayTimeTo = item["DayTimeTo"].ToString();
                    obj1.Formulation = Convert.ToInt32(item["Formulation"].ToString());
                    obj1.EachFee = Convert.ToInt32(item["EachFee"].ToString());
                    obj1.Block0 = Convert.ToInt32(item["Block0"].ToString());
                    obj1.Time0 = Convert.ToInt32(item["Time0"].ToString());
                    obj1.Block1 = Convert.ToInt32(item["Block1"].ToString());
                    obj1.Time1 = Convert.ToInt32(item["Time1"].ToString());
                    obj1.Block2 = Convert.ToInt32(item["Block2"].ToString());
                    obj1.Time2 = Convert.ToInt32(item["Time2"].ToString());
                    obj1.Block3 = Convert.ToInt32(item["Block3"].ToString());
                    obj1.Time3 = Convert.ToInt32(item["Time3"].ToString());
                    obj1.Block4 = Convert.ToInt32(item["Block4"].ToString());
                    obj1.Time4 = Convert.ToInt32(item["Time4"].ToString());
                    obj1.Block5 = Convert.ToInt32(item["Block5"].ToString());
                    obj1.Time5 = Convert.ToInt32(item["Time5"].ToString());
                    obj1.TimePeriods = (item["TimePeriods"].ToString());
                    obj1.Costs = (item["Costs"].ToString());
                    obj1.Inactive = string.IsNullOrWhiteSpace(item["Inactive"].ToString().Trim()) ? false : true;
                    obj1.SortOrder = Convert.ToInt32(item["SortOrder"].ToString());
                    obj1.IsHaveMoneyExcessTime = string.IsNullOrWhiteSpace(item["IsHaveMoneyExcessTime"].ToString().Trim()) ? false : true;
                    obj1.EnableFree = string.IsNullOrWhiteSpace(item["EnableFree"].ToString().Trim()) ? false : true;
                    obj1.FreeTime = Convert.ToInt32(item["FreeTime"].ToString());
                    obj1.IsCheckPlate = string.IsNullOrWhiteSpace(item["IsCheckPlate"].ToString().Trim()) ? true : false;
                    obj1.IsHaveMoneyExpiredDate = string.IsNullOrWhiteSpace(item["IsHaveMoneyExpiredDate"].ToString().Trim()) ? false : true;
                    lst.Add(obj1);
                }

                var collection = _db.GetCollection<BsonDocument>("tblCardGroup");
                foreach (var item in lst)
                {
                    collection.InsertOneAsync(item.ToBsonDocument());
                }
              

                #endregion
                #region Khách hàng
                var lstCus = new List<tblCustomer>();
                var str = new StringBuilder();
                str.AppendLine(string.Format("select * from {0}.dbo.[tblCustomer]", path2));
                var objCus = DBConnect.GetTable(str.ToString(), con2, false);
                foreach (DataRow item1 in objCus.Rows)
                {
                    var obj1 = new tblCustomer();
                    obj1.CustomerID = ObjectId.GenerateNewId().ToString();
                    obj1.Id = obj1.CustomerID;
                    obj1.CustomerCode = item1["CustomerCode"].ToString();
                    obj1.CustomerName = item1["CustomerName"].ToString();
                    obj1.Description = item1["Description"].ToString();
                    obj1.Address = item1["Address"].ToString();
                    obj1.IDNumber = item1["IDNumber"].ToString();
                    obj1.Mobile = item1["Mobile"].ToString();
                    obj1.CustomerGroupID = item1["CustomerGroupID"].ToString();
                    obj1.Description = item1["Description"].ToString();
                    obj1.EnableAccount = string.IsNullOrWhiteSpace(item1["EnableAccount"].ToString().Trim()) ? false : true;
                    obj1.Account = item1["Account"].ToString();
                    obj1.Password = item1["Password"].ToString();
                    obj1.Avatar = item1["Avatar"].ToString();
                    obj1.Inactive = string.IsNullOrWhiteSpace(item1["Inactive"].ToString().Trim()) ? false : true;
                    obj1.SortOrder = Convert.ToInt32(item1["SortOrder"].ToString());
                    obj1.CompartmentId = item1["CompartmentId"].ToString();
                    obj1.AccessLevelID = item1["AccessLevelID"].ToString();
                    obj1.Finger1 = item1["Finger1"].ToString();
                    obj1.Finger2 = item1["Finger2"].ToString();
                    obj1.UserIDofFinger = Convert.ToInt32(item1["UserIDofFinger"].ToString());
                    obj1.AccessExpireDate =Convert.ToDateTime(item1["AccessExpireDate"].ToString());
                    obj1.DevPass = item1["DevPass"].ToString();
                    obj1.ContractStartDate =null ;
                    obj1.ContractEndDate = null;
                    obj1.AddressUnsign = item1["AddressUnsign"].ToString();
                   // obj1.MS_PersonGroupId = item1["MS_PersonGroupId"].ToString();
                   // obj1.UserFaceId = Convert.ToInt32(item1["UserFaceId"].ToString());
                   
                 
                    lstCus.Add(obj1);
                }

                var collectionCus = _db.GetCollection<BsonDocument>("tblCustomer");
                foreach (var item1 in lstCus)
                {
                    collectionCus.InsertOneAsync(item1.ToBsonDocument());
                }

                #endregion
                MessageBox.Show("Chuyển đổi dữ liệu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            
        }
        private void btnTestConnect2_Click(object sender, EventArgs e)
        {
            btnTestConnect2.Enabled = false;

            #region validate
            if (string.IsNullOrEmpty(txtServer2.Text))
            {
                MessageBox.Show("Vui lòng nhập tên server");
                txtServer2.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtLogin2.Text))
            {
                MessageBox.Show("Vui lòng nhập tài khoản");
                txtLogin2.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtPassword2.Text))
            {
                MessageBox.Show("Vui lòng nhập mật khẩu");
                txtPassword2.Focus();
                return;
            }
            if (string.IsNullOrEmpty(txtDatabase2.Text))
            {
                MessageBox.Show("Vui lòng nhập tên database");
                txtDatabase2.Focus();
                return;
            }

            #endregion

            try
            {
                if (ischeck1)
                {
                    // ischeck2 = DBConnect.CheckConnection(txtServer2.Text.Trim(), txtLogin2.Text.Trim(), txtPassword2.Text.Trim(), conn1);
                    ischeck2 = DbMongoConnext.CheckConnection(txtServer2.Text.Trim(), txtLogin2.Text.Trim(), txtPassword2.Text.Trim());
                }

                MessageBox.Show("Kết nối thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ischeck2 = false;
            }

            btnTestConnect2.Enabled = true;

             showExchangeData();
        }
        

       
    }
}
