using Framework.Data.OM;
using Framework.Data.OM.OMAttribute;
using System;

namespace CSIPCommonModel.EntityLayer_new
{
    /// <summary>
    /// EntityL_BATCH_LOG
    /// </summary>
    [Serializable()]
    [AttributeTable("L_AP_LOG")]
    public class EntityL_AP_LOG : Entity
    {

        private string _System_Code;
        /// <summary>
        /// 系統別
        /// </summary>        
        [AttributeField("System_Code", "System.String", false, false, false, "String")]
        public string System_Code
        {
            get
            {
                return this._System_Code;
            }
            set
            {
                this._System_Code = value;
            }
        }
        private string _Login_Account_Nbr;
        /// <summary>
        /// 使用者帳號(員工編號)
        /// </summary>        
        [AttributeField("Login_Account_Nbr", "System.String", false, false, false, "String")]
        public string Login_Account_Nbr
        {
            get
            {
                return this._Login_Account_Nbr;
            }
            set
            {
                this._Login_Account_Nbr = value;
            }
        }
        private string _Query_Datetime;
        /// <summary>
        /// 查詢日期時間
        /// </summary>        
        [AttributeField("Query_Datetime", "System.DateTime", false, false, false, "DateTime")]
        public string Query_Datetime
        {
            get
            {
                return this._Query_Datetime;
            }
            set
            {
                this._Query_Datetime = value;
            }
        }
        private string _AP_Txn_Code;
        /// <summary>
        /// 交易代號或畫面代碼
        /// </summary>        
        [AttributeField("AP_Txn_Code", "System.String", false, false, false, "String")]
        public string AP_Txn_Code
        {
            get
            {
                return this._AP_Txn_Code;
            }
            set
            {
                this._AP_Txn_Code = value;
            }
        }
        private string _Server_Name;
        /// <summary>
        /// 所連線目標資料庫/伺服器IP
        /// </summary>        
        [AttributeField("Server_Name", "System.String", false, false, false, "String")]
        public string Server_Name
        {
            get
            {
                return this._Server_Name;
            }
            set
            {
                this._Server_Name = value;
            }
        }
        private string _User_Terminal;
        /// <summary>
        /// 使用者終端設備Id
        /// </summary>
        [AttributeField("User_Terminal", "System.String", false, false, false, "String")]
        public string User_Terminal
        {
            get
            {
                return this._User_Terminal;
            }
            set
            {
                this._User_Terminal = value;
            }
        }
        private string _AP_Account_Nbr;
        /// <summary>
        /// 使用之AP帳號
        /// </summary>
        [AttributeField("AP_Account_Nbr", "System.String", false, false, false, "String")]
        public string AP_Account_Nbr
        {
            get
            {
                return this._AP_Account_Nbr;
            }
            set
            {
                this._AP_Account_Nbr = value;
            }
        }
        private string _Txn_Type_Code;
        /// <summary>
        /// 動作類別
        /// </summary>
        [AttributeField("Txn_Type_Code", "System.String", false, false, false, "String")]
        public string Txn_Type_Code
        {
            get
            {
                return this._Txn_Type_Code;
            }
            set
            {
                this._Txn_Type_Code = value;
            }
        }
        private string _Statement_Text;
        /// <summary>
        /// 執行參數: 用 | 區隔,若無值仍需帶,切記不要內含 , 
        /// </summary>
        [AttributeField("Statement_Text", "System.String", false, false, false, "String")]
        public string Statement_Text
        {
            get
            {
                return this._Statement_Text;
            }
            set
            {
                this._Statement_Text = value;
            }
        }
        private string _Object_Name;
        /// <summary>
        /// 所存取之檔案/物件名稱
        /// </summary>
        [AttributeField("Object_Name", "System.String", false, false, false, "String")]
        public string Object_Name
        {
            get
            {
                return this._Object_Name;
            }
            set
            {
                this._Object_Name = value;
            }
        }
        private string _Txn_Status_Code;
        /// <summary>
        /// 執行狀態:Y-成功,N-失敗,T:TIME OUT
        /// </summary>
        [AttributeField("Txn_Status_Code", "System.String", false, false, false, "String")]
        public string Txn_Status_Code
        {
            get
            {
                return this._Txn_Status_Code;
            }
            set
            {
                this._Txn_Status_Code = value;
            }
        }
        private string _Customer_Id;
        /// <summary>
        /// Customer_Id:客戶ID/ 統編
        /// </summary>
        [AttributeField("Customer_Id", "System.String", false, false, false, "String")]
        public string Customer_Id
        {
            get
            {
                return this._Customer_Id;
            }
            set
            {
                this._Customer_Id = value;
            }
        }
        private string _Account_Nbr;
        /// <summary>
        /// 交易帳號/卡號:AC_NO
        /// </summary>
        [AttributeField("Account_Nbr", "System.String", false, false, false, "String")]
        public string Account_Nbr
        {
            get
            {
                return this._Account_Nbr;
            }
            set
            {
                this._Account_Nbr = value;
            }
        }
        private string _Branch_Nbr;
        /// <summary>
        /// 帳務分行別
        /// </summary>
        [AttributeField("Branch_Nbr", "System.String", false, false, false, "String")]
        public string Branch_Nbr
        {
            get
            {
                return this._Branch_Nbr;
            }
            set
            {
                this._Branch_Nbr = value;
            }
        }
        private string _Role_Id;
        /// <summary>
        /// 登入系統帳號之角色(需注意長度)
        /// </summary>        
        [AttributeField("Role_Id", "System.String", false, false, false, "String")]
        public string Role_Id
        {
            get
            {
                return this._Role_Id;
            }
            set
            {
                this._Role_Id = value;
            }
        }
        private string _Import_Source;
        /// <summary>
        /// 資料來源
        /// </summary>
        [AttributeField("Import_Source", "System.String", false, false, false, "String")]
        public string Import_Source
        {
            get
            {
                return this._Import_Source;
            }
            set
            {
                this._Import_Source = value;
            }
        }
        private string _As_Of_Date;
        /// <summary>
        /// 資料日期:資料日YYYYMMDD，例如：提供12/2當日的資料，此欄位則填入值20151202
        /// </summary>
        [AttributeField("As_Of_Date", "System.String", false, false, false, "String")]
        public string As_Of_Date
        {
            get
            {
                return this._As_Of_Date;
            }
            set
            {
                this._As_Of_Date = value;
            }
        }

    }    
}
