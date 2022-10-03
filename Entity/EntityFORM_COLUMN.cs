using System;
using System.Collections.Generic;
using System.Text;
using Framework.Data.OM.OMAttribute;
using Framework.Data.OM;
using Framework.Data.OM.Collections;
using System.Data;


namespace CSIPCardMaintain.EntityLayer
{
    /// <summary>
    /// EntityFORM_COLUMN
    /// </summary>
    [Serializable()]
    [AttributeTable("FORM_COLUMN")]
    public class EntityFORM_COLUMN : Entity
    {
        private string _ELEMENT_CODE;

        /// <summary>
        /// 控制項代碼
        /// </summary>
        public static string M_ELEMENT_CODE = "ELEMENT_CODE";

        private string _ELEMENT_NAME;

        /// <summary>
        /// 控制項欄位名稱
        /// </summary>
        public static string M_ELEMENT_NAME = "ELEMENT_NAME";

        private string _ELEMENT_ID;

        /// <summary>
        /// 控制項ID
        /// </summary>
        public static string M_ELEMENT_ID = "ELEMENT_ID";

        private int _SEQUENCE;

        /// <summary>
        /// 順序
        /// </summary>
        public static string M_SEQUENCE = "SEQUENCE";

        private bool _IS_REQUIRED;

        /// <summary>
        /// 是否必填
        /// </summary>
        public static string M_IS_REQUIRED = "IS_REQUIRED";

        private int _CHECK_FLAG;

        /// <summary>
        /// 檢核型態(0.不檢核 1.數字 2.英數)
        /// </summary>
        public static string M_CHECK_FLAG = "CHECK_FLAG";

        private int _VALUE_LENGTH;

        /// <summary>
        /// 長度
        /// </summary>
        public static string M_VALUE_LENGTH = "VALUE_LENGTH";

        private string _DEFAULT_VALUE;

        /// <summary>
        /// 預設值
        /// </summary>
        public static string M_DEFAULT_VALUE = "DEFAULT_VALUE";

        private string _PROPERTY_CODE;

        /// <summary>
        /// 使用類別編號
        /// </summary>
        public static string M_PROPERTY_CODE = "PROPERTY_CODE";

        private string _REMARK;

        /// <summary>
        /// 備註
        /// </summary>
        public static string M_REMARK = "REMARK";

        private string _USER_ID;

        /// <summary>
        /// 員編
        /// </summary>
        public static string M_USER_ID = "USER_ID";

        private string _MOD_DATE;

        /// <summary>
        /// 更新日期
        /// </summary>
        public static string M_MOD_DATE = "MOD_DATE";

        /// <summary>
        /// 控制項代碼
        /// </summary>
        [AttributeField("ELEMENT_CODE", "System.String", false, true, false, "String")]
        public string ELEMENT_CODE
        {
            get
            {
                return this._ELEMENT_CODE;
            }
            set
            {
                this._ELEMENT_CODE = value;
            }
        }

        /// <summary>
        /// 控制項欄位名稱
        /// </summary>
        [AttributeField("ELEMENT_NAME", "System.String", false, false, false, "String")]
        public string ELEMENT_NAME
        {
            get
            {
                return this._ELEMENT_NAME;
            }
            set
            {
                this._ELEMENT_NAME = value;
            }
        }

        /// <summary>
        /// 控制項ID
        /// </summary>
        [AttributeField("ELEMENT_ID", "System.String", false, false, false, "String")]
        public string ELEMENT_ID
        {
            get
            {
                return this._ELEMENT_ID;
            }
            set
            {
                this._ELEMENT_ID = value;
            }
        }

        /// <summary>
        /// 順序
        /// </summary>
        [AttributeField("SEQUENCE", "System.Int16", false, false, false, "Int16")]
        public int SEQUENCE
        {
            get
            {
                return this._SEQUENCE;
            }
            set
            {
                this._SEQUENCE = value;
            }
        }

        /// <summary>
        /// 是否必填
        /// </summary>
        [AttributeField("IS_REQUIRED", "System.Boolean", false, false, false, "Boolean")]
        public bool IS_REQUIRED
        {
            get
            {
                return this._IS_REQUIRED;
            }
            set
            {
                this._IS_REQUIRED = value;
            }
        }

        /// <summary>
        /// 檢核型態(0.不檢核 1.數字 2.英數)
        /// </summary>
        [AttributeField("CHECK_FLAG", "System.Int16", false, false, false, "Int16")]
        public int CHECK_FLAG
        {
            get
            {
                return this._CHECK_FLAG;
            }
            set
            {
                this._CHECK_FLAG = value;
            }
        }

        /// <summary>
        /// 長度
        /// </summary>
        [AttributeField("VALUE_LENGTH", "System.Int16", false, false, false, "Int16")]
        public int VALUE_LENGTH
        {
            get
            {
                return this._VALUE_LENGTH;
            }
            set
            {
                this._VALUE_LENGTH = value;
            }
        }

        /// <summary>
        /// 預設值
        /// </summary>
        [AttributeField("DEFAULT_VALUE", "System.String", false, false, false, "String")]
        public string DEFAULT_VALUE
        {
            get
            {
                return this._DEFAULT_VALUE;
            }
            set
            {
                this._DEFAULT_VALUE = value;
            }
        }

        /// <summary>
        /// 使用類別編號
        /// </summary>
        [AttributeField("PROPERTY_CODE", "System.String", false, false, false, "String")]
        public string PROPERTY_CODE
        {
            get
            {
                return this._PROPERTY_CODE;
            }
            set
            {
                this._PROPERTY_CODE = value;
            }
        }

        /// <summary>
        /// 備註
        /// </summary>
        [AttributeField("REMARK", "System.String", false, false, false, "String")]
        public string REMARK
        {
            get
            {
                return this._REMARK;
            }
            set
            {
                this._REMARK = value;
            }
        }

        /// <summary>
        /// 員編
        /// </summary>
        [AttributeField("USER_ID", "System.String", false, false, false, "String")]
        public string USER_ID
        {
            get
            {
                return this._USER_ID;
            }
            set
            {
                this._USER_ID = value;
            }
        }

        /// <summary>
        /// 更新日期
        /// </summary>
        [AttributeField("MOD_DATE", "System.String", false, false, false, "String")]
        public string MOD_DATE
        {
            get
            {
                return this._MOD_DATE;
            }
            set
            {
                this._MOD_DATE = value;
            }
        }
    }

    /// <summary>
    /// EntityFORM_COLUMN
    /// </summary>
    [Serializable()]
    public class EntityFORM_COLUMNSet : EntitySet<EntityFORM_COLUMN>
    {
    }
}
