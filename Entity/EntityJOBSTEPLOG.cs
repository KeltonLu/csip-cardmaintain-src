//------------------------------------------------------------------------------
// <auto-generated>
//     這段程式碼是由工具產生的。
//     執行階段版本:2.0.50727.42
//
//     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
//     變更將會遺失。
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Text;
using Framework.Data.OM.OMAttribute;
using Framework.Data.OM;
using Framework.Data.OM.Collections;


namespace CSIPCardMaintain.EntityLayer
{
    
    
    /// <summary>
    /// JOBSTEPLOG
    /// </summary>
    [Serializable()]
    [AttributeTable("JOBSTEPLOG")]
    public class EntityJOBSTEPLOG : Entity
    {
        
        private int _SNO;
        
        /// <summary>
        /// SNO
        /// </summary>
        public static string M_SNO = "SNO";
        
        private string _JOBNAME;
        
        /// <summary>
        /// JOBNAME
        /// </summary>
        public static string M_JOBNAME = "JOBNAME";
        
        private object _DT;
        
        /// <summary>
        /// DT
        /// </summary>
        public static string M_DT = "DT";
        
        private string _STEP;
        
        /// <summary>
        /// STEP
        /// </summary>
        public static string M_STEP = "STEP";
        
        private string _CMD;
        
        /// <summary>
        /// CMD
        /// </summary>
        public static string M_CMD = "CMD";
        
        /// <summary>
        /// SNO
        /// </summary>
        [AttributeField("SNO", "System.Int32", false, true, true, "Int32")]
        public int SNO
        {
            get
            {
                return this._SNO;
            }
            set
            {
                this._SNO = value;
            }
        }
        
        /// <summary>
        /// JOBNAME
        /// </summary>
        [AttributeField("JOBNAME", "System.String", false, false, false, "String")]
        public string JOBNAME
        {
            get
            {
                return this._JOBNAME;
            }
            set
            {
                this._JOBNAME = value;
            }
        }
        
        /// <summary>
        /// DT
        /// </summary>
        [AttributeField("DT", "System.DateTime", false, false, false, "DateTime")]
        public object DT
        {
            get
            {
                return this._DT;
            }
            set
            {
                this._DT = value;
            }
        }
        
        /// <summary>
        /// STEP
        /// </summary>
        [AttributeField("STEP", "System.String", false, false, false, "String")]
        public string STEP
        {
            get
            {
                return this._STEP;
            }
            set
            {
                this._STEP = value;
            }
        }
        
        /// <summary>
        /// CMD
        /// </summary>
        [AttributeField("CMD", "System.String", false, false, false, "String")]
        public string CMD
        {
            get
            {
                return this._CMD;
            }
            set
            {
                this._CMD = value;
            }
        }
    }
    
    /// <summary>
    /// JOBSTEPLOG
    /// </summary>
    [Serializable()]
    public class EntityJOBSTEPLOGSet : EntitySet<EntityJOBSTEPLOG>
    {
    }
}
