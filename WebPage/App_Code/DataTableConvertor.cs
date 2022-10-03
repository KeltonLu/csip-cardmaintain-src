//******************************************************************
//*  作    者：林家賜
//*  功能說明：提供公用方法將DATATABLE轉換為指定型別物件集合，並將值對應給屬性
//*  創建日期：2019/01/24
//*  修改記錄：
//*<author>            <time>            <TaskID>                <desc>
//*******************************************************************
using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Reflection;

/// <summary>
/// DataTableConvertr 的摘要描述
/// </summary>
public class DataTableConvertor
{
    public DataTableConvertor()
    {

    }
    /// <summary>
    /// 將TABLE轉成物件
    /// </summary>
    /// <param name="dt"></param>
    /// <returns></returns>
    public static List<T> ConvertCollToObj<T>(DataTable dt)
    {
        if (dt == null)
        {
            return new List<T>();
        }
        List<T> rtn = new List<T>();
        for (int i = 0; i < dt.Rows.Count; i++)
        {
            T Nobj = Activator.CreateInstance<T>();
            convSingRow<T>(ref Nobj, dt.Rows[i]);
            rtn.Add(Nobj);
        }
        return rtn;

    }


    /// <summary>
    /// 以映射方法轉換DataRow 2 Obj
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="obj"></param>
    public static void convSingRow<T>(ref T obj, DataRow dataRow)
    {
        Type v = obj.GetType();  //取的型別實體
        PropertyInfo[] props = v.GetProperties(); //取出所有公開屬性(可以被外部存取得 
        foreach (PropertyInfo prop in props)
        {
            string propName = prop.Name; //屬性名稱，用來對應
            if (dataRow.Table.Columns.Contains(propName))
            {
                switch (prop.PropertyType.Name)
                {
                    case "Int32":
                        prop.SetValue(obj, int.Parse(dataRow[propName].ToString()), null);
                        break;

                    case "String":
                        prop.SetValue(obj, dataRow[propName].ToString(), null);
                        break;
                }

            }
        }
    }
    public static bool IsDate(string anyString)
    {
        if (anyString == null)
        {
            anyString = "";
        }
        if (anyString.Length > 0)
        {
            DateTime dummyDate;
            try
            {
                dummyDate = DateTime.Parse(anyString);
            }
            catch
            {
                return false;
            }
            return true;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// 以映射方式實作值的複製，同型別但附加屬性不同
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <typeparam name="T2"></typeparam>
    /// <param name="inModel"></param>
    /// <param name="OutModel"></param>
    public static void Clone2Other<T, T2>(T inModel,ref T2 OutModel){

        Type inType = inModel.GetType();  //取型別實體
        Type outType = OutModel.GetType();  //取型別實體

        PropertyInfo[] outprops = outType.GetProperties(); //取出所有公開屬性(可以被外部存取得 
        PropertyInfo[] inprops = inType.GetProperties();
        foreach (PropertyInfo prop in inprops)
        { 
                string propName = prop.Name;
            string Ptype = prop.PropertyType.Name;
            //類別為物件，則不取值也不轉
            if (Ptype == "Object") { continue; }
            string exVal = "";
            if (Ptype == "Int32")
            {
                exVal = prop.GetValue(inModel, null).ToString();
            }
            else
            {
                exVal = prop.GetValue(inModel, null) as string;
            }
         
                foreach (PropertyInfo oUprop in outprops)
                {
                    if (propName == oUprop.Name)
                    {
                    // prop.SetValue(OutModel, exVal, null);
                    switch (Ptype)
                    {
                        case "Int32":
                            oUprop.SetValue(OutModel, int.Parse(exVal), null);
                            break;

                        case "String":
                            oUprop.SetValue(OutModel, exVal, null);
                            break;
                    }
                }
                } 
        }

    }
}