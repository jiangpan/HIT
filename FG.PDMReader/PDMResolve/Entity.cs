using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace FG.PDMReader.PDMResolve
{
    public class Entity
    {
        private Dictionary<string, object> properties = new Dictionary<string, object>();
        public Dictionary<string, object> Properties
        {
            get { return properties; }
        }



        public string ObjectType
        {
            get { return GetString("ObjectType"); }
            set { Properties["ObjectType"] = value; }
        }
        public string Name
        {
            get { return GetString("Name"); }
            set { Properties["Name"] = value; }
        }
        [Description("主键")] 
        public string Code
        {
            get { return GetString("Code"); }
            set { Properties["Code"] = value; }
        }
        //[Browsable(false)]
        //public string Id
        //{
        //    get { return GetString("Id"); }
        //    set { Properties["Id"] = value; }
        //}
        //[Browsable(false)]
        //public string ObjectID
        //{
        //    get { return GetString("ObjectID"); }
        //    set { Properties["ObjectID"] = value; }
        //}

        #region helper

        public bool ContainsProperty(string propertyName)
        {
            return properties.ContainsKey(propertyName);
        }
        public object this[string propertyName]
        {
            get
            {
                return properties.ContainsKey(propertyName) ? properties[propertyName] : null;
            }
            set
            {
                if (properties.ContainsKey(propertyName))
                {
                    properties[propertyName] = value;
                }
                else
                {
                    properties.Add(propertyName, value);
                }
            }
        }

        public string GetString(string propertyName)
        {
            return properties.ContainsKey(propertyName) ? (string)properties[propertyName] : null;
        }
        public bool GetBoolean(string propertyName)
        {
            if (properties.ContainsKey(propertyName))
            {
                string s = (string)properties[propertyName];
                if (string.IsNullOrWhiteSpace(s)) return false;
                s = s.ToLower();
                if (s == "1" || s == "true")
                {
                    return true;
                }
                else
                {
                    return Convert.ToBoolean(s);
                }
            }
            else
            {
                return false;
            }
        }
        public int GetInt32(string propertyName)
        {
            return properties.ContainsKey(propertyName) ? Convert.ToInt32(properties[propertyName]) : 0;
        }
        public long GetInt64(string propertyName)
        {
            return properties.ContainsKey(propertyName) ? Convert.ToInt64(properties[propertyName]) : 0;
        }
        public DateTime GetDateTime(string propertyName)
        {
            return (DateTime)this[propertyName];
        }


        public override string ToString()
        {
            if (!string.IsNullOrEmpty(Code))
            {
                return Code;
            }
            else if (!string.IsNullOrEmpty(ObjectType))
            {
                return ObjectType;
            }
            else
            {
                return base.ToString();
            }
        }
        #endregion
    }
}
