using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FG.PDMReader
{

    /// <summary>
    /// pdm列信息实体 20160505 225031
    /// </summary>
    public class ColumnInfo
    {
        public ColumnInfo()
        { }

        string columnId;

        public string ColumnId
        {
            get { return columnId; }
            set { columnId = value; }
        }
        string objectID;

        public string ObjectID
        {
            get { return objectID; }
            set { objectID = value; }
        }
        string name;

        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        string code;

        public string Code
        {
            get { return code; }
            set { code = value; }
        }
        int creationDate;

        public int CreationDate
        {
            get { return creationDate; }
            set { creationDate = value; }
        }
        string creator;

        public string Creator
        {
            get { return creator; }
            set { creator = value; }
        }
        int modificationDate;

        public int ModificationDate
        {
            get { return modificationDate; }
            set { modificationDate = value; }
        }
        string modifier;

        public string Modifier
        {
            get { return modifier; }
            set { modifier = value; }
        }
        string comment;

        public string Comment
        {
            get { return comment; }
            set { comment = value; }
        }
        string dataType;

        public string DataType
        {
            get { return dataType; }
            set { dataType = value; }
        }
        string length;

        public string Length
        {
            get { return length; }
            set { length = value; }
        }
        //是否自增量
        bool identity;

        public bool Identity
        {
            get { return identity; }
            set { identity = value; }
        }
        bool mandatory;
        //禁止为空
        public bool Mandatory
        {
            get { return mandatory; }
            set { mandatory = value; }
        }
        string extendedAttributesText;
        //扩展属性
        public string ExtendedAttributesText
        {
            get { return extendedAttributesText; }
            set { extendedAttributesText = value; }
        }
        string physicalOptions;

        public string PhysicalOptions
        {
            get { return physicalOptions; }
            set { physicalOptions = value; }
        }
    }

}
