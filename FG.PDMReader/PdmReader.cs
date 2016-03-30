using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace FG.PDMReader
{
    public class PdmReader
    {
        public const string a = "attribute", c = "collection", o = "object";

        public const string cClasses = "c:Classes";
        public const string oClass = "o:Class";

        public const string cAttributes = "c:Attributes";
        public const string oAttribute = "o:Attribute";

        public const string cTables = "c:Tables";
        public const string oTable = "o:Table";

        public const string cColumns = "c:Columns";
        public const string oColumn = "o:Column";


        XmlDocument xmlDoc;
        XmlNamespaceManager xmlnsManager;
        /// <summary>构造函数 </summary>
        public PdmReader()
        {
            // TODO: 在此处添加构造函数逻辑
            xmlDoc = new XmlDocument();
        }
        /// <summary>构造函数 </summary>
        public PdmReader(string pdm_file)
        {
            PdmFile = pdm_file;
        }

        string pdmFile;

        public string PdmFile
        {
            get { return pdmFile; }
            set
            {
                pdmFile = value;
                if (xmlDoc == null)
                {
                    xmlDoc = new XmlDocument();
                    xmlDoc.Load(pdmFile);
                    xmlnsManager = new XmlNamespaceManager(xmlDoc.NameTable);
                    xmlnsManager.AddNamespace("a", "attribute");
                    xmlnsManager.AddNamespace("c", "collection");
                    xmlnsManager.AddNamespace("o", "object");
                }
            }
        }

        IList<TableInfo> tables;

        public IList<TableInfo> Tables
        {
            get { return tables; }
            set { tables = value; }
        }

        public void InitData()
        {
            if (Tables == null)
                Tables = new List<TableInfo>();
            XmlNode xnTables = xmlDoc.SelectSingleNode("//" + cTables, xmlnsManager);
            foreach (XmlNode xnTable in xnTables.ChildNodes)
            {
                Tables.Add(GetTable(xnTable));
            }
        }

        //初始化"o:Table"的节点
        private TableInfo GetTable(XmlNode xnTable)
        {
            TableInfo mTable = new TableInfo();
            XmlElement xe = (XmlElement)xnTable;
            mTable.TableId = xe.GetAttribute("Id");
            XmlNodeList xnTProperty = xe.ChildNodes;
            foreach (XmlNode xnP in xnTProperty)
            {
                switch (xnP.Name)
                {
                    case "a:ObjectID": mTable.ObjectID = xnP.InnerText;
                        break;
                    case "a:Name": mTable.Name = xnP.InnerText;
                        break;
                    case "a:Code": mTable.Code = xnP.InnerText;
                        break;
                    case "a:CreationDate": mTable.CreationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Creator": mTable.Creator = xnP.InnerText;
                        break;
                    case "a:ModificationDate": mTable.ModificationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Modifier": mTable.Modifier = xnP.InnerText;
                        break;
                    case "a:Comment": mTable.Comment = xnP.InnerText;
                        break;
                    case "a:PhysicalOptions": mTable.PhysicalOptions = xnP.InnerText;
                        break;
                    case "c:Columns": InitColumns(xnP, mTable);
                        break;
                    case "c:Keys": InitKeys(xnP, mTable);
                        break;
                    case "c:PrimaryKey": InitPrimary(xnP, mTable);
                        break;
                }
            }
            return mTable;
        }
        //初始化"c:Columns"的节点
        private void InitColumns(XmlNode xnColumns, TableInfo pTable)
        {
            foreach (XmlNode xnColumn in xnColumns)
            {
                pTable.AddColumn(GetColumn(xnColumn));
            }
        }

        //初始化c:Keys"的节点
        private void InitKeys(XmlNode xnKeys, TableInfo pTable)
        {
            foreach (XmlNode xnKey in xnKeys)
            {
                pTable.AddKey(GetKey(xnKey));
            }
        }

        //初始化c:PrimaryKey的节点
        private void InitPrimary(XmlNode xnKeys, TableInfo pTable) 
        {
            foreach (XmlNode xnKey in xnKeys)
            {
                PdmKey key = GetPrimary(xnKey);
                foreach (PdmKey pk in pTable.Keys) 
                {
                    if (pk.KeyId.Equals(key.KeyId)) 
                    {
                        foreach (ColumnInfo ci in pk.Columns) 
                        {
                            pTable.AddPrimary(ci.ColumnId);
                        }
                    }
                }
            }
        }

        private ColumnInfo GetColumn(XmlNode xnColumn)
        {
            ColumnInfo mColumn = new ColumnInfo();
            XmlElement xe = (XmlElement)xnColumn;
            mColumn.ColumnId = xe.GetAttribute("Id");
            XmlNodeList xnCProperty = xe.ChildNodes;
            foreach (XmlNode xnP in xnCProperty)
            {
                switch (xnP.Name)
                {
                    case "a:ObjectID": mColumn.ObjectID = xnP.InnerText;
                        break;
                    case "a:Name": mColumn.Name = xnP.InnerText;
                        break;
                    case "a:Code": mColumn.Code = xnP.InnerText;
                        break;
                    case "a:CreationDate": mColumn.CreationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Creator": mColumn.Creator = xnP.InnerText;
                        break;
                    case "a:ModificationDate": mColumn.ModificationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Modifier": mColumn.Modifier = xnP.InnerText;
                        break;
                    case "a:Comment": mColumn.Comment = xnP.InnerText;
                        break;
                    case "a:DataType": mColumn.DataType = xnP.InnerText;
                        break;
                    case "a:Length": mColumn.Length = xnP.InnerText;
                        break;
                    case "a:Identity": mColumn.Identity = xnP.InnerText.Equals("Yes") ? true : false;
                        break;
                    case "a:Mandatory": mColumn.Mandatory = xnP.InnerText.Equals("1") ? true : false;
                        break;
                    case "a:PhysicalOptions": mColumn.PhysicalOptions = xnP.InnerText;
                        break;
                    case "a:ExtendedAttributesText": mColumn.ExtendedAttributesText = xnP.InnerText;
                        break;
                }
            }
            return mColumn;
        }

        private PdmKey GetKey(XmlNode xnKey)
        {
            PdmKey mKey = new PdmKey();
            XmlElement xe = (XmlElement)xnKey;
            mKey.KeyId = xe.GetAttribute("Id");
            XmlNodeList xnKProperty = xe.ChildNodes;
            foreach (XmlNode xnP in xnKProperty)
            {
                switch (xnP.Name)
                {
                    case "a:ObjectID": mKey.ObjectID = xnP.InnerText;
                        break;
                    case "a:Name": mKey.Name = xnP.InnerText;
                        break;
                    case "a:Code": mKey.Code = xnP.InnerText;
                        break;
                    case "a:CreationDate": mKey.CreationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Creator": mKey.Creator = xnP.InnerText;
                        break;
                    case "a:ModificationDate": mKey.ModificationDate = Convert.ToInt32(xnP.InnerText);
                        break;
                    case "a:Modifier": mKey.Modifier = xnP.InnerText;
                        break;
                    //还差 <c:Key.Columns>
                    case "c:Key.Columns": GetKeyColumn(xnP, mKey);
                        break;
                }
            }
            return mKey;
        }

        public void GetKeyColumn(XmlNode xnP ,PdmKey mKey) 
        {
            XmlElement xe = (XmlElement)xnP;
            XmlNodeList nodeList = xe.ChildNodes;
            foreach (XmlNode node in nodeList) 
            {
                ColumnInfo ci = new ColumnInfo();
                ci.ColumnId = ((XmlElement)node).GetAttribute("Ref");
                mKey.AddColumn(ci);
            }
        }

        private PdmKey GetPrimary(XmlNode xnKey) 
        {
            PdmKey mKey = new PdmKey();
            XmlElement xe = (XmlElement)xnKey;
            mKey.KeyId = xe.GetAttribute("Ref");
            return mKey;
        }
    }

}