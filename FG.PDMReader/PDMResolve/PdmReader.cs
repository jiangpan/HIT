using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace FG.PDMReader.PDMResolve
{
    /// <summary>
    /// 新增的pdm读取
    /// 
    /// </summary>
    public class PdmReader 
    {
        private static List<IExtensionReader> extensionReaders = new List<IExtensionReader>() { new TitanExtensionReader() };

         
        public Entity Read(string fileName )
        { 
            


            Entity model = new Entity();
            model.ObjectType = "Model";
            Dictionary<string, Entity> ObjectsById = new Dictionary<string, Entity>();
            Dictionary<string, Entity> ObjectsByObjectID = new Dictionary<string, Entity>();

            XmlDocument doc = new XmlDocument();

            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            doc.Load(fs);
            XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("o", "object");
            ns.AddNamespace("c", "collection");
            ns.AddNamespace("a", "attribute");

            #region sequences
            Dictionary<string, Entity> Sequences = new Dictionary<string, Entity>();
            //XmlNodeList sequenceElements = doc.SelectNodes("/Model/o:RootObject/c:Children/o:Model/c:Sequences/o:Sequence", ns);
            XmlNodeList sequenceElements = doc.SelectNodes("//c:Sequences/o:Sequence", ns);
            foreach (XmlNode sequenceElement in sequenceElements)
            {
                Entity sequence = new Entity(); 
                //sequence自身的属性
                sequence.ObjectType = "Sequence";
                sequence["Id"] = sequenceElement.Attributes["Id"].Value;
                XmlNodeList sequenceAttributeElements = sequenceElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                foreach (XmlNode sequenceAttributeElement in sequenceAttributeElements)
                {
                    sequence[sequenceAttributeElement.LocalName] = sequenceAttributeElement.InnerText;
                }
                Sequences.Add(sequence.Code, sequence);
                ObjectsById.Add((string)sequence["Id"], sequence);
                ObjectsByObjectID.Add((string)sequence["ObjectID"], sequence);
            }
            model.Properties.Add("Sequences", Sequences);
            #endregion

            #region tables
            Dictionary<string, Entity> Tables = new Dictionary<string, Entity>();
            //XmlNodeList tableElements = doc.SelectNodes("/Model/o:RootObject/c:Children/o:Model/c:Tables/o:Table", ns);
            XmlNodeList tableElements = doc.SelectNodes("//c:Tables/o:Table", ns);
            foreach (XmlNode tableElement in tableElements)
            {
                Entity table = new Entity();

                //table自身的属性
                table.ObjectType = "Table";
                table["Id"] = tableElement.Attributes["Id"].Value;
                XmlNodeList tableAttributeElements = tableElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                foreach (XmlNode tableAttributeElement in tableAttributeElements)
                {
                    table[tableAttributeElement.LocalName] = tableAttributeElement.InnerText;
                }

                //columns
                Dictionary<string, Entity> Columns = new Dictionary<string, Entity>();
                XmlNodeList columnElements = tableElement.SelectNodes("c:Columns/o:Column", ns);
                foreach (XmlNode columnElement in columnElements)
                {
                    Entity column = new Entity();

                    //column自身的属性
                    column.ObjectType = "Column";
                    column["Id"] = columnElement.Attributes["Id"].Value;
                    XmlNodeList columnAttributeElements = columnElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                    foreach (XmlNode columnAttributeElement in columnAttributeElements)
                    {
                        column[columnAttributeElement.LocalName] = columnAttributeElement.InnerText;
                    }
                    Columns.Add(column.Code, column);
                    ObjectsById.Add((string)column["Id"], column);
                    ObjectsByObjectID.Add((string)column["ObjectID"], column);

                    //oracle特有的Sequence
                    XmlNode sequenceElement = columnElement.SelectSingleNode("c:Sequence/o:Sequence", ns);
                    if (sequenceElement != null)
                    {
                        string sequenceId = sequenceElement.Attributes["Ref"].Value;
                        //XmlNode rootSequenceElement = doc.SelectSingleNode("//c:Sequences/o:Sequence[@Id='" + sequenceId + "']", ns);
                        column["Sequence"] = ObjectsById[sequenceId];
                    }
                }

                //keys
                XmlNode primaryKeyElement = tableElement.SelectSingleNode("c:PrimaryKey/o:Key", ns);
                if (primaryKeyElement != null)
                {
                    string keyId = primaryKeyElement.Attributes["Ref"].Value;
                    XmlNode keyElement = tableElement.SelectSingleNode("c:Keys/o:Key[@Id='" + keyId + "']", ns);
                    if (keyElement != null)
                    {

                        XmlNodeList keyColumnElements = keyElement.SelectNodes("c:Key.Columns/o:Column", ns);
                        foreach (XmlNode keyColumnElement in keyColumnElements)
                        {
                            Entity column = ObjectsById[keyColumnElement.Attributes["Ref"].Value];
                            column["Primary"] = "1";
                            //table.PrimaryKeys.Add(column.Code, column);
                        }
                    }
                }

                table["Columns"]=Columns;
                Tables.Add(table.Code, table);
                ObjectsById.Add((string)table["Id"], table);
                ObjectsByObjectID.Add((string)table["ObjectID"], table);
            }
            model["Tables"] = Tables;
            #endregion

            #region views
            Dictionary<string, Entity> Views = new Dictionary<string, Entity>();
            //XmlNodeList tableElements = doc.SelectNodes("/Model/o:RootObject/c:Children/o:Model/c:Tables/o:Table", ns);
            tableElements = doc.SelectNodes("//c:Views/o:View", ns);
            foreach (XmlNode tableElement in tableElements)
            {
                Entity view = new Entity();

                //table自身的属性
                view.ObjectType = "View";
                view["Id"] = tableElement.Attributes["Id"].Value;
                XmlNodeList tableAttributeElements = tableElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                foreach (XmlNode tableAttributeElement in tableAttributeElements)
                {
                    view[tableAttributeElement.LocalName] = tableAttributeElement.InnerText;
                }

                //columns
                Dictionary<string, Entity> Columns = new Dictionary<string, Entity>();
                XmlNodeList columnElements = tableElement.SelectNodes("c:Columns/o:ViewColumn", ns);
                foreach (XmlNode columnElement in columnElements)
                {
                    Entity column = new Entity();

                    //column自身的属性
                    column.ObjectType = "Column";
                    column["Id"] = columnElement.Attributes["Id"].Value;
                    XmlNodeList columnAttributeElements = columnElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                    foreach (XmlNode columnAttributeElement in columnAttributeElements)
                    {
                        column[columnAttributeElement.LocalName] = columnAttributeElement.InnerText;
                    }
                    Columns.Add(column.Code, column);
                    ObjectsById.Add((string)column["Id"], column);
                    ObjectsByObjectID.Add((string)column["ObjectID"], column);
                }
                view["Columns"] = Columns;
                Views.Add(view.Code, view);
                ObjectsById.Add((string)view["Id"], view);
                ObjectsByObjectID.Add((string)view["ObjectID"], view);
            }
            model["Views"] = Views;
            #endregion

            #region references
            Dictionary<string, Entity> References = new Dictionary<string, Entity>();
            //XmlNodeList tableElements = doc.SelectNodes("/Model/o:RootObject/c:Children/o:Model/c:Tables/o:Table", ns);
            XmlNodeList referenceElements = doc.SelectNodes("//c:References/o:Reference", ns);
            foreach (XmlNode referenceElement in referenceElements)
            {
                Entity reference = new Entity();

                //reference自身的属性
                reference.ObjectType = "Reference";
                reference["Id"] = referenceElement.Attributes["Id"].Value;
                XmlNodeList referenceAttributeElements = referenceElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                foreach (XmlNode referenceAttributeElement in referenceAttributeElements)
                {
                    reference[referenceAttributeElement.LocalName] = referenceAttributeElement.InnerText;
                }

                //parentTable
                XmlNode parentTableElement = referenceElement.SelectSingleNode("c:ParentTable/o:Table", ns);
                string parentTableId = parentTableElement.Attributes["Ref"].Value;
                reference["ParentTable"] = ObjectsById[parentTableId];
                //childTable
                XmlNode childTableElement = referenceElement.SelectSingleNode("c:ChildTable/o:Table", ns);
                string childTableId = childTableElement.Attributes["Ref"].Value;
                Entity childTable = ObjectsById[childTableId];
                reference["ChildTable"] = childTable;

                //joins
                List<Entity> ReferenceColumns = new List<Entity>();
                XmlNodeList joinElements = referenceElement.SelectNodes("c:Joins/o:ReferenceJoin", ns);
                foreach (XmlNode joinElement in joinElements)
                {
                    Entity referenceColumn = new Entity();
                    referenceColumn.ObjectType = "ReferenceColumn";

                    XmlNode parentColumnElement = joinElement.SelectSingleNode("c:Object1/o:Column", ns);
                    string parentColumnId = parentColumnElement.Attributes["Ref"].Value;
                    referenceColumn["ParentColumn"] = ObjectsById[parentColumnId];
                    XmlNode childColumnElement = joinElement.SelectSingleNode("c:Object2/o:Column", ns);
                    if (childColumnElement == null)
                    {
                        throw new Exception(reference.Code+".ChildColumn is null");
                    }
                    string childColumnId = childColumnElement.Attributes["Ref"].Value;
                    referenceColumn["ChildColumn"] = ObjectsById[childColumnId];

                    ReferenceColumns.Add(referenceColumn);
                }

                reference["ReferenceColumns"] = ReferenceColumns;
                if (childTable["References"] == null)
                {
                    childTable["References"] = new Dictionary<string, Entity>();
                }
                ((Dictionary<string, Entity>)childTable["References"]).Add(reference.Code, reference);
                References.Add(reference.Code, reference);
                ObjectsById.Add((string)reference["Id"], reference);
                ObjectsByObjectID.Add((string)reference["ObjectID"], reference);
            }
            //model["References"] = References;//在下一节赋值
            #endregion

            #region view references
            //XmlNodeList tableElements = doc.SelectNodes("/Model/o:RootObject/c:Children/o:Model/c:Tables/o:Table", ns);
            referenceElements = doc.SelectNodes("//c:ViewReferences/o:ViewReference", ns);
            foreach (XmlNode referenceElement in referenceElements)
            {
                Entity reference = new Entity();

                //reference自身的属性
                reference.ObjectType = "Reference";
                reference["Id"] = referenceElement.Attributes["Id"].Value;
                XmlNodeList referenceAttributeElements = referenceElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                foreach (XmlNode referenceAttributeElement in referenceAttributeElements)
                {
                    reference[referenceAttributeElement.LocalName] = referenceAttributeElement.InnerText;
                }

                //parentTable
                XmlNode parentTableElement = referenceElement.SelectSingleNode("c:TableView1", ns).ChildNodes[0];
                string parentTableId = parentTableElement.Attributes["Ref"].Value;
                reference["ParentTable"] = ObjectsById[parentTableId];
                //childTable
                XmlNode childTableElement = referenceElement.SelectSingleNode("c:TableView2", ns).ChildNodes[0];
                string childTableId = childTableElement.Attributes["Ref"].Value;
                Entity childTable = ObjectsById[childTableId];
                reference["ChildTable"] = childTable;

                //joins
                List<Entity> ReferenceColumns = new List<Entity>();
                XmlNodeList joinElements = referenceElement.SelectNodes("c:ViewReference.Joins/o:ViewReferenceJoin", ns);
                foreach (XmlNode joinElement in joinElements)
                {
                    Entity referenceColumn = new Entity();
                    referenceColumn.ObjectType = "ReferenceColumn";

                    XmlNode parentColumnElement = joinElement.SelectSingleNode("c:Column1", ns).ChildNodes[0];
                    string parentColumnId = parentColumnElement.Attributes["Ref"].Value;
                    referenceColumn["ParentColumn"] = ObjectsById[parentColumnId];
                    XmlNode childColumnElement = joinElement.SelectSingleNode("c:Column2", ns).ChildNodes[0];
                    string childColumnId = childColumnElement.Attributes["Ref"].Value;
                    referenceColumn["ChildColumn"] = ObjectsById[childColumnId];

                    ReferenceColumns.Add(referenceColumn);
                }



                reference["ReferenceColumns"] = ReferenceColumns;
                if (childTable["References"] == null)
                {
                    childTable["References"] = new Dictionary<string, Entity>();
                }
                ((Dictionary<string, Entity>)childTable["References"]).Add(reference.Code, reference);
                References.Add(reference.Code, reference);
                ObjectsById.Add((string)reference["Id"], reference);
                ObjectsByObjectID.Add((string)reference["ObjectID"], reference);
            }
            model["References"] = References;
            #endregion


            #region exrensions
            foreach (IExtensionReader extensionReader in extensionReaders)
            {
                extensionReader.Read(doc, ns, model);
            }
            #endregion
            fs.Close();

            return model;
        }
        internal static void ParseExtendedAttributesText(string text, string startName, Entity entity)
        {
            //text = "{F6DB6067-29D9-4AA0-A339-66CB46DE7EA0},TitanExtension,92={136CCDC6-116C-4BD2-B8E6-D880AC245FF5},EnumMember,37=1,Province,省级 2,City,市级 3,County,县级 ";
            string pattern = @"\{[0123456789ABCDEF\-]{36}\}\," + startName + @"\,([0-9]*?)=";
            MatchCollection ms = Regex.Matches(text, pattern);
            if (ms.Count <= 0) return;
            GroupCollection gs = ms[0].Groups;
            Group g = gs[1];

            string text2 = text.Substring(g.Index + g.Length + 1, Convert.ToInt32(g.Value));
            while (text2 != "")
            {
                string pattern2 = @"\{[0123456789ABCDEF\-]{36}\}\,([^,]*?)\,([0-9]*?)=";
                MatchCollection ms2 = Regex.Matches(text2, pattern2);
                GroupCollection gs2 = ms2[0].Groups;
                Group attributeNameGroup = gs2[1];
                string attributeName = attributeNameGroup.Value;
                Group attributeValueGroup = gs2[2];
                string attributeValue = text2.Substring(attributeValueGroup.Index + attributeValueGroup.Length + 1, Convert.ToInt32(attributeValueGroup.Value));
                entity[attributeName] = attributeValue;
                text2 = text2.Substring(attributeValueGroup.Index + attributeValueGroup.Length + 1 + attributeValue.Length + 2, text2.Length - (attributeValueGroup.Index + attributeValueGroup.Length + 1 + attributeValue.Length + 2));
            }
        }
    }
}
