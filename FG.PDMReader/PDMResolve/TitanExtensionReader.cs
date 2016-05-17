using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace FG.PDMReader.PDMResolve
{
    public class TitanExtensionReader:IExtensionReader
    {
        public void Read(XmlDocument doc, XmlNamespaceManager ns, Entity model)
        {
            Dictionary<string, Entity> ObjectsById = new Dictionary<string, Entity>();
            Dictionary<string, Entity> ObjectsByObjectID = new Dictionary<string, Entity>();
            #region enums
            Dictionary<string, Entity> Enums = new Dictionary<string, Entity>();
            //if (!model.Stereotypes.ContainsKey("Enums"))
            //{
            //    model.Stereotypes.Add("Enums", new Dictionary<string, Entity>());
            //}
            //Dictionary<string, Entity> enums=model.Stereotypes["Enums"];

            XmlNodeList extendedObjectElements = doc.SelectNodes("//c:ExtendedObjects/o:ExtendedObject", ns);
            foreach (XmlNode extendedObjectElement in extendedObjectElements)
            {
                XmlNode stereotypeElement = extendedObjectElement.SelectSingleNode("a:Stereotype", ns);
                if (stereotypeElement.InnerText == "Enum")
                {
                    Entity entity = new Entity();

                    //column自身的属性
                    entity.ObjectType = "Enum";
                    entity["Id"] = extendedObjectElement.Attributes["Id"].Value;
                    XmlNodeList entityAttributeElements = extendedObjectElement.SelectNodes("*[namespace-uri()='attribute']", ns);
                    foreach (XmlNode entityAttributeElement in entityAttributeElements)
                    {
                        entity[entityAttributeElement.LocalName] = entityAttributeElement.InnerText;
                    }

                    //ExtendedAttributesText
                    PdmReader.ParseExtendedAttributesText(entity.GetString("ExtendedAttributesText"), "TitanExtension", entity);


                    Enums.Add(entity.Code, entity);
                    ObjectsById.Add((string)entity["Id"], entity);
                    ObjectsByObjectID.Add((string)entity["ObjectID"], entity); 
                }
            }
            model["Enums"] = Enums;
            #endregion

            #region column.Enums
            //columns
            XmlNodeList columnElements = doc.SelectNodes("//c:Columns/o:Column", ns);
            foreach (XmlNode columnElement in columnElements)
            {

                XmlNode extendedCollectionElement = columnElement.SelectSingleNode("c:ExtendedCollections/o:ExtendedCollection", ns);
                if (extendedCollectionElement != null)
                {
                    XmlNode nameElement = extendedCollectionElement.SelectSingleNode("a:Name", ns);
                    XmlNode refElement = extendedCollectionElement.SelectSingleNode("c:Content/o:ExtendedObject", ns);
                    if (nameElement != null && refElement != null)
                    {
                        string name = nameElement.InnerText;
                        if (name == "Enum")
                        {
                            string columnId = columnElement.Attributes["Id"].Value;
                            string refId = refElement.Attributes["Ref"].Value;
                            Entity column = FindColumn(model, columnId);
                            column["Enum"] = ObjectsById[refId];
                        }
                    }
                }
            }
            #endregion


            #region column.ExtendedAttributesText
            Dictionary<string, Entity> Tables = (Dictionary<string, Entity>)model["Tables"];
            foreach (Entity table in Tables.Values)
            {
                Dictionary<string, Entity> Columns = (Dictionary<string, Entity>)table["Columns"];
                foreach (Entity column in Columns.Values)
                {
                    if (column.ContainsProperty("ExtendedAttributesText"))
                    {
                        PdmReader.ParseExtendedAttributesText(column.GetString("ExtendedAttributesText"), "TitanExtension", column);
                    } 
                }
            }
            Dictionary<string, Entity> Views = (Dictionary<string, Entity>)model["Views"];
            foreach (Entity view in Views.Values)
            {
                Dictionary<string, Entity> Columns = (Dictionary<string, Entity>)view["Columns"];
                foreach (Entity column in Columns.Values)
                {
                    if (column.ContainsProperty("ExtendedAttributesText"))
                    {
                        PdmReader.ParseExtendedAttributesText(column.GetString("ExtendedAttributesText"), "TitanExtension", column);
                    }
                }
            }
            #endregion
        }

        private Entity FindColumn(Entity model, string columnId)
        {
            Dictionary<string, Entity> Tables = (Dictionary<string, Entity>)model["Tables"];
            foreach (Entity table in Tables.Values)
            {
                Dictionary<string, Entity> Columns = (Dictionary<string, Entity>)table["Columns"];
                foreach (Entity column in Columns.Values)
                {
                    if ((string)column["Id"] == columnId) return column;
                }
            }
            Dictionary<string, Entity> Views = (Dictionary<string, Entity>)model["Views"];
            foreach (Entity view in Views.Values)
            {
                Dictionary<string, Entity> Columns = (Dictionary<string, Entity>)view["Columns"];
                foreach (Entity column in Columns.Values)
                {
                    if ((string)column["Id"] == columnId) return column;
                }
            }
            throw new InvalidOperationException(string.Format("没有找到id={0}的Column",columnId));
        }
    }
}
