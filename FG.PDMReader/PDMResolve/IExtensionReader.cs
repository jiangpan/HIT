using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace FG.PDMReader.PDMResolve
{
    public interface IExtensionReader
    {
        void Read(XmlDocument doc, XmlNamespaceManager ns, Entity model);
    }
}
