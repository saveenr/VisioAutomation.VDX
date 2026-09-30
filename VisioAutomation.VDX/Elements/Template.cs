using VisioAutomation.VDX.Internal.Extensions;
using System.Xml.Linq;

namespace VisioAutomation.VDX
{
    public class Template
    {
        private readonly string xml;

        public Template() : this(Elements.Drawing.DefaultTemplateXML)
        {
        }

        public Template(string xml)
        {
            this.xml = xml ?? throw new System.ArgumentNullException(nameof(xml));
        }

        internal XDocument LoadCleanDOM()
        {
            var dom = XDocument.Parse(this.xml);
            CleanUpTemplate(dom);
            return dom;
        }

        public static void CleanUpTemplate(XDocument vdx_xml_doc)
        {
            if (vdx_xml_doc == null)
            {
                throw new System.ArgumentNullException(nameof(vdx_xml_doc));
            }
            string ns = Internal.Constants.VisioXmlNamespace2003;
            var root = vdx_xml_doc.Root;
            if (root == null || root.Name != ns + "VisioDocument")
            {
                throw new System.ArgumentException("Template must contain a Visio 2003 VisioDocument root", nameof(vdx_xml_doc));
            }

            root.Elements(ns + "DocumentProperties").Remove();
            root.Elements(ns + "Windows").Remove();
            root.Elements(ns + "DocumentSettings").Remove();

            // Preserve resources referenced by masters, but replace template pages.
            foreach (string name in new[] { "FaceNames", "Colors", "Masters", "Pages" })
            {
                if (root.Element(ns + name) == null)
                {
                    root.Add(new XElement(ns + name));
                }
            }
            root.Element(ns + "Pages").RemoveNodes();
        }
    }
}
