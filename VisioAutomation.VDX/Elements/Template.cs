using VisioAutomation.VDX.Internal.Extensions;

namespace VisioAutomation.VDX
{
    public class Template
    {
        private string xml;

        public Template()
        {
            this.xml = Elements.Drawing.DefaultTemplateXML;
        }

        public Template(string xml)
        {
            this.xml = xml;
        }

        internal System.Xml.Linq.XDocument LoadCleanDOM()
        {
            var dom = System.Xml.Linq.XDocument.Parse(this.xml);
            Template.CleanUpTemplate(dom);
            return dom;                
        }

        public static void CleanUpTemplate(System.Xml.Linq.XDocument vdx_xml_doc)
        {
            var root = vdx_xml_doc.Root;

            string ns_2003 = Internal.Constants.VisioXmlNamespace2003;

            // set document properties
            var docprops = root.ElementVisioSchema2003("DocumentProperties");
            docprops.RemoveElement(ns_2003 + "PreviewPicture");
            docprops.SetElementValue(ns_2003 + "Creator", "");
            docprops.SetElementValue(ns_2003 + "Company", "");

            // remove any pages
            var pages = root.ElementVisioSchema2003("Pages");
            pages.RemoveNodes();

            // Do not remove the FaceNames node - it contains fonts to which the template may be referring
            root.RemoveElement(ns_2003 + "Windows");
            root.RemoveElement(ns_2003 + "DocumentProperties");


            // TODO Add DocumentSettings to VDX
            var docsettings = root.ElementsVisioSchema2003("DocumentSettings");
            if (docsettings != null)
            {
                System.Xml.Linq.Extensions.Remove(docsettings);
            }
        }

    }
}