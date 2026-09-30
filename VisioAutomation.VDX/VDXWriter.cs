using SXL = System.Xml.Linq;
using VisioAutomation.VDX.Internal.Extensions;

namespace VisioAutomation.VDX;

internal static class VDXWriter
{
    public static SXL.XDocument CreateVDX(Elements.Drawing drawing, SXL.XDocument template)
    {
        foreach (var window in drawing.Windows)
        {
            if (window is Elements.DocumentWindow document_window)
            {
                document_window.ValidatePage(drawing);
            }
        }

        // Each serialization owns its DOM; neither writes nor callers can alter the template.
        var dom = new SXL.XDocument(template);
        var root = dom.Root;
        root.AddFirst(drawing.DocumentProperties.ToXml());

        var faces = root.ElementVisioSchema2003("FaceNames");
        faces.RemoveNodes();
        foreach (var face in drawing.Faces.Items)
        {
            face.ToXml(faces);
        }

        var colors = root.ElementVisioSchema2003("Colors");
        colors.RemoveNodes();
        int index = 0;
        foreach (var color in drawing.Colors)
        {
            color.AddToElement(colors, index++);
        }

        var pages = root.ElementVisioSchema2003("Pages");
        foreach (var page in drawing.Pages.Items)
        {
            page.AddToElement(pages);
        }

        if (drawing.Windows.Count > 0)
        {
            var windows = Internal.XMLUtil.CreateVisioSchema2003Element("Windows");
            root.Add(windows);
            foreach (var window in drawing.Windows)
            {
                window.AddToElement(windows);
            }
        }
        return dom;
    }
}
