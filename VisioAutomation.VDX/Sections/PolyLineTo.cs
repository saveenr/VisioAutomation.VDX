using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class PolyLineTo : GeomRow
{
    public DistanceCell X = new();
    public DistanceCell Y = new();
    public StringCell A = new();

    public override void AddToElement(SXL.XElement parent, int index)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("PolylineTo");
        el.SetAttributeValueInt("IX", index);
        el.Add(this.X.ToXml("X"));
        el.Add(this.Y.ToXml("Y"));
        el.Add(this.A.ToXml("A"));
        parent.Add(el);
    }
}