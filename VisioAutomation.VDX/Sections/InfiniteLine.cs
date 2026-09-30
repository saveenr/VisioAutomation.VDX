using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class InfiniteLine : GeomRow
{
    public DistanceCell X = new();
    public DistanceCell Y = new();
    public StringCell A = new();
    public StringCell B = new();

    public override void AddToElement(SXL.XElement parent, int index)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("InfiniteLine");
        el.SetAttributeValueInt("IX", index);
        el.Add(this.X.ToXml("X"));
        el.Add(this.Y.ToXml("Y"));
        el.Add(this.A.ToXml("A"));
        el.Add(this.B.ToXml("B"));
        parent.Add(el);
    }
}