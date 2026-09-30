using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class NURBSTo : GeomRow
{
    public DistanceCell X = new();
    public DistanceCell Y = new();
    public StringCell A = new();
    public StringCell B = new();
    public StringCell C = new();
    public StringCell D = new();
    public StringCell E = new();

    public override void AddToElement(SXL.XElement parent, int index)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("NURBSTo");
        el.SetAttributeValueInt("IX", index);
        el.Add(this.X.ToXml("X"));
        el.Add(this.Y.ToXml("Y"));
        el.Add(this.A.ToXml("A"));
        el.Add(this.B.ToXml("B"));
        el.Add(this.C.ToXml("C"));
        el.Add(this.D.ToXml("D"));
        el.Add(this.E.ToXml("E"));
        parent.Add(el);
    }
}