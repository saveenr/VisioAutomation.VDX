using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class XForm
{
    public DistanceCell PinX = new();
    public DistanceCell PinY = new();
    public DistanceCell Width = new();
    public DistanceCell Height = new();
    public DistanceCell LocPinX = new();
    public DistanceCell LocPinY = new();
    public AngleCell Angle = new();
    public IntCell FlipX = new();
    public IntCell FlipY = new();
    public IntCell FlipMode = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("XForm");
        el.Add(this.PinX.ToXml("PinX"));
        el.Add(this.PinY.ToXml("PinY"));
        el.Add(this.Width.ToXml("Width"));
        el.Add(this.Height.ToXml("Height"));
        el.Add(this.LocPinX.ToXml("LocPinX"));
        el.Add(this.LocPinY.ToXml("LocPinY"));
        el.Add(this.Angle.ToXml("Angle"));
        el.Add(this.FlipX.ToXml("FlipX"));
        el.Add(this.FlipY.ToXml("FlipY"));
        el.Add(this.FlipMode.ToXml("FlipMode"));

        parent.Add(el);
    }
}