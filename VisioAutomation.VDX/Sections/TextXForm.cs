using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class TextXForm
{
    public DistanceCell PinX = new();
    public DistanceCell PinY = new();
    public DistanceCell Width = new();
    public DistanceCell Height = new();
    public DistanceCell LocPinX = new();
    public DistanceCell LocPinY = new();
    public AngleCell Angle = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("TextXForm");
        el.Add(this.PinX.ToXml("TxtPinX"));
        el.Add(this.PinY.ToXml("TxtPinY"));
        el.Add(this.Width.ToXml("TxtWidth"));
        el.Add(this.Height.ToXml("TxtHeight"));
        el.Add(this.LocPinX.ToXml("TxtLocPinX"));
        el.Add(this.LocPinY.ToXml("TxtLocPinY"));
        el.Add(this.Angle.ToXml("TxtAngle"));
        parent.Add(el);
    }
}