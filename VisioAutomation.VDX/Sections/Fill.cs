using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Fill
{
    public ColorCell ForegroundColor = new();
    public ColorCell BackgroundColor = new();
    public IntCell Pattern = new();
    public ColorCell ShadowForegroundColor = new();
    public ColorCell ShadowBackgroundColor = new();
    public IntCell ShadowPattern = new();
    public TransparencyCell ForegroundTransparency = new();
    public TransparencyCell BackgroundTransparency = new();
    public TransparencyCell ShadowForegroundTransparency = new();
    public TransparencyCell ShadowBackgroundTransparency = new();

    public IntCell ShadowType = new();
    public DistanceCell ShadowOffsetX = new();
    public DistanceCell ShadowOffsetY = new();
    public AngleCell ShadowObliqueAngle = new();
    public DoubleCell ShadowScale = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("Fill");
        el.Add(this.ForegroundColor.ToXml("FillForegnd"));
        el.Add(this.BackgroundColor.ToXml("FillBkgnd"));
        el.Add(this.Pattern.ToXml("FillPattern"));
        el.Add(this.ShadowForegroundColor.ToXml("ShdwForegnd"));
        el.Add(this.ShadowBackgroundColor.ToXml("ShdwBkgnd"));
        el.Add(this.ShadowPattern.ToXml("ShdwPattern"));

        el.Add(this.ForegroundTransparency.ToXml("FillForegndTrans"));
        el.Add(this.BackgroundTransparency.ToXml("FillBkgndTrans"));

        el.Add(
            this.ShadowForegroundTransparency.ToXml("ShdwForegndTrans"));
        el.Add(
            this.ShadowBackgroundTransparency.ToXml("ShdwBkgndTrans"));

        el.Add(this.ShadowType.ToXml("ShapeShdwType"));
        el.Add(this.ShadowOffsetX.ToXml("ShapeShdwOffsetX"));
        el.Add(this.ShadowOffsetY.ToXml("ShapeShdwOffsetY"));
        el.Add(this.ShadowObliqueAngle.ToXml("ShapeShdwObliqueAngle"));
        el.Add(this.ShadowScale.ToXml("ShapeShdwScaleFactor"));

        parent.Add(el);
    }
}