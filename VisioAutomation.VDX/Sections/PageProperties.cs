using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class PageProperties
{
    public DistanceCell PageWidth = new();
    public DistanceCell PageHeight = new();
    public DistanceCell ShdwOffsetX = new();
    public DistanceCell ShdwOffsetY = new();
    public DistanceCell PageScale = new();
    public IntCell DrawingSizeType = new();
    public IntCell DrawingScaleType = new();
    public IntCell InhibitSnap = new();
    public IntCell UIVisibility = new();
    public IntCell ShdwType = new();

    public IntCell ShdwObliqueAngle = new();
    public IntCell ShdwScaleFactor = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("PageProps");
        el.Add(this.PageWidth.ToXml("PageWidth"));
        el.Add(this.PageHeight.ToXml("PageHeight"));
        el.Add(this.ShdwOffsetX.ToXml("ShdwOffsetX"));
        el.Add(this.ShdwOffsetY.ToXml("ShdwOffsetY"));
        el.Add(this.PageScale.ToXml("PageScale"));
        el.Add(this.DrawingSizeType.ToXml("DrawingSizeType"));
        el.Add(this.DrawingScaleType.ToXml("DrawingScaleType"));
        el.Add(this.InhibitSnap.ToXml("InhibitSnap"));
        el.Add(this.UIVisibility.ToXml("UIVisibility"));
        el.Add(this.ShdwType.ToXml("ShdwType"));
        el.Add(this.ShdwObliqueAngle.ToXml("ShdwObliqueAngle"));
        el.Add(this.ShdwScaleFactor.ToXml("ShdwScaleFactor"));
        parent.Add(el);
    }
}