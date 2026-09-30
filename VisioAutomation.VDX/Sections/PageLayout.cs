using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class PageLayout
{
    public BoolCell ResizePage = new();
    public BoolCell EnableGrid = new();
    public BoolCell DynamicsOff = new();
    public IntCell PlaceStyle = new();

    public IntCell RouteStyle = new();
    public DoubleCell PlaceDepth = new();
    public IntCell PlowCode = new();

    public IntCell LineJumpCode = new();
    public IntCell LineJumpStyle = new();

    public IntCell PageLineJumpDirX = new();
    public IntCell PageLineJumpDirY = new();

    public IntCell LineToNodeX = new();
    public IntCell LineToNodeY = new();

    public IntCell BlockSizeX = new();
    public IntCell BlockSizeY = new();

    public IntCell AvenueSizeX = new();
    public IntCell AvenueSizeY = new();

    public IntCell LineToLineX = new();
    public IntCell LineToLineY = new();

    public IntCell LineJumpFactorX = new();
    public IntCell LineJumpFactorY = new();

    public IntCell LineAdjustFrom = new();
    public IntCell LineAdjustTo = new();

    public IntCell PlaceFlip = new();
    public IntCell LineRouteExt = new();
    public IntCell PageShapeSplit = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("PageLayout");
        el.Add(this.ResizePage.ToXml("ResizePage"));
        el.Add(this.EnableGrid.ToXml("EnableGrid"));
        el.Add(this.DynamicsOff.ToXml("DynamicsOff"));
        el.Add(this.PlaceStyle.ToXml("PlaceStyle"));
        el.Add(this.RouteStyle.ToXml("RouteStyle"));
        el.Add(this.PlaceDepth.ToXml("PlaceDepth"));
        el.Add(this.PlowCode.ToXml("PlowCode"));
        el.Add(this.LineJumpCode.ToXml("LineJumpCode"));
        el.Add(this.LineJumpStyle.ToXml("LineJumpStyle"));
        el.Add(this.PageLineJumpDirX.ToXml("PageLineJumpDirX"));
        el.Add(this.PageLineJumpDirY.ToXml("PageLineJumpDirY"));
        el.Add(this.LineToNodeX.ToXml("LineToNodeX"));
        el.Add(this.LineToNodeY.ToXml("LineToNodeY"));
        el.Add(this.BlockSizeX.ToXml("BlockSizeX"));
        el.Add(this.BlockSizeY.ToXml("BlockSizeY"));
        el.Add(this.AvenueSizeX.ToXml("AvenueSizeX"));
        el.Add(this.AvenueSizeY.ToXml("AvenueSizeY"));
        el.Add(this.LineToLineX.ToXml("LineToLineX"));
        el.Add(this.LineToLineY.ToXml("LineToLineY"));
        el.Add(this.LineJumpFactorX.ToXml("LineJumpFactorX"));
        el.Add(this.LineJumpFactorY.ToXml("LineJumpFactorY"));
        el.Add(this.LineAdjustFrom.ToXml("LineAdjustFrom"));
        el.Add(this.LineAdjustTo.ToXml("LineAdjustTo"));

        el.Add(this.PlaceFlip.ToXml("PlaceFlip"));
        el.Add(this.LineRouteExt.ToXml("LineRouteExt"));
        el.Add(this.PageShapeSplit.ToXml("PageShapeSplit"));

        parent.Add(el);
    }
}