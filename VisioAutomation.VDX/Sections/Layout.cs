using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Layout
{
    public BoolCell ShapePermeableX = new();
    public BoolCell ShapePermeableY = new();
    public BoolCell ShapePermeablePlace = new();
    public EnumCell<ShapeFixedCodeType> ShapeFixedCode = new(v => (int)v);
    public EnumCell<ShapePlowCodeType> ShapePlowCode = new(v => (int)v);
    public EnumCell<RouteStyle> ShapeRouteStyle = new(v => (int)v);
    public EnumCell<ConFixedCode> ConFixedCode = new(v => (int)v);
    public EnumCell<ConLineJumpCode> ConLineJumpCode = new(v => (int)v);
    public EnumCell<ConLineJumpStyle> ConLineJumpStyle = new(v => (int)v);
    public EnumCell<ConLineJumpDirX> ConLineJumpDirX = new(v => (int)v);
    public EnumCell<ConLineJumpDirY> ConLineJumpDirY = new(v => (int)v);
    public EnumCell<PlaceFlip> ShapePlaceFlip = new(v => (int)v);
    public EnumCell<ConLineRouteExt> ConLineRouteExt = new(v => (int)v);
    public EnumCell<PageShapeSplit> ShapeSplit = new(v => (int)v);

    public EnumCell<ShapeSplittable> ShapeSplittable = new(v => (int)v);

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("Layout");

        el.Add(this.ShapePermeableX.ToXml("ShapePermeableX"));
        el.Add(this.ShapePermeableY.ToXml("ShapePermeableY"));
        el.Add(this.ShapePermeablePlace.ToXml("ShapePermeablePlace"));
        el.Add(this.ShapeFixedCode.ToXml("ShapeFixedCode"));
        el.Add(this.ShapePlowCode.ToXml("ShapePlowCode"));
        el.Add(this.ShapeRouteStyle.ToXml("ShapeRouteStyle"));
        el.Add(this.ConFixedCode.ToXml("ConFixedCode"));
        el.Add(this.ConLineJumpCode.ToXml("ConLineJumpCode"));
        el.Add(this.ConLineJumpStyle.ToXml("ConLineJumpStyle"));
        el.Add(this.ConLineJumpDirX.ToXml("ConLineJumpDirX"));
        el.Add(this.ConLineJumpDirY.ToXml("ConLineJumpDirY"));
        el.Add(this.ShapePlaceFlip.ToXml("ShapePlaceFlip"));
        el.Add(this.ConLineRouteExt.ToXml("ConLineRouteExt"));

        el.Add(this.ShapeSplit.ToXml("ShapeSplit"));
        el.Add(this.ShapeSplittable.ToXml("ShapeSplittable"));

        parent.Add(el);
    }
}