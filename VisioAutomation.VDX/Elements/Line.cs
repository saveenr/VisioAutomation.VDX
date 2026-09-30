using VisioAutomation.VDX.Internal;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Line
{
    public ShapeSheet.PointCell Weight = new();
    public ShapeSheet.ColorCell Color = new();
    public ShapeSheet.IntCell Pattern = new();
    public ShapeSheet.DoubleCell Rounding = new();
    public ShapeSheet.IntCell EndArrowSize = new();
    public ShapeSheet.IntCell BeginArrowSize = new();
    public ShapeSheet.IntCell EndArrow = new();
    public ShapeSheet.IntCell BeginArrow = new();
    public ShapeSheet.IntCell Cap = new();
    public ShapeSheet.TransparencyCell Transparency = new();

    public void AddToElement(SXL.XElement parent)
    {
        var line_el = XMLUtil.CreateVisioSchema2003Element("Line");
        line_el.Add(this.Weight.ToXml("LineWeight"));
        line_el.Add(this.Color.ToXml("LineColor"));
        line_el.Add(this.Pattern.ToXml("LinePattern"));
        line_el.Add(this.Rounding.ToXml("Rounding"));
        line_el.Add(this.EndArrowSize.ToXml("EndArrowSize"));
        line_el.Add(this.BeginArrowSize.ToXml("BeginArrowSize"));
        line_el.Add(this.EndArrow.ToXml("EndArrow"));
        line_el.Add(this.BeginArrow.ToXml("BeginArrow"));
        line_el.Add(this.Cap.ToXml("LineCap"));
        line_el.Add(this.Transparency.ToXml("LineColorTrans"));
        parent.Add(line_el);
    }
}