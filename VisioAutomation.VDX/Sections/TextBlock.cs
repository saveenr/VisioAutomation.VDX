using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class TextBlock
{
    public DistanceCell LeftMargin = new();
    public DistanceCell RightMargin = new();
    public DistanceCell TopMargin = new();
    public DistanceCell BottomMargin = new();

    public IntCell VerticalAlign = new();
    public ColorCell TextBkgnd = new();

    public DistanceCell DefaultTabStop = new();
    public IntCell TextDirection = new();
    public TransparencyCell TextBkgndTrans = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el1 = XMLUtil.CreateVisioSchema2003Element("TextBlock");
        el1.Add(this.LeftMargin.ToXml("LeftMargin"));
        el1.Add(this.RightMargin.ToXml("RightMargin"));
        el1.Add(this.TopMargin.ToXml("TopMargin"));
        el1.Add(this.BottomMargin.ToXml("BottomMargin"));
        el1.Add(this.VerticalAlign.ToXml("VerticalAlign"));
        el1.Add(this.TextBkgnd.ToXml("TextBkgnd"));

        el1.Add(this.DefaultTabStop.ToXml("DefaultTabStop"));
        el1.Add(this.TextDirection.ToXml("TextDirection"));
        el1.Add(this.TextBkgndTrans.ToXml("TextBkgndTrans"));

        parent.Add(el1);
    }
}