using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class PrintProperties
{
    public DistanceCell PageLeftMargin = new();
    public DistanceCell PageRightMargin = new();
    public DistanceCell PageTopMargin = new();
    public DistanceCell PageBottomMargin = new();

    public DoubleCell ScaleX = new();
    public DoubleCell ScaleY = new();

    public IntCell PagesX = new();
    public IntCell PagesY = new();

    public DistanceCell CenterX = new();
    public DistanceCell CenterY = new();

    public BoolCell OnPage = new();
    public BoolCell PrintGrid = new();
    public BoolCell PrintPageOrientation = new();
    public IntCell PaperKind = new();
    public IntCell PaperSource = new();

    public IntCell ShdwObliqueAngle = new();
    public IntCell ShdwScaleFactor = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("PrintProps");
        el.Add(this.PageLeftMargin.ToXml("PageLeftMargin"));
        el.Add(this.PageRightMargin.ToXml("PageRightMargin"));
        el.Add(this.PageTopMargin.ToXml("PageTopMargin"));
        el.Add(this.PageBottomMargin.ToXml("PageBottomMargin"));

        el.Add(this.ScaleX.ToXml("ScaleX"));
        el.Add(this.ScaleY.ToXml("ScaleY"));

        el.Add(this.PagesX.ToXml("PagesX"));
        el.Add(this.PagesY.ToXml("PagesY"));

        el.Add(this.CenterX.ToXml("CenterX"));
        el.Add(this.CenterY.ToXml("CenterY"));

        el.Add(this.OnPage.ToXml("OnPage"));
        el.Add(this.PrintGrid.ToXml("PrintGrid"));

        el.Add(this.PrintPageOrientation.ToXml("PrintPageOrientation"));
        el.Add(this.PaperKind.ToXml("PaperKind"));

        el.Add(this.PaperSource.ToXml("PaperSource"));
        el.Add(this.ShdwObliqueAngle.ToXml("ShdwObliqueAngle"));

        el.Add(this.ShdwScaleFactor.ToXml("ShdwScaleFactor"));

        parent.Add(el);
    }
}