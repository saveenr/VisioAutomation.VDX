using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Protection
{
    public BoolCell Width = new();
    public BoolCell Height = new();
    public BoolCell MoveX = new();
    public BoolCell MoveY = new();
    public BoolCell Aspect = new();
    public BoolCell Delete = new();
    public BoolCell Begin = new();
    public BoolCell Rotate = new();
    public BoolCell Crop = new();
    public BoolCell VtxEdit = new();

    public BoolCell TextEdit = new();
    public BoolCell Format = new();
    public BoolCell Group = new();
    public BoolCell CalcWH = new();
    public BoolCell Select = new();
    public BoolCell CustProp = new();

    //<vx:Protection xmlns:vx="http://schemas.microsoft.com/visio/2006/extension">
    //<vx:LockFromGroupFormat>0</vx:LockFromGroupFormat>
    //<vx:LockThemeColors>0</vx:LockThemeColors>
    //<vx:LockThemeEffects>0</vx:LockThemeEffects>
    //</vx:Protection>

    public BoolCell FromGroupFormat = new();
    public BoolCell ThemeColors = new();
    public BoolCell ThemeEffects = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el1 = XMLUtil.CreateVisioSchema2003Element("Protection");
        el1.Add(this.Width.ToXml("LockWidth"));
        el1.Add(this.Height.ToXml("LockHeight"));

        el1.Add(this.MoveX.ToXml("LockMoveX"));
        el1.Add(this.MoveY.ToXml("LockMoveY"));

        el1.Add(this.Aspect.ToXml("LockAspect"));
        el1.Add(this.Delete.ToXml("LockDelete"));

        el1.Add(this.Begin.ToXml("LockBegin"));
        el1.Add(this.Rotate.ToXml("LockRotate"));

        el1.Add(this.Crop.ToXml("LockCrop"));
        el1.Add(this.VtxEdit.ToXml("LockVtxEdit"));

        el1.Add(this.TextEdit.ToXml("LockTextEdit"));
        el1.Add(this.Format.ToXml("LockFormat"));

        el1.Add(this.Group.ToXml("LockGroup"));
        el1.Add(this.CalcWH.ToXml("LockCalcWH"));

        el1.Add(this.Select.ToXml("LockSelect"));
        el1.Add(this.CustProp.ToXml("LockCustProp"));

        parent.Add(el1);

        var el2 = XMLUtil.CreateVisioSchema2006Element("Protection");
        el2.Add(this.FromGroupFormat.ToXml2006("LockFromGroupFormat"));
        el2.Add(this.ThemeColors.ToXml2006("LockThemeColors"));
        el2.Add(this.ThemeEffects.ToXml2006("LockThemeEffects"));
        parent.Add(el2);
    }

    public void SetAll(bool v)
    {
        this.Width.Result = v;
        this.Height.Result = v;
        this.MoveX.Result = v;
        this.MoveY.Result = v;
        this.Aspect.Result = v;
        this.Delete.Result = v;
        this.Begin.Result = v;
        this.Rotate.Result = v;
        this.Crop.Result = v;
        this.VtxEdit.Result = v;
        this.TextEdit.Result = v;
        this.Format.Result = v;
        this.Group.Result = v;
        this.CalcWH.Result = v;
        this.Select.Result = v;
        this.CustProp.Result = v;
    }
}