using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Misc
{
    public IntCell NoObjHandles = new();
    public IntCell NonPrinting = new();
    public IntCell NoCtlHandles = new();
    public IntCell NoAlignBox = new();

    public IntCell UpdateAlignBox = new();
    public IntCell HideText = new();
    public IntCell DynFeedback = new();
    public IntCell GlueType = new();
    public IntCell WalkPreference = new();

    public DoubleCell BegTrigger = new();
    public DoubleCell EndTrigger = new();

    public IntCell ObjType = new();
    public IntCell Comment = new();
    public IntCell IsDropSource = new();
    public IntCell NoLiveDynamics = new();
    public IntCell LocalizeMerge = new();

    public IntCell Calendar = new();
    public IntCell LangID = new();
    public DoubleCell ShapeKeywords = new();
    public IntCell DropOnPageScale = new();

    public void AddToElement(SXL.XElement parent)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("Misc");
        el.Add(this.NoObjHandles.ToXml("NoObjHandles"));
        el.Add(this.NonPrinting.ToXml("NonPrinting"));
        el.Add(this.NoCtlHandles.ToXml("NoCtlHandles"));
        el.Add(this.NoAlignBox.ToXml("NoAlignBox"));

        el.Add(this.UpdateAlignBox.ToXml("UpdateAlignBox"));
        el.Add(this.HideText.ToXml("HideText"));
        el.Add(this.DynFeedback.ToXml("DynFeedback"));
        el.Add(this.GlueType.ToXml("GlueType"));
        el.Add(this.WalkPreference.ToXml("WalkPreference"));

        el.Add(this.BegTrigger.ToXml("BegTrigger"));
        el.Add(this.EndTrigger.ToXml("EndTrigger"));

        el.Add(this.ObjType.ToXml("ObjType"));
        el.Add(this.Comment.ToXml("Comment"));
        el.Add(this.IsDropSource.ToXml("IsDropSource"));
        el.Add(this.NoLiveDynamics.ToXml("NoLiveDynamics"));
        el.Add(this.LocalizeMerge.ToXml("LocalizeMerge"));

        el.Add(this.Calendar.ToXml("Calendar"));
        el.Add(this.LangID.ToXml("LangID"));
        el.Add(this.ShapeKeywords.ToXml("ShapeKeywords"));
        el.Add(this.DropOnPageScale.ToXml("DropOnPageScale"));

        parent.Add(el);
    }
}