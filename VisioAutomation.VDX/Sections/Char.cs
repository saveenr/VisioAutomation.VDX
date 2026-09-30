using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class Char
{
    public IntCell Font = new();
    public IntCell Color = new();
    public EnumCell<CharStyle> Style = new(v => (int)v);
    public EnumCell<CharCase> Case = new(v => (int)v);
    public IntCell Pos = new();
    public DoubleCell FontScale = new();
    public PointCell Size = new();
    public BoolCell DoubleUnderline = new();
    public BoolCell Overline = new();
    public BoolCell Strikethru = new();
    public IntCell Highlight = new();
    public BoolCell DoubleStrikethrough = new();
    public BoolCell RTLText = new();
    public BoolCell UseVertical = new();
    public DoubleCell Letterspace = new();
    public TransparencyCell Transparency = new();
    public IntCell AsianFont = new();
    public IntCell ComplexScriptFont = new();
    public IntCell LocalizeFont = new();
    public IntCell ComplexScriptSize = new();
    public IntCell LangID = new();

    public void AddToElement(SXL.XElement parent, int index)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("Char");

        el.SetAttributeValueInt("IX", index);
        el.Add(this.Font.ToXml("Font"));
        el.Add(this.Color.ToXml("Color"));

        el.Add(this.Style.ToXml("Style"));
        el.Add(this.Case.ToXml("Case"));
        el.Add(this.Pos.ToXml("Pos"));

        el.Add(this.FontScale.ToXml("FontScale"));
        el.Add(this.Size.ToXml("Size"));

        el.Add(this.DoubleUnderline.ToXml("DblUnderline"));
        el.Add(this.Overline.ToXml("Overline"));

        el.Add(this.Strikethru.ToXml("Strikethru"));
        el.Add(this.Highlight.ToXml("Highlight"));

        el.Add(this.DoubleStrikethrough.ToXml("DoubleStrikethrough"));
        el.Add(this.RTLText.ToXml("RTLText"));
        el.Add(this.UseVertical.ToXml("UseVertical"));

        el.Add(this.Letterspace.ToXml("Letterspace"));
        el.Add(this.Transparency.ToXml("ColorTrans"));

        el.Add(this.AsianFont.ToXml("AsianFont"));
        el.Add(this.ComplexScriptFont.ToXml("ComplexScriptFont"));

        el.Add(this.LocalizeFont.ToXml("LocalizeFont"));
        el.Add(this.ComplexScriptSize.ToXml("ComplexScriptSize"));
        el.Add(this.LangID.ToXml("LangID"));

        parent.Add(el);
    }
}