using VisioAutomation.VDX.Internal.Extensions;
using VisioAutomation.VDX.Enums;
using VisioAutomation.VDX.Internal;
using VisioAutomation.VDX.ShapeSheet;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Sections;

public class ParagraphFormat
{
    public DoubleCell IndFirst = new();
    public DoubleCell IndLeft = new();
    public DoubleCell IndRight = new();
    public DoubleCell SpLine = new();
    public DoubleCell SpBefore = new();
    public DoubleCell SpAfter = new();
    public EnumCell<ParaHorizontalAlignment> HorzAlign = new(v => (int)v);
    public IntCell Bullet = new();
    public DoubleCell BulletStr = new();
    public IntCell BulletFont = new();
    public BoolCell LocalizeBulletFont = new();
    public IntCell BulletFontSize = new();
    public DoubleCell TextPosAfterBullet = new();
    public IntCell Flags = new();

    public void AddToElement(SXL.XElement parent, int ix)
    {
        var el = XMLUtil.CreateVisioSchema2003Element("Para");

        el.SetAttributeValueInt("IX", ix);
        el.Add(this.IndFirst.ToXml("IndFirst"));
        el.Add(this.IndLeft.ToXml("IndLeft"));
        el.Add(this.IndRight.ToXml("IndRight"));

        el.Add(this.SpLine.ToXml("SpLine"));
        el.Add(this.SpBefore.ToXml("SpBefore"));
        el.Add(this.SpAfter.ToXml("SpAfter"));

        el.Add(this.HorzAlign.ToXml("HorzAlign"));
        el.Add(this.Bullet.ToXml("Bullet"));
        el.Add(this.BulletStr.ToXml("BulletStr"));
        el.Add(this.BulletFont.ToXml("BulletFont"));
        el.Add(this.LocalizeBulletFont.ToXml("LocalizeBulletFont"));
        el.Add(this.BulletFontSize.ToXml("BulletFontSize"));
        el.Add(this.TextPosAfterBullet.ToXml("TextPosAfterBullet"));
        el.Add(this.Flags.ToXml("Flags"));

        parent.Add(el);
    }
}