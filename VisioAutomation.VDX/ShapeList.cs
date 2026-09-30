namespace VisioAutomation.VDX;

public class ShapeList : NamedNodeList<Elements.Shape>
{
    private readonly Elements.Page page_el;

    public ShapeList(Elements.Page page_el) :
        base(shape => shape.Name)
    {
        this.page_el = page_el;
    }

    public override void Add(Elements.Shape shape)
    {
        this.ValidateItem(shape);
        if (this.page_el.Drawing == null)
        {
            throw new System.ArgumentException(
                "page must to added to a drawing before shapes can be added to the page");
        }

        var master_md = this.page_el.Drawing.GetMasterMetData(shape.Master);


        var culture = System.Globalization.CultureInfo.InvariantCulture;
        string name = string.Format(culture, "Shape.{0}", this.page_el.Drawing.CurrentShapeID);
        this.ValidateName(name);
        shape.Page = this.page_el;
        shape._id = this.page_el.Drawing.GetNextShapeID();
        shape.Name = name;

        base.Add(shape);

        this.page_el.Drawing.AccountForMasteSubshapes(master_md.SubShapeCount);
    }
}
