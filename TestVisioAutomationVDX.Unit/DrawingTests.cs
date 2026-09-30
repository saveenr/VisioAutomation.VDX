using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.VDX;
using VisioAutomation.VDX.Elements;

namespace TestVisioAutomationVDX.Unit;

[TestClass]
public class DrawingTests
{
    private static readonly XNamespace Vdx = "http://schemas.microsoft.com/visio/2003/core";
    private string directory;

    [TestInitialize]
    public void Initialize()
    {
        this.directory = Path.Combine(Path.GetTempPath(), "vdx-tests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(this.directory);
    }

    [TestCleanup]
    public void Cleanup()
    {
        Directory.Delete(this.directory, true);
    }

    private static Drawing NewDrawing()
    {
        var drawing = new Drawing(new Template());
        drawing.Pages.Add(new Page(8, 5));
        drawing.Windows.Add(new DocumentWindow { Page = drawing.Pages.Items.Single().ID });
        return drawing;
    }

    private XDocument Save(Drawing drawing, string name)
    {
        string path = Path.Combine(this.directory, name + ".vdx");
        drawing.Save(path);
        return XDocument.Load(path, LoadOptions.PreserveWhitespace);
    }

    [TestMethod]
    public void Save_Twice_ProducesIdenticalDocuments()
    {
        var drawing = NewDrawing();
        var first = this.Save(drawing, "first");
        var second = this.Save(drawing, "second");
        Assert.AreEqual(1, second.Root.Elements(Vdx + "DocumentProperties").Count());
        Assert.AreEqual(1, second.Root.Element(Vdx + "Pages").Elements(Vdx + "Page").Count());
        Assert.AreEqual(1, second.Root.Elements(Vdx + "Windows").Count());
        Assert.IsTrue(XNode.DeepEquals(first, second));
    }

    [TestMethod]
    public void ToXml_ReturnsIndependentSnapshots()
    {
        var drawing = NewDrawing();
        var expected = drawing.ToXml();
        var modified = drawing.ToXml();
        modified.Root.RemoveAll();
        Assert.IsTrue(XNode.DeepEquals(expected, drawing.ToXml()));
        Assert.IsTrue(XNode.DeepEquals(expected, this.Save(drawing, "snapshot")));
    }

    [TestMethod]
    public void ToXml_InvalidWindow_RejectsMissingPage()
    {
        var drawing = NewDrawing();
        ((DocumentWindow)drawing.Windows.Single()).Page = int.MaxValue;
        Assert.ThrowsExactly<ArgumentException>(() => drawing.ToXml());
        Assert.ThrowsExactly<ArgumentException>(() => this.Save(drawing, "invalid"));
    }

    [TestMethod]
    public void Save_ValidatesFilename()
    {
        var drawing = NewDrawing();
        Assert.ThrowsExactly<ArgumentNullException>(() => drawing.Save(null));
        Assert.ThrowsExactly<ArgumentException>(() => drawing.Save("drawing.vsdx"));
        drawing.Save(Path.Combine(this.directory, "uppercase.VDX"));
    }

    [TestMethod]
    public void AddFace_EmptyTemplate_StartsAtOne()
    {
        var drawing = new Drawing(new Template("<VisioDocument xmlns='http://schemas.microsoft.com/visio/2003/core'/>"));
        Assert.AreEqual(1, drawing.AddFace("First").ID);
    }

    [TestMethod]
    public void Save_AfterEditingPage_ContainsOnlyCurrentState()
    {
        var drawing = NewDrawing();
        var page = drawing.Pages.Items.Single();
        this.Save(drawing, "before");
        page.Name = "Changed";
        var after = this.Save(drawing, "after");
        var pages = after.Root.Element(Vdx + "Pages").Elements(Vdx + "Page").ToList();
        Assert.AreEqual(1, pages.Count);
        Assert.AreEqual("Changed", (string)pages[0].Attribute("NameU"));
    }

    [TestMethod]
    public void Save_FailedWrite_DoesNotPolluteNextSave()
    {
        var drawing = NewDrawing();
        Assert.ThrowsExactly<DirectoryNotFoundException>(() =>
            drawing.Save(Path.Combine(this.directory, "missing", "drawing.vdx")));
        var xml = this.Save(drawing, "retry");
        Assert.AreEqual(1, xml.Root.Elements(Vdx + "DocumentProperties").Count());
        Assert.AreEqual(1, xml.Root.Element(Vdx + "Pages").Elements(Vdx + "Page").Count());
    }

    [TestMethod]
    public void Save_FractionalCells_AreCultureIndependent()
    {
        var original = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
            var drawing = NewDrawing();
            drawing.Pages.Items.Single().PageProperties.PageWidth.Result = 8.25;
            var xml = this.Save(drawing, "culture");
            Assert.AreEqual("8.25", xml.Descendants(Vdx + "PageWidth").Last().Value);
        }
        finally { Thread.CurrentThread.CurrentCulture = original; }
    }

    [TestMethod]
    public void Save_FormattedText_PreservesWhitespaceAndEscapesXml()
    {
        var drawing = NewDrawing();
        var page = drawing.Pages.Items.Single();
        var shape = new Shape(drawing.GetMasterMetaData("Rectangle").ID, 1, 1);
        page.Shapes.Add(shape);
        shape.Text.Add(" left & <\n", 0, 0, null);
        shape.Text.Add("right > ", 1, null, null);
        var xml = this.Save(drawing, "text");
        var text = xml.Root.Element(Vdx + "Pages").Descendants(Vdx + "Text").Single();
        Assert.AreEqual(" left & <\nright > ", text.Value);
        Assert.AreEqual(2, text.Elements(Vdx + "cp").Count());
    }

    [TestMethod]
    public void Template_MissingOptionalSections_CanGenerateDrawing()
    {
        var xml = XDocument.Parse(Drawing.DefaultTemplateXML);
        foreach (var name in new[] { "DocumentProperties", "Pages", "FaceNames", "Colors" })
            xml.Root.Element(Vdx + name)?.Remove();
        var drawing = new Drawing(new Template(xml.ToString()));
        drawing.Pages.Add(new Page(8, 5));
        var saved = this.Save(drawing, "minimal");
        Assert.AreEqual(1, saved.Root.Element(Vdx + "Pages").Elements().Count());
    }

    [TestMethod]
    public void Template_WrongNamespace_RejectsNonVdxXml()
    {
        Assert.ThrowsExactly<ArgumentException>(() => new Drawing(new Template("<VisioDocument />")));
    }

    [TestMethod]
    public void AddFace_SparseIds_AllocatesAboveExistingMaximum()
    {
        var drawing = NewDrawing();
        drawing.Faces.Add(new Face(50000, "Sparse font"));
        var added = drawing.AddFace("New font");
        Assert.AreEqual(50001, added.ID);
        Assert.AreSame(added, drawing.AddFace("NEW FONT"));
    }

    [TestMethod]
    public void FaceList_DuplicateId_IsRejectedWithoutTakingOwnership()
    {
        var faces = new FaceList();
        faces.Add(new Face(7, "First"));
        var duplicate = new Face(7, "Second");
        Assert.ThrowsExactly<ArgumentException>(() => faces.Add(duplicate));
        Assert.IsNull(duplicate.Parent);
        Assert.AreEqual(1, faces.Count);
    }
}
