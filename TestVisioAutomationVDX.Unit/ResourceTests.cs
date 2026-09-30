using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.VDX.Elements;

namespace TestVisioAutomationVDX.Unit;

[TestClass]
public class ResourceTests
{
    [TestMethod]
    public void DefaultTemplate_IsTheOnlyEmbeddedResource()
    {
        CollectionAssert.AreEqual(
            new[] { "VisioAutomation.VDX.DefaultTemplate.xml" },
            typeof(Drawing).Assembly.GetManifestResourceNames());
        Assert.IsTrue(new Drawing(new VisioAutomation.VDX.Template()).Faces.Count > 0);
    }

    [TestMethod]
    public void DefaultTemplate_HasNoPersonalOrMachineMetadata()
    {
        var xml = XDocument.Parse(Drawing.DefaultTemplateXML);
        XNamespace ns = "http://schemas.microsoft.com/visio/2003/core";
        Assert.AreEqual(string.Empty, xml.Root.Element(ns + "DocumentProperties").Element(ns + "Creator").Value);
        var values = xml.Descendants().Where(element => !element.HasElements).Select(element => element.Value)
            .Concat(xml.Descendants().Attributes().Select(attribute => attribute.Value));
        foreach (var value in values)
        {
            Assert.IsFalse(Regex.IsMatch(value, @"^(?:[A-Za-z]:[\\/]|\\\\|file:)", RegexOptions.IgnoreCase),
                "The embedded template must not refer to an external local file.");
        }
    }
}
