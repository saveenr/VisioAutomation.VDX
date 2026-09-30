using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.Extensions;
using VACUSTPROP=VisioAutomation.Shapes;
using IVisio = Microsoft.Office.Interop.Visio;
using VA = VisioAutomation;
using SXL = System.Xml.Linq;

namespace TestVisioAutomationVDX
{
    [TestClass]
    public class VDX_Tests : BaseVDXTest
    {
        public IVisio.Document TryOpen(IVisio.Documents docs, string filename)
        {
            using (var scope = new VA.Application.AlertResponseScope(docs.Application, VA.Application.AlertResponseCode.No))
            {
                var doc = docs.Open(filename);
                return doc;
            }
        }

        public void VerifyDocCanBeLoaded(string filename)
        {
            var app = this.GetApplication();
            string source = this.CopyForOpen(filename);
            var document = this.TryOpen(app.Documents, source);
            Assert.IsNotNull(document);
            var records = this.GetLoadRecords(app, source);
            Assert.AreEqual(0, records.Count(record => record.Type == "Error"),
                "Visio reported errors while loading " + source);
        }

        private List<VA.Application.Logging.LogRecord> GetLoadRecords(IVisio.Application app, string source)
        {
            string log_path = VA.Application.Logging.LoggingHelper.GetXmlErrorLogFilename(app);
            if (!File.Exists(log_path))
            {
                return new List<VA.Application.Logging.LogRecord>();
            }
            var log = this.ReadLog(log_path);
            // Each open uses a unique path, so old sessions cannot satisfy this check.
            return log.LogSessions
                .Where(session => string.Equals(session.Source, source, StringComparison.OrdinalIgnoreCase))
                .SelectMany(session => session.LogRecords).ToList();
        }

        private VA.Application.Logging.XmlErrorLog ReadLog(string log_path)
        {
            string snapshot = this.NewOutputPath() + ".log";
            string text;
            using (var stream = new FileStream(log_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                text = reader.ReadToEnd();
            }
            // VisioAutomation2010 3.0.0 cannot parse space-padded single-digit days.
            // Normalize a private copy until the upstream parser fix is published.
            text = System.Text.RegularExpressions.Regex.Replace(text,
                @"(?m)^(\w{3} \w{3}) +(\d{1,2})( \d{2}:\d{2}:\d{2} \d{4} Begin Session)",
                match => match.Groups[1].Value + " " + match.Groups[2].Value.PadLeft(2, '0') + match.Groups[3].Value);
            File.WriteAllText(snapshot, text);
            return new VA.Application.Logging.XmlErrorLog(snapshot);
        }

        [TestMethod]
        public void LogCompatibility_SpacePaddedDate_PreservesWarnings()
        {
            string source = this.NewOutputPath() + ".log";
            File.WriteAllText(source, "Open VDX Processing Log\n\nSource: sample.vdx\n\n" +
                "Mon May  4 05:41:19 2026 Begin Session\n\n[Warning] DataType:\nDescription: Example\n\n");
            var session = this.ReadLog(source).LogSessions.Single();
            Assert.AreEqual(new DateTime(2026, 5, 4, 5, 41, 19), session.StartTime);
            Assert.AreEqual("Warning", session.LogRecords.Single().Type);
        }

        [TestMethod]
        public void VDX_MultiPageDocument()
        {
            string output_filename = this.NewOutputPath();

            var template = new VisioAutomation.VDX.Template(); // the default template
            var doc = new VisioAutomation.VDX.Elements.Drawing(template);

            var Page01 = VDX_Files.GetPage01_Simple_Fill_Format(doc);
            var Page02 = VDX_Files.GetPage02_Locking(doc);
            var Page03 = VDX_Files.GetPage03_Text_Block(doc);
            var Page04 = VDX_Files.GetPage04_Simple_Text(doc);
            var Page05 = VDX_Files.GetPage05_Formatted_Text(doc);
            var Page06 = VDX_Files.GetPage06_All_FillPatterns(doc);
            var Page08 = VDX_Files.GetPage08_Connector_With_Geometry(doc);
            var Page09 = VDX_Files.GetPage09_Layout(doc);
            var Page10 = VDX_Files.GetPage10_layers(doc);
            var Page11 = VDX_Files.GetPage11_Add_color(doc);
            var Page12 = VDX_Files.GetPage12_AdjustToTextSize(doc);
            var Page13 = VDX_Files.GetPage13_MultipleConnectors(doc);
            var Page14 = VDX_Files.GetPage14_Hyperlinks(doc);

            var w1 = new VisioAutomation.VDX.Elements.DocumentWindow();
            w1.ShowGrid = false;
            w1.ShowGuides = false;
            w1.ShowConnectionPoints = false;
            w1.ShowPageBreaks = false;
            w1.Page = Page01.ID; // point to first page we created
            
            doc.Windows.Add(w1);

            doc.Save(output_filename);
            
            // Verify this file can be loaded
            this.VerifyDocCanBeLoaded(output_filename);
        }

        [TestMethod]
        public void VDX_CustomProperties()
        {
            string filename = this.NewOutputPath();

            var template = new VisioAutomation.VDX.Template();
            var doc_node = new VisioAutomation.VDX.Elements.Drawing(template);

            int rect_id = doc_node.GetMasterMetaData("REctAngle").ID;

            var node_page = new VisioAutomation.VDX.Elements.Page(8, 5);
            doc_node.Pages.Add(node_page);

            var node_shape = new VisioAutomation.VDX.Elements.Shape(rect_id, 4, 2, 3, 2);
            node_shape.CustomProps = new VisioAutomation.VDX.Elements.CustomProps();

            var node_custprop0 = new VisioAutomation.VDX.Elements.CustomProp("PROP1");
            node_custprop0.Value = "VALUE1";
            node_shape.CustomProps.Add(node_custprop0);

            var node_custprop1 = new VisioAutomation.VDX.Elements.CustomProp("PROP2");
            node_custprop1.Value = "123";
            node_custprop1.Type.Result = VisioAutomation.VDX.Enums.CustomPropType.String;
            node_shape.CustomProps.Add(node_custprop1);

            var node_custprop2 = new VisioAutomation.VDX.Elements.CustomProp("PROP3");
            node_custprop2.Value = "456";
            node_custprop2.Type.Result = VisioAutomation.VDX.Enums.CustomPropType.Number;
            node_shape.CustomProps.Add(node_custprop2);

            node_page.Shapes.Add(node_shape);

            doc_node.Save(filename);

            var app = this.GetApplication();
            var docs = app.Documents;
            var doc = docs.Add(filename);

            var page = app.ActivePage;
            var shapes = page.Shapes;
            Assert.AreEqual(1,page.Shapes.Count);

            var shape = page.Shapes[1];
            var customprops = VACUSTPROP.CustomPropertyHelper.GetDictionary(shape, VA.Core.CellValueType.Formula);

            Assert.IsTrue(customprops.ContainsKey("PROP1"));
            Assert.AreEqual("\"VALUE1\"",customprops["PROP1"].Formula.Value);


            Assert.IsTrue(customprops.ContainsKey("PROP2"));
            Assert.AreEqual("\"123\"", customprops["PROP2"].Formula.Value);
            Assert.AreEqual("0", customprops["PROP2"].Type.Value);

            Assert.IsTrue(customprops.ContainsKey("PROP3"));
            Assert.AreEqual("\"456\"", customprops["PROP3"].Formula.Value);
            Assert.AreEqual("2", customprops["PROP3"].Type.Value);

        }

        [TestMethod]
        public void VDX_CustomTemplate()
        {
            string input_filename = this.GetTestResultsOutPath(@"datafiles\template_router.vdx");
            string output_filename = this.NewOutputPath();
            
            // Load the template
            string template_xml = File.ReadAllText(input_filename);

            var template = new VisioAutomation.VDX.Template(template_xml);
            var doc = new VisioAutomation.VDX.Elements.Drawing(template);
            var page = new VisioAutomation.VDX.Elements.Page(8, 4);

            doc.Pages.Add(page);

            // add layers
            var layer0 = page.AddLayer("Layer0", 0);
            var layer1 = page.AddLayer("Layer1", 1);
            var layer2 = page.AddLayer("Layer2", 2);

            // create layout
            var layout = new VisioAutomation.VDX.Sections.Layout();
            layout.ShapeRouteStyle.Result = VisioAutomation.VDX.Enums.RouteStyle.TreeEW;

            // find the id of the master for rounded rectangles
            int shapeMasterNameId = doc.GetMasterMetaData("Router").ID;
            bool shapeMasterNameGroup = doc.GetMasterMetaData("Router").IsGroup;

            // add shape1
            var shape1 = new VisioAutomation.VDX.Elements.Shape(shapeMasterNameId, shapeMasterNameGroup, 1, 3);
            page.Shapes.Add(shape1);
            shape1.Text.Add("Router1");
            shape1.Layout = layout;

            // add shape2
            var shape2 = new VisioAutomation.VDX.Elements.Shape(shapeMasterNameId, shapeMasterNameGroup, 5, 3);
            page.Shapes.Add(shape2);
            shape2.Text.Add("Router2");
            shape2.Layout = shape1.Layout;

            // add shape3 - this is the dynamic connector
            var shape3 = VisioAutomation.VDX.Elements.Shape.CreateDynamicConnector(doc);
            shape3.XForm1D.BeginX.Result = 1;
            shape3.XForm1D.EndX.Result = 5;
            shape3.XForm1D.BeginY.Result = 3;
            shape3.XForm1D.EndY.Result = 3;
            page.Shapes.Add(shape3);
            shape3.Geom = new VisioAutomation.VDX.Sections.Geom();
            shape3.Geom.Rows.Add(new VisioAutomation.VDX.Sections.MoveTo(1, 3));
            shape3.Geom.Rows.Add(new VisioAutomation.VDX.Sections.LineTo(5, 3));

            shape3.Layout = shape1.Layout;
            page.ConnectShapesViaConnector(shape3, shape1, shape2);

            // handle layers
            shape3.LayerMembership = new List<int> {layer0.Index, layer2.Index};
            shape1.LayerMembership = new List<int> {layer1.Index};
            shape2.LayerMembership = new List<int> {layer2.Index};

            // write document to disk as .vdx file

            doc.Save(output_filename);

            this.VerifyDocCanBeLoaded(output_filename);
        }

        [TestMethod]
        public void VDX_CheckNoErrorOnLoad()
        {
            var folder = this.TestResultsOutFolder;
            
            string input_filename = this.GetTestResultsOutPath(@"datafiles\template_router.vdx");

            this.VerifyDocCanBeLoaded(input_filename);
        }


        [TestMethod]
        public void VDX_DetectLoadWarnings()
        {
            string input_filename = this.CopyForOpen(this.GetTestResultsOutPath(@"datafiles\vdx_with_warnings_1.vdx"));
 
            // Load the VDX
            var app = this.GetApplication();
            var version = VA.Application.ApplicationHelper.GetVersion(app);

            var doc = this.TryOpen(app.Documents, input_filename);
            
            // See what happened
            var records = this.GetLoadRecords(app, input_filename);
            var warnings = records.Where(r => r.Type == "Warning").ToList();
            var errors = records.Where(r => r.Type == "Error").ToList();

            // Verify
            int expected_errors = 0;  // this VDX should not report any errors
            int expected_warnings = 4; // this VDX should contain four warnings for Visio2010 and two warnings for Visio 2013         
            if (version.Major >= 15)
            {
                expected_warnings = 2;
            }

            Assert.AreEqual(expected_errors, errors.Count); // this VDX should not report any errors
            Assert.AreEqual(expected_warnings, warnings.Count); // this VDX should contain exactly two warnings                                
            Assert.AreEqual(1, app.Documents.Count);

        }
    }
}
