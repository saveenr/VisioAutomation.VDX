using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VisioAutomation.VDX;
using VisioAutomation.VDX.Elements;

namespace TestVisioAutomationVDX.Unit
{
    [TestClass]
    public class OwnershipTests
    {
        private static Drawing NewDrawing()
        {
            return new Drawing(new Template());
        }

        [TestMethod]
        public void Add_DuplicateName_DoesNotTakeOwnershipOfRejectedItem()
        {
            var faces = new FaceList();
            faces.Add(new Face(1, "Arial"));
            var rejected = new Face(2, "ARIAL");
            Assert.ThrowsExactly<ArgumentException>(() => faces.Add(rejected));
            Assert.IsNull(rejected.Parent);
            Assert.IsFalse(faces.Contains(rejected));
            var other = new FaceList();
            other.Add(rejected);
            Assert.IsTrue(other.Contains(rejected));
        }

        [TestMethod]
        public void Lookup_AfterRename_UsesCurrentName()
        {
            var faces = new FaceList();
            var face = new Face(1, "Before");
            faces.Add(face);
            face.Name = "After";
            Assert.IsFalse(faces.ContainsName("Before"));
            Assert.AreSame(face, faces["AFTER"]);
        }

        [TestMethod]
        public void PageAdd_AlreadyOwned_DoesNotChangeDrawing()
        {
            var first = NewDrawing();
            var second = NewDrawing();
            var page = new Page(8, 5);
            first.Pages.Add(page);
            Assert.ThrowsExactly<ArgumentException>(() => second.Pages.Add(page));
            Assert.AreSame(first, page.Drawing);
            Assert.AreEqual(1, first.Pages.Count);
            Assert.AreEqual(0, second.Pages.Count);
        }

        [TestMethod]
        public void PageAdd_DuplicateName_LeavesRejectedPageDetached()
        {
            var drawing = NewDrawing();
            drawing.Pages.Add(new Page(8, 5) { Name = "Same" });
            var rejected = new Page(8, 5) { Name = "Same" };
            Assert.ThrowsExactly<ArgumentException>(() => drawing.Pages.Add(rejected));
            Assert.IsNull(rejected.Drawing);
            Assert.IsNull(rejected.Parent);
        }

        [TestMethod]
        public void ShapeAdd_AlreadyOwned_PreservesIdentityAndOwner()
        {
            var drawing = NewDrawing();
            var first = new Page(8, 5);
            var second = new Page(8, 5);
            drawing.Pages.Add(first);
            drawing.Pages.Add(second);
            var shape = new Shape(drawing.GetMasterMetaData("Rectangle").ID, 1, 1);
            first.Shapes.Add(shape);
            int id = shape.ID;
            string name = shape.Name;
            Assert.ThrowsExactly<ArgumentException>(() => second.Shapes.Add(shape));
            Assert.AreSame(first, shape.Page);
            Assert.AreEqual(id, shape.ID);
            Assert.AreEqual(name, shape.Name);
            Assert.AreEqual(0, second.Shapes.Count);
        }

        [TestMethod]
        public void ShapeAdd_Null_ThrowsArgumentNullException()
        {
            var drawing = NewDrawing();
            var page = new Page(8, 5);
            drawing.Pages.Add(page);
            Assert.ThrowsExactly<ArgumentNullException>(() => page.Shapes.Add(null));
        }

        [TestMethod]
        public void Connect_CrossPageEndpoint_LeavesConnectionsUnchanged()
        {
            var drawing = NewDrawing();
            var first = new Page(8, 5);
            var second = new Page(8, 5);
            drawing.Pages.Add(first);
            drawing.Pages.Add(second);
            int master = drawing.GetMasterMetaData("Rectangle").ID;
            var source = new Shape(master, 1, 1);
            var target = new Shape(master, 2, 2);
            var connector = Shape.CreateDynamicConnector(drawing);
            first.Shapes.Add(source);
            first.Shapes.Add(connector);
            second.Shapes.Add(target);
            Assert.ThrowsExactly<ArgumentException>(() =>
                first.ConnectShapesViaConnector(connector, source, target));
            Assert.AreEqual(0, first.Connects.Count);
        }

        [TestMethod]
        public void Connect_SamePage_CreatesBothEndpoints()
        {
            var drawing = NewDrawing();
            var page = new Page(8, 5);
            drawing.Pages.Add(page);
            int master = drawing.GetMasterMetaData("Rectangle").ID;
            var source = new Shape(master, 1, 1);
            var target = new Shape(master, 2, 2);
            var connector = Shape.CreateDynamicConnector(drawing);
            page.Shapes.Add(source);
            page.Shapes.Add(target);
            page.Shapes.Add(connector);
            page.ConnectShapesViaConnector(connector, source, target);
            Assert.AreEqual(2, page.Connects.Count);
            CollectionAssert.AreEqual(new[] { source.ID, target.ID }, page.Connects.Select(c => c.ToSheet).ToArray());
            Assert.IsTrue(page.Connects.All(c => c.FromSheet == connector.ID));
        }
    }
}
