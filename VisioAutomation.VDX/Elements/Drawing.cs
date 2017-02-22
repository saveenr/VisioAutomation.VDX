using System.Collections.Generic;
using System.Linq;
using VisioAutomation.VDX.Internal.Extensions;
using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Elements
{
    public class Drawing : Node
    {
        private readonly Dictionary<string, MasterMetadata> master_metadata =
            new Dictionary<string, MasterMetadata>(System.StringComparer.InvariantCultureIgnoreCase);

        public Sections.DocumentProperties DocumentProperties = new Sections.DocumentProperties();

        internal int CurrentShapeID = -100;

        private readonly SXL.XDocument dom;

        public Drawing(Template template)
        {
            if(template == null)
            {
                throw new System.ArgumentNullException(nameof(template));
            }

            this.dom = template.LoadCleanDOM();

            this.Pages = new PageList(this);
            this.Faces = new FaceList();
            this.Windows = new List<Window>();
            this.Colors = new List<ColorEntry>();

            var masters_el = this.dom.Root.ElementVisioSchema2003("Masters");
            if(masters_el == null)
            {
                throw new System.InvalidOperationException();
            }

            // Store information about each master found in the drawing
            foreach(var master_el in masters_el.ElementsVisioSchema2003("Master"))
            {
                var name = master_el.Attribute("NameU").Value;
                var id = int.Parse(master_el.Attribute("ID").Value);

                var subshapes = master_el.Descendants()
                    .Where(el => el.Name.LocalName == "Shape")
                    .ToList();

                var count_groups = subshapes.Count(shape_el => shape_el.Attribute("Type").Value == "Group");

                bool master_is_group = count_groups > 0;

                var md = new MasterMetadata();
                md.Name = name;
                md.ID = id;
                md.IsGroup = master_is_group;
                md.SubShapeCount = subshapes.Count();

                this.master_metadata[md.Name] = md;

                this.CurrentShapeID = 1;
            }

            var facenames_el = this.dom.Root.ElementVisioSchema2003("FaceNames");
            foreach(var face_el in facenames_el.ElementsVisioSchema2003("FaceName"))
            {
                var id = int.Parse(face_el.Attribute("ID").Value);
                var name = face_el.Attribute("Name").Value;
                var face = new Face(id, name);
                this.Faces.Add(face);
            }

            var colors_el = this.dom.Root.ElementVisioSchema2003("Colors");
            foreach(var color_el in colors_el.ElementsVisioSchema2003("ColorEntry"))
            {
                var rgb_s = color_el.Attribute("RGB").Value;
                int rgb = int.Parse(rgb_s.Substring(1), System.Globalization.NumberStyles.AllowHexSpecifier);
                var ce = new ColorEntry();
                ce.RGB = rgb;
                this.Colors.Add(ce);
            }
        }

        public PageList Pages { get; }

        public FaceList Faces { get; }

        public List<Window> Windows { get; }

        public List<ColorEntry> Colors { get; }

        internal int GetNextShapeID()
        {
            int id = this.CurrentShapeID;
            this.CurrentShapeID++;
            return id;
        }

        public MasterMetadata GetMasterMetData(int id)
        {
            foreach(var m in this.master_metadata)
            {
                if(m.Value.ID == id)
                {
                    return m.Value;
                }
            }

            throw new System.ArgumentException("no such master id", nameof(id));
        }

        public MasterMetadata GetMasterMetaData(string name)
        {
            return this.master_metadata[name];
        }

        public Face AddFace(string name)
        {
            if(!this.Faces.ContainsName(name))
            {
                var new_face = new Face(this.Faces.Count + 1, name);
                this.Faces.Add(new_face);
                return new_face;
            }
            else
            {
                return this.Faces[name];
            }
        }

        public static string DefaultTemplateXML => Properties.Resources.DefaultVDXTemplate;

        public void Save(string filename)
        {
            string ext = System.IO.Path.GetExtension(filename).ToLower();
            if(ext != ".vdx")
            {
                throw new System.ArgumentException("only .vdx extension is supported", nameof(filename));
            }
            var vdxWriter = new VDXWriter();
            vdxWriter.CreateVDX(this, this.dom, filename);
        }

        public void Save(System.IO.Stream objStream)
        {
            var vdxWriter = new VDXWriter();
            vdxWriter.CreateVDX(this, this.dom, objStream);
        }

        internal void AccountForMasteSubshapes(int n)
        {
            this.CurrentShapeID += n + 1;
        }
    }
}