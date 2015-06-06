using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VDX=VisioAutomation.VDX;

namespace VisioFontCompare
{
    class Program
    {
        static void Main(string[] args)
        {
            string left_font = "Tahoma";
            string right_font = "Arial";

            var template = new VDX.Template();
            var doc = new VDX.Elements.Drawing(template);

            var faces = doc.Faces;

            var left_face = Helpers.GetFontSafe(doc.Faces, left_font);
            var right_face = Helpers.GetFontSafe(doc.Faces, left_font);

        }

    }

    public static class Helpers
    {
        public static VDX.Elements.Face GetFontSafe(VDX.FaceList faces,string name)
        {
            if (faces.ContainsName(name))
            {
                return faces[name];
            }

            int max_id = faces.Items.Select(f => f.ID).Max();
            int new_id = max_id +1;

            var face = new VDX.Elements.Face(new_id,name);

            faces.Add(face);

            return face;
        }
    }


}
