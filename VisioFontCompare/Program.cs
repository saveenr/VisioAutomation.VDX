using System;
using System.Collections.Generic;
using System.Linq;

using VDX=VisioAutomation.VDX;

namespace VisioFontCompare
{
    class Program
    {
        static void Main(string[] args)
        {
            string left_font = "Calibri";

            var template = new VDX.Template();
            var doc = new VDX.Elements.Drawing(template);
            var page = new VDX.Elements.Page(10,10);
            doc.Pages.Add(page);
            var faces = doc.Faces;

            var left_face = Helpers.GetFontSafe(doc.Faces, left_font);

            double face_size = 32.0;
            int rect_id = doc.GetMasterMetaData("REctAngle").ID;

            string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string letters2 = "abcdefghijklmnopqrstuvwxyz";
            string numbers = "01234567890";
            string punctuation = "!@#$%^&*()_+-=<>,.[]{}\\|``/?;:'\"";

            var text = letters + letters2 + numbers + punctuation;

            int count = text.Length;
            int cols = (int) System.Math.Ceiling(System.Math.Sqrt(count));
            double face_size_inches = face_size / 72.0;

            double box_size = face_size_inches*2.0;
            double box_sep = box_size/8.0;

            double grid_delta = box_size + box_sep;

            page.PageProperties.PageWidth.Formula = (cols*grid_delta).ToString();
            page.PageProperties.PageHeight.Formula = (cols * grid_delta).ToString();

            int n = 0;
            foreach (int row in Enumerable.Range(0, cols))
            {
                foreach (int col in Enumerable.Range(0,cols))
                {
                    if (n >= count)
                    {
                        break;
                    }


                    var pinx = col* grid_delta;
                    var piny = (cols*grid_delta) - (row * grid_delta);

                    var shape = new VDX.Elements.Shape(rect_id,pinx, piny, box_size, box_size);
                    page.Shapes.Add(shape);

                    string s = text[n].ToString();
                    shape.Text.Add(s);

                    shape.CharFormats = new List<VDX.Sections.Char>();

                    var charfmt = new VDX.Sections.Char();
                    charfmt.Font.Formula = left_face.ID.ToString();
                    charfmt.Size.Formula = face_size.ToString()+"pt"; 

                    shape.CharFormats.Add(charfmt);

                    shape.Fill = new VDX.Sections.Fill();
                    shape.Fill.Pattern.Formula = "0";
                    shape.Fill.ForegroundColor.Formula = "rgb(255,255,255)";
                    shape.Fill.ForegroundTransparency.Formula = "1.0";

                    shape.Line  = new VDX.Sections.Line();
                    shape.Line.Pattern.Formula = "0";
                    shape.Line.Weight.Formula = "0";
                    shape.Line.Color.Formula = "rgb(240,240,240)";
                    n++;
                }

            }

            doc.Save("D:\\specimen.vdx");
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
