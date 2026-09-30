using System;
using System.Collections.Generic;
using System.Linq;

using VDX=VisioAutomation.VDX;

namespace VisioFontCompare
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length > 2)
            {
                Console.Error.WriteLine("Usage: VisioFontCompare [output.vdx] [font-name]");
                return 1;
            }
            string output = System.IO.Path.GetFullPath(args.Length > 0 ? args[0] : "specimen.vdx");
            string left_font = args.Length > 1 ? args[1] : "Calibri";

            var template = new VDX.Template();
            var doc = new VDX.Elements.Drawing(template);
            var page = new VDX.Elements.Page(10,10);
            doc.Pages.Add(page);

            var left_face = doc.AddFace(left_font);

            double face_size = 32.0;
            int rect_id = doc.GetMasterMetaData("REctAngle").ID;

            string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string letters2 = "abcdefghijklmnopqrstuvwxyz";
            string numbers = "01234567890";
            string punctuation = "!@#$%^&*()_+-=<>,.[]{}\\|``/?;:'\"";

            var text = letters + letters2 + numbers + punctuation;

            int count = text.Length;
            int cols = (int) System.Math.Ceiling(System.Math.Sqrt(count));
            double face_size_inches = VDX.ShapeSheet.Converter.PointsToInches(face_size);

            double box_size = face_size_inches*2.0;
            double box_sep = box_size/8.0;

            double grid_delta = box_size + box_sep;

            page.PageProperties.PageWidth.Result = cols * grid_delta;
            page.PageProperties.PageHeight.Result = cols * grid_delta;

            int n = 0;
            foreach (int row in Enumerable.Range(0, cols))
            {
                foreach (int col in Enumerable.Range(0,cols))
                {
                    if (n >= count)
                    {
                        break;
                    }


                    var pinx = (col + 0.5) * grid_delta;
                    var piny = (cols - row - 0.5) * grid_delta;

                    var shape = new VDX.Elements.Shape(rect_id,pinx, piny, box_size, box_size);
                    page.Shapes.Add(shape);

                    string s = text[n].ToString();
                    shape.Text.Add(s);

                    shape.CharFormats = new List<VDX.Sections.Char>();

                    var charfmt = new VDX.Sections.Char();
                    charfmt.Font.Result = left_face.ID;
                    charfmt.Size.Result = face_size;

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

            doc.Save(output);
            Console.WriteLine(output);
            return 0;
        }
        
    }

}
