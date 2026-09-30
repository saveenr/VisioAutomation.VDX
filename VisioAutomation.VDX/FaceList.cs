using System.Linq;

namespace VisioAutomation.VDX
{
    public class FaceList : NamedNodeList<Elements.Face>
    {
        public FaceList() :
            base(face => face.Name)
        {

        }

        public override void Add(Elements.Face face)
        {
            this.ValidateItem(face);
            if (this.Items.Any(existing => existing.ID == face.ID))
            {
                throw new System.ArgumentException("Already contains a face with that ID", nameof(face));
            }
            base.Add(face);
        }
    }
}
