using SXL = System.Xml.Linq;

namespace VisioAutomation.VDX.Internal;

public static class XMLUtil
{
    public static SXL.XElement CreateVisioSchema2003Element(string name)
        => new(Constants.VisioXmlNamespace2003 + name);

    public static SXL.XElement CreateVisioSchema2006Element(string name)
        => new(Constants.VisioXmlNamespace2006 + name);

}
