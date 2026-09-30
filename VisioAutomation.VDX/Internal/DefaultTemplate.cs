using System;
using System.IO;

namespace VisioAutomation.VDX.Internal;

internal static class DefaultTemplate
{
    private static readonly Lazy<string> xml = new(Read);

    internal static string Xml => xml.Value;

    private static string Read()
    {
        using var stream = typeof(DefaultTemplate).Assembly.GetManifestResourceStream("VisioAutomation.VDX.DefaultTemplate.xml")
            ?? throw new InvalidOperationException("The embedded VDX template is missing.");
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}
