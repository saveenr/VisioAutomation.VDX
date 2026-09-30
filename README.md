# VisioAutomation.VDX

Create simple Visio VDX files without installing or automating Microsoft Visio.

**Status:** feature development remains paused; correctness, tests, and build maintenance are included in the VisioAutomation ownership handoff.

The library targets .NET Framework 4.5.2, matching [VisioAutomation](https://github.com/saveenr/VisioAutomation). It has no Visio or COM runtime dependency. These source changes are unreleased; see [CHANGELOG.md](https://github.com/saveenr/VisioAutomation.VDX/blob/master/CHANGELOG.md) before upgrading from the existing [NuGet package](https://www.nuget.org/packages/VisioAutomation.VDX/).

```csharp
var drawing = new VisioAutomation.VDX.Elements.Drawing(new VisioAutomation.VDX.Template());
var page = new VisioAutomation.VDX.Elements.Page(8, 5);
drawing.Pages.Add(page);
var master = drawing.GetMasterMetaData("Rectangle");
var shape = new VisioAutomation.VDX.Elements.Shape(master.ID, 4, 2.5, 3, 1);
page.Shapes.Add(shape);
shape.Text.Add("Hello, Visio");
drawing.Save("hello.vdx");
```

`drawing.ToXml()` returns an independent `XDocument` for inspection or custom output. Preserve text whitespace when writing it; `Save` handles that automatically. Save may be called repeatedly after editing the model.

- [Build, tests, font tool, and packaging](https://github.com/saveenr/VisioAutomation.VDX/blob/master/BUILDING.md)
- [Architecture, maintenance, and handover boundaries](https://github.com/saveenr/VisioAutomation.VDX/blob/master/ARCHITECTURE.md)
- [MIT license](https://github.com/saveenr/VisioAutomation.VDX/blob/master/LICENSE.txt)
- [VisioAutomation portfolio](https://github.com/saveenr/VisioAutomation/blob/master/docs/RELATED-REPOS.md)

The library and pure tests need no Visio installation. The separate integration suite uses real Visio to check generated documents. VDX generation is distinct from automating Visio through VisioAutomation or VisioPowerShell.
