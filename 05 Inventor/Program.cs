using DemoInventorApi;
using Inventor;
using System.Net;

var app = (Application)Marshal2.GetActiveObject("Inventor.Application");
var path = "your_inventor_file_path";
loadExitingPart(app, path);

createNewPart(app);

Console.ReadLine();

static void loadExitingPart(Application app, string path)
{
    var doc = app.Documents.Open(path);
    System.Threading.Thread.Sleep(200);
    app.CommandManager.ControlDefinitions["AppIsometricViewCmd"].Execute();
    app.ActiveView.DisplayMode = DisplayModeEnum.kTechnicalIllustrationRendering;

    doc.SaveAs("C:\\temp\\thumbnail.png", true);
}

static void createNewPart(Application app)
{
    var doc = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject);
    
    var trans = app.TransactionManager.StartTransaction((_Document)doc, "My Command");
    drawCube(app, (PartDocument)trans.Document, 5, 5, 5);
    Thread.Sleep(200);
    app.CommandManager.ControlDefinitions["AppIsometricViewCmd"].Execute();
    trans.End();

    app.ActiveView.DisplayMode = DisplayModeEnum.kTechnicalIllustrationRendering;
}

static void drawCube(Application app, PartDocument doc, int x, int y, int z)
{
    var comp = doc.ComponentDefinition;

    var sketch = comp.Sketches.Add(comp.WorkPlanes[3]);

    var geo = app.TransientGeometry;
    var lines = sketch.SketchLines.AddAsTwoPointRectangle(geo.CreatePoint2d(0, 0), geo.CreatePoint2d(3, 3));

    var profile = sketch.Profiles.AddForSolid();

    var extrudeDef = comp.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, PartFeatureOperationEnum.kJoinOperation);
    extrudeDef.SetDistanceExtent(3, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection);

    var extrude = comp.Features.ExtrudeFeatures.Add(extrudeDef);

    WorkAxis XAxis = comp.WorkAxes[1];
    WorkAxis YAxis = comp.WorkAxes[2];
    WorkAxis ZAxis = comp.WorkAxes[3];

    ObjectCollection objCol = app.TransientObjects.CreateObjectCollection();
    objCol.Add(extrude);

    var pattern = comp.Features.RectangularPatternFeatures.Add(objCol, XAxis, true, 5, 4, YDirectionEntity: YAxis, YCount: 5, YSpacing: 4);

    ObjectCollection objCol2 = app.TransientObjects.CreateObjectCollection();
    objCol2.Add(extrude);
    objCol2.Add(pattern);

    var pattern2 = comp.Features.RectangularPatternFeatures.Add(objCol2, ZAxis, true, 5, 4);
}