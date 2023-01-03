
using DemoSolidWorksApi;
using SldWorks;
using System.Runtime.InteropServices;

var progId = "SldWorks.Application";
var progType = System.Type.GetTypeFromProgID(progId);

var app = Marshal2.GetActiveObject(progId) as ISldWorks;

var part = app.NewPart() as PartDoc;

var box = CreateBox(app, 0.2, 0.2);
box.Name = "MyBox";

var cyl = CreateCylinder(app, 1, 1);
cyl.Name = "MyCylinder";

Console.ReadLine();

static IFeature CreateBox(ISldWorks app, double diam, double height)
{
    var part = app.ActiveDoc as IPartDoc;

    var modeler = app.IGetModeler();

    var boxBody = modeler.CreateBodyFromBox(new double[]
    {
                0, 0, 0,
                1, 0, 0,
                1, 1, 1
    }) as Body;
    if (boxBody != null)
    {
        var feat = part.CreateFeatureFromBody3(boxBody, false, 1) as IFeature;
        return feat;
    }
    else
    {
        throw new NullReferenceException("Failed to create body. Make sure that the parameters are valid");
    }
}

static IFeature CreateCylinder(ISldWorks app, double diam, double height)
{
    var part = app.ActiveDoc as IPartDoc;

    var modeler = app.IGetModeler();

    var cylBody = modeler.CreateBodyFromCyl(new double[]
   {
                2, 0, 0,
                0, 1, 0,
                diam / 2, height
   }) as Body;

    if (cylBody != null)
    {
        var feat = part.CreateFeatureFromBody3(cylBody, false, 1) as IFeature;
        return feat;
    }
    else
    {
        throw new NullReferenceException("Failed to create body. Make sure that the parameters are valid");
    }
}