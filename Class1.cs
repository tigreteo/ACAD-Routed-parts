using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

//Build a 3D solid in AutoCAD from a 2D drawing for the router

//User selects block referances containg drawings to convert
//for each valid blockref (CNC borders):
// 1)try to read tag info to find board thickness, orientation, degree of deviation for orient
//   *should that info not be available, request from user
// 2)select the polylines,circles from inside of the block
//      create solids from the various selections based on layer
//      delete solids from the parts created on through cut layer
//      delete solid holes from main part?? if polyline is inside polyline?

//parts should be created on top of current 2dDrawings

//should the part fail to create (poly line not closed, made wrong)
// then we need flag the part and move on to next piece

namespace RoutedParts

{
    public class Class1
    {
        [CommandMethod("ConvertToCNCPart")]
        public static void FindInterferenceBetweenSolids()
        {
            // Get the current document and database, and start a transaction
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                //select blockrefs - cancel if they opt out of choosing
                ObjectIdCollection blocksToProcess = chooseParts(ed, tr);
                if (blocksToProcess == null || blocksToProcess.Count == 0)
                { return; }                                
                
                //collect the polylines, circles
                foreach (ObjectId borderId in blocksToProcess)
                {   
                    BlockReference blkRef = tr.GetObject(borderId, OpenMode.ForRead) as BlockReference;
                    try
                    {
                        ObjectIdCollection twoDDrawing = routePaths(blkRef, ed, borderId);

                        //get board thickness from block
                        double thickness = getBoardThickness(blkRef, tr, ed);

                        if (twoDDrawing == null || twoDDrawing.Count == 0)
                        { return; }

                        //collect all of the solids into a collection. subtract the routes away from the main part
                        //main part should be the largest volume of all the pieces..at least you would think so
                        ObjectIdCollection routedSolids = new ObjectIdCollection();
                        ObjectId solidID = ObjectId.Null;

                        //extrude the entities based on layer
                        foreach (ObjectId id in twoDDrawing)
                        {
                            solidID = extrudePart(id, tr, btr, db, ed, thickness);
                            if (solidID != ObjectId.Null)
                            { routedSolids.Add(solidID); }
                        }

                        //subtract routes from main part
                        cutOutParts(tr, routedSolids);
                    }
                    catch(Autodesk.AutoCAD.Runtime.Exception e)
                    {
                        ed.WriteMessage("\nPart failed to be extruded. \n Check for open polylines");
                        Entity entCurrent = tr.GetObject(borderId, OpenMode.ForRead) as Entity;
                        Extents3d target = entCurrent.GeometricExtents;
                        xOUt(tr, bt, target);
                    }
                }
                // Save the new objects to the database
                tr.Commit();
            }
        }

        //select blocks to process
        private static ObjectIdCollection chooseParts(Editor ed, Transaction tr)
        {
            // Create a TypedValue array to define the filter criteria
            TypedValue[] filterList = new TypedValue[1];
            filterList.SetValue(new TypedValue(0, "INSERT"), 0);
            SelectionFilter setFilter = new SelectionFilter(filterList);
            //get selection of blocks to use
            SelectionSet acSSet;
            PromptSelectionResult selection = ed.GetSelection(setFilter);
            if (selection.Status == PromptStatus.OK)
            { acSSet = selection.Value; }
            else
            { return null; }

            //collection to hold valid Blkrefs
            ObjectIdCollection blocksToProcess = new ObjectIdCollection();
            //loop through blockrefs
            foreach (ObjectId msId in acSSet.GetObjectIds())
            {
                //validate this is a block we care about
                BlockReference blkRef = tr.GetObject(msId, OpenMode.ForRead) as BlockReference;
                if(blkRef.Name.Contains("CNC_BORDER"))
                { blocksToProcess.Add(msId); }
            }

            return blocksToProcess;
        }

        //collect entities that are used for routes (polylines, circles)
        private static ObjectIdCollection routePaths(BlockReference blkRef, Editor ed, ObjectId cncBorder)
        {
            ObjectIdCollection partCol = new ObjectIdCollection();
            //Extents3d window = (Extents3d)blkRef.Bounds;
            Extents3d window = blkRef.GeometricExtents;

            SelectionSet select;
            PromptSelectionResult res = ed.SelectCrossingWindow(window.MinPoint, window.MaxPoint);
            if(res.Status == PromptStatus.OK)
                select = res.Value;
            else
                return null;

            ObjectIdCollection objIdCol = new ObjectIdCollection(select.GetObjectIds());
            //remove CNC frame from that collection
            objIdCol.Remove(cncBorder);

            foreach (ObjectId msId in objIdCol)
            {
                if (msId.ObjectClass.DxfName.ToUpper() == "POLYLINE" ||
                    msId.ObjectClass.DxfName.ToUpper() == "LWPOLYLINE" ||
                    msId.ObjectClass.DxfName.ToUpper() == "CIRCLE")
                { partCol.Add(msId); }
            }

            return objIdCol;
        }

        //clones a collection of objects into a new collection
        private static ObjectIdCollection cloneCollection(ObjectIdCollection partCol, Transaction tr, BlockTableRecord btr)
        {
            ObjectIdCollection clonedCol = new ObjectIdCollection();

            foreach (ObjectId id in partCol)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                Entity copyEnt = ent.Clone() as Entity;
                //add cloned object to the BTR and DB
                btr.AppendEntity(copyEnt);
                tr.AddNewlyCreatedDBObject(copyEnt, true);
                clonedCol.Add(copyEnt.Id);
            }

            return clonedCol;
        }

        //determine if part is a circle or polyline extrude accordingly
        private static ObjectId extrudePart(ObjectId id, Transaction tr, BlockTableRecord btr, Database db, Editor ed,
            double BoardThickness)
        {
            ObjectId solidID = ObjectId.Null;
            string objecttype = id.ObjectClass.DxfName.ToUpper();
            if(objecttype == "LWPOLYLINE")
            { solidID = extrudePline(id, tr, btr, db, ed, BoardThickness);}
            else if(objecttype == "CIRCLE")
            { solidID = makeCircles(id, tr, btr, db, BoardThickness); }

            return solidID;
        }         

        //makes a 3d cylinder
        private static ObjectId makeCircles(ObjectId id, Transaction tr, BlockTableRecord btr,Database db, double depth)
        {
            Circle cir = tr.GetObject(id, OpenMode.ForWrite) as Circle;
            if(cir.Layer.ToString().Contains("Drill") == false)
            { depth = layerThickness(id, tr, depth); }
            double radius = cir.Radius;
            //Create a solid cylinder
            using(Solid3d acSolid = new Solid3d())
            {
                acSolid.RecordHistory = true;
                acSolid.CreateFrustum(depth , radius, radius, radius);
                acSolid.Color = cir.Color;

                //move to appropriate place
                //NEED TO MOVE TO FIX z AXIS
                Point3d moveTo = new Point3d(cir.Center.X, cir.Center.Y, (depth * -.5));
                acSolid.TransformBy(Matrix3d.Displacement(moveTo - Point3d.Origin));

                ObjectId savedExtrusionId = ObjectId.Null;

                ObjectId modelId;
                modelId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                savedExtrusionId = btr.AppendEntity(acSolid);
                tr.AddNewlyCreatedDBObject(acSolid, true);

                return acSolid.Id;
            }
        }

        //makes a 3d polyline
        private static ObjectId extrudePline(ObjectId id, Transaction tr, BlockTableRecord btr, Database db, Editor ed,
            double depth)
        {
            //find depth of the layer
            double layerdepth = layerThickness(id, tr, depth) * -1;
            ObjectId savedExtrusionId = ObjectId.Null;
            //if it isnt a designated layer then we'll skip over
            if (layerdepth != 0)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                Polyline pPoly = ent as Polyline;                

                if (pPoly == null) return savedExtrusionId;
                DBObjectCollection lines = new DBObjectCollection();
                pPoly.Explode(lines);

                // Create a region from the set of lines.
                DBObjectCollection regions = new DBObjectCollection();
                regions = Region.CreateFromCurves(lines);

                if (regions.Count == 0)
                {
                    ed.WriteMessage("\nFailed to create region\n");
                    return savedExtrusionId;
                }

                Region pRegion = (Region)regions[0];
                // Extrude the region to create a solid.

                Solid3d pSolid = new Solid3d();
                pSolid.RecordHistory = true;
                pSolid.Extrude(pRegion, layerdepth, 0.0);
                

                ObjectId modelId;
                modelId = SymbolUtilityServices.GetBlockModelSpaceId(db);
                savedExtrusionId = btr.AppendEntity(pSolid);
                tr.AddNewlyCreatedDBObject(pSolid, true);

                return pSolid.Id;
            }
            return savedExtrusionId;
        }

        private static double layerThickness(ObjectId id, Transaction tr, double boardthickness)
        {
            Entity acObject = tr.GetObject(id, OpenMode.ForRead) as Entity;
            switch(acObject.Layer.ToString().ToUpper())
            {
                case "1-8 ROUTE":
                case "EIGHTH ROUTE":
                case "POCKET CUT 1-8":
                    return .125;
                case "1-4 ROUTE":
                case "QUARTER ROUTE":
                case "POCKET CUT 1-4":
                case "1-4 ROUTER 1-4 DEPTH":
                    return .25;
                case "3-8 ROUTE":
                case "POCKET CUT 3-8":
                    return .375;
                case "7-16 ROUTE":
                case "POCKET CUT 7-16":
                    return .4375;
                case "FULL DEPTH":
                    return boardthickness;
                default:
                    return 0;
            }
        }

        //scan through attributes and look tags if they exist
        //grab the board thickness (only on cnc border v2 and later
        private static double getBoardThickness(BlockReference blkRef, Transaction tr, Editor ed)
        {
            AttributeCollection atts = blkRef.AttributeCollection;
            double quantity = 0;

            PromptDoubleOptions pdo = new PromptDoubleOptions("\nPlease enter valid thickness:");
            pdo.AllowNone = false;
            pdo.AllowNegative = false;
            pdo.AllowZero = false;
            pdo.DefaultValue = .75;

            //iterate throught he attributes to find the tag if it exists
            foreach(ObjectId objId in atts)
            {
                DBObject dbObj = tr.GetObject(objId, OpenMode.ForRead) as DBObject;
                AttributeReference attRef = dbObj as AttributeReference;
                if (attRef.Tag.Contains("THICK"))
                {
                    var holder = attRef.TextString;
                    try
                    { quantity = Convert.ToDouble(holder);}
                    catch
                    { ed.WriteMessage("\nCheck if thickness is valid number"); }
                }
            }

            if(quantity == 0)
            {
                ed.WriteMessage("\nNo board thickness specified");
                PromptDoubleResult pdr = ed.GetDouble(pdo);
                if(pdr.Status == PromptStatus.OK)
                { quantity = pdr.Value; }
            }
            return quantity;
        }

        //loop through collection of solids to find the largest volume. subtract all other parts from that piece
        private static void cutOutParts(Transaction tr,ObjectIdCollection routedSolids)
        {
            double volume = 0;
            ObjectId basePart = ObjectId.Null;
            //find largest part by volume.
            foreach(ObjectId id in routedSolids)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                Solid3d part = ent as Solid3d;

                if(part.MassProperties.Volume > volume)
                {
                    volume = part.MassProperties.Volume;
                    basePart = id;
                }
            }

            //removed the base part from the list before boolean subtracting all the parts of the list from it
            routedSolids.Remove(basePart);
            Entity largeEnt = tr.GetObject(basePart, OpenMode.ForWrite) as Entity;
            Solid3d baseSolid = largeEnt as Solid3d;

            foreach(ObjectId id in routedSolids)
            {
                Entity acEnt = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                Solid3d cuttingPart = acEnt as Solid3d;

                if(baseSolid.CheckInterference(cuttingPart) == true)
                { baseSolid.BooleanOperation(BooleanOperationType.BoolSubtract, cuttingPart); }
            }
        }

        private static void xOUt(Transaction tr, BlockTable bt, Extents3d target)
        {
            Line line1 = new Line(new Point3d(target.MinPoint.X,target.MaxPoint.Y,0),
                new Point3d(target.MaxPoint.X, target.MinPoint.Y, 0));
            Line line2 = new Line(new Point3d(target.MinPoint.X, target.MinPoint.Y, 0),
                new Point3d(target.MaxPoint.X, target.MaxPoint.Y, 0));

            BlockTableRecord btr = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

            btr.AppendEntity(line1);
            tr.AddNewlyCreatedDBObject(line1, true);
            btr.AppendEntity(line2);
            tr.AddNewlyCreatedDBObject(line2, true);
        }
    }
}
