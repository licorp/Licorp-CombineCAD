using System;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace Licorp_MergeSheets
{
    public static class ExtentsHelper
    {
        public static Extents3d GetBlockExtents(BlockTableRecord btr, Transaction trans)
        {
            double minX = double.MaxValue, minY = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue;

            foreach (ObjectId id in btr)
            {
                try
                {
                    var ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null)
                    {
                        var ext = ent.GeometricExtents;
                        if (ext.MinPoint.X < minX) minX = ext.MinPoint.X;
                        if (ext.MinPoint.Y < minY) minY = ext.MinPoint.Y;
                        if (ext.MaxPoint.X > maxX) maxX = ext.MaxPoint.X;
                        if (ext.MaxPoint.Y > maxY) maxY = ext.MaxPoint.Y;
                    }
                }
                catch { }
            }

            if (minX == double.MaxValue)
                return new Extents3d(new Point3d(0, 0, 0), new Point3d(0, 0, 0));

            return new Extents3d(new Point3d(minX, minY, 0), new Point3d(maxX, maxY, 0));
        }
    }
}
