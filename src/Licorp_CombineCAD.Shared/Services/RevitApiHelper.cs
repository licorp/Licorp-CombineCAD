using System;
using Autodesk.Revit.DB;

namespace Licorp_CombineCAD.Services
{
    /// <summary>
    /// Cross-version Revit API helpers. ElementId property names changed between
    /// Revit 2024 (IntegerValue) and Revit 2025+ (Value), so reflection is used
    /// to support all versions from a single shared assembly.
    /// </summary>
    internal static class RevitApiHelper
    {
        public static string GetElementIdString(ElementId id)
        {
            if (id == null)
                return "";

            try
            {
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                    return Convert.ToInt64(valueProperty.GetValue(id, null)).ToString();

                var integerValueProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (integerValueProperty != null)
                    return Convert.ToInt64(integerValueProperty.GetValue(id, null)).ToString();
            }
            catch { }

            return "";
        }

        public static long GetElementIdLong(ElementId id)
        {
            if (id == null)
                return 0;

            try
            {
                var valueProperty = typeof(ElementId).GetProperty("Value");
                if (valueProperty != null)
                    return Convert.ToInt64(valueProperty.GetValue(id, null));

                var integerValueProperty = typeof(ElementId).GetProperty("IntegerValue");
                if (integerValueProperty != null)
                    return Convert.ToInt64(integerValueProperty.GetValue(id, null));
            }
            catch { }

            return 0;
        }

        public static ElementId CreateElementId(long value)
        {
            try
            {
                var longCtor = typeof(ElementId).GetConstructor(new[] { typeof(long) });
                if (longCtor != null)
                    return (ElementId)longCtor.Invoke(new object[] { value });

                var intCtor = typeof(ElementId).GetConstructor(new[] { typeof(int) });
                if (intCtor != null)
                    return (ElementId)intCtor.Invoke(new object[] { Convert.ToInt32(value) });
            }
            catch { }

            return null;
        }
    }
}
