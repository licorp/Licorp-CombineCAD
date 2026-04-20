using System;
using System.Diagnostics;
using System.Reflection;

namespace Licorp_CombineCAD.Helpers
{
    /// <summary>
    /// Reflection-based property setter for cross-version Revit API compatibility.
    /// Extracted from Export+ DWGExportManager.TrySetProperty pattern.
    /// </summary>
    public static class ReflectionHelper
    {
        /// <summary>
        /// Try to set a property using reflection (works across Revit API versions
        /// where properties may not exist in older versions).
        /// </summary>
        public static bool TrySetProperty(object target, string propertyName, object value)
        {
            try
            {
                var property = target.GetType().GetProperty(propertyName,
                    BindingFlags.Public | BindingFlags.Instance);

                if (property != null && property.CanWrite)
                {
                    property.SetValue(target, value);
                    Debug.WriteLine($"[CombineCAD] Set {propertyName} = {value}");
                    return true;
                }
                else
                {
                    Debug.WriteLine($"[CombineCAD] Property not found or read-only: {propertyName}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CombineCAD] Failed to set {propertyName}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Try to get an enum value from the Revit API assembly by type name and value name.
        /// Used for enums that may not exist in all Revit versions.
        /// </summary>
        public static object TryGetEnumValue(string enumTypeName, string valueName)
        {
            try
            {
                var revitAssembly = typeof(Autodesk.Revit.DB.DWGExportOptions).Assembly;
                var enumType = revitAssembly.GetType($"Autodesk.Revit.DB.{enumTypeName}");

                if (enumType != null && enumType.IsEnum)
                {
                    return Enum.Parse(enumType, valueName);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[CombineCAD] Enum lookup failed: {enumTypeName}.{valueName}: {ex.Message}");
            }

            return null;
        }
    }
}
