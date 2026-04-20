using System.Collections.Generic;

namespace Licorp_CombineCAD.Models
{
    /// <summary>
    /// Layer mapping configuration for DWG export.
    /// Stores layer mappings from category/subcategory to layer name, color, and linetype.
    /// </summary>
    public class LayerMapping
    {
        public string SetupName { get; set; }
        public List<LayerMappingEntry> Entries { get; set; } = new List<LayerMappingEntry>();
    }

    public class LayerMappingEntry
    {
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string LayerName { get; set; }
        public int ColorId { get; set; }
        public string Linetype { get; set; }

        public static LayerMappingEntry Parse(string line)
        {
            var parts = line.Split('\t');
            if (parts.Length < 5) return null;

            return new LayerMappingEntry
            {
                Category = parts[0].Trim(),
                SubCategory = parts[1].Trim(),
                LayerName = parts[2].Trim(),
                ColorId = int.TryParse(parts[3].Trim(), out var c) ? c : 0,
                Linetype = parts[4].Trim()
            };
        }

        public override string ToString()
        {
            return $"{Category}\t{SubCategory}\t{LayerName}\t{ColorId}\t{Linetype}";
        }
    }
}