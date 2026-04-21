using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.DB;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Manages tag types with dynamic shelf length.
    /// Creates new tag types with appropriate shelf length based on text width.
    /// </summary>
    public class TagTypeManager
    {
        private readonly Document _document;
        
        // Shelf length step (in mm)
        private const double SHELF_LENGTH_STEP_MM = 2.0;
        
        // Maximum shelf length to prevent unreasonable values from 3D view bounding boxes
        private const double MAX_SHELF_LENGTH_MM = 50.0;
        
        // Default shelf length to use when calculation is unreliable (e.g., in 3D views)
        private const double DEFAULT_SHELF_LENGTH_MM = 10.0;
        
        // Parameter name for shelf length
        private const string SHELF_LENGTH_PARAM_NAME = "Длина полки";
        
        // Default tag types for each element category
        public static readonly Dictionary<string, (string FamilyName, string TypeName)> DefaultTagTypes = new Dictionary<string, (string, string)>
        {
            // Round ducts
            { "DuctRound", ("ADSK_Марка_Воздуховодов_ОсновныеОбозначения", "Круглый воздуховод_Размер и расход (2)") },
            // Rectangular ducts
            { "DuctRect", ("ADSK_Марка_Воздуховодов_ОсновныеОбозначения", "Прямоугольный воздуховод_Размер и расход (2)") },
            // Air terminals (diffusers)
            { "AirTerminal", ("ADSK_M_Воздухораспределители", "Имя типа / Расход (25)") },
            // Air terminals - short name
            { "AirTerminalShort", ("ADSK_M_Воздухораспределители", "ADSK_Наименование краткое / Расход (2)") },
            // Duct accessories
            { "DuctAccessory", ("ADSK_M_Арматура воздуховодов", "ADSK_Марка (2)") },
            // Equipment
            { "Equipment", ("ADSK_M_Оборудование", "ADSK_Марка (2)") },
            // Spot dimensions (elevation marks)
            { "SpotDimension", ("Сист. семейство: Высотные отметки", "ADSK_Стрелка_Относительная_Вниз") },
        };
        
        public TagTypeManager(Document document)
        {
            _document = document;
        }
        
        /// <summary>
        /// Get or create a tag type with the appropriate shelf length for the given text width.
        /// IMPORTANT: This method must be called from within an existing transaction.
        /// </summary>
        /// <param name="baseTagType">The base tag type to use as template</param>
        /// <param name="textWidthInViewUnits">The width of the text in view units (from bounding box)</param>
        /// <param name="viewScale">The view scale (e.g., 100 for 1:100)</param>
        /// <returns>A tag type with appropriate shelf length</returns>
        public FamilySymbol GetOrCreateTagTypeWithShelfLength(FamilySymbol baseTagType, double textWidthInViewUnits, int viewScale)
        {
            if (baseTagType == null)
                return null;
            
            // The textWidthInViewUnits is in Revit internal units (feet)
            // The shelf length parameter in the family is in mm
            // We need to convert and calculate the required shelf length
            
            // The textWidthInViewUnits is in Revit internal units (feet)
            // For 2D views: bounding box returns model-space coordinates, so we need to divide by viewScale
            // to get paper-space size. For 3D views: the width was estimated in mm then converted to feet
            // with viewScale included, so we also need to divide by viewScale for consistency.
            // Formula: textWidthMm = textWidthInViewUnits * 304.8 / viewScale
            // This gives the paper-space width in mm regardless of view type.
            
            double effectiveScale = viewScale > 0 ? viewScale : 1;
            double textWidthMm = textWidthInViewUnits * 304.8 / effectiveScale;
            
            // VALIDATION: Check if the calculated text width is reasonable
            // In 3D views, bounding boxes can be in model coordinates and give huge values
            // A typical annotation text is 5-50mm wide, anything over 100mm is suspicious
            if (textWidthMm > MAX_SHELF_LENGTH_MM)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] WARNING: Calculated text width {textWidthMm:F1}mm exceeds maximum {MAX_SHELF_LENGTH_MM}mm. " +
                    $"This likely indicates a 3D view bounding box issue. Using default shelf length {DEFAULT_SHELF_LENGTH_MM}mm.");
                
                // Use the default shelf length instead of creating an unreasonably large one
                // Check if current type already has a reasonable shelf length
                double? currentShelfLength = GetShelfLengthFromTypeName(baseTagType.Name);
                if (currentShelfLength.HasValue && currentShelfLength.Value <= MAX_SHELF_LENGTH_MM)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Using existing type '{baseTagType.Name}' with shelf length {currentShelfLength.Value:F0}mm");
                    return baseTagType;
                }
                
                // Try to find an existing type with default shelf length
                string baseTypeName = GetBaseTypeName(baseTagType.Name);
                string familyName = baseTagType.FamilyName;
                var existingType = FindExistingTypeWithShelfLength(familyName, baseTypeName, DEFAULT_SHELF_LENGTH_MM);
                if (existingType != null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Found existing type '{existingType.Name}' with default shelf length {DEFAULT_SHELF_LENGTH_MM}mm");
                    return existingType;
                }
                
                // Create new type with default shelf length
                var defaultType = CreateNewTypeWithShelfLength(baseTagType, DEFAULT_SHELF_LENGTH_MM);
                if (defaultType != null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Created new type '{defaultType.Name}' with default shelf length {DEFAULT_SHELF_LENGTH_MM}mm");
                }
                return defaultType ?? baseTagType;
            }
            
            // For annotations, the view scale doesn't affect the tag size
            // because tags are annotation elements that always display at the same paper size
            // The shelf length should accommodate the text width with minimal padding
            
            // Add minimal padding (1mm total for margin)
            double requiredShelfLengthMm = textWidthMm + 1.0;
            
            // Round up to next step (2mm increments)
            requiredShelfLengthMm = Math.Ceiling(requiredShelfLengthMm / SHELF_LENGTH_STEP_MM) * SHELF_LENGTH_STEP_MM;
            
            // Minimum shelf length
            requiredShelfLengthMm = Math.Max(requiredShelfLengthMm, SHELF_LENGTH_STEP_MM);
            
            // Cap maximum shelf length
            requiredShelfLengthMm = Math.Min(requiredShelfLengthMm, MAX_SHELF_LENGTH_MM);
            
            DebugLogger.Log($"[TAG-TYPE-MANAGER] Text width: {textWidthMm:F1}mm, required shelf: {requiredShelfLengthMm:F0}mm");
            
            // Get base type name pattern (without the shelf length in parentheses)
            string baseTypeName2 = GetBaseTypeName(baseTagType.Name);
            string familyName2 = baseTagType.FamilyName;
            
            // Check if current type already has the right shelf length
            double? currentShelfLength2 = GetShelfLengthFromTypeName(baseTagType.Name);
            if (currentShelfLength2.HasValue && Math.Abs(currentShelfLength2.Value - requiredShelfLengthMm) < 0.5)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Current type '{baseTagType.Name}' already has correct shelf length: {currentShelfLength2.Value:F0}mm");
                return baseTagType;
            }
            
            // Look for existing type with the required shelf length
            var existingType2 = FindExistingTypeWithShelfLength(familyName2, baseTypeName2, requiredShelfLengthMm);
            if (existingType2 != null)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Found existing type '{existingType2.Name}' with shelf length {requiredShelfLengthMm:F0}mm");
                return existingType2;
            }
            
            // Create new type with the required shelf length
            // IMPORTANT: We're inside an existing transaction, so we just do the work directly
            var newType = CreateNewTypeWithShelfLength(baseTagType, requiredShelfLengthMm);
            if (newType != null)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Created new type '{newType.Name}' with shelf length {requiredShelfLengthMm:F0}mm");
            }
            
            return newType ?? baseTagType;
        }
        
        /// <summary>
        /// Get shelf length from type name (extracted from parentheses).
        /// E.g., "Круглый воздуховод_Размер и расход (14)" -> 14.0
        /// Public version for use by other classes.
        /// </summary>
        public static double? GetShelfLengthFromTypeNamePublic(string typeName)
        {
            return GetShelfLengthFromTypeName(typeName);
        }
        
        /// <summary>
        /// Get shelf length from type name (extracted from parentheses).
        /// E.g., "Круглый воздуховод_Размер и расход (14)" -> 14.0
        /// </summary>
        private static double? GetShelfLengthFromTypeName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return null;
            
            // Look for pattern "(number)" at the end
            var match = Regex.Match(typeName, @"\((\d+(?:\.\d+)?)\)\s*$");
            if (match.Success && double.TryParse(match.Groups[1].Value, out double shelfLength))
            {
                return shelfLength;
            }
            
            return null;
        }
        
        /// <summary>
        /// Get base type name without the shelf length in parentheses.
        /// E.g., "Круглый воздуховод_Размер и расход (14)" -> "Круглый воздуховод_Размер и расход"
        /// </summary>
        public static string GetBaseTypeName(string typeName)
        {
            if (string.IsNullOrEmpty(typeName))
                return typeName;
            
            // Remove trailing "(number)" pattern
            return Regex.Replace(typeName.Trim(), @"\s*\(\d+(?:\.\d+)?\)\s*$", "");
        }
        
        /// <summary>
        /// Get the shelf length parameter value from a tag type.
        /// </summary>
        private double? GetShelfLengthParameter(FamilySymbol tagType)
        {
            try
            {
                // Look for the shelf length parameter
                var param = tagType.LookupParameter(SHELF_LENGTH_PARAM_NAME);
                if (param != null && param.StorageType == StorageType.Double)
                {
                    return param.AsDouble();
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Error getting shelf length parameter: {ex.Message}");
            }
            
            return null;
        }
        
        /// <summary>
        /// Find an existing tag type with the specified shelf length in the type name.
        /// </summary>
        private FamilySymbol FindExistingTypeWithShelfLength(string familyName, string baseTypeName, double shelfLengthMm)
        {
            try
            {
                // Construct the expected type name
                string expectedTypeName = $"{baseTypeName} ({(int)shelfLengthMm})";
                
                // Get all symbols in the family
                var collector = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .Where(s => s.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase));
                
                foreach (var symbol in collector)
                {
                    // Check if the type name matches exactly
                    if (symbol.Name.Equals(expectedTypeName, StringComparison.OrdinalIgnoreCase))
                    {
                        return symbol;
                    }
                    
                    // Also check if the base name matches and shelf length is correct
                    string symbolBaseName = GetBaseTypeName(symbol.Name);
                    if (symbolBaseName.Equals(baseTypeName, StringComparison.OrdinalIgnoreCase))
                    {
                        double? symbolShelf = GetShelfLengthFromTypeName(symbol.Name);
                        if (symbolShelf.HasValue && Math.Abs(symbolShelf.Value - shelfLengthMm) < 0.5)
                        {
                            return symbol;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Error finding existing type: {ex.Message}");
            }
            
            return null;
        }
        
        /// <summary>
        /// Create a new tag type with the specified shelf length.
        /// IMPORTANT: Must be called within an existing transaction.
        /// </summary>
        private FamilySymbol CreateNewTypeWithShelfLength(FamilySymbol baseTagType, double shelfLengthMm)
        {
            try
            {
                // Get base type name
                string baseTypeName = GetBaseTypeName(baseTagType.Name);
                string newTypeName = $"{baseTypeName} ({(int)shelfLengthMm})";
                
                // Check if type with this name already exists
                var existingType = new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.FamilyName.Equals(baseTagType.FamilyName, StringComparison.OrdinalIgnoreCase) &&
                                         s.Name.Equals(newTypeName, StringComparison.OrdinalIgnoreCase));
                
                if (existingType != null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Type '{newTypeName}' already exists, using it");
                    SetShelfLengthParameter(existingType, shelfLengthMm);
                    return existingType;
                }
                
                // Duplicate the type (we're inside an existing transaction)
                var newType = baseTagType.Duplicate(newTypeName) as FamilySymbol;
                
                if (newType != null)
                {
                    // Set the shelf length parameter
                    SetShelfLengthParameter(newType, shelfLengthMm);
                    return newType;
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Error creating new type: {ex.Message}");
            }
            
            return null;
        }
        
        /// <summary>
        /// Set the shelf length parameter on a tag type.
        /// IMPORTANT: The parameter expects internal units (feet), not mm.
        /// </summary>
        private void SetShelfLengthParameter(FamilySymbol tagType, double shelfLengthMm)
        {
            try
            {
                var param = tagType.LookupParameter(SHELF_LENGTH_PARAM_NAME);
                if (param != null && !param.IsReadOnly && param.StorageType == StorageType.Double)
                {
                    // CRITICAL: Convert mm to internal units (feet) before setting
                    // Revit parameters expect internal units, not display units
                    double shelfLengthFeet = shelfLengthMm / 304.8; // mm to feet
                    param.Set(shelfLengthFeet);
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Set shelf length to {shelfLengthMm:F0}mm ({shelfLengthFeet:F6} feet) on '{tagType.Name}'");
                }
                else if (param == null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Parameter '{SHELF_LENGTH_PARAM_NAME}' not found on '{tagType.Name}'");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[TAG-TYPE-MANAGER] Error setting shelf length: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Find the default tag type for an annotation plan based on element category.
        /// </summary>
        public FamilySymbol FindDefaultTagType(AnnotationPlan plan, View view)
        {
            string key = GetDefaultTypeKey(plan);
            
            if (key != null && DefaultTagTypes.TryGetValue(key, out var defaultType))
            {
                var tagType = FindTagTypeByName(defaultType.FamilyName, defaultType.TypeName);
                if (tagType != null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Found default type for {key}: {tagType.FamilyName} - {tagType.Name}");
                    return tagType;
                }
                
                // Try to find by base name (without shelf length)
                string baseTypeName = GetBaseTypeName(defaultType.TypeName);
                tagType = FindTagTypeByBaseName(defaultType.FamilyName, baseTypeName);
                if (tagType != null)
                {
                    DebugLogger.Log($"[TAG-TYPE-MANAGER] Found type by base name for {key}: {tagType.FamilyName} - {tagType.Name}");
                    return tagType;
                }
            }
            
            return null;
        }
        
        private string GetDefaultTypeKey(AnnotationPlan plan)
        {
            return plan.AnnotationType switch
            {
                AnnotationType.DuctRoundSizeFlow => "DuctRound",
                AnnotationType.DuctRectSizeFlow => "DuctRect",
                AnnotationType.AirTerminalTypeFlow => "AirTerminal",
                AnnotationType.AirTerminalShortNameFlow => "AirTerminalShort",
                AnnotationType.DuctAccessory => "DuctAccessory",
                AnnotationType.EquipmentMark => "Equipment",
                AnnotationType.SpotDimension => "SpotDimension",
                _ => null
            };
        }
        
        private FamilySymbol FindTagTypeByName(string familyName, string typeName)
        {
            try
            {
                return new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                         s.Name.Equals(typeName, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return null;
            }
        }
        
        private FamilySymbol FindTagTypeByBaseName(string familyName, string baseTypeName)
        {
            try
            {
                return new FilteredElementCollector(_document)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.FamilyName.Equals(familyName, StringComparison.OrdinalIgnoreCase) &&
                                         GetBaseTypeName(s.Name).Equals(baseTypeName, StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return null;
            }
        }
    }
}
