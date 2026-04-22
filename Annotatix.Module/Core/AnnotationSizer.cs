using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using PluginsManager.Core;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Represents the size of an annotation
    /// All values are in Revit internal units (feet) after conversion from paper space.
    /// </summary>
    public class AnnotationSize
    {
        public double Width { get; set; }
        public double Height { get; set; }
        
        /// <summary>Padding around the annotation for collision detection (in model feet)</summary>
        public double Padding { get; set; } = 0.05;
        
        public double TotalWidth => Width + Padding * 2;
        public double TotalHeight => Height + Padding * 2;
    }
    
    /// <summary>
    /// Measures annotation sizes by creating them temporarily and getting BoundingBox.
    /// Sizes are calculated in paper space (mm) then converted to model space (feet)
    /// using the view scale, since annotations display at fixed paper size regardless of zoom.
    /// 
    /// Conversion formula: modelFeet = paperMm / 304.8 * viewScale
    /// 
    /// Width = actual shelf length (from tag family type name)
    /// Height = calculated from character height and line count
    ///   Single line: charHeightMm
    ///   Double line: 2 * charHeightMm + lineGapMm
    /// </summary>
    public class AnnotationSizer
    {
        private readonly Document _document;
        private readonly View _view;
        
        // Character height in mm for standard Revit annotation text
        private const double CHAR_HEIGHT_MM = 2.5;
        // Line gap in mm between text lines in multi-line tags
        private const double LINE_GAP_MM = 1.5;
        
        // Paper-space sizes in mm (what the annotation looks like on printed sheet)
        // These are independent of view scale - annotations always display at the same paper size
        // WidthMm = default shelf length for this annotation type
        // HeightMm = text height (charHeight * lineCount + lineGap * (lineCount-1))
        // LineCount: 2 for duct/air terminal tags (size+flow or name+flow), 1 for others
        private readonly Dictionary<AnnotationType, (double ShelfLengthMm, double TextHeightMm, int LineCount)> _paperSpaceSizes = 
            new Dictionary<AnnotationType, (double, double, int)>
        {
            { AnnotationType.DuctRoundSizeFlow, (8, 2 * CHAR_HEIGHT_MM + LINE_GAP_MM, 2) },   // "ø100" / "L0" - 2 lines
            { AnnotationType.DuctRectSizeFlow, (10, 2 * CHAR_HEIGHT_MM + LINE_GAP_MM, 2) },   // "200x100" / "L0" - 2 lines
            { AnnotationType.AirTerminalTypeFlow, (12, 2 * CHAR_HEIGHT_MM + LINE_GAP_MM, 2) }, // "Ecoline 2" / "100" - 2 lines
            { AnnotationType.AirTerminalShortNameFlow, (10, 2 * CHAR_HEIGHT_MM + LINE_GAP_MM, 2) }, // 2 lines
            { AnnotationType.DuctAccessory, (10, CHAR_HEIGHT_MM, 1) },                         // "ДКК 200" - 1 line
            { AnnotationType.SpotDimension, (10, CHAR_HEIGHT_MM, 1) },                         // Elevation mark - 1 line
            { AnnotationType.EquipmentMark, (12, CHAR_HEIGHT_MM, 1) }                          // "Ecoline 2" - 1 line
        };
        
        // Paper-space padding in mm (clearance for collision detection)
        // Reduced from 2mm to 1mm since edge-based bbox already correctly sizes the occupied area
        private const double PADDING_PAPER_MM = 1.0;
        
        public AnnotationSizer(Document document, View view)
        {
            _document = document;
            _view = view;
        }
        
        /// <summary>
        /// Get sizes for all annotation plans (uses defaults for now)
        /// </summary>
        public Dictionary<AnnotationPlan, AnnotationSize> MeasureSizes(List<AnnotationPlan> plans)
        {
            var sizes = new Dictionary<AnnotationPlan, AnnotationSize>();
            
            foreach (var plan in plans)
            {
                sizes[plan] = GetDefaultSize(plan.AnnotationType);
            }
            
            return sizes;
        }
        
        /// <summary>
        /// Get default size for an annotation type, converted from paper-space mm to model-space feet.
        /// Conversion: modelFeet = paperMm / 304.8 * viewScale
        /// Width = shelf length (actual underline length from tag family)
        /// Height = text height (character height * line count + line gaps)
        /// </summary>
        public AnnotationSize GetDefaultSize(AnnotationType type)
        {
            double viewScale = _view?.Scale ?? 100;
            if (viewScale < 1) viewScale = 1;
            
            double shelfLengthMm, textHeightMm;
            if (_paperSpaceSizes.TryGetValue(type, out var paperSize))
            {
                shelfLengthMm = paperSize.ShelfLengthMm;
                textHeightMm = paperSize.TextHeightMm;
            }
            else
            {
                shelfLengthMm = 15;
                textHeightMm = CHAR_HEIGHT_MM;
            }
            
            // Convert paper-space mm to model-space feet
            return new AnnotationSize
            {
                Width = shelfLengthMm / 304.8 * viewScale,   // shelf length
                Height = textHeightMm / 304.8 * viewScale,    // text height
                Padding = PADDING_PAPER_MM / 304.8 * viewScale
            };
        }
    }
}
