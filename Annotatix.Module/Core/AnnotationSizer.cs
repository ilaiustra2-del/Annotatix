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
    /// </summary>
    public class AnnotationSizer
    {
        private readonly Document _document;
        private readonly View _view;
        
        // Paper-space sizes in mm (what the annotation looks like on printed sheet)
        // These are independent of view scale - annotations always display at the same paper size
        // Widths are the LONGEST LINE width (for multi-line tags, not the combined width)
        // This is the text content width only (no padding - padding is added separately)
        private readonly Dictionary<AnnotationType, (double WidthMm, double HeightMm)> _paperSpaceSizes = 
            new Dictionary<AnnotationType, (double, double)>
        {
            { AnnotationType.DuctRoundSizeFlow, (9, 6) },      // "ø100" (size line is usually longest)
            { AnnotationType.DuctRectSizeFlow, (10, 6) },     // "200x100" (size line)
            { AnnotationType.AirTerminalTypeFlow, (12, 6) },   // "Ecoline 2" (name line)
            { AnnotationType.AirTerminalShortNameFlow, (10, 6) }, // Short name line
            { AnnotationType.DuctAccessory, (10, 6) },        // "ДКК 200"
            { AnnotationType.SpotDimension, (10, 8) },        // Elevation mark
            { AnnotationType.EquipmentMark, (12, 6) }         // "Ecoline 2"
        };
        
        // Paper-space padding in mm (minimal clearance for collision detection)
        private const double PADDING_PAPER_MM = 2.0;
        
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
        /// </summary>
        public AnnotationSize GetDefaultSize(AnnotationType type)
        {
            double viewScale = _view?.Scale ?? 100;
            if (viewScale < 1) viewScale = 1;
            
            double widthMm, heightMm;
            if (_paperSpaceSizes.TryGetValue(type, out var paperSize))
            {
                widthMm = paperSize.WidthMm;
                heightMm = paperSize.HeightMm;
            }
            else
            {
                widthMm = 35;
                heightMm = 8;
            }
            
            // Convert paper-space mm to model-space feet
            return new AnnotationSize
            {
                Width = widthMm / 304.8 * viewScale,
                Height = heightMm / 304.8 * viewScale,
                Padding = PADDING_PAPER_MM / 304.8 * viewScale
            };
        }
    }
}
