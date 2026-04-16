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
    /// </summary>
    public class AnnotationSize
    {
        public double Width { get; set; }
        public double Height { get; set; }
        
        /// <summary>Padding around the annotation for collision detection</summary>
        public double Padding { get; set; } = 0.05;
        
        public double TotalWidth => Width + Padding * 2;
        public double TotalHeight => Height + Padding * 2;
    }
    
    /// <summary>
    /// Measures annotation sizes by creating them temporarily and getting BoundingBox
    /// </summary>
    public class AnnotationSizer
    {
        private readonly Document _document;
        private readonly View _view;
        
        // Default sizes for different annotation types (in internal units - feet)
        private readonly Dictionary<AnnotationType, AnnotationSize> _defaultSizes = new Dictionary<AnnotationType, AnnotationSize>
        {
            { AnnotationType.DuctRoundSizeFlow, new AnnotationSize { Width = 0.2, Height = 0.05 } },
            { AnnotationType.DuctRectSizeFlow, new AnnotationSize { Width = 0.25, Height = 0.05 } },
            { AnnotationType.AirTerminalTypeFlow, new AnnotationSize { Width = 0.15, Height = 0.04 } },
            { AnnotationType.AirTerminalShortNameFlow, new AnnotationSize { Width = 0.15, Height = 0.04 } },
            { AnnotationType.DuctAccessory, new AnnotationSize { Width = 0.1, Height = 0.04 } },
            { AnnotationType.SpotDimension, new AnnotationSize { Width = 0.1, Height = 0.08 } },
            { AnnotationType.EquipmentMark, new AnnotationSize { Width = 0.2, Height = 0.05 } }
        };
        
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
        
        public AnnotationSize GetDefaultSize(AnnotationType type)
        {
            if (_defaultSizes.TryGetValue(type, out var size))
                return size;
            
            return new AnnotationSize { Width = 0.15, Height = 0.05, Padding = 0.05 };
        }
    }
}
