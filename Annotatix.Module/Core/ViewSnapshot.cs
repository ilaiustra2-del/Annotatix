using System;
using System.Collections.Generic;

namespace Annotatix.Module.Core
{
    /// <summary>
    /// Represents a snapshot of a view at a point in time
    /// </summary>
    public class ViewSnapshot
    {
        public string SessionId { get; set; }
        public string DocumentName { get; set; }
        public long ViewId { get; set; }
        public string ViewName { get; set; }
        public DateTime Timestamp { get; set; }
        public string SnapshotType { get; set; } // "start" or "end"
        
        /// <summary>
        /// View scale (e.g., 100 for 1:100, 50 for 1:50)
        /// </summary>
        public double ViewScale { get; set; }
        
        /// <summary>
        /// View scale as string (e.g., "1:100", "1:50")
        /// </summary>
        public string ViewScaleString { get; set; }
        
        /// <summary>
        /// View type (e.g., "FloorPlan", "Section", "ThreeD")
        /// </summary>
        public string ViewType { get; set; }
        
        public List<ElementData> Elements { get; set; }
        public List<AnnotationData> Annotations { get; set; }
        public List<SystemData> Systems { get; set; }
        /// <summary>
        /// Available annotation family types in the project (for ML to know what tags are available)
        /// </summary>
        public List<AnnotationTypeData> AvailableAnnotationTypes { get; set; }

        public ViewSnapshot()
        {
            Elements = new List<ElementData>();
            Annotations = new List<AnnotationData>();
            Systems = new List<SystemData>();
            AvailableAnnotationTypes = new List<AnnotationTypeData>();
        }
    }

    /// <summary>
    /// Data for a model element (pipe, duct, fitting, etc.)
    /// </summary>
    public class ElementData
    {
        public long ElementId { get; set; }
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
            
        // Start coordinates (always present)
        public Coordinates3D ModelStart { get; set; }
        public Coordinates2D ViewStart { get; set; }
            
        // End coordinates (for linear elements like pipes)
        public Coordinates3D ModelEnd { get; set; }
        public Coordinates2D ViewEnd { get; set; }
            
        // HasEndPoint indicates if element has a second point
        public bool HasEndPoint { get; set; }
            
        // Size dimensions
        public double? Diameter { get; set; } // For round pipes/ducts (in internal units - feet)
        public double? Width { get; set; }    // For rectangular ducts
        public double? Height { get; set; }   // For rectangular ducts
        public string SizeDisplay { get; set; } // Human-readable size string
        
        /// <summary>
        /// Pipe slope in percentage (e.g., 2.0 for 2%)
        /// </summary>
        public double? Slope { get; set; }
        
        /// <summary>
        /// Pipe slope as string (e.g., "2.0%")
        /// </summary>
        public string SlopeDisplay { get; set; }
            
        public long? SystemId { get; set; }
        public string SystemName { get; set; }
        public string BelongTo { get; set; } // "system:{name}" or empty
    
        public ElementData()
        {
            ModelStart = new Coordinates3D();
            ViewStart = new Coordinates2D();
            ModelEnd = new Coordinates3D();
            ViewEnd = new Coordinates2D();
            HasEndPoint = false;
        }
    }

    /// <summary>
    /// Data for an annotation element (tag, text note, etc.)
    /// </summary>
    public class AnnotationData
    {
        public long ElementId { get; set; }
        public string Category { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        
        // Head position in model coordinates (full 3D position)
        public Coordinates3D HeadModelPosition { get; set; }
        
        // Head position in view coordinates (2D projection on view plane)
        public Coordinates2D HeadViewPosition { get; set; }
        
        // Leader end in model coordinates (where leader points to in 3D)
        public Coordinates3D LeaderEndModel { get; set; }
        
        // Leader end in view coordinates (2D projection)
        public Coordinates2D LeaderEndView { get; set; }
        
        // Elbow position in model coordinates (break point on leader line)
        public Coordinates3D ElbowModelPosition { get; set; }
        
        // Elbow position in view coordinates
        public Coordinates2D ElbowViewPosition { get; set; }
        
        // SpotDimension-specific: Origin (arrow position) in model coordinates
        public Coordinates3D SpotOriginModel { get; set; }
        
        // SpotDimension-specific: Origin (arrow position) in view coordinates
        public Coordinates2D SpotOriginView { get; set; }
                
        // SpotDimension-specific: LeaderShoulderPosition (bend point) in model coordinates
        public Coordinates3D LeaderShoulderModel { get; set; }
        
        // Legacy properties for backward compatibility
        public Coordinates2D HeadPosition { get; set; }
        public Coordinates2D LeaderEnd { get; set; }
        public Coordinates2D ElbowPosition { get; set; }
        
        // HasLeader indicates if annotation has a leader
        public bool HasLeader { get; set; }
        
        // HasElbow indicates if leader has a break point
        public bool HasElbow { get; set; }
        
        // Leader type: "Free", "Attached"
        public string LeaderType { get; set; }
        
        // Orientation: "Horizontal", "Vertical", etc.
        public string Orientation { get; set; }
        
        // Leader angle in degrees
        public double LeaderAngle { get; set; }
        
        public long? TaggedElementId { get; set; } // ID of the element this annotation tags
        public string BelongTo { get; set; } // "element:{id}"
        
        /// <summary>
        /// Text content of the annotation (for tags)
        /// </summary>
        public string TagText { get; set; }

        public AnnotationData()
        {
            HeadModelPosition = new Coordinates3D();
            HeadViewPosition = new Coordinates2D();
            LeaderEndModel = new Coordinates3D();
            LeaderEndView = new Coordinates2D();
            ElbowModelPosition = new Coordinates3D();
            ElbowViewPosition = new Coordinates2D();
            // SpotDimension-specific
            SpotOriginModel = new Coordinates3D();
            SpotOriginView = new Coordinates2D();
            LeaderShoulderModel = new Coordinates3D();
            // Legacy
            HeadPosition = new Coordinates2D();
            LeaderEnd = new Coordinates2D();
            ElbowPosition = new Coordinates2D();
            HasLeader = false;
            HasElbow = false;
            LeaderAngle = 0;
        }
    }

    /// <summary>
    /// Data for a system (piping, duct, etc.)
    /// </summary>
    public class SystemData
    {
        public long SystemId { get; set; }
        public string SystemName { get; set; }
        public string SystemType { get; set; }
        public List<long> ElementIds { get; set; }

        public SystemData()
        {
            ElementIds = new List<long>();
        }
    }

    /// <summary>
    /// Data for an available annotation family type in the project
    /// </summary>
    public class AnnotationTypeData
    {
        /// <summary>
        /// Element ID of the family symbol (type)
        /// </summary>
        public long TypeId { get; set; }
        
        /// <summary>
        /// Category name (e.g., "Mechanical Equipment Tags", "Pipe Tags")
        /// </summary>
        public string Category { get; set; }
        
        /// <summary>
        /// Family name
        /// </summary>
        public string FamilyName { get; set; }
        
        /// <summary>
        /// Type/symbol name
        /// </summary>
        public string TypeName { get; set; }
        
        /// <summary>
        /// Built-in category id for filtering (e.g., BuiltInCategory.OST_PipeTags)
        /// </summary>
        public int BuiltInCategoryId { get; set; }
        
        /// <summary>
        /// Whether this is the default type for the category
        /// </summary>
        public bool IsDefault { get; set; }
    }

    /// <summary>
    /// 3D coordinates (model space)
    /// </summary>
    public class Coordinates3D
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    /// <summary>
    /// 2D coordinates (view space)
    /// </summary>
    public class Coordinates2D
    {
        public double X { get; set; }
        public double Y { get; set; }
    }
}
