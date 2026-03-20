using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using PluginsManager.Core;

namespace ClashResolve.Module.Core
{
    /// <summary>
    /// A pair of intersecting pipes/ducts where PipeA will be rerouted
    /// to bypass PipeB which remains unchanged.
    /// </summary>
    public class ClashPair
    {
        /// <summary>Pipe/duct that will be split and lowered to bypass PipeB.</summary>
        public ElementId PipeAId { get; set; }

        /// <summary>Pipe/duct that remains untouched (the obstacle).</summary>
        public ElementId PipeBId { get; set; }

        /// <summary>Clearance below PipeB bottom outer edge in mm (default 50).</summary>
        public double ClearanceMm { get; set; } = 50.0;

        /// <summary>
        /// When true, ClearanceMm is ignored and the resolver automatically
        /// computes the minimum clearance required to fit elbow fittings:
        /// clearance = 2 * radiusA (one diameter of pipe A).
        /// </summary>
        public bool AutoClearance { get; set; } = false;

        /// <summary>
        /// Half-length of the detached middle segment in mm (default 300).
        /// Split point 1 = intersection - HalfLengthMm along pipe A axis.
        /// Split point 2 = intersection + HalfLengthMm along pipe A axis.
        /// Total middle segment = 2 × HalfLengthMm = 600 mm by default.
        /// </summary>
        public double HalfLengthMm { get; set; } = 300.0;

        /// <summary>
        /// Bypass angle in degrees: 90 (vertical, default), 45, 60, or 30.
        /// Values other than 90 create a diagonal transition segment.
        /// The split points are shifted outward by |dropZ| × tan(angle) so the
        /// junction segments run at exactly the chosen angle to the horizontal.
        /// </summary>
        public double AngleDegrees { get; set; } = 90.0;

        /// <summary>
        /// When true, HalfLengthMm is ignored and computed automatically:
        /// 90° mode: 2.5 × max(radiusA, radiusB)
        /// angled mode: 3.0 × max(radiusA, radiusB)
        /// </summary>
        public bool AutoHalfLength { get; set; } = false;

        /// <summary>
        /// When true, pipe A bypasses pipe B from above (A routes up over B).
        /// When false (default), pipe A bypasses pipe B from below.
        /// </summary>
        public bool BypassUp { get; set; } = false;

        /// <summary>
        /// When true, ClashResolver will look up drop/segment values in ClashLookupService
        /// before falling back to formula-based auto calculation.
        /// </summary>
        public bool UseTable { get; set; } = false;
    }

    /// <summary>
    /// A single pipe A that must bypass multiple pipes B in one combined operation.
    /// The resolver will find all intersections between A and every B,
    /// take the two outermost intersection points as the bypass zone boundaries,
    /// and drop the middle segment deep enough to clear ALL B pipes.
    /// </summary>
    public class ClashMultiPair
    {
        /// <summary>Pipe/duct that will be rerouted to bypass all pipes B.</summary>
        public ElementId PipeAId { get; set; }

        /// <summary>Pipes/ducts that remain untouched (the obstacles).</summary>
        public List<ElementId> PipeBIds { get; set; } = new List<ElementId>();

        public double ClearanceMm    { get; set; } = 50.0;
        public bool   AutoClearance  { get; set; } = false;
        public double HalfLengthMm   { get; set; } = 300.0;
        public bool   AutoHalfLength { get; set; } = false;

        /// <summary>Bypass angle in degrees: 90 (vertical), 45, 60, 30.</summary>
        public double AngleDegrees   { get; set; } = 90.0;

        /// <summary>When true, pipe A bypasses all B pipes from above.</summary>
        public bool BypassUp { get; set; } = false;

        /// <summary>
        /// When true, ClashResolver will look up drop/segment values in ClashLookupService
        /// before falling back to formula-based auto calculation.
        /// </summary>
        public bool UseTable { get; set; } = false;
    }

    /// <summary>Result of a clash resolve operation.</summary>
    public class ClashResolveResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public List<ElementId> CreatedElements { get; set; } = new List<ElementId>();

        // ── Diagnostics ─────────────────────────────────────────────────
        public ElementId PipeAId { get; set; }
        public ElementId PipeBId { get; set; }
        public XYZ PipeAStart { get; set; }
        public XYZ PipeAEnd   { get; set; }
        public double PipeARadiusMm { get; set; }
        public XYZ PipeBStart { get; set; }
        public XYZ PipeBEnd   { get; set; }
        public double PipeBRadiusMm { get; set; }
        public XYZ IntersectionPoint { get; set; }
        public double DropMm { get; set; }

        /// <summary>Actual clearance used (mm). Populated after resolve.</summary>
        public double UsedClearanceMm { get; set; }

        /// <summary>Actual half-length used (mm). Populated after resolve.</summary>
        public double UsedHalfLengthMm { get; set; }
    }

    /// <summary>
    /// Core logic for resolving a pipe/duct clash by splitting PipeA at two
    /// points around the intersection, lowering the middle segment, and
    /// connecting the joints with fittings.
    /// </summary>
    public class ClashResolver
    {
        // 1 foot = 304.8 mm
        private const double FeetPerMm = 1.0 / 304.8;

        /// <summary>MEP element kind used for choosing calculation multipliers.</summary>
        // MepKind enum is now defined in ClashLookupTable.cs (public, same namespace)

        /// <summary>
        /// Determine the kind of MEP element: round duct, rectangular duct, or pipe.
        /// Rectangular ducts have both RBS_CURVE_WIDTH_PARAM and RBS_CURVE_HEIGHT_PARAM.
        /// </summary>
        private MepKind GetMepKind(Element elem)
        {
            if (elem is Pipe) return MepKind.Pipe;
            if (elem is Duct duct)
            {
                Parameter w = duct.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                Parameter h = duct.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (w != null && h != null && w.HasValue && h.HasValue)
                    return MepKind.RectDuct;
            }
            return MepKind.RoundDuct;
        }

        /// <summary>
        /// Returns the nominal size in mm for use in lookup table matching.
        /// For Pipe: uses RBS_PIPE_DIAMETER_PARAM (nominal inner diameter).
        /// For Duct: uses height param for RectDuct, diameter for RoundDuct.
        /// Falls back to outer-radius * 2 if nominal param is unavailable.
        /// </summary>
        private static double GetNominalSizeMm(Element elem, MepKind kind, double radiusFt)
        {
            MEPCurve mep = elem as MEPCurve;
            if (mep != null)
            {
                if (kind == MepKind.Pipe)
                {
                    // Try nominal (inner) diameter first — matches DN sizing
                    Parameter nomDiam = mep.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                                     ?? mep.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    if (nomDiam != null && nomDiam.HasValue && nomDiam.AsDouble() > 0)
                        return nomDiam.AsDouble() * 304.8;
                }
                else if (kind == MepKind.RectDuct)
                {
                    // Use height param for rect duct (H dimension)
                    Parameter h = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                    if (h != null && h.HasValue && h.AsDouble() > 0)
                        return h.AsDouble() * 304.8;
                }
                else // RoundDuct
                {
                    Parameter diam = mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                                  ?? mep.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    if (diam != null && diam.HasValue && diam.AsDouble() > 0)
                        return diam.AsDouble() * 304.8;
                }
            }
            // Fallback: outer diameter from radius
            return radiusFt * 2.0 * 304.8;
        }

        /// <summary>
        /// Returns the clearance multiplier applied to radiusA (= height/2) for rectangular
        /// ducts at angled connections:
        ///   30° → 2.0  (1.0 × height = 2.0 × radiusA)
        ///   45° → 6.0  (3.0 × height = 6.0 × radiusA)
        ///   60° → 10.0 (5.0 × height = 10.0 × radiusA)
        ///   other → 6.0 (default / 45°-equivalent)
        /// Multipliers doubled compared to previous version to provide enough
        /// room for elbow fittings at the angled joints.
        /// </summary>
        private static double GetRectDuctAngledClearanceMult(double angleDegrees)
        {
            if (Math.Abs(angleDegrees - 30.0) < 0.5) return 2.0;
            if (Math.Abs(angleDegrees - 45.0) < 0.5) return 6.0;
            if (Math.Abs(angleDegrees - 60.0) < 0.5) return 10.0;
            return 6.0; // fallback
        }

        public ClashResolveResult ResolveClash(Document doc, ClashPair pair)
        {
            var result = new ClashResolveResult();
            try
            {
                // -------------------------------------------------------
                // A. Collect geometry
                // -------------------------------------------------------
                Element elemA = doc.GetElement(pair.PipeAId);
                Element elemB = doc.GetElement(pair.PipeBId);

                if (elemA == null || elemB == null)
                {
                    result.Message = "Один из элементов не найден в документе.";
                    return result;
                }

                if (!TryGetAxisAndRadius(elemA, out XYZ pt1, out XYZ pt2, out double radiusA))
                {
                    result.Message = "Элемент A не является трубой или воздуховодом (не удалось получить ось).";
                    return result;
                }

                if (!TryGetAxisAndRadius(elemB, out XYZ pt3, out XYZ pt4, out double radiusB))
                {
                    result.Message = "Элемент B не является трубой или воздуховодом (не удалось получить ось).";
                    return result;
                }

                // Fill diagnostics
                result.PipeAId       = pair.PipeAId;
                result.PipeBId       = pair.PipeBId;
                result.PipeAStart    = pt1;
                result.PipeAEnd      = pt2;
                result.PipeARadiusMm = radiusA * 304.8;
                result.PipeBStart    = pt3;
                result.PipeBEnd      = pt4;
                result.PipeBRadiusMm = radiusB * 304.8;

                DebugLogger.Log($"[CLASH-RESOLVER] PipeA: {pt1} -> {pt2}, radius={radiusA * 304.8:F1}mm");
                DebugLogger.Log($"[CLASH-RESOLVER] PipeB: {pt3} -> {pt4}, radius={radiusB * 304.8:F1}mm");

                // Determine element kind (round duct / rect duct / pipe) for multiplier selection
                MepKind kindA = GetMepKind(elemA);
                DebugLogger.Log($"[CLASH-RESOLVER] ElemA kind: {kindA}, radius={radiusA * 304.8:F1}mm");

                // -------------------------------------------------------
                // Lookup table override (applied before auto-formula logic)
                // -------------------------------------------------------
                if (pair.UseTable)
                {
                    double sizeMm = GetNominalSizeMm(elemA, kindA, radiusA);
                    double? tableDropMm = ClashLookupService.Instance.LookupDrop(kindA, sizeMm, pair.AngleDegrees);
                    double? tableSegMm  = ClashLookupService.Instance.LookupSegment(kindA, sizeMm, pair.AngleDegrees);
                    DebugLogger.Log($"[CLASH-RESOLVER] Lookup: kind={kindA} sizeMm={sizeMm:F1} angle={pair.AngleDegrees:F0}° drop={tableDropMm?.ToString("F0") ?? "null"} seg={tableSegMm?.ToString("F0") ?? "null"}");

                    if (tableDropMm.HasValue)
                    {
                        pair.AutoClearance = false;
                        pair.ClearanceMm   = tableDropMm.Value;
                        DebugLogger.Log($"[CLASH-RESOLVER] Lookup drop override: {tableDropMm.Value:F0}mm");
                    }
                    if (tableSegMm.HasValue)
                    {
                        pair.AutoHalfLength = false;
                        pair.HalfLengthMm   = tableSegMm.Value / 2.0;
                        DebugLogger.Log($"[CLASH-RESOLVER] Lookup segment override: {tableSegMm.Value:F0}mm (half={pair.HalfLengthMm:F0}mm)");
                    }
                }

                // -------------------------------------------------------
                // B. Find intersection point on axis of A (closest approach)
                // -------------------------------------------------------
                XYZ ptOnA = GetClosestPointOnLine(pt1, pt2, pt3, pt4, clamped: false);
                result.IntersectionPoint = ptOnA;
                DebugLogger.Log($"[CLASH-RESOLVER] Intersection point on A: {ptOnA}");

                // -------------------------------------------------------
                // C. Calculate split points along A's axis
                // -------------------------------------------------------
                XYZ dirA = (pt2 - pt1).Normalize();

                // splitOffset = user-defined half-length (default 300 mm)
                // splitPt1 = intersection - halfLength, splitPt2 = intersection + halfLength
                double minSegmentFt = 0.15; // ~46 mm — safe margin from pipe endpoints for BreakCurve

                // Auto half-length multipliers depend on element kind and angle:
                //   RoundDuct  90°: 2.5×   angled: 3.0×
                //   Pipe       90°: 2.5×   angled: 1.5×  (halved per user feedback — was too long)
                //   RectDuct   90°: 6.0×   angled: 6.0×  (increased to give fittings enough room)
                bool isAngled = Math.Abs(pair.AngleDegrees - 90.0) > 0.5;
                double halfLengthMultiplier = isAngled ? 3.0 : 2.5;
                if (kindA == MepKind.RectDuct && !isAngled)
                    halfLengthMultiplier = 6.0;
                if (kindA == MepKind.RectDuct && isAngled)
                    halfLengthMultiplier = 6.0; // same as 90° — rect duct fittings need more room
                if (kindA == MepKind.Pipe && isAngled)
                    halfLengthMultiplier = 1.5; // reduced: was 3.0 + 1D bonus, now 1.5× only

                double effectiveHalfLengthMm = pair.AutoHalfLength
                    ? Math.Ceiling(Math.Max(radiusA, radiusB) * halfLengthMultiplier * 304.8)
                    : pair.HalfLengthMm;

                // Pipe bonus in 90° mode only (angled mode already handled via multiplier above)
                if (kindA == MepKind.Pipe && pair.AutoHalfLength && !isAngled)
                    effectiveHalfLengthMm += radiusA * 2.0 * 304.8;

                if (pair.AutoHalfLength)
                    DebugLogger.Log($"[CLASH-RESOLVER] AutoHalfLength ({pair.AngleDegrees:F0}°, {kindA}): → halfLength={effectiveHalfLengthMm:F0}mm");

                double splitOffset = effectiveHalfLengthMm * FeetPerMm; // convert mm → feet
                result.UsedHalfLengthMm = effectiveHalfLengthMm;

                // -------------------------------------------------------
                // D. Calculate vertical shift for middle segment
                // -------------------------------------------------------
                double effectiveClearanceMm;
                if (pair.AutoClearance)
                {
                    // Base clearance depends on element kind and angle:
                    //   RoundDuct  90°: 3.5×max(R)   angled: 50mm
                    //   Pipe       90°: 3.5×max(R)   angled: 50mm  + 1 diameter bonus
                    //   RectDuct   90°: 9.0×max(R)
                    //              angled: enough to make the diagonal segment at least 3×width long
                    //              so Revit can fit elbow fittings at both ends of the diagonal.
                    //              minDiagMm = 3 × (2×radiusA_mm);  minShiftMm = minDiagMm × sin(angle)
                    //              baseClearance = max(50, minShiftMm - radiusA_mm - radiusB_mm)
                    double baseClearanceMm;
                    if (kindA == MepKind.RectDuct)
                    {
                        if (isAngled)
                        {
                            // Compute minimum shift so the diagonal segment remains long enough
                            // after Revit inserts elbow fittings at both ends.
                            // Each elbow "consumes" ~1.88×radiusA from the diagonal.
                            // minDiag = 2 × elbowSize + 1 width clearance = 2×(1.88×W) + W ≈ 4.76×W.
                            // Use 5.0× for safety margin.
                            // diagonal length = shift / sin(angle); shift = radiusA + clearance + radiusB.
                            double radiusA_mm = radiusA * 304.8;
                            double radiusB_mm = radiusB * 304.8;
                            double minDiagMm  = 5.0 * (2.0 * radiusA_mm);
                            double sinAngle   = Math.Sin(pair.AngleDegrees * Math.PI / 180.0);
                            double minShiftMm = minDiagMm * sinAngle;
                            baseClearanceMm   = Math.Max(50.0, minShiftMm - radiusA_mm - radiusB_mm);
                        }
                        else
                        {
                            baseClearanceMm = Math.Ceiling(Math.Max(radiusA, radiusB) * 9.0 * 304.8);
                        }
                    }
                    else
                    {
                        // Angled: 50mm base (same for both RoundDuct and Pipe — pipe diameter
                        // bonus is NOT applied in angled mode because it makes expandFt too large).
                        // 90°: 3.5×max(R) for RoundDuct; 3.5×max(R) + 1D for Pipe.
                        baseClearanceMm = isAngled
                            ? 50.0
                            : Math.Ceiling(Math.Max(radiusA, radiusB) * 3.5 * 304.8);
                        // Pipe bonus only in 90° (vertical) mode
                        if (kindA == MepKind.Pipe && !isAngled)
                            baseClearanceMm += radiusA * 2.0 * 304.8;
                    }

                    double pipeBAxisZ_check = GetZAtXY(pt3, pt4, ptOnA.X, ptOnA.Y);
                    double extraClearanceMm = 0.0;

                    if (!pair.BypassUp)
                    {
                        // Bypass down: B may be just above A — check if fitting won't clear B's underside.
                        // Extra clearance is only added when B's axis is clearly above A's axis
                        // (at least by more than the sum of outer radii), meaning B sits above A
                        // and the fitting elbow would collide with B's bottom surface.
                        // When B and A intersect in the same horizontal plane (standard collision),
                        // headroomMm is negative and adding extra clearance would grossly over-shift.
                        double clearanceGapMm = (pipeBAxisZ_check - ptOnA.Z) * 304.8;
                        bool bIsAboveA = clearanceGapMm > (radiusA + radiusB) * 304.8;
                        if (bIsAboveA)
                        {
                            double headroomMm      = (pipeBAxisZ_check - radiusB - ptOnA.Z) * 304.8;
                            double fittingHeightMm = Math.Max(radiusA, radiusB) * 304.8;
                            if (headroomMm < fittingHeightMm + baseClearanceMm)
                            {
                                extraClearanceMm = fittingHeightMm + baseClearanceMm - headroomMm;
                                DebugLogger.Log($"[CLASH-RESOLVER] B above A — adding extra clearance {extraClearanceMm:F0}mm");
                            }
                        }
                    }
                    // Bypass up: fitting goes upward away from B, no extra clearance needed.

                    effectiveClearanceMm = baseClearanceMm + extraClearanceMm;
                                        DebugLogger.Log($"[CLASH-RESOLVER] AutoClearance ({pair.AngleDegrees:F0}°, {(pair.BypassUp ? "сверху" : "снизу")}): base={baseClearanceMm:F0}mm extra={extraClearanceMm:F0}mm total={effectiveClearanceMm:F0}mm");
                }
                else
                {
                    effectiveClearanceMm = pair.ClearanceMm;
                }

                // For RectDuct angled bypass, enforce a minimum clearance regardless of manual/auto mode.
                // Revit elbow fittings each consume ~1.88×halfWidth from the diagonal segment.
                // The diagonal must be >= 5×width (minDiag) so after both elbows enough remains.
                // minShift = minDiag × sin(angle); minClearance = minShift - halfWidthA - halfWidthB.
                if (kindA == MepKind.RectDuct && isAngled)
                {
                    double rAmm   = radiusA * 304.8;
                    double rBmm   = radiusB * 304.8;
                    double minDiag = 5.0 * (2.0 * rAmm);
                    double sinA    = Math.Sin(pair.AngleDegrees * Math.PI / 180.0);
                    double minClearance = Math.Max(50.0, minDiag * sinA - rAmm - rBmm);
                    if (effectiveClearanceMm < minClearance)
                    {
                        DebugLogger.Log($"[CLASH-RESOLVER] RectDuct angled: user clearance {effectiveClearanceMm:F0}mm < required {minClearance:F0}mm — overriding");
                        effectiveClearanceMm = minClearance;
                    }
                }

                result.UsedClearanceMm = effectiveClearanceMm;

                double clearanceFeet = effectiveClearanceMm * FeetPerMm;
                double pipeBAxisZ    = GetZAtXY(pt3, pt4, ptOnA.X, ptOnA.Y);

                double middleAxisZ;
                if (!pair.BypassUp)
                {
                    // A goes below B
                    double pipeBBottomOuterZ = pipeBAxisZ - radiusB;
                    middleAxisZ = pipeBBottomOuterZ - clearanceFeet - radiusA;
                }
                else
                {
                    // A goes above B
                    double pipeBTopOuterZ = pipeBAxisZ + radiusB;
                    middleAxisZ = pipeBTopOuterZ + clearanceFeet + radiusA;
                }

                double currentMiddleZ = ptOnA.Z;
                double dropZ          = middleAxisZ - currentMiddleZ;
                XYZ moveVector        = new XYZ(0, 0, dropZ);

                result.DropMm = dropZ * 304.8;

                DebugLogger.Log($"[CLASH-RESOLVER] PipeB axis Z at intersection: {pipeBAxisZ * 304.8:F1}mm");
                DebugLogger.Log($"[CLASH-RESOLVER] Direction: {(pair.BypassUp ? "сверху" : "снизу")}, shift: {dropZ * 304.8:F1}mm");

                if (!pair.BypassUp && dropZ >= 0)
                {
                    result.Message = "Труба A уже находится ниже трубы B. Обход не требуется.";
                    return result;
                }
                if (pair.BypassUp && dropZ <= 0)
                {
                    result.Message = "Труба A уже находится выше трубы B. Обход не требуется.";
                    return result;
                }

                // -------------------------------------------------------
                // C2. Compute final split offsets.
                // Angle is measured from horizontal: dxy/dz = cot(angle) = 1/tan(angle).
                // 30° from horiz → cot(30°)=1.732 (long/shallow diagonal)
                // 45° from horiz → cot(45°)=1.0   (equal dxy and dz)
                // 60° from horiz → cot(60°)=0.577 (short/steep diagonal)
                // 90° mode: no expansion (vertical duct between levels).
                //
                // splitOffset = flatHalfLength + expandFt
                //   where flatHalfLength is the half-length of the flat lowered section,
                //   and expandFt is the horizontal footprint of the diagonal connector.
                // t1 = tCenter - splitOffset - expandFt  (= tCenter - flatHalf - 2*expand)
                // t2 = tCenter + splitOffset + expandFt  (= tCenter + flatHalf + 2*expand)
                // After moving middle down, it is trimmed inward by expandFt on each side,
                // leaving flatHalfLength*2 as the actual flat lowered section.
                // -------------------------------------------------------
                double expandFt = isAngled
                    ? Math.Abs(dropZ) / Math.Tan(pair.AngleDegrees * Math.PI / 180.0)
                    : 0.0;
                if (isAngled)
                    DebugLogger.Log($"[CLASH-RESOLVER] {pair.AngleDegrees:F0}° mode: expandFt={expandFt * 304.8:F1}mm each side");

                if (isAngled)
                {
                    // For auto mode: splitOffset already represents the desired flat half-length.
                    // We promote it to splitOffset = flatHalf + expandFt so the flat section
                    // comes out as expected, instead of being eroded by expandFt.
                    // For manual mode: user-set HalfLengthMm is treated as the flat section too.
                    // Minimum flat section: max(effectiveHalfLengthMm, 1*diameter) = max(user, 2*radiusA).
                    double minFlatHalfFt = 2.0 * radiusA; // at least 1 diameter flat each side
                    double flatHalfFt    = Math.Max(splitOffset, minFlatHalfFt);
                    splitOffset = flatHalfFt + expandFt;
                    result.UsedHalfLengthMm = flatHalfFt / FeetPerMm;
                    DebugLogger.Log($"[CLASH-RESOLVER] flatHalf={flatHalfFt * 304.8:F1}mm + expand={expandFt * 304.8:F1}mm → splitOffset={splitOffset * 304.8:F1}mm");
                }

                // Parametric positions along A's axis (from pt1)
                double totalLength = pt1.DistanceTo(pt2);
                double tCenter     = (ptOnA - pt1).DotProduct(dirA);
                double t1 = tCenter - splitOffset;
                double t2 = tCenter + splitOffset;

                DebugLogger.Log($"[CLASH-RESOLVER] tCenter={tCenter * 304.8:F1}mm (totalLength={totalLength * 304.8:F1}mm)");

                // Clamp to safe range away from pipe endpoints
                t1 = Math.Max(minSegmentFt, Math.Min(t1, totalLength - minSegmentFt));
                t2 = Math.Max(minSegmentFt, Math.Min(t2, totalLength - minSegmentFt));

                if (t1 >= t2 - minSegmentFt)
                {
                    result.Message = $"Труба A слишком короткая для обхода. " +
                        $"Длина трубы A: {totalLength * 304.8:F0}мм, " +
                        $"требуется минимум {(2 * (splitOffset + expandFt) + 2 * minSegmentFt) * 304.8:F0}мм.";
                    return result;
                }

                XYZ splitPt1 = pt1 + dirA.Multiply(t1);
                XYZ splitPt2 = pt1 + dirA.Multiply(t2);

                DebugLogger.Log($"[CLASH-RESOLVER] HalfLength={effectiveHalfLengthMm:F0}mm → splitOffset={splitOffset * 304.8:F1}mm (expand={expandFt * 304.8:F1}mm)");
                DebugLogger.Log($"[CLASH-RESOLVER] Split pt1: t={t1 * 304.8:F1}mm, pt2: t={t2 * 304.8:F1}mm from start");
                DebugLogger.Log($"[CLASH-RESOLVER] Middle segment length: {(t2 - t1) * 304.8:F1}mm");

                // -------------------------------------------------------
                // E. Execute in transaction: split → move → connect
                // -------------------------------------------------------
                using (var tx = new Transaction(doc, "Clash Resolve: обход трубы"))
                {
                    tx.Start();

                    // -------------------------------------------------------
                    // E1. First cut at splitPt1
                    // BreakCurve returns the NEW segment ID. Revit may return
                    // either the left or right piece depending on internal direction.
                    // We must determine which piece contains splitPt2 by geometry.
                    // -------------------------------------------------------
                    ElementId returnedIdAfterCut1 = SplitCurve(doc, pair.PipeAId, splitPt1);
                    if (returnedIdAfterCut1 == null || returnedIdAfterCut1 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"Не удалось выполнить первый разрез трубы A " +
                            $"(точка: X={splitPt1.X * 304.8:F0}, Y={splitPt1.Y * 304.8:F0}, Z={splitPt1.Z * 304.8:F0}мм).";
                        return result;
                    }
                    DebugLogger.Log($"[CLASH-RESOLVER] First cut returned new ID={returnedIdAfterCut1}, original ID={pair.PipeAId}");

                    // Determine which of the two resulting segments contains splitPt2
                    // by checking which one's axis contains the point within its length.
                    ElementId segContainingSplit2 = FindSegmentContainingPoint(doc,
                        new[] { pair.PipeAId, returnedIdAfterCut1 }, splitPt2, minSegmentFt);

                    if (segContainingSplit2 == null || segContainingSplit2 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"После первого разреза не удалось найти сегмент, содержащий вторую точку разреза " +
                            $"(X={splitPt2.X * 304.8:F0}, Y={splitPt2.Y * 304.8:F0}, Z={splitPt2.Z * 304.8:F0}мм).";
                        return result;
                    }

                    // The "left" segment is the one that does NOT contain splitPt2
                    ElementId leftId = (segContainingSplit2 == pair.PipeAId)
                        ? returnedIdAfterCut1
                        : pair.PipeAId;

                    if (TryGetAxisAndRadius(doc.GetElement(segContainingSplit2), out XYZ mrStart, out XYZ mrEnd, out _))
                    {
                        double mrLength = mrStart.DistanceTo(mrEnd);
                        double tSplit2check = (splitPt2 - mrStart).DotProduct((mrEnd - mrStart).Normalize());
                        DebugLogger.Log($"[CLASH-RESOLVER] Segment for 2nd cut: ID={segContainingSplit2}, length={mrLength * 304.8:F1}mm, tSplit2={tSplit2check * 304.8:F1}mm from its start");
                    }

                    // -------------------------------------------------------
                    // E2. Second cut at splitPt2 on the correct segment
                    // -------------------------------------------------------
                    ElementId returnedIdAfterCut2 = SplitCurve(doc, segContainingSplit2, splitPt2);
                    if (returnedIdAfterCut2 == null || returnedIdAfterCut2 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"Не удалось выполнить второй разрез трубы A " +
                            $"(точка: X={splitPt2.X * 304.8:F0}, Y={splitPt2.Y * 304.8:F0}, Z={splitPt2.Z * 304.8:F0}мм).";
                        return result;
                    }
                    DebugLogger.Log($"[CLASH-RESOLVER] Second cut returned new ID={returnedIdAfterCut2}, cut segment ID={segContainingSplit2}");

                    // After second cut we have three segments:
                    //   leftId            – from original pt1 to splitPt1
                    //   segContainingSplit2 – from splitPt1 to splitPt2  (the middle = to be lowered)
                    //   returnedIdAfterCut2 – from splitPt2 to original pt2 (the right)
                    // But again BreakCurve may have returned left or right piece.
                    // Middle = the piece that contains ptOnA (the intersection point).
                    ElementId middleId = FindSegmentContainingPoint(doc,
                        new[] { segContainingSplit2, returnedIdAfterCut2 }, ptOnA, 0);
                    ElementId rightId  = (middleId == segContainingSplit2)
                        ? returnedIdAfterCut2
                        : segContainingSplit2;

                    if (middleId == null || middleId == ElementId.InvalidElementId)
                    {
                        // Fallback: assume segContainingSplit2 is the middle
                        middleId = segContainingSplit2;
                        rightId  = returnedIdAfterCut2;
                        DebugLogger.Log("[CLASH-RESOLVER] Warning: could not determine middle by intersection point, using fallback.");
                    }

                    DebugLogger.Log($"[CLASH-RESOLVER] Segments: left={leftId}, middle={middleId}, right={rightId}");

                    // -------------------------------------------------------
                    // E3. Move middle segment down
                    // -------------------------------------------------------
                    ElementTransformUtils.MoveElement(doc, middleId, moveVector);
                    DebugLogger.Log($"[CLASH-RESOLVER] Middle segment moved by {dropZ * 304.8:F1}mm");

                    bool conn1, conn2;
                    if (isAngled && expandFt > 1e-9)
                    {
                        // -------------------------------------------------------
                        // E3b. Angled mode: trim middle inward by expandFt on each side,
                        // then create diagonal segments + elbows via Connect45Fixed.
                        //
                        // After lowering, middle spans splitPt1..splitPt2 @ Z_low.
                        // We trim it to (splitPt1+expand)...(splitPt2-expand) @ Z_low.
                        // Gap: dxy=expandFt, dz=|dropZ| → angle = atan(dz/dxy) matches chosen angle.
                        // We create a diagonal segment ptA→ptB, then:
                        //   • elbow(connA_horiz, diagConnTop)  — angle fitting at top
                        //   • ConnectTo(diagConnBot, connB_middle) — coincident, no elbow needed
                        // -------------------------------------------------------
                        try
                        {
                            MEPCurve middleMep = doc.GetElement(middleId) as MEPCurve;
                            if (middleMep != null)
                            {
                                var lc = middleMep.Location as LocationCurve;
                                Line oldLine = lc?.Curve as Line;
                                if (oldLine != null)
                                {
                                    XYZ mStart = oldLine.GetEndPoint(0);
                                    XYZ mEnd   = oldLine.GetEndPoint(1);
                                    bool startIsLeft = mStart.DistanceTo(splitPt1) < mEnd.DistanceTo(splitPt1);
                                    XYZ leftEnd  = startIsLeft ? mStart : mEnd;
                                    XYZ rightEnd = startIsLeft ? mEnd   : mStart;
                                    XYZ newLeft  = new XYZ(leftEnd.X  + dirA.X * expandFt, leftEnd.Y  + dirA.Y * expandFt, leftEnd.Z);
                                    XYZ newRight = new XYZ(rightEnd.X - dirA.X * expandFt, rightEnd.Y - dirA.Y * expandFt, rightEnd.Z);
                                    lc.Curve = Line.CreateBound(
                                        startIsLeft ? newLeft : newRight,
                                        startIsLeft ? newRight : newLeft);
                                    DebugLogger.Log($"[CLASH-RESOLVER] {pair.AngleDegrees:F0}° middle trimmed: left={newLeft.X*304.8:F0},{newLeft.Y*304.8:F0},{newLeft.Z*304.8:F0}  right={newRight.X*304.8:F0},{newRight.Y*304.8:F0},{newRight.Z*304.8:F0}");
                                }
                            }
                        }
                        catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] {pair.AngleDegrees:F0}° trim failed: {ex.Message}"); }

                        conn1 = Connect45Fixed(doc, leftId,  middleId, result);
                        conn2 = Connect45Fixed(doc, rightId, middleId, result);
                    }
                    else
                    {
                        // 90° mode: gap is purely vertical → Case B (vertical duct + 2×90° elbows)
                        conn1 = ConnectOpenEnds(doc, leftId,  middleId, result);
                        conn2 = ConnectOpenEnds(doc, middleId, rightId, result);
                    }

                    if (!conn1 || !conn2)
                        DebugLogger.Log("[CLASH-RESOLVER] Warning: some connections could not be created automatically.");

                    tx.Commit();
                }

                result.Success = true;
                result.Message = "Обход успешно создан. Труба A перестроена под трубу B.";
                return result;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] ERROR: {ex.Message}\n{ex.StackTrace}");
                result.Message = $"Ошибка: {ex.Message}";
                return result;
            }
        }

        // -----------------------------------------------------------------------
        // Multi-B: one pipe A bypasses ALL selected pipes B in a single operation
        // -----------------------------------------------------------------------

        /// <summary>
        /// Resolve a clash where pipe A must bypass multiple pipes B simultaneously.
        /// Steps:
        ///   1. Find all B pipes that physically intersect A.
        ///   2. Project intersections onto A's axis → take t_min / t_max as bypass zone.
        ///   3. Compute the deepest required drop across all intersecting B pipes.
        ///   4. Split A at (t_min - halfLength) and (t_max + halfLength), lower middle, connect.
        /// If pipe A intersects no B pipe, returns Success=false with empty Message (silent skip).
        /// </summary>
        public ClashResolveResult ResolveClashMultiB(Document doc, ClashMultiPair pair)
        {
            var result = new ClashResolveResult();
            try
            {
                // -------------------------------------------------------
                // A. Geometry of pipe A
                // -------------------------------------------------------
                Element elemA = doc.GetElement(pair.PipeAId);
                if (elemA == null)
                {
                    result.Message = "Труба A не найдена в документе.";
                    return result;
                }

                if (!TryGetAxisAndRadius(elemA, out XYZ pt1, out XYZ pt2, out double radiusA))
                {
                    result.Message = "Элемент A не является трубой или воздуховодом.";
                    return result;
                }

                XYZ dirA        = (pt2 - pt1).Normalize();
                double totalLen = pt1.DistanceTo(pt2);
                result.PipeAId    = pair.PipeAId;
                result.PipeAStart = pt1;
                result.PipeAEnd   = pt2;
                result.PipeARadiusMm = radiusA * 304.8;

                DebugLogger.Log($"[MULTI-RESOLVER] PipeA ID={pair.PipeAId.Value}: {pt1} -> {pt2}, r={radiusA * 304.8:F1}mm");

                // Determine element kind for multiplier selection
                MepKind kindA = GetMepKind(elemA);
                DebugLogger.Log($"[MULTI-RESOLVER] ElemA kind: {kindA}, radius={radiusA * 304.8:F1}mm");

                // -------------------------------------------------------
                // Lookup table override (applied before auto-formula logic)
                // -------------------------------------------------------
                if (pair.UseTable)
                {
                    double sizeMm = GetNominalSizeMm(elemA, kindA, radiusA);
                    double? tableDropMm = ClashLookupService.Instance.LookupDrop(kindA, sizeMm, pair.AngleDegrees);
                    double? tableSegMm  = ClashLookupService.Instance.LookupSegment(kindA, sizeMm, pair.AngleDegrees);
                    DebugLogger.Log($"[MULTI-RESOLVER] Lookup: kind={kindA} sizeMm={sizeMm:F1} angle={pair.AngleDegrees:F0}° drop={tableDropMm?.ToString("F0") ?? "null"} seg={tableSegMm?.ToString("F0") ?? "null"}");

                    if (tableDropMm.HasValue)
                    {
                        pair.AutoClearance = false;
                        pair.ClearanceMm   = tableDropMm.Value;
                        DebugLogger.Log($"[MULTI-RESOLVER] Lookup drop override: {tableDropMm.Value:F0}mm");
                    }
                    if (tableSegMm.HasValue)
                    {
                        pair.AutoHalfLength = false;
                        pair.HalfLengthMm   = tableSegMm.Value / 2.0;
                        DebugLogger.Log($"[MULTI-RESOLVER] Lookup segment override: {tableSegMm.Value:F0}mm (half={pair.HalfLengthMm:F0}mm)");
                    }
                }

                // -------------------------------------------------------
                // B. Find all intersecting B pipes, compute outermost t values
                //    and the deepest required drop.
                // -------------------------------------------------------
                double tMin      = double.MaxValue;
                double tMax      = double.MinValue;
                double dropZFinal = 0.0;   // most negative = deepest required

                // Parameters object compatible with ComputeDropZ
                var dummyPair = new ClashPair
                {
                    AngleDegrees  = pair.AngleDegrees,
                    AutoClearance = pair.AutoClearance,
                    ClearanceMm   = pair.ClearanceMm,
                    BypassUp      = pair.BypassUp,
                };

                // For bypass-down: keep the most-negative (deepest) dropZ
                // For bypass-up:   keep the most-positive (highest) dropZ
                dropZFinal = pair.BypassUp ? double.MinValue : double.MaxValue;

                int bCount = 0;
                foreach (var pipeBId in pair.PipeBIds)
                {
                    Element elemB = doc.GetElement(pipeBId);
                    if (elemB == null) continue;

                    if (!TryGetAxisAndRadius(elemB, out XYZ pb1, out XYZ pb2, out double radiusB))
                        continue;

                    // Point on A's axis closest to B's axis
                    XYZ ptOnA = GetClosestPointOnLine(pt1, pt2, pb1, pb2);
                    // Corresponding point on B's axis closest to ptOnA
                    XYZ ptOnB = GetClosestPointOnLine(pb1, pb2, pt1, pt2);

                    double gapDist = ptOnA.DistanceTo(ptOnB);
                    DebugLogger.Log($"[MULTI-RESOLVER]   B ID={pipeBId.Value}: gapDist={gapDist * 304.8:F1}mm, sumR={(radiusA + radiusB) * 304.8:F1}mm");

                    // Physical intersection check: axes must be closer than sum of radii
                    if (gapDist > radiusA + radiusB + 0.01) // 0.01 ft ≈ 3mm tolerance
                        continue;

                    double t = (ptOnA - pt1).DotProduct(dirA);
                    if (t < tMin) { tMin = t; }
                    if (t > tMax) { tMax = t; }

                    double dropZ_i = ComputeDropZ(pt1, pt2, radiusA, pb1, pb2, radiusB, ptOnA, dummyPair, kindA);
                    DebugLogger.Log($"[MULTI-RESOLVER]   B ID={pipeBId.Value}: t={t * 304.8:F1}mm, shift={dropZ_i * 304.8:F1}mm");

                    // Keep the "worst" shift: most negative (down) or most positive (up)
                    if (!pair.BypassUp && dropZ_i < dropZFinal)
                        dropZFinal = dropZ_i;
                    else if (pair.BypassUp && dropZ_i > dropZFinal)
                        dropZFinal = dropZ_i;

                    bCount++;
                }

                if (dropZFinal == double.MaxValue || dropZFinal == double.MinValue)
                    dropZFinal = 0.0; // safety reset if no B was processed

                if (bCount == 0)
                {
                    // No actual intersections — silent skip
                    DebugLogger.Log($"[MULTI-RESOLVER] PipeA ID={pair.PipeAId.Value}: no intersecting B pipes found, skipping.");
                    result.Message = "";
                    result.Success = false;
                    return result;
                }

                DebugLogger.Log($"[MULTI-RESOLVER] Found {bCount} intersecting B pipes. tMin={tMin * 304.8:F1}mm tMax={tMax * 304.8:F1}mm dropZFinal={dropZFinal * 304.8:F1}mm");

                if (!pair.BypassUp && dropZFinal >= 0)
                {
                    result.Message = "Труба A уже находится ниже всех труб B. Обход не требуется.";
                    return result;
                }
                if (pair.BypassUp && dropZFinal <= 0)
                {
                    result.Message = "Труба A уже находится выше всех труб B. Обход не требуется.";
                    return result;
                }

                // -------------------------------------------------------
                // C. Compute half-length and expand (angled mode)
                // -------------------------------------------------------
                bool isAngledMulti = Math.Abs(pair.AngleDegrees - 90.0) > 0.5;

                // Half-length multipliers (same rules as ResolveClash)
                double halfLengthMultiplierM = isAngledMulti ? 3.0 : 2.5;
                if (kindA == MepKind.RectDuct && !isAngledMulti)
                    halfLengthMultiplierM = 6.0;
                if (kindA == MepKind.RectDuct && isAngledMulti)
                    halfLengthMultiplierM = 6.0; // same as 90° — rect duct fittings need more room
                if (kindA == MepKind.Pipe && isAngledMulti)
                    halfLengthMultiplierM = 1.5; // reduced: was 3.0 + 1D bonus

                double effectiveHalfLengthMm = pair.AutoHalfLength
                    ? Math.Ceiling(radiusA * halfLengthMultiplierM * 304.8)
                    : pair.HalfLengthMm;
                if (kindA == MepKind.Pipe && pair.AutoHalfLength && !isAngledMulti)
                    effectiveHalfLengthMm += radiusA * 2.0 * 304.8;

                double splitOffset = effectiveHalfLengthMm * FeetPerMm;
                double expandFt    = isAngledMulti
                    ? Math.Abs(dropZFinal) / Math.Tan(pair.AngleDegrees * Math.PI / 180.0)
                    : 0.0;

                // For angled mode: splitOffset = flatHalf + expandFt
                // so the flat section equals effectiveHalfLengthMm (or at least 2 diameters).
                if (isAngledMulti)
                {
                    double minFlatHalfFt = 2.0 * radiusA; // at least 1 diameter flat each side
                    double flatHalfFt    = Math.Max(splitOffset, minFlatHalfFt);
                    splitOffset = flatHalfFt + expandFt;
                    DebugLogger.Log($"[MULTI-RESOLVER] flatHalf={flatHalfFt * 304.8:F1}mm + expand={expandFt * 304.8:F1}mm → splitOffset={splitOffset * 304.8:F1}mm");
                }

                double minSegmentFt = 0.15;

                double t1split = tMin - splitOffset;
                double t2split = tMax + splitOffset;

                t1split = Math.Max(minSegmentFt, Math.Min(t1split, totalLen - minSegmentFt));
                t2split = Math.Max(minSegmentFt, Math.Min(t2split, totalLen - minSegmentFt));

                if (t1split >= t2split - minSegmentFt)
                {
                    result.Message = $"Труба A слишком короткая для обхода нескольких препятствий. " +
                        $"Длина трубы A: {totalLen * 304.8:F0}мм, " +
                        $"требуется минимум {(t2split - t1split + 2 * minSegmentFt) * 304.8:F0}мм.";
                    return result;
                }

                XYZ splitPt1 = pt1 + dirA.Multiply(t1split);
                XYZ splitPt2 = pt1 + dirA.Multiply(t2split);

                // A representative centre point for FindSegmentContainingPoint (middle of bypass zone)
                XYZ ptCenter = pt1 + dirA.Multiply((tMin + tMax) / 2.0);

                XYZ moveVector = new XYZ(0, 0, dropZFinal);
                result.DropMm = dropZFinal * 304.8;
                result.UsedClearanceMm = pair.ClearanceMm;
                result.UsedHalfLengthMm = effectiveHalfLengthMm;

                DebugLogger.Log($"[MULTI-RESOLVER] splitPt1 t={t1split * 304.8:F1}mm  splitPt2 t={t2split * 304.8:F1}mm  middle span={(t2split - t1split) * 304.8:F1}mm");
                DebugLogger.Log($"[MULTI-RESOLVER] drop={dropZFinal * 304.8:F1}mm  expand={expandFt * 304.8:F1}mm");

                // -------------------------------------------------------
                // D. Execute in transaction: split → move → connect
                //    (identical logic to ResolveClash E1..E3)
                // -------------------------------------------------------
                using (var tx = new Transaction(doc, "Multi Clash Resolve: обход труб"))
                {
                    tx.Start();

                    ElementId returnedIdAfterCut1 = SplitCurve(doc, pair.PipeAId, splitPt1);
                    if (returnedIdAfterCut1 == null || returnedIdAfterCut1 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"Не удалось выполнить первый разрез трубы A " +
                            $"(X={splitPt1.X * 304.8:F0}, Y={splitPt1.Y * 304.8:F0}, Z={splitPt1.Z * 304.8:F0}мм).";
                        return result;
                    }

                    ElementId segContainingSplit2 = FindSegmentContainingPoint(doc,
                        new[] { pair.PipeAId, returnedIdAfterCut1 }, splitPt2, minSegmentFt);

                    if (segContainingSplit2 == null || segContainingSplit2 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"После первого разреза не удалось найти сегмент, содержащий вторую точку разреза.";
                        return result;
                    }

                    ElementId leftId = (segContainingSplit2 == pair.PipeAId)
                        ? returnedIdAfterCut1
                        : pair.PipeAId;

                    ElementId returnedIdAfterCut2 = SplitCurve(doc, segContainingSplit2, splitPt2);
                    if (returnedIdAfterCut2 == null || returnedIdAfterCut2 == ElementId.InvalidElementId)
                    {
                        tx.RollBack();
                        result.Message = $"Не удалось выполнить второй разрез трубы A.";
                        return result;
                    }

                    // Middle segment is the one containing the centre of the bypass zone
                    ElementId middleId = FindSegmentContainingPoint(doc,
                        new[] { segContainingSplit2, returnedIdAfterCut2 }, ptCenter, 0);
                    ElementId rightId  = (middleId == segContainingSplit2)
                        ? returnedIdAfterCut2
                        : segContainingSplit2;

                    if (middleId == null || middleId == ElementId.InvalidElementId)
                    {
                        middleId = segContainingSplit2;
                        rightId  = returnedIdAfterCut2;
                        DebugLogger.Log("[MULTI-RESOLVER] Warning: could not determine middle by centre point, using fallback.");
                    }

                    DebugLogger.Log($"[MULTI-RESOLVER] Segments: left={leftId} middle={middleId} right={rightId}");

                    ElementTransformUtils.MoveElement(doc, middleId, moveVector);

                    bool conn1, conn2;
                    if (isAngledMulti && expandFt > 1e-9)
                    {
                        // Angled mode: trim middle inward by expandFt then insert diagonal segments
                        try
                        {
                            MEPCurve middleMep = doc.GetElement(middleId) as MEPCurve;
                            if (middleMep != null)
                            {
                                var lc = middleMep.Location as LocationCurve;
                                Line oldLine = lc?.Curve as Line;
                                if (oldLine != null)
                                {
                                    XYZ mStart = oldLine.GetEndPoint(0);
                                    XYZ mEnd   = oldLine.GetEndPoint(1);
                                    bool startIsLeft = mStart.DistanceTo(splitPt1) < mEnd.DistanceTo(splitPt1);
                                    XYZ leftEnd  = startIsLeft ? mStart : mEnd;
                                    XYZ rightEnd = startIsLeft ? mEnd   : mStart;
                                    XYZ newLeft  = new XYZ(leftEnd.X  + dirA.X * expandFt, leftEnd.Y  + dirA.Y * expandFt, leftEnd.Z);
                                    XYZ newRight = new XYZ(rightEnd.X - dirA.X * expandFt, rightEnd.Y - dirA.Y * expandFt, rightEnd.Z);
                                    lc.Curve = Line.CreateBound(
                                        startIsLeft ? newLeft : newRight,
                                        startIsLeft ? newRight : newLeft);
                                }
                            }
                        }
                        catch (Exception ex) { DebugLogger.Log($"[MULTI-RESOLVER] 45° trim failed: {ex.Message}"); }

                        conn1 = Connect45Fixed(doc, leftId,  middleId, result);
                        conn2 = Connect45Fixed(doc, rightId, middleId, result);
                    }
                    else
                    {
                        conn1 = ConnectOpenEnds(doc, leftId,  middleId, result);
                        conn2 = ConnectOpenEnds(doc, middleId, rightId, result);
                    }

                    if (!conn1 || !conn2)
                        DebugLogger.Log("[MULTI-RESOLVER] Warning: some connections could not be created.");

                    tx.Commit();
                }

                result.Success = true;
                result.Message = "Обход успешно создан. Труба A перестроена под все трубы B.";
                return result;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[MULTI-RESOLVER] ERROR: {ex.Message}\n{ex.StackTrace}");
                result.Message = $"Ошибка: {ex.Message}";
                return result;
            }
        }

        /// <summary>
        /// Compute the vertical drop required to position pipe A's axis below pipe B,
        /// respecting clearance settings. Returns a negative value (downward shift in feet).
        /// </summary>
        private double ComputeDropZ(XYZ pt1, XYZ pt2, double radiusA,
            XYZ pb1, XYZ pb2, double radiusB,
            XYZ ptOnA, ClashPair pair, MepKind kindA = MepKind.RoundDuct)
        {
            double effectiveClearanceMm;
            if (pair.AutoClearance)
            {
                bool isAngledC = Math.Abs(pair.AngleDegrees - 90.0) > 0.5;
                double baseClearanceMm;
                if (kindA == MepKind.RectDuct)
                {
                    if (isAngledC)
                    {
                        // Same formula as ResolveClash: ensure diagonal >= 5×width after elbow insertion.
                        double radiusA_mmC = radiusA * 304.8;
                        double radiusB_mmC = radiusB * 304.8;
                        double minDiagMmC  = 5.0 * (2.0 * radiusA_mmC);
                        double sinAngleC   = Math.Sin(pair.AngleDegrees * Math.PI / 180.0);
                        double minShiftMmC = minDiagMmC * sinAngleC;
                        baseClearanceMm    = Math.Max(50.0, minShiftMmC - radiusA_mmC - radiusB_mmC);
                    }
                    else
                    {
                        baseClearanceMm = Math.Ceiling(Math.Max(radiusA, radiusB) * 9.0 * 304.8);
                    }
                }
                else
                {
                    baseClearanceMm = isAngledC
                        ? 50.0
                        : Math.Ceiling(Math.Max(radiusA, radiusB) * 3.5 * 304.8);
                    // Pipe diameter bonus only in 90° (vertical) mode
                    if (kindA == MepKind.Pipe && !isAngledC)
                        baseClearanceMm += radiusA * 2.0 * 304.8;
                }

                double pipeBAxisZ_check = GetZAtXY(pb1, pb2, ptOnA.X, ptOnA.Y);

                if (!pair.BypassUp)
                {
                    // Bypass down: check if B is above A and fittings might clip it.
                    // Extra clearance only when B sits clearly above A (not just coplanar intersection).
                    double clearanceGapMm2 = (pipeBAxisZ_check - ptOnA.Z) * 304.8;
                    bool   bIsAboveA        = clearanceGapMm2 > (radiusA + radiusB) * 304.8;
                    double extraClearanceMm = 0.0;
                    if (bIsAboveA)
                    {
                        double headroomMm      = (pipeBAxisZ_check - radiusB - ptOnA.Z) * 304.8;
                        double fittingHeightMm = Math.Max(radiusA, radiusB) * 304.8;
                        if (headroomMm < fittingHeightMm + baseClearanceMm)
                            extraClearanceMm = fittingHeightMm + baseClearanceMm - headroomMm;
                    }
                    effectiveClearanceMm = baseClearanceMm + extraClearanceMm;
                }
                else
                {
                    // Bypass up: fitting goes upward away from B, no extra clearance needed.
                    effectiveClearanceMm = baseClearanceMm;
                }
            }
            else
            {
                effectiveClearanceMm = pair.ClearanceMm;
            }

            // For RectDuct angled bypass, enforce minimum clearance regardless of manual/auto mode.
            bool isAngledEnforce = Math.Abs(pair.AngleDegrees - 90.0) > 0.5;
            if (kindA == MepKind.RectDuct && isAngledEnforce)
            {
                double rAmm2    = radiusA * 304.8;
                double rBmm2    = radiusB * 304.8;
                double minDiag2 = 5.0 * (2.0 * rAmm2);
                double sinA2    = Math.Sin(pair.AngleDegrees * Math.PI / 180.0);
                double minClearance2 = Math.Max(50.0, minDiag2 * sinA2 - rAmm2 - rBmm2);
                if (effectiveClearanceMm < minClearance2)
                    effectiveClearanceMm = minClearance2;
            }

            double clearanceFeet = effectiveClearanceMm * FeetPerMm;
            double pipeBAxisZ    = GetZAtXY(pb1, pb2, ptOnA.X, ptOnA.Y);

            if (!pair.BypassUp)
            {
                // A goes below B: target = below B bottom outer edge
                double pipeBBottomOuter = pipeBAxisZ - radiusB;
                double middleAxisZ      = pipeBBottomOuter - clearanceFeet - radiusA;
                return middleAxisZ - ptOnA.Z; // negative = must go down
            }
            else
            {
                // A goes above B: target = above B top outer edge
                double pipeBTopOuter = pipeBAxisZ + radiusB;
                double middleAxisZ   = pipeBTopOuter + clearanceFeet + radiusA;
                return middleAxisZ - ptOnA.Z; // positive = must go up
            }
        }

        // -----------------------------------------------------------------------
        // Geometry helpers
        // -----------------------------------------------------------------------

        /// <summary>
        /// Extract axis endpoints and outer radius from a Pipe or Duct element.
        /// Returns false if element is not a linear MEP curve.
        /// </summary>
        private bool TryGetAxisAndRadius(Element elem, out XYZ start, out XYZ end, out double radius)
        {
            start = end = null;
            radius = 0;

            MEPCurve mep = elem as MEPCurve;
            if (mep == null) return false;

            LocationCurve lc = mep.Location as LocationCurve;
            if (lc == null) return false;

            Line line = lc.Curve as Line;
            if (line == null) return false;

            start = line.GetEndPoint(0);
            end = line.GetEndPoint(1);

            // Diameter parameter — works for round Pipe and round Duct
            Parameter diamParam = mep.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                               ?? mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                               ?? mep.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

            if (diamParam != null && diamParam.HasValue && diamParam.AsDouble() > 1e-9)
            {
                radius = diamParam.AsDouble() / 2.0;
            }
            else if (mep is Duct)
            {
                // Rectangular duct: use height/2 as effective radius (vertical clearance dimension)
                Parameter hParam = mep.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (hParam != null && hParam.HasValue)
                    radius = hParam.AsDouble() / 2.0;
                else
                {
                    // Fallback: use bounding box half-height
                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    radius = bb != null ? (bb.Max.Z - bb.Min.Z) / 2.0 : 0.05;
                }
            }
            else
            {
                // Fallback: use bounding box half-height
                BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                if (bb != null)
                    radius = (bb.Max.Z - bb.Min.Z) / 2.0;
                else
                    radius = 0.05; // 15mm fallback
            }

            return true;
        }

        /// <summary>
        /// Find the point on Line(a1→a2) that is closest to Line(b1→b2).
        /// Works for both intersecting and skew lines.
        /// </summary>
        /// <summary>
        /// Returns the closest point on line A (a1→a2) to line B (b1→b2).
        /// When clamped=true the result is constrained to the segment [a1,a2].
        /// When clamped=false the full infinite line is used (for finding true intersection).
        /// </summary>
        private XYZ GetClosestPointOnLine(XYZ a1, XYZ a2, XYZ b1, XYZ b2,
                                           bool clamped = true)
        {
            XYZ d1 = (a2 - a1);
            XYZ d2 = (b2 - b1);
            XYZ r = a1 - b1;

            double a = d1.DotProduct(d1);
            double e = d2.DotProduct(d2);
            double f = d2.DotProduct(r);

            double t;
            if (a <= 1e-10 && e <= 1e-10)
            {
                t = 0;
            }
            else if (a <= 1e-10)
            {
                t = 0;
            }
            else
            {
                double c = d1.DotProduct(r);
                if (e <= 1e-10)
                {
                    t = -c / a;
                }
                else
                {
                    double b = d1.DotProduct(d2);
                    double denom = a * e - b * b;
                    if (Math.Abs(denom) > 1e-10)
                        t = (b * f - c * e) / denom;
                    else
                        t = 0;
                }
            }

            if (clamped)
                t = Math.Max(0.0, Math.Min(1.0, t));
            return a1 + d1.Multiply(t);
        }

        /// <summary>
        /// Get the Z coordinate of a line (b1→b2) at a given X,Y by projecting.
        /// Used to find PipeB's axis Z at the intersection X,Y.
        /// </summary>
        private double GetZAtXY(XYZ b1, XYZ b2, double x, double y)
        {
            XYZ d = b2 - b1;
            double lenXY = Math.Sqrt(d.X * d.X + d.Y * d.Y);
            if (lenXY < 1e-10)
                return b1.Z; // vertical pipe - return its Z

            // Project (x,y) onto the line's XY direction
            double dx = x - b1.X;
            double dy = y - b1.Y;
            double t = (dx * d.X + dy * d.Y) / (d.X * d.X + d.Y * d.Y);
            t = Math.Max(0.0, Math.Min(1.0, t));
            return b1.Z + t * d.Z;
        }

        /// <summary>
        /// Among the given candidate element IDs, finds the one whose axis curve
        /// contains <paramref name="point"/> within its length (with <paramref name="tolerance"/> margin).
        /// Returns ElementId.InvalidElementId if none qualifies.
        /// </summary>
        private ElementId FindSegmentContainingPoint(Document doc, ElementId[] candidates,
            XYZ point, double tolerance)
        {
            foreach (ElementId id in candidates)
            {
                if (!TryGetAxisAndRadius(doc.GetElement(id), out XYZ s, out XYZ e, out _))
                    continue;

                XYZ dir = (e - s);
                double len = dir.GetLength();
                if (len < 1e-9) continue;
                dir = dir.Normalize();

                // Project point onto segment axis
                double t = (point - s).DotProduct(dir);
                // Distance from the infinite line
                XYZ projected = s + dir.Multiply(t);
                double distFromAxis = point.DistanceTo(projected);

                DebugLogger.Log($"[CLASH-RESOLVER] FindSegment ID={id}: len={len * 304.8:F1}mm, t={t * 304.8:F1}mm, distFromAxis={distFromAxis * 304.8:F1}mm");

                // Accept if t is within [tolerance, len-tolerance] and point is near the axis
                if (t >= tolerance - 1e-6 && t <= len - tolerance + 1e-6 && distFromAxis < 0.01)
                    return id;
            }
            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Split a MEP curve element at the given point using MechanicalUtils or PlumbingUtils.
        /// Returns the ElementId of the newly created second segment (or InvalidElementId on failure).
        /// </summary>
        private ElementId SplitCurve(Document doc, ElementId elemId, XYZ splitPoint)
        {
            try
            {
                Element elem = doc.GetElement(elemId);

                if (elem is Pipe)
                    return PlumbingUtils.BreakCurve(doc, elemId, splitPoint);

                if (elem is Duct)
                    return MechanicalUtils.BreakCurve(doc, elemId, splitPoint);

                // FlexPipe / FlexDuct - not supported
                DebugLogger.Log($"[CLASH-RESOLVER] SplitCurve: unsupported element type {elem.GetType().Name}");
                return ElementId.InvalidElementId;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] SplitCurve error: {ex.Message}");
                return ElementId.InvalidElementId;
            }
        }

        /// <summary>
        /// 45° mode connection:
        /// After the middle segment has been lowered vertically, the gap between
        /// the left (or right) segment and the middle is purely vertical (dxy=0, dz>0).
        /// We bridge this gap with an explicit diagonal segment whose endpoints are:
        ///   ptDiagTop = connA.Origin  (on the horizontal segment, original height)
        ///   ptDiagBot = connB.Origin  (on the lowered middle segment)
        /// These two points differ both in the axis direction and in Z, giving a
        /// diagonal duct. We then call NewElbowFitting at each joint so Revit
        /// inserts the correct angle fitting (45° when dxy≈dz).
        /// </summary>
        private bool Connect45Diagonal(Document doc, ElementId elemAId, ElementId elemBId,
            ClashResolveResult result)
        {
            try
            {
                MEPCurve mepA = doc.GetElement(elemAId) as MEPCurve;
                MEPCurve mepB = doc.GetElement(elemBId) as MEPCurve;
                if (mepA == null || mepB == null) return false;

                Connector connA = GetOpenEndClosestTo(mepA, GetCenter(mepB));
                Connector connB = GetOpenEndClosestTo(mepB, GetCenter(mepA));

                if (connA == null || connB == null)
                {
                    DebugLogger.Log("[CLASH-RESOLVER] Connect45Diagonal: could not find open connectors.");
                    return false;
                }

                XYZ ptA = connA.Origin;  // on horizontal segment (original Z)
                XYZ ptB = connB.Origin;  // on lowered middle segment

                double dxyMm = Math.Sqrt(
                    Math.Pow((ptA.X - ptB.X) * 304.8, 2) +
                    Math.Pow((ptA.Y - ptB.Y) * 304.8, 2));
                double dzMm  = Math.Abs((ptA.Z - ptB.Z) * 304.8);

                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal: dxy={dxyMm:F1}mm dz={dzMm:F1}mm");

                // Case A: already coincident — direct connect
                if (dxyMm < 2.0 && dzMm < 2.0)
                {
                    try { connA.ConnectTo(connB); return true; }
                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal ConnectTo: {ex.Message}"); }
                }

                // The gap is vertical (dxy=0) because the middle moved straight down.
                // Build a diagonal transition point:
                //   ptDiagBot = ptA shifted along dirA (horizontal) by dzMm,
                //               then down to ptB.Z
                // This gives dxy_diag = dzMm and dz_diag = dzMm → exactly 45°.
                //
                // dirA of the horizontal segment:
                if (!TryGetAxisAndRadius(mepA, out XYZ axStart, out XYZ axEnd, out _)) return false;
                XYZ dirA = (axEnd - axStart).Normalize();

                // Determine sign: the diagonal should go TOWARD the middle segment
                // (i.e. away from the far end of the horizontal segment)
                XYZ midCenter = GetCenter(mepB);
                // Project midCenter onto dirA relative to ptA
                double signDot = (midCenter - ptA).DotProduct(dirA);
                // If signDot < 0 we need to go in -dirA direction
                XYZ diagDir = signDot >= 0 ? dirA : dirA.Negate();

                double shiftFt = Math.Abs(ptA.Z - ptB.Z);  // same magnitude as dz
                XYZ ptDiag = new XYZ(
                    ptA.X + diagDir.X * shiftFt,
                    ptA.Y + diagDir.Y * shiftFt,
                    ptB.Z);  // drop to the same Z as the middle connector

                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal: ptDiag={ptDiag.X*304.8:F0},{ptDiag.Y*304.8:F0},{ptDiag.Z*304.8:F0}");

                // Create the diagonal segment ptA → ptDiag
                ElementId diagId = CreateMepSegmentByPoints(doc, mepA, ptA, ptDiag);
                if (diagId == null || diagId == ElementId.InvalidElementId)
                {
                    DebugLogger.Log("[CLASH-RESOLVER] Connect45Diagonal: CreateMepSegmentByPoints failed.");
                    return false;
                }
                result.CreatedElements.Add(diagId);

                MEPCurve diagSeg = doc.GetElement(diagId) as MEPCurve;
                if (diagSeg == null) return false;

                Connector diagConnTop = GetOpenEndClosestTo(diagSeg, ptA);
                Connector diagConnBot = GetOpenEndClosestTo(diagSeg, ptDiag);

                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal: diagonal segment {diagId} created");

                // Elbow at top junction: horizontal ↔ diagonal
                bool ok = true;
                try
                {
                    Element elbowTop = doc.Create.NewElbowFitting(connA, diagConnTop);
                    if (elbowTop != null)
                    {
                        result.CreatedElements.Add(elbowTop.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal: top elbow {elbowTop.Id} inserted.");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal top elbow failed: {ex.Message}");
                    ok = false;
                }

                // Elbow at bottom junction: diagonal ↔ middle
                // ptDiag is at the same Z as ptB but offset in XY.
                // Connect diagConnBot to connB via elbow.
                try
                {
                    Element elbowBot = doc.Create.NewElbowFitting(diagConnBot, connB);
                    if (elbowBot != null)
                    {
                        result.CreatedElements.Add(elbowBot.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal: bottom elbow {elbowBot.Id} inserted.");
                    }
                }
                catch (Exception ex)
                {
                    DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal bottom elbow failed: {ex.Message}");
                    ok = false;
                }

                return ok;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Diagonal error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Connects a horizontal segment (elemHoriz) to the trimmed middle segment (elemMiddle)
        /// for 45° bypass mode.
        ///
        /// After trim the gap is: dxy = expandFt = |dropZ|, dz = |dropZ| — exact 45°.
        /// We create a diagonal segment from connHoriz → connMiddle, then:
        ///   • elbow(connHoriz, diagTop) — 45° fitting at the horizontal-to-diagonal junction
        ///   • elbow(diagBot, connMiddle) — 45° fitting at the diagonal-to-middle junction
        /// </summary>
        private bool Connect45Fixed(Document doc, ElementId elemHorizId, ElementId elemMiddleId,
            ClashResolveResult result)
        {
            try
            {
                MEPCurve mepH = doc.GetElement(elemHorizId)  as MEPCurve;
                MEPCurve mepM = doc.GetElement(elemMiddleId) as MEPCurve;
                if (mepH == null || mepM == null) return false;

                // Get open connectors facing each other
                Connector connH = GetOpenEndClosestTo(mepH, GetCenter(mepM));
                Connector connM = GetOpenEndClosestTo(mepM, GetCenter(mepH));
                if (connH == null || connM == null)
                {
                    DebugLogger.Log("[CLASH-RESOLVER] Connect45Fixed: could not find open connectors.");
                    return false;
                }

                XYZ ptH = connH.Origin;  // on horizontal segment @ Z_orig
                XYZ ptM = connM.Origin;  // on trimmed middle  @ Z_low, shifted in XY

                double dxyMm = Math.Sqrt(Math.Pow((ptH.X-ptM.X)*304.8,2)+Math.Pow((ptH.Y-ptM.Y)*304.8,2));
                double dzMm  = Math.Abs((ptH.Z-ptM.Z)*304.8);
                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: ptH={ptH.X*304.8:F0},{ptH.Y*304.8:F0},{ptH.Z*304.8:F0}  ptM={ptM.X*304.8:F0},{ptM.Y*304.8:F0},{ptM.Z*304.8:F0}  dxy={dxyMm:F1}mm dz={dzMm:F1}mm");

                // Case A: already coincident
                if (dxyMm < 2.0 && dzMm < 2.0)
                {
                    try { connH.ConnectTo(connM); return true; }
                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed ConnectTo: {ex.Message}"); }
                }

                // Create diagonal segment ptH → ptM (dxy=dz → 45°)
                ElementId diagId = CreateMepSegmentByPoints(doc, mepH, ptH, ptM);
                if (diagId == null || diagId == ElementId.InvalidElementId)
                {
                    DebugLogger.Log("[CLASH-RESOLVER] Connect45Fixed: CreateMepSegmentByPoints failed.");
                    return false;
                }
                result.CreatedElements.Add(diagId);
                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: diagonal segment {diagId} created ({dxyMm:F1}mm horizontal, {dzMm:F1}mm vertical)");

                MEPCurve diagSeg = doc.GetElement(diagId) as MEPCurve;
                if (diagSeg == null) return false;

                Connector diagTop = GetOpenEndClosestTo(diagSeg, ptH);
                Connector diagBot = GetOpenEndClosestTo(diagSeg, ptM);

                bool ok = true;
                // Elbow at top: horizontal ↔ diagonal
                try
                {
                    // Log horizontal segment endpoints before elbow insertion
                    if (TryGetAxisAndRadius(mepH, out XYZ hS0, out XYZ hE0, out _))
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: horiz BEFORE top-elbow: ({hS0.X*304.8:F1},{hS0.Y*304.8:F1},{hS0.Z*304.8:F1})->({hE0.X*304.8:F1},{hE0.Y*304.8:F1},{hE0.Z*304.8:F1})");
                    if (TryGetAxisAndRadius(diagSeg, out XYZ dS0, out XYZ dE0, out _))
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: diag  BEFORE top-elbow: ({dS0.X*304.8:F1},{dS0.Y*304.8:F1},{dS0.Z*304.8:F1})->({dE0.X*304.8:F1},{dE0.Y*304.8:F1},{dE0.Z*304.8:F1})");

                    Element elbowTop = doc.Create.NewElbowFitting(connH, diagTop);
                    if (elbowTop != null)
                    {
                        result.CreatedElements.Add(elbowTop.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: top elbow {elbowTop.Id}");
                        // Log positions after top-elbow insertion to detect repositioning
                        if (TryGetAxisAndRadius(mepH, out XYZ hS1, out XYZ hE1, out _))
                            DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: horiz AFTER  top-elbow: ({hS1.X*304.8:F1},{hS1.Y*304.8:F1},{hS1.Z*304.8:F1})->({hE1.X*304.8:F1},{hE1.Y*304.8:F1},{hE1.Z*304.8:F1})");
                        if (TryGetAxisAndRadius(diagSeg, out XYZ dS1, out XYZ dE1, out _))
                            DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: diag  AFTER  top-elbow: ({dS1.X*304.8:F1},{dS1.Y*304.8:F1},{dS1.Z*304.8:F1})->({dE1.X*304.8:F1},{dE1.Y*304.8:F1},{dE1.Z*304.8:F1})");
                    }
                }
                catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed top elbow failed: {ex.Message}"); ok = false; }

                // Elbow at bottom: diagonal ↔ middle
                try
                {
                    // Log middle segment endpoints before bottom elbow insertion
                    if (TryGetAxisAndRadius(mepM, out XYZ mS0, out XYZ mE0, out _))
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: mid   BEFORE bot-elbow: ({mS0.X*304.8:F1},{mS0.Y*304.8:F1},{mS0.Z*304.8:F1})->({mE0.X*304.8:F1},{mE0.Y*304.8:F1},{mE0.Z*304.8:F1})");
                    if (TryGetAxisAndRadius(diagSeg, out XYZ dS2, out XYZ dE2, out _))
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: diag  BEFORE bot-elbow: ({dS2.X*304.8:F1},{dS2.Y*304.8:F1},{dS2.Z*304.8:F1})->({dE2.X*304.8:F1},{dE2.Y*304.8:F1},{dE2.Z*304.8:F1})");

                    // Re-fetch connectors after top-elbow repositioning
                    diagBot = GetOpenEndClosestTo(diagSeg, ptM);
                    connM   = GetOpenEndClosestTo(mepM, GetCenter(mepH));

                    Element elbowBot = doc.Create.NewElbowFitting(diagBot, connM);
                    if (elbowBot != null)
                    {
                        result.CreatedElements.Add(elbowBot.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: bottom elbow {elbowBot.Id}");
                        if (TryGetAxisAndRadius(mepM, out XYZ mS1, out XYZ mE1, out _))
                            DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: mid   AFTER  bot-elbow: ({mS1.X*304.8:F1},{mS1.Y*304.8:F1},{mS1.Z*304.8:F1})->({mE1.X*304.8:F1},{mE1.Y*304.8:F1},{mE1.Z*304.8:F1})");
                        if (TryGetAxisAndRadius(diagSeg, out XYZ dS3, out XYZ dE3, out _))
                            DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: diag  AFTER  bot-elbow: ({dS3.X*304.8:F1},{dS3.Y*304.8:F1},{dS3.Z*304.8:F1})->({dE3.X*304.8:F1},{dE3.Y*304.8:F1},{dE3.Z*304.8:F1})");
                    }
                }
                catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed bottom elbow failed: {ex.Message}"); ok = false; }

                return ok;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Connects the open end of elemA to the open end of elemB.
        ///
        /// Case A — connectors coincident (&lt;2mm): ConnectTo directly.
        /// Case B — purely vertical offset (same XY, different Z):
        ///   Create a vertical transition duct ptA->ptB, then insert elbow fittings
        ///   at both ends: NewElbowFitting(connHoriz, connVertTop) and
        ///   NewElbowFitting(connVertBottom, connHoriz).
        /// Case C — general angular: NewElbowFitting(connA, connB) directly.
        /// </summary>
        private bool ConnectOpenEnds(Document doc, ElementId elemAId, ElementId elemBId,
            ClashResolveResult result)
        {
            try
            {
                MEPCurve mepA = doc.GetElement(elemAId) as MEPCurve;
                MEPCurve mepB = doc.GetElement(elemBId) as MEPCurve;
                if (mepA == null || mepB == null) return false;

                Connector connA = GetOpenEndClosestTo(mepA, GetCenter(mepB));
                Connector connB = GetOpenEndClosestTo(mepB, GetCenter(mepA));

                if (connA == null || connB == null)
                {
                    DebugLogger.Log("[CLASH-RESOLVER] ConnectOpenEnds: could not find open connectors.");
                    return false;
                }

                XYZ ptA = connA.Origin;
                XYZ ptB = connB.Origin;
                double distMm = ptA.DistanceTo(ptB) * 304.8;
                DebugLogger.Log($"[CLASH-RESOLVER] ConnectOpenEnds: dist={distMm:F1}mm  ptA={ptA.X*304.8:F0},{ptA.Y*304.8:F0},{ptA.Z*304.8:F0}  ptB={ptB.X*304.8:F0},{ptB.Y*304.8:F0},{ptB.Z*304.8:F0}");

                // Case A: coincident
                if (distMm < 2.0)
                {
                    try
                    {
                        connA.ConnectTo(connB);
                        DebugLogger.Log("[CLASH-RESOLVER] Connected via ConnectTo (coincident).");
                        return true;
                    }
                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] ConnectTo failed: {ex.Message}"); }
                }

                // Determine geometry of the gap between connectors:
                // - sameAxis: purely vertical (same XY, different Z) → Case B: vertical duct + 2×90° elbows
                // - diagonal45: dxy ≈ dz (within 5mm) → Case D: direct NewElbowFitting →45°
                // - other: Case C: direct NewElbowFitting
                double dxyMm = Math.Sqrt(
                    Math.Pow((ptA.X - ptB.X) * 304.8, 2) +
                    Math.Pow((ptA.Y - ptB.Y) * 304.8, 2));
                double dzMm = Math.Abs((ptA.Z - ptB.Z) * 304.8);
                bool sameAxis  = dxyMm < 1.0 && dzMm > 1.0;
                bool diagonal45 = Math.Abs(dxyMm - dzMm) < 5.0 && dxyMm > 1.0 && dzMm > 1.0;

                DebugLogger.Log($"[CLASH-RESOLVER] ConnectOpenEnds: dxy={dxyMm:F1}mm dz={dzMm:F1}mm sameAxis={sameAxis} diag45={diagonal45}");

                if (sameAxis)
                {
                    // Case B: horizontal end -> vertical duct -> horizontal end
                    // Create vertical segment, then elbow at each junction
                    try
                    {
                        ElementId vertId = CreateMepSegmentByPoints(doc, mepA, ptA, ptB);
                        if (vertId == null || vertId == ElementId.InvalidElementId)
                        {
                            DebugLogger.Log("[CLASH-RESOLVER] Case B: CreateMepSegment returned invalid id.");
                        }
                        else
                        {
                            result.CreatedElements.Add(vertId);
                            DebugLogger.Log($"[CLASH-RESOLVER] Case B: vertical segment {vertId} created.");

                            MEPCurve vertSeg = doc.GetElement(vertId) as MEPCurve;
                            if (vertSeg != null)
                            {
                                Connector vertConnA = GetOpenEndClosestTo(vertSeg, ptA);
                                Connector vertConnB = GetOpenEndClosestTo(vertSeg, ptB);

                                if (vertConnA != null)
                                {
                                    try
                                    {
                                        Element elbowA = doc.Create.NewElbowFitting(connA, vertConnA);
                                        if (elbowA != null)
                                        {
                                            result.CreatedElements.Add(elbowA.Id);
                                            DebugLogger.Log($"[CLASH-RESOLVER] ElbowA inserted: {elbowA.Id}");
                                        }
                                    }
                                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] ElbowA failed: {ex.Message}"); }
                                }

                                if (vertConnB != null)
                                {
                                    try
                                    {
                                        Element elbowB = doc.Create.NewElbowFitting(vertConnB, connB);
                                        if (elbowB != null)
                                        {
                                            result.CreatedElements.Add(elbowB.Id);
                                            DebugLogger.Log($"[CLASH-RESOLVER] ElbowB inserted: {elbowB.Id}");
                                        }
                                    }
                                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] ElbowB failed: {ex.Message}"); }
                                }

                                return true;
                            }
                        }
                    }
                    catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Case B failed: {ex.Message}"); }
                }

                // Case C/D: general angular or exact 45° diagonal — direct elbow
                // For 45° diagonal (dxy ≈ dz), Revit auto-selects the 45° fitting family.
                // For general angular, it picks whatever fitting fits the angle.
                try
                {
                    Element fitting = doc.Create.NewElbowFitting(connA, connB);
                    if (fitting != null)
                    {
                        result.CreatedElements.Add(fitting.Id);
                        string caseLabel = diagonal45 ? "D (45° diagonal)" : "C (angular)";
                        DebugLogger.Log($"[CLASH-RESOLVER] Case {caseLabel}: elbow fitting {fitting.Id}");
                        return true;
                    }
                }
                catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Case C/D elbow failed: {ex.Message}"); }

                DebugLogger.Log("[CLASH-RESOLVER] ConnectOpenEnds: all methods failed.");
                return false;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] ConnectOpenEnds error: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates a duct/pipe segment from ptA to ptB, copying type, system and diameter
        /// from the source element.
        /// </summary>
        private ElementId CreateMepSegmentByPoints(Document doc, MEPCurve source, XYZ ptA, XYZ ptB)
        {
            // Level
            ElementId levelId = source.LevelId;
            if (levelId == null || levelId == ElementId.InvalidElementId)
            {
                using (var col = new FilteredElementCollector(doc).OfClass(typeof(Level)))
                    levelId = col.FirstElementId();
            }

            if (source is Duct ductSrc)
            {
                ElementId ductTypeId = ductSrc.GetTypeId();

                // Get system type id from duct's own parameter — most reliable method
                ElementId sysTypeId = ElementId.InvalidElementId;

                // Try built-in parameter RBS_DUCT_SYSTEM_TYPE_PARAM
                Parameter sysParam = ductSrc.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM);
                if (sysParam != null && sysParam.HasValue && sysParam.StorageType == StorageType.ElementId)
                    sysTypeId = sysParam.AsElementId();

                // Fallback: MEPSystem
                if (sysTypeId == null || sysTypeId == ElementId.InvalidElementId)
                    sysTypeId = ductSrc.MEPSystem?.GetTypeId() ?? ElementId.InvalidElementId;

                // Last resort: first MechanicalSystemType from document
                if (sysTypeId == null || sysTypeId == ElementId.InvalidElementId)
                {
                    using (var col = new FilteredElementCollector(doc).OfClass(typeof(MechanicalSystemType)))
                        sysTypeId = col.FirstElementId();
                }

                DebugLogger.Log($"[CLASH-RESOLVER] CreateMepSegment duct: typeId={ductTypeId}, sysTypeId={sysTypeId}, levelId={levelId}");

                if (sysTypeId == null || sysTypeId == ElementId.InvalidElementId)
                    throw new InvalidOperationException("No MechanicalSystemType found in document.");

                Duct d = Duct.Create(doc, sysTypeId, ductTypeId, levelId, ptA, ptB);
                if (d == null) return ElementId.InvalidElementId;

                // For rectangular ducts copy width+height; for round ducts copy diameter
                if (GetMepKind(ductSrc) == MepKind.RectDuct)
                    CopyDuctShape(ductSrc, d);
                else
                    CopyDiameter(ductSrc, d);

                return d.Id;
            }

            if (source is Pipe pipeSrc)
            {
                ElementId pipeTypeId = pipeSrc.GetTypeId();

                ElementId sysTypeId = ElementId.InvalidElementId;

                Parameter sysParam = pipeSrc.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                if (sysParam != null && sysParam.HasValue && sysParam.StorageType == StorageType.ElementId)
                    sysTypeId = sysParam.AsElementId();

                if (sysTypeId == null || sysTypeId == ElementId.InvalidElementId)
                    sysTypeId = pipeSrc.MEPSystem?.GetTypeId() ?? ElementId.InvalidElementId;

                if (sysTypeId == null || sysTypeId == ElementId.InvalidElementId)
                {
                    using (var col = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)))
                        sysTypeId = col.FirstElementId();
                }

                DebugLogger.Log($"[CLASH-RESOLVER] CreateMepSegment pipe: typeId={pipeTypeId}, sysTypeId={sysTypeId}, levelId={levelId}");

                Pipe p = Pipe.Create(doc, sysTypeId, pipeTypeId, levelId, ptA, ptB);
                if (p == null) return ElementId.InvalidElementId;

                CopyDiameter(pipeSrc, p);

                return p.Id;
            }

            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Copies diameter (outer diameter / diameter) from source MEPCurve to target.
        /// </summary>
        private void CopyDiameter(MEPCurve source, MEPCurve target)
        {
            try
            {
                // Try outer diameter first (ducts)
                BuiltInParameter[] diamParams = new[]
                {
                    BuiltInParameter.RBS_PIPE_OUTER_DIAMETER,
                    BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
                    BuiltInParameter.RBS_PIPE_DIAMETER_PARAM
                };

                foreach (var bip in diamParams)
                {
                    Parameter srcP = source.get_Parameter(bip);
                    if (srcP == null || !srcP.HasValue) continue;
                    double val = srcP.AsDouble();

                    Parameter dstP = target.get_Parameter(bip);
                    if (dstP != null && !dstP.IsReadOnly)
                    {
                        dstP.Set(val);
                        DebugLogger.Log($"[CLASH-RESOLVER] CopyDiameter: {bip}={val * 304.8:F0}mm");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] CopyDiameter failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Copies width and height from a rectangular source duct to a target duct.
        /// </summary>
        private void CopyDuctShape(Duct source, Duct target)
        {
            try
            {
                // Width
                Parameter srcW = source.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                Parameter dstW = target.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM);
                if (srcW != null && dstW != null && !dstW.IsReadOnly)
                {
                    dstW.Set(srcW.AsDouble());
                    DebugLogger.Log($"[CLASH-RESOLVER] CopyDuctShape: width={srcW.AsDouble() * 304.8:F0}mm");
                }
                // Height
                Parameter srcH = source.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                Parameter dstH = target.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM);
                if (srcH != null && dstH != null && !dstH.IsReadOnly)
                {
                    dstH.Set(srcH.AsDouble());
                    DebugLogger.Log($"[CLASH-RESOLVER] CopyDuctShape: height={srcH.AsDouble() * 304.8:F0}mm");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"[CLASH-RESOLVER] CopyDuctShape failed: {ex.Message}");
            }
        }

        private Connector GetOpenEndClosestTo(MEPCurve mep, XYZ target)
        {
            Connector best = null;
            double bestDist = double.MaxValue;

            foreach (Connector c in mep.ConnectorManager.Connectors)
            {
                if (c.IsConnected) continue;
                double d = c.Origin.DistanceTo(target);
                if (d < bestDist)
                {
                    bestDist = d;
                    best = c;
                }
            }
            return best;
        }

        private XYZ GetCenter(MEPCurve mep)
        {
            LocationCurve lc = mep.Location as LocationCurve;
            if (lc == null) return XYZ.Zero;
            return (lc.Curve.GetEndPoint(0) + lc.Curve.GetEndPoint(1)) / 2.0;
        }
    }
}
