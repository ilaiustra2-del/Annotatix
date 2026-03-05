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
        /// When true, use 45° diagonal transitions instead of 90° vertical ducts.
        /// The split points are shifted inward by |dropZ| so the junction segments
        /// run at exactly 45° to the horizontal axis.
        /// </summary>
        public bool UseAngle45 { get; set; } = false;

        /// <summary>
        /// When true, HalfLengthMm is ignored and computed automatically:
        /// 90° mode: 2.5 × max(radiusA, radiusB)
        /// 45° mode: 3.0 × max(radiusA, radiusB)
        /// </summary>
        public bool AutoHalfLength { get; set; } = false;
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

                // -------------------------------------------------------
                // B. Find intersection point on axis of A (closest approach)
                // -------------------------------------------------------
                XYZ ptOnA = GetClosestPointOnLine(pt1, pt2, pt3, pt4);
                result.IntersectionPoint = ptOnA;
                DebugLogger.Log($"[CLASH-RESOLVER] Intersection point on A: {ptOnA}");

                // -------------------------------------------------------
                // C. Calculate split points along A's axis
                // -------------------------------------------------------
                XYZ dirA = (pt2 - pt1).Normalize();

                // splitOffset = user-defined half-length (default 300 mm)
                // splitPt1 = intersection - halfLength, splitPt2 = intersection + halfLength
                double minSegmentFt = 0.15; // ~46 mm — safe margin from pipe endpoints for BreakCurve

                // Auto half-length:
                // 90°: 2.5 × max(R_A, R_B)
                // 45°: 3.0 × max(R_A, R_B)
                double effectiveHalfLengthMm = pair.AutoHalfLength
                    ? Math.Ceiling(Math.Max(radiusA, radiusB) * (pair.UseAngle45 ? 3.0 : 2.5) * 304.8)
                    : pair.HalfLengthMm;
                if (pair.AutoHalfLength)
                    DebugLogger.Log($"[CLASH-RESOLVER] AutoHalfLength ({(pair.UseAngle45 ? "45°" : "90°")}): → halfLength={effectiveHalfLengthMm:F0}mm");

                double splitOffset = effectiveHalfLengthMm * FeetPerMm; // convert mm → feet
                result.UsedHalfLengthMm = effectiveHalfLengthMm;

                // -------------------------------------------------------
                // D. Calculate vertical drop for middle segment
                // -------------------------------------------------------
                double effectiveClearanceMm;
                if (pair.AutoClearance)
                {
                    // Base clearance between the bottom of B and the top of A
                    double baseClearanceMm = pair.UseAngle45
                        ? 50.0
                        : Math.Ceiling(Math.Max(radiusA, radiusB) * 3.5 * 304.8);

                    double pipeBAxisZ_check = GetZAtXY(pt3, pt4, ptOnA.X, ptOnA.Y);
                    bool   bIsAboveA        = pipeBAxisZ_check > ptOnA.Z;

                    // When B is well above A the diagonal/vertical sections gain extra
                    // height naturally. But when B is only slightly above A (close to
                    // ptOnA.Z) the fittings physically overlap B — add 1× diameter extra.
                    double extraClearanceMm = 0.0;
                    if (bIsAboveA)
                    {
                        // Available head-room = (pipeBAxisZ - radiusB) - ptOnA.Z  (in mm)
                        double headroomMm      = (pipeBAxisZ_check - radiusB - ptOnA.Z) * 304.8;
                        double fittingHeightMm = Math.Max(radiusA, radiusB) * 304.8; // 1× max-radius extra
                        if (headroomMm < fittingHeightMm + baseClearanceMm)
                        {
                            extraClearanceMm = fittingHeightMm + baseClearanceMm - headroomMm;
                            DebugLogger.Log($"[CLASH-RESOLVER] B is above A — adding extra clearance {extraClearanceMm:F0}mm for fittings");
                        }
                    }

                    effectiveClearanceMm = baseClearanceMm + extraClearanceMm;
                    DebugLogger.Log($"[CLASH-RESOLVER] AutoClearance ({(pair.UseAngle45 ? "45°" : "90°")}): base={baseClearanceMm:F0}mm extra={extraClearanceMm:F0}mm total={effectiveClearanceMm:F0}mm");
                }
                else
                {
                    effectiveClearanceMm = pair.ClearanceMm;
                }
                result.UsedClearanceMm = effectiveClearanceMm;

                double clearanceFeet = effectiveClearanceMm * FeetPerMm;

                double pipeBAxisZ        = GetZAtXY(pt3, pt4, ptOnA.X, ptOnA.Y);
                double pipeBBottomOuterZ = pipeBAxisZ - radiusB;
                double middleTopZ        = pipeBBottomOuterZ - clearanceFeet;
                double middleAxisZ       = middleTopZ - radiusA;
                double currentMiddleZ    = ptOnA.Z;
                double dropZ             = middleAxisZ - currentMiddleZ;
                XYZ moveVector           = new XYZ(0, 0, dropZ);

                result.DropMm = dropZ * 304.8;

                DebugLogger.Log($"[CLASH-RESOLVER] PipeB axis Z at intersection: {pipeBAxisZ * 304.8:F1}mm");
                DebugLogger.Log($"[CLASH-RESOLVER] PipeB bottom outer Z: {pipeBBottomOuterZ * 304.8:F1}mm");
                DebugLogger.Log($"[CLASH-RESOLVER] Middle segment drop: {dropZ * 304.8:F1}mm");

                if (dropZ >= 0)
                {
                    result.Message = "Труба A уже находится ниже трубы B. Обход не требуется.";
                    return result;
                }

                // -------------------------------------------------------
                // C2. Compute final split offsets.
                // 45° mode: expand split points outward by |dropZ| (=expandFt) so
                // the middle segment is longer. After lowering, we will trim it
                // inward by expandFt on each side via LocationCurve, creating a
                // gap of dxy=expandFt, dz=expandFt between each horizontal end
                // and the trimmed middle — giving natural 45° for ConnectOpenEnds.
                // 90° mode: no expansion.
                // -------------------------------------------------------
                double expandFt = pair.UseAngle45 ? Math.Abs(dropZ) : 0.0;
                if (pair.UseAngle45)
                    DebugLogger.Log($"[CLASH-RESOLVER] 45° mode: expandFt={expandFt * 304.8:F1}mm each side");

                // Parametric positions along A's axis (from pt1)
                double totalLength = pt1.DistanceTo(pt2);
                double tCenter     = (ptOnA - pt1).DotProduct(dirA);
                double t1 = tCenter - splitOffset - expandFt;
                double t2 = tCenter + splitOffset + expandFt;

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
                    if (pair.UseAngle45 && expandFt > 1e-9)
                    {
                        // -------------------------------------------------------
                        // E3b. 45° mode: trim middle inward by expandFt on each side,
                        // then create diagonal segments + elbows via Connect45Fixed.
                        //
                        // After lowering, middle spans splitPt1..splitPt2 @ Z_low.
                        // We trim it to (splitPt1+expand)...(splitPt2-expand) @ Z_low.
                        // Gap left↔middle: ptA=(splitPt1@Z_orig), ptB=(splitPt1+expand@Z_low)
                        //   → dxy=expand=|dropZ|, dz=|dropZ| → exact 45°.
                        // We create a diagonal segment ptA→ptB, then:
                        //   • elbow(connA_horiz, diagConnTop)  — 45° at top
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
                                    DebugLogger.Log($"[CLASH-RESOLVER] 45° middle trimmed: left={newLeft.X*304.8:F0},{newLeft.Y*304.8:F0},{newLeft.Z*304.8:F0}  right={newRight.X*304.8:F0},{newRight.Y*304.8:F0},{newRight.Z*304.8:F0}");
                                }
                            }
                        }
                        catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] 45° trim failed: {ex.Message}"); }

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

            // Diameter parameter - works for both Pipe and Duct
            Parameter diamParam = mep.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                               ?? mep.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                               ?? mep.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

            if (diamParam != null && diamParam.HasValue)
            {
                radius = diamParam.AsDouble() / 2.0;
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
        private XYZ GetClosestPointOnLine(XYZ a1, XYZ a2, XYZ b1, XYZ b2)
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

            // Clamp t to [0, 1] so point stays within segment A
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
                    Element elbowTop = doc.Create.NewElbowFitting(connH, diagTop);
                    if (elbowTop != null)
                    {
                        result.CreatedElements.Add(elbowTop.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: top elbow {elbowTop.Id}");
                    }
                }
                catch (Exception ex) { DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed top elbow failed: {ex.Message}"); ok = false; }

                // Elbow at bottom: diagonal ↔ middle
                try
                {
                    Element elbowBot = doc.Create.NewElbowFitting(diagBot, connM);
                    if (elbowBot != null)
                    {
                        result.CreatedElements.Add(elbowBot.Id);
                        DebugLogger.Log($"[CLASH-RESOLVER] Connect45Fixed: bottom elbow {elbowBot.Id}");
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

                // Copy diameter from source
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
