using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    internal class TerminalUtils
    {
        public static Space GetSpaceByDuctTerminal(Document doc, FamilyInstance ductTerminal)
        {
            return ductTerminal.Space;
        }
        public static (List<FamilyInstance> supplyDuctTerminals, List<FamilyInstance> exhaustDuctTerminals) GroupTerminalByType(
            List<FamilyInstance> ductTerminals)
        {
            List<FamilyInstance> supplyDuctTerminals = ductTerminals
                .Where(IsSupplyTerminal)
                .ToList();
            List<FamilyInstance> exhaustDuctTerminals = ductTerminals
                .Where(IsExhaustTerminal)
                .ToList();
            return (supplyDuctTerminals, exhaustDuctTerminals);
        }
        public static bool IsSupplyTerminal(FamilyInstance instance)
        {
            Parameter systemClassParam = instance.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            return systemClassParam.AsString() == Constants.KW_SUPPLY_AIR;
        }
        public static bool IsExhaustTerminal(FamilyInstance instance)
        {
            Parameter systemClassParam = instance.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);
            return systemClassParam.AsString() == Constants.KW_EXHAUST_AIR;
        }
        public static bool IsDuctTerminal(Element element)
        {
            return element.Category.Id.IntegerValue == (int)BuiltInCategory.OST_DuctTerminal;
        }
    }
}
