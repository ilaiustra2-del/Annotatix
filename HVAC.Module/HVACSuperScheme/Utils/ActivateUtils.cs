using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class ActivateUtils
    {
        public static void ActivateFamilySymbol(Document doc, ElementId familySymbolId)
        {
            FamilySymbol familySymbol = doc.GetElement(familySymbolId) as FamilySymbol;
            using var t = new Transaction(doc);
            t.Start("Активация типоразмера");
            familySymbol.Activate();
            doc.Regenerate();
            t.Commit();  
        }
        public static bool IsSymbolActive(Document doc, ElementId familySymbolId)
        {
            FamilySymbol familySymbol = doc.GetElement(familySymbolId) as FamilySymbol;
            return familySymbol.IsActive;
        }
    }
}
