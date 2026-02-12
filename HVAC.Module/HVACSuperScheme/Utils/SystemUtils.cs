using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public class SystemUtils
    {
        public SystemUtils()
        {

        }
        public List<List<string>> GetKeysForSpaces(List<List<string>> systemNamesSpaces)
        {
            var copySystemNamesSpaces = systemNamesSpaces.ToList();
            var visual = systemNamesSpaces.Select(s => string.Join("/", s)).ToList();
            List<List<string>> globalKeys = new List<List<string>>();
            do
            {
                List<List<string>> systemNamesSpacesAlreadyInGroupToDelete = new List<List<string>>();
                int countKeysInCurrentGroup = 0;
                List<string> keysForCurrentGroup = copySystemNamesSpaces.First();
                copySystemNamesSpaces.Remove(copySystemNamesSpaces.First());
                do
                {
                    countKeysInCurrentGroup = keysForCurrentGroup.Count;
                    foreach (var systemNamesSpace in copySystemNamesSpaces)
                    {
                        if (keysForCurrentGroup.Any(key => systemNamesSpace.Contains(key)))
                        {
                            systemNamesSpacesAlreadyInGroupToDelete.Add(systemNamesSpace);
                            foreach (var systemName in systemNamesSpace)
                            {
                                keysForCurrentGroup.Add(systemName);
                            }
                        }
                    }
                    foreach (var systemNamesSpaceAlready in systemNamesSpacesAlreadyInGroupToDelete)
                    {
                        if (copySystemNamesSpaces.Contains(systemNamesSpaceAlready))
                        {
                            copySystemNamesSpaces.Remove(systemNamesSpaceAlready);
                        }
                    }
                } while (keysForCurrentGroup.Count != countKeysInCurrentGroup);
                globalKeys.Add(keysForCurrentGroup.Distinct().OrderBy(el => el).ToList());
            } while (copySystemNamesSpaces.Count != 0);
            return globalKeys.OrderBy(list => string.Join("/", list)).ToList();
        }
    }
}
