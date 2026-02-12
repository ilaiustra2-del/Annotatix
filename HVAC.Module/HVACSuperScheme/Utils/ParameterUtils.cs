using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace HVACSuperScheme.Utils
{
    public static class ParameterUtils
    {
        public static void CleanParameters(FamilyInstance annotationInstance)
        {
            foreach (var parametersPair in Constants.MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM)
            {
                Parameter parameter = annotationInstance.LookupParameter(parametersPair.Value);
                if (parameter == null)
                    throw new CustomException($"Нет параметра {parametersPair.Value} у семейства аннотаций");
                ClearValue(parameter);
            }
        }
        public static void ClearValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Integer:
                    parameter.Set(0);
                    break;

                case StorageType.Double:
                    if (parameter.Definition.GetDataType() == SpecTypeId.HvacTemperature)
                        parameter.Set(273.15);
                    else
                        parameter.Set(0);
                    break;

                case StorageType.String:
                    parameter.Set("");
                    break;

                case StorageType.ElementId:
                case StorageType.None:
                default:
                    throw new CustomException($"Не поддерживаемый тип хранилища параметра у параметра {parameter.Definition.Name}");
            }
        }
        public static object GetParameterValue(Parameter parameter)
        {
            switch (parameter.StorageType)
            {
                case StorageType.Integer:
                    return parameter.AsInteger();

                case StorageType.Double:
                    if (parameter.Definition.GetDataType() == SpecTypeId.HvacTemperature)
                    {
                        if (parameter.HasValue)
                            return parameter.AsDouble();
                        else
                            return 273.15;
                    }
                    else
                    {
                        return parameter.AsDouble();
                    }

                case StorageType.String:
                    return parameter.AsString();

                case StorageType.ElementId:
                    return parameter.AsValueString();

                case StorageType.None:
                default:
                    throw new CustomException($"Не поддерживаемый тип данных имя параметра:{parameter.Definition.Name}");
            }
        }
        public static bool IsValueEmpty(object value)
        {
            return value == null
                || string.IsNullOrWhiteSpace(value.ToString())
                || (int.TryParse(value.ToString(), out int intResult) && intResult == 0);
        }
        public static string NormalizeValue(object value)
        {
            if (value == null)
                return string.Empty;
            return value.ToString();
        }
        public static void SetParameterValueWithEqualCheck(Parameter parameter, object newValue, ElementId id)
        {
            string oldValueNormalize = NormalizeValue(GetParameterValue(parameter));
            string newValueNormalize = NormalizeValue(newValue);
            if (oldValueNormalize == newValueNormalize)
            {
                LoggingUtils.Logging(Warnings.ElementParameterValueNotChanged(parameter, id, oldValueNormalize), "-");
                return;
            }
            else
            {
                LoggingUtils.Logging(Warnings.ElementParameterValueChanged(parameter, id, oldValueNormalize, newValueNormalize), "-");
                SetValue(parameter, newValue);
            }
        }
        public static void SetValue(Parameter parameter, object value)
        {
            string parameterName = parameter.Definition.Name;
            switch (parameter.StorageType)
            {
                case StorageType.Integer:
                    parameter.Set(int.Parse(value.ToString()));
                    break;

                case StorageType.Double:
                    parameter.Set(double.Parse(value.ToString()));
                    break;

                case StorageType.String:
                    parameter.Set(value.ToString());
                    break;

                case StorageType.ElementId:
                case StorageType.None:
                default:
                    throw new CustomException($"Не поддерживаемый тип хранилища параметра у параметра {parameterName}");
            }
        }
    }
}
