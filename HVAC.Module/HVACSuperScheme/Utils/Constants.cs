using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HVACSuperScheme.Utils
{
    public static class Constants
    {
        public const double HEIGHT_ANNOTATION = 1000 / 304.8;
        public const double WIDTH_ANNOTATION = 800 / 304.8;
        public const double VERTICAL_DISTANCE_BETWEEN_ROW = 675 / 304.8;
        public const double HORIZONTAL_DISTANCE_BETWEEN_GROUP_IN_ROW = 250 / 304.8;
        public const double VERTICAL_DISTANCE_TO_SYSTEM_NAME_TEXT = 560 / 304.8;

        public const string ANNOTATION_SYMBOL_TYPE_NAME = "Блок";
        public const string ANNOTATION_SYMBOL_FAMILY_NAME = "Блок ver2.0";
        public const string TEXT_NOTE_TYPE_NAME = "ADSK_Основной текст_7";
        public const string DRAFT_VIEW_NAME = "Схема вентиляции";

        public const string PN_VISIBLE_SUPPLY = "Приток";
        public const string PN_VISIBLE_EXHAUST = "Вытяжка";
        public const string PN_VISIBLE_LOCAL_SUCTION = "МО";

        public const string PN_LEADER_LENGTH_SUPPLY = "Длина_выноски П";
        public const string PN_LEADER_LENGTH_LOCAL_SUCTION = "Длина_выноски МО";
        public const string PN_LEADER_LENGTH_EXHAUST = "Длина_выноски В";

        public const string PN_COUNT_LOCAL_SUCTION = "Количество МО";
        public const string PN_NAME = "Имя";
        public const string PN_NUMBER = "Номер";
        public const string PN_LEVEL = "Уровень";
        public const string PN_ADSK_NUMBER = "ADSK_Номер";
        public const string PN_ADSK_NAME = "ADSK_Наименование";

        public const string PN_ADSK_NAME_EXHAUST_SYSTEM = "ADSK_Наименование вытяжной системы";
        public const string PN_ADSK_NAME_SUPPLY_SYSTEM = "ADSK_Наименование приточной системы";
        public const string PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION = "ADSK_Наименование МО1/МО2/МО3/МО4/МО5";
        public const string PN_ADSK_CALCULATED_EXHAUST = "ADSK_Расчетная вытяжка";
        public const string PN_ADSK_CALCULATED_SUPPLY = "ADSK_Расчетный приток";
        public const string PN_ADSK_ROOM_CATEGORY = "ADSK_Категория помещения";
        public const string PN_ADSK_ROOM_TEMPERATURE = "ADSK_Температура в помещении";
        public const string PN_ADSK_UPDATER = "ADSK_Апдейтер";
        public const string PN_CLEAN_CLASS = "Класс чистоты";

        public const string PN_ADSK_AIR_FLOW_EXHAUST = "ADSK_Расход воздуха вытяжной";
        public const string PN_ADSK_AIR_FLOW_SUPPLY = "ADSK_Расход воздуха приточный";
        public const string PN_AIR_FLOW = "Расход воздуха";

        public const string PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE = "Расход_МО1";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE = "Расход_МО2";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE = "Расход_МО3";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE = "Расход_МО4";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE = "Расход_МО5";

        public const string PN_AIR_FLOW_LOCAL_SUCTION_1 = "Расход МО1";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_2 = "Расход МО2";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_3 = "Расход МО3";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_4 = "Расход МО4";
        public const string PN_AIR_FLOW_LOCAL_SUCTION_5 = "Расход МО5";

        public const string KW_ANNOTATION_ID = "AnnotationId";
        public const string KW_SPACE_ID = "SpaceId";
        public const string KW_ANNOTATION_SPACE_LINK= "AnnotationSpaceLink";
        public const string KW_SUPPLY_AIR = "Приточный воздух";
        public const string KW_EXHAUST_AIR = "Отработанный воздух";

        public static readonly List<string> REQUIRED_PARAMETERS_FOR_SPACE = new List<string>()
        {
            PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION,
            PN_ADSK_CALCULATED_EXHAUST,
            PN_ADSK_NAME_EXHAUST_SYSTEM,
            PN_ADSK_CALCULATED_SUPPLY,
            PN_ADSK_NAME_SUPPLY_SYSTEM,
            PN_ADSK_UPDATER,
            PN_ADSK_ROOM_CATEGORY,
            PN_ADSK_ROOM_TEMPERATURE,
            PN_CLEAN_CLASS,
            PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE,
            PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE,
            PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE,
            PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE,
            PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE
        };

        public static readonly List<string> REQUIRED_PARAMETERS_FOR_DUCT_TERMINAL = new List<string>()
            {
                PN_AIR_FLOW,
                PN_ADSK_UPDATER
            };
        public static readonly Dictionary<string, string> MATCH_SPACE_PARAM_AND_ANNOTATION_PARAM = new Dictionary<string, string>()
        {
            { PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION, PN_ADSK_NAME_EXHAUST_SYSTEM_FROM_LOCAL_SUCTION },//string
            { PN_ADSK_CALCULATED_EXHAUST, PN_ADSK_CALCULATED_EXHAUST },// double
            { PN_ADSK_NAME_EXHAUST_SYSTEM, PN_ADSK_NAME_EXHAUST_SYSTEM },//string
            { PN_ADSK_CALCULATED_SUPPLY, PN_ADSK_CALCULATED_SUPPLY },// double
            { PN_ADSK_NAME_SUPPLY_SYSTEM, PN_ADSK_NAME_SUPPLY_SYSTEM },//string
            { PN_NUMBER, PN_ADSK_NUMBER }, //string
            { PN_NAME, PN_ADSK_NAME }, // string
            { PN_ADSK_UPDATER, PN_ADSK_UPDATER }, // bool
            { PN_ADSK_ROOM_CATEGORY, PN_ADSK_ROOM_CATEGORY }, //string
            { PN_ADSK_ROOM_TEMPERATURE, PN_ADSK_ROOM_TEMPERATURE }, // double
            { PN_CLEAN_CLASS, PN_CLEAN_CLASS }, //string
            { PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE, PN_AIR_FLOW_LOCAL_SUCTION_1_WITH_UNDERSCORE }, //double
            { PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE, PN_AIR_FLOW_LOCAL_SUCTION_2_WITH_UNDERSCORE }, //double
            { PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE, PN_AIR_FLOW_LOCAL_SUCTION_3_WITH_UNDERSCORE }, //double
            { PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE, PN_AIR_FLOW_LOCAL_SUCTION_4_WITH_UNDERSCORE }, //double
            { PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE, PN_AIR_FLOW_LOCAL_SUCTION_5_WITH_UNDERSCORE }, //double
        };
        public static readonly Dictionary<string, string> MATCH_SPACE_PARAM_AND_DUCT_TERMINAL_PARAM = new Dictionary<string, string>()
        {
            { PN_ADSK_CALCULATED_EXHAUST, PN_AIR_FLOW },// double
            { PN_ADSK_CALCULATED_SUPPLY, PN_AIR_FLOW },// double
        };
    }
}
