using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace KABService.Object
{
    static class BDOEnum
    {
        public enum FileMoveOption { Processed, Archive, Error, Manual}

        public enum LogLevel { Info, Warning, Error }

        public enum MaalerTypeLong
        {
            [Description("Varme")]
            HCA = 0,
            [Description("Koldt vand")]
            CW = 1,
            [Description("Varmt vand")]
            HW = 2
        }


        public static string GetMaalerTypeDescription(string _value)
        {
            MaalerTypeLong maaler = (MaalerTypeLong)Enum.Parse(typeof(MaalerTypeLong), _value);
            return GetAttributeOfType<DescriptionAttribute>(maaler).Description;
        }

        private static T GetAttributeOfType<T>(this Enum enumVal) where T : Attribute
        {
            var type = enumVal.GetType();
            var memInfo = type.GetMember(enumVal.ToString());
            var attributes = memInfo[0].GetCustomAttributes(typeof(T), false);
            return (attributes.Length > 0) ? (T)attributes[0] : null;
        }
    }
}