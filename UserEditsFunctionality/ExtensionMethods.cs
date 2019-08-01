using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

namespace UserEditsFunctionality
{
    public static class ExtensionMethods
    {
        public static string GetDescription(this Enum constant)
        {
            Type enumType = constant.GetType();
            var memberInfo = enumType.GetMember(enumType.GetEnumName(constant));
            var description = memberInfo[0].GetCustomAttribute<DescriptionAttribute>();

            return description?.Description;
        }

        public static Range Cell(this Range range, int columnIndex, object rowIndex = null)
        {
            return (Range)range.Cells[rowIndex ?? Type.Missing, columnIndex];
        }

        public static Range LastNonEmptyCell(this Range range)
        {
            Worksheet worksheet = range.Worksheet;
            return ((Range)worksheet.Cells[range.Row, worksheet.Columns.Count]).End[XlDirection.xlToLeft];
        }

        public static List<Range> ToList(this Range source)
        {
            List<Range> ranges = new List<Range>();
            foreach (Range range in source.Rows)
                ranges.Add(range);

            return ranges;
        }
    }
}
