using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class Formatting
    {
        public static void RemoveBorders(Range range)
        {
            range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlLineStyleNone;
            range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlLineStyleNone;
        }

        public static void UnderlineBorder(Range range)
        {
            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            range.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
        }

        public static void TableHeaderRow(Range header)
        {
            header.Interior.Color = Color.White;
            header.Font.Bold = true;
            header.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            header.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

            header.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            header.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

            header.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            header.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
        }

        public static void TableHeaderCell(Range headerCell)
        {
        }

        public static void TableDataCell(Range dataCell)
        {
        }

        public static void TableDataRow(Range rowRange, bool isEven)
        {
            rowRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rowRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            rowRange.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

            rowRange.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

            rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            rowRange.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            rowRange.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;

            rowRange.Interior.Color = isEven ? Color.LightYellow : Color.White;
        }

        public static void VerticalTableHeaderCell(Range fieldNameCell)
        {
            fieldNameCell.Font.Bold = true;
            fieldNameCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            fieldNameCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
        }

        public static void VerticalTableDataCell(Range dataCell)
        {
            dataCell.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        public static void VerticalTableRow(Range row, bool isEven)
        {
            row.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            row.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            row.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            row.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            row.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            row.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            row.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            row.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
            row.Interior.Color = isEven ? Color.LightYellow : Color.White;
        }

        public static void ActivateErrorCell(Range cell)
        {
            cell.Interior.Color = Color.Red;
            cell.Font.Bold = true;
        }
        
        public static void DeactivateErrorCell(Range cell)
        {
            cell.Interior.Color = Color.White;
            cell.Font.Bold = true;
        }
    }
}