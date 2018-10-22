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
            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            header.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
        }

        public static void TableHeaderCell(Range headerCell)
        {
            headerCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            headerCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            headerCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            headerCell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public static void TableDataCell(Range dataCell)
        {
            dataCell.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            dataCell.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
            dataCell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            dataCell.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
            dataCell.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            dataCell.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        public static void TableDataRow(Range rowRange, bool isEven)
        {
            rowRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            rowRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
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