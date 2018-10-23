﻿using System.Drawing;
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
            
            VerticalBorders(header);

            header.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            header.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;

            header.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            header.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;
        }

        public static void TableDataRow(Range cells, bool isEven)
        {
            cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            cells.VerticalAlignment = XlVAlign.xlVAlignCenter;

            cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            cells.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;

            VerticalBorders(cells);

            cells.Interior.Color = isEven ? Color.LightYellow : Color.White;
            cells.NumberFormat = "0";
        }

        public static void VerticalBorders(Range cells)
        {
            cells.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;

            cells.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;

            cells.Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlThin;
        }

        public static void AllThinBorders(Range cells)
        {
            cells.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
            
            cells.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;

            cells.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            cells.Borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;
            
            VerticalBorders(cells);
        }

        public static void VerticalTable(Range cells)
        {
            AllThinBorders(cells);
        }

        public static void VerticalTableHeaderColumn(Range cells)
        {
            cells.Font.Bold = true;
        }

        public static void VerticalTableDataColumn(Range cells)
        {
            cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            cells.NumberFormat = "0";
        }

        public static void VerticalTableRow(Range cells, bool isEven)
        {
            cells.Interior.Color = isEven ? Color.LightYellow : Color.White;
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