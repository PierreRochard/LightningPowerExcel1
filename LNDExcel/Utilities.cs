using System.Threading;
using System.Windows.Forms;
using Google.Protobuf;
using Google.Protobuf.Collections;
using Google.Protobuf.Reflection;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class Utilities
    {
        public static void MarkAsLoadingTable(Worksheet ws)
        {
            ws.Cells[1, 2].Value2 = "Loading...";
        }

        public static void RemoveLoadingMark(Worksheet ws)
        {
            ws.Cells[1, 2].Value2 = "";
        }

        public static string FormatFieldName(string fieldName)
        {
            return Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(fieldName.Replace("_", " "));
        }

        public static void EnableButton(Worksheet ws, string buttonName, bool enable)
        {
            var worksheet = Globals.Factory.GetVstoObject(ws);
            foreach (Control control in worksheet.Controls)
            {
                if (control.Name == buttonName)
                {
                    control.Enabled = enable;
                }
            }
        }

        public static Microsoft.Office.Tools.Excel.Controls.Button CreateButton(string buttonName, Worksheet ws, Range selection, string buttonText)
        {
            var button = new Microsoft.Office.Tools.Excel.Controls.Button();
            var worksheet = Globals.Factory.GetVstoObject(ws);
            worksheet.Controls.AddControl(button, selection, buttonName);
            button.Text = buttonText;
            return button;
        }

        public static void AssignCellValue<TMessageClass>(TMessageClass newMessage, FieldDescriptor field, string newValue, dynamic dataCell) where TMessageClass : IMessage
        {
            var value = "";

            if (field.IsRepeated && field.FieldType != FieldType.Message)
            {
                var items = (RepeatedField<string>)field.Accessor.GetValue(newMessage);
                for (var i = 0; i < items.Count; i++)
                {
                    value += items[i];
                    if (i < items.Count - 1)
                    {
                        value += ",\n";
                    }
                }
                dataCell.Value2 = value;
            }
            else switch (field.FieldType)
            {
                case FieldType.UInt64:
                    dataCell.NumberFormat = "@";
                    dataCell.Value2 = newValue;
                    break;

                default:
                    //dataCell.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)";
                    dataCell.Value2 = newValue;
                    break;
            }
        }

    }
}