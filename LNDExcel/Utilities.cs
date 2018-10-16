using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace LNDExcel
{
    public class Utilities
    {
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

    }
}