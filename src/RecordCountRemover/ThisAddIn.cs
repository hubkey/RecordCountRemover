using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace RecordCountRemover
{
    public partial class ThisAddIn
    {
        private static Excel.Application _xl;
        private const string PageTag = "Page";
        private static Regex _pageRx = new Regex(@"^Page\s\d+\sof\s\d+\.\sTotal:\s\d+\s.+\.$", RegexOptions.Compiled);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _xl = Application;
            _xl.WorkbookOpen += OnXlOnWorkbookOpen;
        }

        private static void OnXlOnWorkbookOpen(Workbook wb)
        {
            if (LooksLikeCRMExport(wb))
                RemovePageRecords(wb);
        }

        private static void RemovePageRecords(Workbook wb)
        {
            var ws = wb.Sheets[1] as Excel.Worksheet;
            if (ws == null)
                return;

            var va = ws.UsedRange.Value2 as object[,];
            if (va == null)
                return;

            var rows = va.GetUpperBound(0);
            var cols = va.GetUpperBound(1);
            var pageCol = -1;
            Range delRange = null;

            for (var row = 1; row <= rows; row++)
            for (var col = (pageCol < 0 ? 1 : pageCol); col <= (pageCol < 0 ? cols : pageCol); col++)
                if (IsPageTotal(va[row, col]))
                {
                    pageCol = col;
                    delRange = delRange == null ? ws.Rows[row] : _xl.Union(delRange, ws.Rows[row]) as Range;
                }

            if (delRange != null && delRange.Rows.Count > 0)
                delRange.EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);

        }

        private  static bool IsPageTotal(object val)
        {
            if (val == null)
                return false;

            var stringVal = val as string;
            if (stringVal == null)
                return false;

            if (!stringVal.StartsWith(PageTag, StringComparison.Ordinal))
                return false;

            if (!_pageRx.Match(stringVal).Success)
                return false;

            return true;
        }

        private static bool LooksLikeCRMExport(Workbook wb)
        {
            return wb.Sheets.Count == 1 && LooksLikeXmlOrHtml(wb.FullName);
        }

        private static bool LooksLikeXmlOrHtml(string fullName)
        {
            try
            {
                using (var fileStream = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.Delete | FileShare.ReadWrite))
                using (var streamReader = new StreamReader(fileStream))
                return streamReader.Read() == '<';
            }
            catch
            {
                return false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
