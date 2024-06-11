using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace wordProtect
{
    public partial class Form1 : Form
    {
        public static Application ap_ = null;
        public static Application ap
        {
            get
            {
                try
                {
                    if (ap_ != null)
                    {
                        ap_.Visible = false;
                        ap_.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                        ap_.Options.ConfirmConversions = false;
                    }
                    else
                    {
                        ap_ = new Application();
                        ap_.Visible = false;
                        ap_.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                        ap_.Options.ConfirmConversions = false;
                    }
                }
                catch
                {
                    ap_ = new Application();
                    ap_.Visible = false;
                    ap_.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    ap_.Options.ConfirmConversions = false;
                }
                return ap_;
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void openDoc_Click(object sender, EventArgs e)
        {
            var filePath = "G:\\multiMatrixReport.docx";
            var document = OpenFileInWord(filePath);
            var bookmarks = new List<string> { "calibration_table", "curve_param_table", "isr_table", "qc_repCrossAccNoRep", "assay", "blank_matrix_table" };
            object oMoveCharacter = WdUnits.wdCharacter;
            object oOne = -1;
            object OSix = -6;
            //add section breaks before each bookmark
            foreach (var bookmark in bookmarks)
            {
                var range = document.Bookmarks[bookmark].Range;
                range.Collapse(WdCollapseDirection.wdCollapseStart);
                range.MoveStart(ref oMoveCharacter, ref oOne);
                range.InsertBreak(WdBreakType.wdSectionBreakContinuous);
            }
            foreach (var bookmark in bookmarks)
            {
                var range = document.Bookmarks[bookmark].Range;

                range.Collapse(WdCollapseDirection.wdCollapseEnd);
                range.InsertBreak(WdBreakType.wdSectionBreakContinuous);
            }
            //get the sections
            var sections = document.Sections;
            //protect each section
            foreach (Section section in sections)
            {
                var range = section.Range;

                var text = range.Text;
                ;
            }
            List<(int, int)> sectionsToProtect = new List<(int, int)>();

            foreach (string bookmarkName in bookmarks)
            {
                Bookmark bookmark = document.Bookmarks[bookmarkName];
                Range bookmarkRange = bookmark.Range;
                var bookmarkStart = bookmarkRange.Start;
                var bookmarkEnd = bookmarkRange.End;
                if (bookmarkRange.Sections.Count > 0)
                {
                    sectionsToProtect.Add((bookmarkRange.Sections.First.Range.Start, bookmarkRange.Sections.First.Range.End));
                }
                foreach (Section section in bookmarkRange.Sections)
                {
                    var range = section.Range;

                    var text = range.Text;
                    ;
                }
            }

            document.Protect(WdProtectionType.wdAllowOnlyReading, NoReset: true);
            foreach (Section section in sections)
            {
                if (!sectionsToProtect.Contains((section.Range.Start, section.Range.End)))
                {
                    section.Range.Editors.Add(WdEditorType.wdEditorEveryone);
                }
            }


        }
        public static Document OpenFileInWord(string filePath)
        {
            Document document = ap.Documents.Open(filePath);
            document.ActiveWindow.Visible = true;
            return document;
        }
        //protect the whole document
        private void protectDoc_Click(Document document)
        {
            document.Protect(WdProtectionType.wdAllowOnlyReading, NoReset: true);
        }
    }
}
