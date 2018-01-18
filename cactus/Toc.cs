using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cactus
{
    class Toc
    {
        // 制作汇编材料目录
        public void doFormat()
        {
            Application app = Globals.ThisAddIn.Application;
            Document thisDoc = app.ActiveDocument;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;
            Selection cursor = Globals.ThisAddIn.Application.Selection;

            cursor.WholeStory();

            cursor.ParagraphFormat.SpaceBeforeAuto = 0;
            cursor.ParagraphFormat.SpaceAfterAuto = 0;
            cursor.ParagraphFormat.LeftIndent = 0;
            cursor.ParagraphFormat.FirstLineIndent = 0;
            cursor.ParagraphFormat.CharacterUnitFirstLineIndent = 0;
            cursor.ParagraphFormat.WordWrap = 0;

            cursor.ParagraphFormat.FirstLineIndent = 0;

            cursor.ParagraphFormat.TabStops.ClearAll();
            cursor.ParagraphFormat.TabStops.Add(Position: app.CentimetersToPoints(14.2F),
                Alignment: WdHorizontalLineAlignment.wdHorizontalLineAlignLeft, 
                Leader: WdTabLeader.wdTabLeaderDots);

          

            System.Windows.Forms.MessageBox.Show("目录制作完成。");
        }
    }
}
