using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace cactus
{
    class CommentFinder : IFinder
    {
        // 查找注释并提取出来

        public void getContent()
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string fileName = doc.Name;

            int rows = doc.Comments.Count;

            if (rows == 0)
            {
                System.Windows.Forms.MessageBox.Show("文档中没有批注。");
                return;
            }

            string[,] commentsArray = new string[rows, 5];

            int p = 0;
            foreach (Comment c in doc.Comments)
            {
                //页码;
                commentsArray[p, 0] = Convert.ToString(c.Scope.Information[WdInformation.wdActiveEndPageNumber]);
                //行号;
                commentsArray[p, 1] = Convert.ToString(c.Scope.Information[WdInformation.wdFirstCharacterLineNumber]);
                //批注引用内容;
                commentsArray[p, 2] = c.Scope.Text;
                //批注内容;
                commentsArray[p, 3] = c.Range.Text;
                //作者;
                commentsArray[p, 4] = c.Author;
                p = p + 1;
            }

            Document newDoc = null;
            // Create An New Word   
            newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.ParagraphFormat.CharacterUnitFirstLineIndent = 2;
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            Paragraph par = newDoc.Content.Paragraphs.Add();
            par.Range.Text = "导出批注工具";
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("以下是来自文档：“" + fileName + "”的批注。");
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("--------------分割线------------------");
            par.Range.InsertParagraphAfter();

            for (int i = 0; i < rows; i++)
            {
                par.Range.InsertAfter(i + 1 + "、" + "第" + commentsArray[i, 0] + "页，第"
                    + commentsArray[i, 1] + "行 || " + "原文：" + commentsArray[i, 2]
                    + " || 意见：" + commentsArray[i, 3]);

                par.Range.InsertParagraphAfter();
            }

        }
    }
}
