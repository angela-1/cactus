using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace cactus
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string fileName = doc.Name;
           
            int rows = doc.Comments.Count;
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

        private bool _match_regex(string text, Regex re)
        {
            if (re.IsMatch(text))
                return true;
            else return false;
        }

        private int _get_format_end(Selection cursor, Paragraphs pars)
        {
            int endFormatPar = 0;
            foreach (Paragraph par in pars)
            {
                if (par.Range.Start > cursor.End)
                    break;
                endFormatPar += 1;
            }
            return endFormatPar;
        }

        private int _get_format_start(Selection cursor, Paragraphs pars)
        {
            int startFormatPar = 0;
            foreach (Paragraph par in pars)
            {
                if (par.Range.Start > cursor.Start)
                    break;
                startFormatPar += 1;
            }
            return startFormatPar;
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;
            Selection cursor = Globals.ThisAddIn.Application.Selection;

            // 判断是否选择了段落
            if (cursor.Paragraphs.Count <= 1)
            {
                System.Windows.Forms.MessageBox.Show("请选中要格式化的段落。");
                return;
            }

            // 获取选中的要修改样式的开头段落和结尾段落
            // 开头通过光标位置获取，结尾通过检测“抄送：”字符获取
            int startFormatPar = this._get_format_start(cursor, pars);
            int endFormatPar = this._get_format_end(cursor, pars);

            //System.Windows.Forms.MessageBox.Show("start par: " + startFormatPar
            //    + "end par:" + endFormatPar + "total par:" + parCount);

            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];
                par.Range.Text = par.Range.Text.Replace("　", "").Replace(" ", "");
            }

            // 各级标题通过正则表达式检测
            Regex reLevel1 = new Regex("^[一二三四五六七八九十]+、");
            Regex reLevel2 = new Regex("^（[一二三四五六七八九十]+）");
            //Regex reLevel3 = new Regex("^[0-9]. ");

            // 正则检测每段开头对应修改样式
            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];

                // 需要格式化的从开头到结尾都是首行缩进2字符，三号字，固定行距28磅
                par.CharacterUnitFirstLineIndent = 2;
                par.Range.Font.Size = 16;
                par.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                par.LineSpacing = 28;

                string lineStart = par.Range.Text;
                // 一级标题黑体
                if (reLevel1.IsMatch(lineStart))
                {
                    par.Range.Font.Name = "黑体";
                    par.Range.Font.NameAscii = "Tmimes New Roman";
                }
                // 二级标题楷体
                else if (reLevel2.IsMatch(lineStart))
                {
                    par.Range.Font.Name = "楷体";
                    par.Range.Font.NameAscii = "Tmimes New Roman";
                }
                // 三级和其他全部都是方正仿宋
                else
                {
                    par.Range.Font.Size = 16;
                    par.Range.Font.Name = "方正仿宋_GBK";
                    par.Range.Font.NameAscii = "Tmimes New Roman";
                }
            }

            System.Windows.Forms.MessageBox.Show("格式化完成。");
        }
    }
}
