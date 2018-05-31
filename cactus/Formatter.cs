using Microsoft.Office.Interop.Word;
using System;
using System.Text.RegularExpressions;

namespace cactus
{
    class Formatter
    {
        // 按照公文要求修改文档格式
        public void doFormat()
        {
            Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;
            Selection cursor = Globals.ThisAddIn.Application.Selection;

            thisDoc.Content.ListFormat.ConvertNumbersToText();

            // 判断是否选择了段落
            if (cursor.Paragraphs.Count <= 1)
            {
                System.Windows.Forms.MessageBox.Show("请选中要套用格式的段落。");
                return;
            }

            // 获取选中的要修改样式的开头段落和结尾段落
            //
            int startFormatPar = this._get_format_start(cursor, pars);
            int endFormatPar = this._get_format_end(cursor, pars);

            //System.Windows.Forms.MessageBox.Show("start par: " + startFormatPar
            //    + "end par:" + endFormatPar + "total par:" + parCount);

            // 各级标题通过正则表达式检测
            Regex reLevel1 = new Regex("^[一二三四五六七八九十]+、");
            Regex reLevel2 = new Regex("^（[一二三四五六七八九十]+）");
            //Regex reLevel3 = new Regex("^[0-9]. ");

            Boolean level1FontLock = true;
            Boolean level2FontLock = true;

            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];
                par.Range.Text = par.Range.Text.Replace("　", "").Replace(" ", "");
            }

            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];
                string lineStart = par.Range.Text;
                if (reLevel1.IsMatch(lineStart, 0) && lineStart.Length > 24)
                {
                    level1FontLock = false;
                    break;
                }
                if (reLevel2.IsMatch(lineStart) && lineStart.Length > 24)
                {
                    level2FontLock = false;
                    break;
                }
            }

            // 正则检测每段开头对应修改样式
            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];

                // 需要格式化的从开头到结尾都是首行缩进2字符，三号字，固定行距28磅
                par.CharacterUnitFirstLineIndent = 2;
                par.Range.Font.NameAscii = "Tmimes New Roman";
                par.Range.Font.Size = 16;
                par.Range.Font.Bold = 0;
                par.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
                par.LineSpacing = 28;
                par.SpaceBefore = 0;
                par.SpaceAfter = 0;

                string lineStart = par.Range.Text;

                // 一级标题黑体
                if (level1FontLock && reLevel1.IsMatch(lineStart))
                {
                    par.Range.Font.Name = "黑体";
                    par.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;
                }
                // 二级标题楷体
                else if (level2FontLock && reLevel2.IsMatch(lineStart))
                {
                    par.Range.Font.Name = "方正楷体_GBK";
                    par.Range.ParagraphFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevel2;
                }
                // 三级和其他全部都是方正仿宋
                else
                {
                    par.Range.Font.Size = 16;
                    par.Range.Font.Name = "方正仿宋_GBK";
                }
            }

            System.Windows.Forms.MessageBox.Show("格式化完成。");
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
            return endFormatPar + 1;
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
    }
}
