using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace cactus
{
    class DivisionFinder : AFinder
    {
        private String _org;

        public DivisionFinder(String org) : base()
        {
           
            _org = org;
        }

        public override void GetContent()
        {
            List<String> final_list = _search();
            if (final_list.Count > 0)
            {
                _print_to_file(final_list);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("未找到符合条件的项。");
            }
        }


        private List<String> _search()
        {
            Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;

            List<String> draft_list = new List<String>();
            foreach (Paragraph par in pars)
            {
                Regex reg = new Regex(_org);
                Match match = reg.Match(par.Range.Text);
                if (match.Success)
                {
                    draft_list.Add(par.Range.Text);
                    Debug.WriteLine(par.Range.Text);
                }
            }
            return draft_list;
        }

        private void _print_to_file(List<String> final_list)
        {
            Document newDoc = null;
            newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            newDoc.Content.Paragraphs[1].Range.Font.NameAscii = "Times New Roman";
            Paragraph par = newDoc.Content.Paragraphs.Add();

            par.Range.Text = "来自文档：“" + src_file + "”中" + _org + "涉及的任务共" + final_list.Count + "项。";
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("--------------------------------分割线-------------------------------");
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter(" ");
            par.Range.InsertParagraphAfter();

            foreach (String item in final_list)
            {
                par.Range.InsertAfter(item);
            }

            Paragraphs ps = newDoc.Paragraphs;
            foreach (Paragraph p in ps)
            {
                p.Range.Select();
                Selection ss = Globals.ThisAddIn.Application.Selection;
                ss.Find.Text = _org;//查询的文字
                Boolean is_find = ss.Find.Execute(Forward: true, Wrap: WdFindWrap.wdFindContinue, Format: false);
                //ss.Font.Color = WdColor.wdColorRed;//设置颜色为红
                if (is_find)
                {
                    ss.Range.HighlightColorIndex = WdColorIndex.wdYellow;
                }
            }
        }
    }
}
