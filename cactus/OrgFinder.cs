using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace cactus
{
    class OrgFinder : AFinder
    {
        public override void GetContent()
        {
            List<String> a = _search();
            if (a.Count > 0)
            {
                SortedSet<String> b = _strip_brackets(a);
                _print_to_file(b);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("未找到符合条件的项。");
            }
        }

        private List<String> _search()
        {
            Document this_doc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = this_doc.Paragraphs;

            List<String> draft_list = new List<String>();
            Regex reg = new Regex(@"（(牵头|责任|配合)(单位|部门)：\S+）$");

            foreach (Paragraph par in pars)
            {
                String line = par.Range.Text.Trim();
                Match match = reg.Match(line);
                if (match.Success)
                {
                    String b = match.Groups[0].ToString();
                    draft_list.Add(b);
                }
            }
            return draft_list;
        }

        private SortedSet<String> _strip_brackets(List<String> draft_list)
        {
            SortedSet<String> final_list = new SortedSet<String>();
            char[] trimChars = { '（', '）', ' ' };
            foreach (String item in draft_list)
            {
                Regex sp = new Regex(@"(牵头|责任|配合)(单位|部门)：");
                Regex sp2 = new Regex(@"[：，。]");
                String newitem = sp2.Replace(sp.Replace(item, ""), "、");
                String a = newitem.Trim(trimChars);
                String[] b = a.Split('、');
                foreach (String c in b)
                {
                    final_list.Add(c);
                    Debug.WriteLine(c);

                }
            }
            //System.Windows.Forms.MessageBox.Show("dd");
            return final_list;
        }

        private void _print_to_file(SortedSet<String> final_list)
        {
            Document newDoc = null;
            // Create An New Word   
            newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            newDoc.Content.Paragraphs[1].Range.Font.NameAscii = "Times New Roman";
            Paragraph par = newDoc.Content.Paragraphs.Add();

            par.Range.Text = "来自文档：“" + src_file + "”中出现的单位共" + final_list.Count + "家。";
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("--------------------------------分割线-------------------------------");
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter(" ");
            par.Range.InsertParagraphAfter();

            foreach (String item in final_list)
            {
                par.Range.InsertAfter(item + "、");
            }
        }
    }
}
