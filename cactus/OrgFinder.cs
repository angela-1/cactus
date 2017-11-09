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
    class OrgFinder : AFinder
    {
        // 查找责任单位/牵头部门并提取
        private Regex reg;

        public OrgFinder()
        {
            reg = new Regex(@"（[责牵]\S+）$");
        }

        public override void GetContent()
        {
            // 新建文档填入单位名称
            //System.Windows.Forms.MessageBox.Show("新建文件了。");

            List<String> a = __search_file();

            if (a.Count > 0)
            {
                SortedSet<String> b = __strip_brackets(a);
                
                __print_to_file(b);

            } else
            {
                System.Windows.Forms.MessageBox.Show("未找到符合条件的项。");
            }

            ClearTmp();
        }

        private List<String> __search_file()
        {
            StreamReader sr = new StreamReader(tmp_file, Encoding.Default);
            String line;
            List<String> draft_list = new List<String>();

            while ((line = sr.ReadLine()) != null)
            {
                Match match = reg.Match(line);
                if (match.Success)
                {
                    String b = match.Groups[0].ToString();
                    draft_list.Add(b);
                    //System.Windows.Forms.MessageBox.Show("bb" + b);

                }
                //Debug.WriteLine(match.ToString(), line.ToString());
            }
            sr.Close();
            return draft_list;
        }

        private SortedSet<String> __strip_brackets(List<String> draft_list)
        {
            SortedSet<String> final_list = new SortedSet<String>();
            char[] trimChars = { '（', '）' };
            foreach (String item in draft_list)
            {
                String[] a = item.Trim(trimChars).Split('：');
                String[] b = a[1].Split('、');
                foreach (String c in b)
                {
                    final_list.Add(c);
                    //Debug.WriteLine(c);

                }
            }
            //System.Windows.Forms.MessageBox.Show("dd");

            return final_list;
        }

        private void __print_to_file(SortedSet<String> final_list)
        {
            Document newDoc = null;
            // Create An New Word   
            newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            newDoc.Content.Paragraphs[1].Range.Font.NameAscii = "Times New Roman";
            Paragraph par = newDoc.Content.Paragraphs.Add();

            par.Range.Text = "来自文档：“" + src_file+ "”中出现的单位共" + final_list.Count + "家。";
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
