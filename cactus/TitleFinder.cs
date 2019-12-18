using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace cactus
{
    class TitleFinder : AFinder
    {
        // 查找一级标题并提取
        private readonly int type;
        private readonly Regex reLevel1;
        private readonly Regex reLevel2;

        public TitleFinder(int find_type)
        {
            type = find_type;
            reLevel1 = new Regex("^[一二三四五六七八九十]+、");
            reLevel2 = new Regex("^（[一二三四五六七八九十]+）");
        }


        private List<string> Search()
        {
            Document this_doc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = this_doc.Paragraphs;

            List<string> draft_list = new List<string>();
            Regex reg;
            WdOutlineLevel level;
            if (type == 1)
            {
                reg = reLevel1;
                level = WdOutlineLevel.wdOutlineLevel1;
            }
            else
            {
                reg = reLevel2;
                level = WdOutlineLevel.wdOutlineLevel2;
            }

            foreach (Paragraph par in pars)
            {
                string line = par.Range.Text.Trim();
                Match match = reg.Match(line);
                if (match.Success)
                {
                    string b = match.Groups[0].ToString();
                    draft_list.Add(line);
                }

                if (par.OutlineLevel == level)
                {
                    draft_list.Add(line);
                }
            }
            return draft_list;
        }

        private List<string> Search2()
        {
            Document this_doc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = this_doc.Paragraphs;

            List<string> draft_list = new List<string>();
            //WdOutlineLevel level = WdOutlineLevel.wdOutlineLevel1;



            //if (type == 1)
            //{
            //    reg = reLevel1;
            //    level = WdOutlineLevel.wdOutlineLevel1;
            //}
            //else
            //{
            //    reg = reLevel2;
            //    level = WdOutlineLevel.wdOutlineLevel2;
            //}

            foreach (Paragraph par in pars)
            {
                string line = par.Range.Text.Trim();
                Match match1 = reLevel1.Match(line);
                Match match2 = reLevel2.Match(line);
                if (match1.Success || par.OutlineLevel == WdOutlineLevel.wdOutlineLevel1)
                {
                    draft_list.Add(line);
                } else if ( match2.Success || par.OutlineLevel == WdOutlineLevel.wdOutlineLevel2) {
                    draft_list.Add(line);
                }

                //if (par.OutlineLevel == level)
                //{
                //    draft_list.Add(line);
                //}
            }
            return draft_list;
        }
        //private List<String> __parse_file()
        //{
        //    StreamReader sr = new StreamReader(tmp_file, Encoding.Default);
        //    String line;
        //    List<String> draft_list = new List<String>();

        //    Regex reg;
        //    if (type == 1)
        //    {
        //        reg = reLevel1;
        //    }
        //    else
        //    {
        //        reg = reLevel2;
        //    }

        //    while ((line = sr.ReadLine()) != null)
        //    {
        //        Match match = reg.Match(line.TrimStart());
        //        if (match.Success)
        //        {
        //            String b = match.Groups[0].ToString();
        //            draft_list.Add(line);
        //            //System.Windows.Forms.MessageBox.Show("bb" + line.Trim() + "\n");

        //        }
        //        //Debug.WriteLine(match.ToString(), line.ToString());
        //    }
        //    sr.Close();
        //    return draft_list;
        //}

        //private List<String> __parse_title()
        //{
        //    List<String> title_list = new List<String>();

        //    Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
        //    Paragraphs pars = thisDoc.Paragraphs;
        //    foreach (Paragraph item in pars)
        //    {
        //        if (item.OutlineLevel == WdOutlineLevel.wdOutlineLevel2)
        //        {
        //            title_list.Add(item.Range.Text);
        //        }
        //    }
        //    return title_list;
        //}

        private void PrintToFile(List<string> final_list)
        {
            // Create An New Word   
            Document newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            newDoc.Content.Paragraphs[1].Range.Font.NameAscii = "Times New Roman";
            Paragraph par = newDoc.Content.Paragraphs.Add();
            par.OutlineLevel = WdOutlineLevel.wdOutlineLevel1;

            par.Range.Text = "来自文档：“" + src_file + "”中的标题共" + final_list.Count + "项。";
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("--------------------------------分割线-------------------------------");
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter(" ");
            par.Range.InsertParagraphAfter();

            foreach (string item in final_list)
            {
                par.Range.InsertAfter(item);
                par.Range.InsertParagraphAfter();
            }
        }

        public override void GetContent()
        {
            List<string> list = Search();
            if (list.Count > 0)
            {
                PrintToFile(list);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("未找到符合条件的项。");
            }
            
        }


        public void GetContent2()
        {
            List<string> list = Search2();
            if (list.Count > 0)
            {
                PrintToFile(list);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("未找到符合条件的项。");
            }

        }

    }
}
