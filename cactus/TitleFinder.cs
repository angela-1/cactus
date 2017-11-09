using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace cactus
{
    class TitleFinder : IFinder
    {
        // 查找一级标题并提取

        private String src_file_path;
        private int type;

        public TitleFinder(int find_type)
        {
            type = find_type; 
            src_file_path = "";
        }

        private List<String> __parse_file()
        {
            Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
            src_file_path = thisDoc.Path + "\\" + thisDoc.Name;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;
 
            // 各级标题通过正则表达式检测
            Regex reLevel1 = new Regex("^[一二三四五六七八九十]+、");
            Regex reLevel2 = new Regex("^（[一二三四五六七八九十]+）");

            List<String> title_list = new List<String>();

            int startFormatPar = 1;
            int endFormatPar = pars.Count;

            for (int i = startFormatPar; i < endFormatPar; i++)
            {
                Paragraph par = pars[i];
                par.Range.Text = par.Range.Text.Replace("　", "").Replace(" ", "");
            }

            Regex reg;
            if (type == 1)
            {
                reg = reLevel1;
            } else
            {
                reg = reLevel2;
            }

            foreach (Paragraph item in pars)
            {
                String lineStart = item.Range.Text;
                
                if (reg.IsMatch(lineStart, 0))
                {
                    title_list.Add(lineStart);
                }
            }
            return title_list;
        }

        private void __print_to_file(List<String> final_list)
        {
            Document newDoc = null;
            // Create An New Word   
            newDoc = Globals.ThisAddIn.Application.Documents.Add();
            newDoc.Content.Paragraphs[1].Range.Font.Size = 16;
            newDoc.Content.Paragraphs[1].Range.Font.Name = "方正仿宋_GBK";
            newDoc.Content.Paragraphs[1].Range.Font.NameAscii = "Times New Roman";
            Paragraph par = newDoc.Content.Paragraphs.Add();

            par.Range.Text = "来自文档：“" + src_file_path + "”中的一级标题共" + final_list.Count + "项。";
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter("--------------------------------分割线-------------------------------");
            par.Range.InsertParagraphAfter();

            par.Range.InsertAfter(" ");
            par.Range.InsertParagraphAfter();

            foreach (String item in final_list)
            {
                par.Range.InsertAfter(item);
            }
        }

        public void getContent()
        {
            List<String> list = __parse_file();
            __print_to_file(list);
        }
    }
}
