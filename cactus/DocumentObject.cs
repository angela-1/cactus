using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace cactus
{
    class DocumentObject : AFinder
    {
        public string title = "";
        public string code = "";
        public string sendBy = "";
        public string sendTo = "";
        public string sendDate = "";

        public override void GetContent()
        {
            int flag = 0b0000;

            List<String> contents = _parse_contents();
            if (contents.Count == 1)
            {
                MessageBox.Show("文件为空。");
                return;
            }
            bool isTitleLine = false;
            string title = "";

            foreach (var line in contents)
            {
                if ((flag & 1) == 0 && (flag & 2) == 0)
                {
                    string code = _get_code(line);
                    if (code.Length > 0)
                    {
                        this.code = code;
                        flag = flag | 1;
                        continue;
                    }
                }

                if ((flag & 2) == 0)
                {
                    if (isTitleLine)
                    {
                        title += line.Trim();
                    }

                    if (_is_title_line(line))
                    {
                        title += line;
                        isTitleLine = true;
                    }

                    if (_is_white_line(line) && isTitleLine)
                    {
                        this.title = title;
                        flag = flag | 2;
                        isTitleLine = false;
                        continue;
                    }
                }

                if ((flag & 4) == 0)
                {
                    string send_to = _get_send_to(line);
                    if (send_to.Length > 0)
                    {
                        this.sendTo = send_to;
                        flag = flag | 4;
                        continue;
                    }
                }

                if ((flag & 8) == 0)
                {
                    string send_date = _get_send_date(line);
                    if (send_date.Length > 0)
                    {
                        int ind = contents.IndexOf(line);
                        this.sendBy = contents[ind - 1];
                        this.sendDate = send_date;
                        flag = flag | 8;
                        continue;
                    }
                }

                if (flag == 0b1111)
                {
                    break;
                }
            }

            string json = JsonConvert.SerializeObject(this);
            Clipboard.SetDataObject(json);
            MessageBox.Show("文件对象已经存入剪贴板。请使用 Ctrl+v 粘贴。");
        }

        private string _get_code(string par)
        {
            string value = "";
            Regex reg = new Regex(@"\S+〔\d{4}〕\d+号");
            Match match = reg.Match(par);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }

        private string _get_send_to(string par)
        {
            string value = "";
            Regex reg = new Regex(@"\S+[：:]$");
            Match match = reg.Match(par);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }

        private string _get_send_date(string par)
        {
            string value = "";
            Regex reg = new Regex(@"^\d{4}年\d{1,2}月\d{1,2}日$");
            Match match = reg.Match(par);
            if (match.Success)
            {
                value = match.Value;
            }
            return value;
        }



        private bool _is_title_line(string par)
        {
            Regex reg = new Regex(@"关于\S+");
            Match match = reg.Match(par);
            return match.Success;
        }

        private bool _is_white_line(string par)
        {
            Regex reg = new Regex(@"^\s*$");
            Match match = reg.Match(par);
            return match.Success;
        }

        private List<String> _parse_contents()
        {
            Document thisDoc = Globals.ThisAddIn.Application.ActiveDocument;
            Paragraphs pars = thisDoc.Paragraphs;
            int parCount = thisDoc.Paragraphs.Count;

            List<String> draft_list = new List<String>();
            foreach (Paragraph par in pars)
            {
                draft_list.Add(par.Range.Text.Trim());
            }
            return draft_list;
        }
    }
}
