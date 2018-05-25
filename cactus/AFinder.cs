using Microsoft.Office.Interop.Word;
using System;

namespace cactus
{
    abstract class AFinder
    {
        protected String src_file;
        //protected String copy_file;
        //protected String tmp_file;

        public AFinder()
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            src_file = doc.Path + "\\" + doc.Name;
            //copy_file = Path.GetTempFileName();
            //init();
        }

        //public void init()
        //{
    
        //    File.Copy(src_file, copy_file, true);

        //    Document newDoc = Globals.ThisAddIn.Application.Documents.Open(copy_file);
        //    tmp_file = Path.GetTempFileName();

        //    newDoc.SaveAs2(tmp_file, WdSaveFormat.wdFormatText);
        //    newDoc.Close();
        //}

        public abstract void GetContent();

        //public void ClearTmp()
        //{
        //    File.Delete(copy_file);
        //    File.Delete(tmp_file);
        //}
    }
}
