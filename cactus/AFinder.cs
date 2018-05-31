using Microsoft.Office.Interop.Word;
using System;

namespace cactus
{
    abstract class AFinder
    {
        protected String src_file;
      
        public AFinder()
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            src_file = doc.Path + "\\" + doc.Name;
        }

        public abstract void GetContent();
    }
}
