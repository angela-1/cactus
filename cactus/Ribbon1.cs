using Microsoft.Office.Tools.Ribbon;

namespace cactus
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            splitButton1.Label = "提取内容\n";
            button1.Label = "导出批注\n";
            button2.Label = "套用格式\n";
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            CommentFinder comment = new CommentFinder();
            comment.GetContent();

        }


     
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Formatter format = new Formatter();
            format.doFormat();
        }

        private void splitButton1_Click(object sender, RibbonControlEventArgs e)
        {
            OrgFinder of = new OrgFinder();
            of.GetContent();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            splitButton1_Click(sender, e);
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            TitleFinder t1 = new TitleFinder(1);
            t1.GetContent();
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            TitleFinder t2 = new TitleFinder(2);
            t2.GetContent();
        }
    }
}
