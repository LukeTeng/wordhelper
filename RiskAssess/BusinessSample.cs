using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Ganemo.WordHelper;

namespace Ganemo.Business.Sample
{

    public partial class BusinessSample : Form
    {

        private static string ROOT_PATH = "../../files/";

        public BusinessSample()
        {
            InitializeComponent();
        }


        private void CreateWord_Click(object sender, EventArgs e)
        {
            
            Document doc = WordHelper.WordTool.OpenDocument(ROOT_PATH + "demo_doc.docx");

            //replace the logo in footer
            doc.ReplaceFooterLogo("<instance logo>", ROOT_PATH + "logo2.png");
            doc.ReplaceFooterLogo("<application logo>", ROOT_PATH + "logo2.png");
            

            

            //add page number
            doc.AddPageNumber();
            
            
            //doc.ReplaceParagraph(@"<background>", @"The founders initially limited the website's membership to Harvard students. Later they expanded it to higher education institutions in the Boston area, the Ivy League schools, and Stanford University. Facebook gradually added support for students at various other universities, and eventually to high school students. Since 2006, anyone who claims to be at least 13 years old has been allowed to become a registered user of Facebook, though variations exist in this requirement, depending on local laws. The name comes from the face book directories often given to American university students. Facebook held its initial public offering (IPO) in February 2012, valuing the company at $104 billion, the largest valuation to date for a newly listed public company. It began selling stock to the public three months later. Facebook makes most of its revenue from advertisements that appear onscreen.");


            doc.ReplaceParagraph(@"<background>", @"The founders initially limited the website's membership to Harvard students. Later they expanded it to higher education institutions in the Boston area, the Ivy League schools, and Stanford University. 
Facebook gradually added support for students at various other universities, and eventually to high school students. Since 2006, anyone who claims to be at least 13 years old has been allowed to become a registered user of Facebook, though variations exist in this requirement, depending on local laws. The name comes from the face book directories often given to American university students. Facebook held its initial public offering (IPO) in February 2012, valuing the company at $104 billion, the largest valuation to date for a newly listed public company. It began selling stock to the public three months later. Facebook makes most of its revenue from advertisements that appear onscreen.");

            ////需求4：添加heading并设置样式
            var format = doc.GetCharacterFormat("Heading1");
            doc.AddHeading("4 Heading test", format);

            ////需求5：添加表格
            //AddTable(section);

            ////需求6：在指定位置替换图片，我没看到你的模板文件含有diagram图像。
            ////您是想替换原始文档中某个图片为新的图片吗?如果是的话，请使用下面的代码：
            ////foreach (Section sec in doc.Sections)
            ////{
            ////    foreach (Paragraph paragraph in sec.Paragraphs)
            ////    {
            ////        foreach (DocumentObject docObj in paragraph.ChildObjects)
            ////        {
            ////            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
            ////            {
            ////                DocPicture picture = docObj as DocPicture;
            ////                if (picture.Title == "Figure 1")
            ////                {
            ////                    //替换图像
            ////                    picture.LoadImage(Image.FromFile("../../files/logo1.png"));
            ////                }
            ////            }
            ////        }
            ////    }
            ////}
            ////或者您是想替换某个文本为图像，这个实现的原理和您的"需求1:替换页脚<>里面的文本为图片"一样。

            ////需求7：添加链接
            //AddLink(doc);

            //update contents
            doc.UpdateTableOfContents();

            //save document
            doc.SaveToFile("result.docx", FileFormat.Docx);
            System.Diagnostics.Process.Start("result.docx");
        }


    }
}
