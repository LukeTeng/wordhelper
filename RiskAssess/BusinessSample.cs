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

            ////add heading and use character format
            var format = doc.GetCharacterFormat("Heading1");
            doc.AddHeading("4 Heading test", format);

  

            //update contents
            doc.UpdateTableOfContents();

            //save document
            doc.SaveToFile("result.docx", FileFormat.Docx);
            System.Diagnostics.Process.Start("result.docx");
        }


    }
}
