using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;


namespace Ganemo.WordHelper
{
    public static class WordTool
    {

        public static Document OpenDocument(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath))
            {
                throw new ArgumentNullException($"some parameter is null - docPath: {docPath}.");
            }

            try
            {
                return new Document(docPath);
            }
            catch (Exception)
            {
                throw;
            }

        }

        /// <summary>
        /// Replace the logo in footer
        /// </summary>
        /// <param name="doc">opened document</param>
        /// <param name="logoName">the place holder name </param>
        /// <param name="imagePath">full path of image</param>
        /// <returns>
        ///    true: replaced successfully
        ///    false: logoName was not found
        ///    exception: various reasons as described in detail
        /// </returns>
        public static Boolean ReplaceFooterLogo(this Document doc, string logoName, string imagePath)
        {
            if(doc == null || string.IsNullOrWhiteSpace(logoName) || string.IsNullOrWhiteSpace(imagePath))
            {
                throw new ArgumentNullException($"some parameter is null - " +
                    $"doc: {doc}, logoName: {logoName}, imagePath: {imagePath}.");
            }

            try
            {
                TextSelection[] tsl = doc.FindAllString(logoName, true, false);

                if(tsl == null || tsl.Count() <= 0 )
                {
                    return false;
                }

                foreach (TextSelection select in tsl)
                {
                    TextRange tr = select.GetAsOneRange();
                    Paragraph par = tr.OwnerParagraph;
                    int index = par.GetIndex(tr);
                    Section sec = par.Owner.Owner as Section;
                    //add a temporary paragraph
                    Paragraph parTem = sec.AddParagraph();
                    DocPicture picture = parTem.AppendPicture(Image.FromFile(imagePath));
                    //insert image
                    par.ChildObjects.Insert(index, picture);
                    //set wrapping style
                    picture.TextWrappingStyle = TextWrappingStyle.Inline;
                    //remove temporary paragraph
                    sec.Body.ChildObjects.Remove(parTem);
                    //remove the space holder logName
                    par.ChildObjects.Remove(tr);
                }

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        public static void AddPageNumber(this Document doc)
        {
            if (doc == null)
            {
                throw new ArgumentNullException($"some parameter is null - doc: {doc}.");
            }

            HeaderFooter footer = doc.Sections[0].HeadersFooters.Footer;
            Paragraph footerParagraph = footer.FirstParagraph as Paragraph;
            footerParagraph.AppendField("page number", FieldType.FieldPage);
            footerParagraph.AppendText(" of ");
            footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);

            //Section section = doc.Sections[0];
            //HeaderFooter footer = section.HeadersFooters.Footer;
            //Paragraph footerParagraph = footer.AddParagraph();
            //footerParagraph.AppendField("page number", FieldType.FieldPage);
            //footerParagraph.AppendText(" of ");
            //footerParagraph.AppendField("number of pages", FieldType.FieldNumPages);
            //footerParagraph.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
        }


        public static Boolean ReplaceParagraph(this Document doc, string placeHolder, string newParaContent)
        {
            if (doc == null || string.IsNullOrWhiteSpace(placeHolder) || string.IsNullOrWhiteSpace(newParaContent))
            {
                throw new ArgumentNullException($"some parameter is null - " +
                    $"doc: {doc}, placeHolder: {placeHolder}, newParaContent: {newParaContent}.");
            }


            //search for the place holder
            TextSelection ts = doc.FindString(placeHolder, true, false);

            if(ts == null || ts.Count < 1)
            {
                return false;
            }


            //get the paragram which contains the place holder
            Paragraph para = ts.GetAsOneRange().OwnerParagraph;
            Section sec = para.Owner.Owner as Section;
            //get index of this paragraph
            int index = sec.Body.Paragraphs.IndexOf(para);
            //remove paragraph
            //sec.Body.Paragraphs.RemoveAt(index);
            sec.Body.Paragraphs.RemoveAt(index);
            //add a new paragram to the current location
            Paragraph NewPara = new Paragraph(doc);
            NewPara.AppendText(newParaContent);
            sec.Body.ChildObjects.Insert(index, NewPara);

            return true;
        }

        
        public static Spire.Doc.Formatting.CharacterFormat GetCharacterFormat(this Document doc, string styleName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(styleName))
            {
                throw new ArgumentNullException($"some parameter is null - " +
                    $"doc: {doc}, styleName: {styleName}.");
            }

            foreach (Section sec in doc.Sections)
            {
                foreach (DocumentObject obj in sec.Body.ChildObjects)
                {
                    if (obj is Paragraph)
                    {
                        Paragraph para = obj as Paragraph;
                        if (para.StyleName == styleName)
                        {
                            return para.BreakCharacterFormat;
                        }
                    }
                }
            }

            return null;
        }

        public static void AddHeading(this Document doc,  string headingContent, Spire.Doc.Formatting.CharacterFormat characterFormat = null)
        {

            if (doc == null || string.IsNullOrWhiteSpace(headingContent))
            {
                throw new ArgumentNullException($"some parameter is null - " +
                    $"doc: {doc}, headingContent: {headingContent}.");
            }

            
            //create heading
            Section section = doc.Sections[0];
            Paragraph para1 = section.AddParagraph();
            para1.AppendText(headingContent);
            para1.ApplyStyle(BuiltinStyle.Heading1);

            if(characterFormat != null)
            {
                foreach (DocumentObject obj in para1.ChildObjects)
                {
                    TextRange tr = obj as TextRange;
                    tr.ApplyCharacterFormat(characterFormat);
                }
            }


            //Paragraph para2 = section.AddParagraph();
            //para2.AppendText("3.1 Head2");
            //para2.ApplyStyle(BuiltinStyle.Heading2);
            //foreach (DocumentObject obj in para2.ChildObjects)
            //{
            //    TextRange tr = obj as TextRange;
            //    tr.ApplyCharacterFormat(format2);
            //}

            ////use self-defined style
            //Paragraph para3 = section.AddParagraph();
            //para3.AppendText("3.1.1 Head3");
            //para3.ApplyStyle(BuiltinStyle.Heading3);
            //foreach (DocumentObject obj in para3.ChildObjects)
            //{
            //    TextRange tr = obj as TextRange;
            //    tr.CharacterFormat.FontSize = 10.0f;
            //    tr.CharacterFormat.Bold = true;
            //    tr.CharacterFormat.TextColor = format2.TextColor;
            //}

        }

    }
}

