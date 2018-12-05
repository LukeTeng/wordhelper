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


        //public void CreateWord_Click(object sender, EventArgs e)
        //{
        //    Document doc = new Document("../../files/demo_doc.docx");

        //    //需求1：替换页脚<>里面的文本为图片
        //    TextSelection[] ts1 = doc.FindAllString("instance logo", true, false);
        //    replaceFooterLogo(ts1, "../../files/logo1.png");

        //    TextSelection[] ts2 = doc.FindAllString("application logo", true, false);
        //    replaceFooterLogo(ts2, "../../files/logo2.png");

        //    //获取第一个section
        //    Section section = doc.Sections[0];

        //    //需求2：添加页码
        //    AddPageNumber(section);

        //    //需求3：替换一大段文本,例如这里替换掉<Domain> Name所在段落和<Sub-Name>所在段落
        //    ReplaceParagraph(doc);

        //    //需求4：添加heading并设置样式
        //    AddHeading(doc);

        //    //需求5：添加表格
        //    AddTable(section);

        //    //需求6：在指定位置替换图片，我没看到你的模板文件含有diagram图像。
        //    //您是想替换原始文档中某个图片为新的图片吗?如果是的话，请使用下面的代码：
        //    //foreach (Section sec in doc.Sections)
        //    //{
        //    //    foreach (Paragraph paragraph in sec.Paragraphs)
        //    //    {
        //    //        foreach (DocumentObject docObj in paragraph.ChildObjects)
        //    //        {
        //    //            if (docObj.DocumentObjectType == DocumentObjectType.Picture)
        //    //            {
        //    //                DocPicture picture = docObj as DocPicture;
        //    //                if (picture.Title == "Figure 1")
        //    //                {
        //    //                    //替换图像
        //    //                    picture.LoadImage(Image.FromFile("../../files/logo1.png"));
        //    //                }
        //    //            }
        //    //        }
        //    //    }
        //    //}
        //    //或者您是想替换某个文本为图像，这个实现的原理和您的"需求1:替换页脚<>里面的文本为图片"一样。

        //    //需求7：添加链接
        //    AddLink(doc);

        //    //保存文档
        //    doc.SaveToFile("result.docx", FileFormat.Docx);
        //    System.Diagnostics.Process.Start("result.docx");
        //}

        
        public static Document OpenDocument(string docPath)
        {
            if (string.IsNullOrWhiteSpace(docPath))
            {
                throw new ArgumentNullException($"some parameter is null - docPath: {docPath}.");
            }

            try
            {
                return new Document("../../files/demo_doc.docx");
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


        public static void AddTable(Section section)
        {
            Table table = section.AddTable(true);

            //表头和表格数据
            String[] Header = { "Group", "item1", "item2", "item3", "item4" };
            String[][] data = {
                                  new String[]{ "Group1 name","name1","name1","name1","name1"},
                                  new String[]{"Group2 name","name2","name2","name2","name2"},
                                  new String[]{"Group3 name","name3","name3","name3","name3"},
                              };
            //添加表
            table.ResetCells(data.Length + 1, Header.Length);

            //表头样式
            TableRow FRow = table.Rows[0];
            FRow.IsHeader = true;
            FRow.Height = 23;
            FRow.RowFormat.BackColor = Color.Blue;
            for (int i = 0; i < Header.Length; i++)
            {
                Paragraph p = FRow.Cells[i].AddParagraph();
                FRow.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                TextRange TR = p.AppendText(Header[i]);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.White;
                TR.CharacterFormat.Bold = true;
            }
            //行样式
            for (int r = 0; r < data.Length; r++)
            {
                TableRow DataRow = table.Rows[r + 1];
                DataRow.Height = 20;
                for (int c = 0; c < data[r].Length; c++)
                {
                    DataRow.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    Paragraph p2 = DataRow.Cells[c].AddParagraph();
                    TextRange TR2 = p2.AppendText(data[r][c]);
                    p2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    TR2.CharacterFormat.FontName = "Calibri";
                    TR2.CharacterFormat.FontSize = 12;
                }
            }
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

            ////自定义heading3样式
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

            ////更新目录
            //doc.UpdateTableOfContents();

        }



        public static void AddLink(Document doc)
        {
            //设置链接样式
            ParagraphStyle style = new ParagraphStyle(doc);
            style.Name = "linkStyle";
            style.CharacterFormat.TextColor = Color.Blue;
            style.CharacterFormat.UnderlineStyle = UnderlineStyle.Single;
            doc.Styles.Add(style);
            //循环表格找到内容为<Link>的段落，添加hyperlink域
            foreach (Section section in doc.Sections)
            {
                foreach (Table table in section.Tables)
                {
                    foreach (TableRow row in table.Rows)
                    {
                        foreach (TableCell col in row.Cells)
                        {
                            foreach (DocumentObject obj in col.ChildObjects)
                            {
                                if (obj is Paragraph)
                                {
                                    Paragraph para = obj as Paragraph;
                                    if (para.Text == "<Link>")
                                    {
                                        int index = para.ChildObjects.IndexOf(para.ChildObjects[0]);
                                        Field field = new Field(doc);
                                        field.Code = "HYPERLINK \"" + "http://www.e-iceblue.com" + "\"";
                                        field.Type = FieldType.FieldHyperlink;

                                        para.ChildObjects.Insert(index, field);

                                        FieldMark fm = new FieldMark(doc, Spire.Doc.Documents.FieldMarkType.FieldSeparator);
                                        para.ChildObjects.Insert(index + 1, fm);

                                        FieldMark fmend = new FieldMark(doc, FieldMarkType.FieldEnd);
                                        para.ChildObjects.Add(fmend);
                                        field.End = fmend;
                                        //添加样式
                                        para.ApplyStyle("linkStyle");
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

