using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace AutoPTD
{
    public partial class Form1 : Form
    {
        public DocConfig docConfig = new DocConfig();
        
        public Form1()
        {
            InitializeComponent();
            #region 生成配置文件

            textBox1.Text = docConfig.PicturePath;
            textBox2.Text = docConfig.DocName;
            #endregion
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Please select pictures save directory";
            fbd.RootFolder = Environment.SpecialFolder.MyComputer;
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                docConfig.PicturePath = textBox1.Text = fbd.SelectedPath + @"\";
                docConfig.SetNodeValue("PicturePath", docConfig.PicturePath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                DirectoryInfo di = new DirectoryInfo(textBox1.Text);
                Document document = new Document();
                document.LoadFromFile(Application.StartupPath + "\\TT.docx");
                FileInfo[] _allFiles = di.GetFiles("*.*", SearchOption.TopDirectoryOnly);

                #region 配置信息
                int HorizontalPosition = 70;
                int VerticalPosition = 70;
                int count = 0;
                #endregion

                foreach (FileInfo fi in _allFiles)
                {
                    if (count >= 3) break;
                    count++;

                    #region 创建文本框
                    Spire.Doc.Fields.TextBox TB = document.Sections[0].Paragraphs[0].AppendTextBox(220, 300);
                    TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
                    TB.Format.HorizontalPosition = HorizontalPosition;
                    TB.Format.VerticalOrigin = VerticalOrigin.Page;
                    TB.Format.VerticalPosition = VerticalPosition;
                    TB.Format.TextWrappingStyle = TextWrappingStyle.Tight;
                    TB.Format.TextWrappingType = TextWrappingType.Both;
                    #endregion

                    #region 设置文本框框的颜色，内部边距，图片填充。
                    TB.Format.LineStyle = TextBoxLineStyle.Simple;
                    TB.Format.LineColor = Color.Transparent;
                    TB.Format.LineDashing = LineDashing.Solid;
                    TB.Format.LineWidth = 3;
                    TB.Format.FillEfects.Type = BackgroundType.Picture;
                    //TB.Format.FillEfects.Picture = Image.FromFile(Application.StartupPath + "\\2.jpg");
                    #endregion

                    #region 在文本框内添加段落文本，图片，设置字体，字体颜色，行间距，段后距，对齐方式等。然后保存文档，打开查看效果。
                    Paragraph para1 = TB.Body.AddParagraph();
                    para1.Format.AfterSpacing = 6;
                    para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    TextRange TR1 = para1.AppendText("图" + count);
                    TR1.CharacterFormat.FontName = "华文新魏";
                    TR1.CharacterFormat.FontSize = 16;
                    TR1.CharacterFormat.Bold = true;

                    Paragraph para2 = TB.Body.AddParagraph();
                    Image image = Image.FromFile(fi.FullName);
                    DocPicture picture = para2.AppendPicture(image);
                    picture.Width = 200;
                    picture.Height = 250;
                    para2.Format.AfterSpacing = 8;
                    para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    //Paragraph para3 = TB.Body.AddParagraph();
                    //TextRange TR2 = para3.AppendText("描述--盛唐最杰出的诗人，中国历史最伟大的浪漫主义诗人杜甫赞其文章“笔落惊风雨，诗成泣鬼神”");
                    //TR2.CharacterFormat.FontName = "华文新魏";
                    //TR2.CharacterFormat.FontSize = 11;
                    //para3.Format.LineSpacing = 15;
                    //para3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                    //para3.Format.SuppressAutoHyphens = true;


                    #endregion

                    #region 更改配置变量
                    if (count == 1)
                    {
                        HorizontalPosition = 300;
                        VerticalPosition = 70;
                    }
                    if (count == 2)
                    {
                        HorizontalPosition = 70;
                        VerticalPosition = 400;
                    }
                    if (count == 3)
                    {
                        HorizontalPosition = 300;
                        VerticalPosition = 400;

                        Spire.Doc.Fields.TextBox _TB = document.Sections[0].Paragraphs[0].AppendTextBox(220, 300);
                        _TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
                        _TB.Format.HorizontalPosition = HorizontalPosition;
                        _TB.Format.VerticalOrigin = VerticalOrigin.Page;
                        _TB.Format.VerticalPosition = VerticalPosition;
                        _TB.Format.TextWrappingStyle = TextWrappingStyle.Tight;
                        _TB.Format.TextWrappingType = TextWrappingType.Both;

                        _TB.Format.LineStyle = TextBoxLineStyle.Simple;
                        _TB.Format.LineColor = Color.Transparent;
                        _TB.Format.LineDashing = LineDashing.Solid;
                        _TB.Format.LineWidth = 3;
                        _TB.Format.FillEfects.Type = BackgroundType.Color;
                        _TB.Format.FillEfects.Color = Color.Green;

                        Paragraph para3 = _TB.Body.AddParagraph();
                        TextRange TR2 = para3.AppendText("  描述--盛唐最杰出的诗人，中国历史最伟大的浪漫主义诗人杜甫赞其文章“笔落惊风雨，诗成泣鬼神”");
                        TR2.CharacterFormat.FontName = "幼圆";
                        TR2.CharacterFormat.FontSize = 15;
                        para3.Format.LineSpacing = 15;
                        para3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                        para3.Format.SuppressAutoHyphens = true;

                    }
                    #endregion
                }

                //Save and Launch  
                document.SaveToFile(textBox1.Text + textBox2.Text, FileFormat.Docx);
                //System.Diagnostics.Process.Start(textBox2.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:" + ex.Message, "Tips", MessageBoxButtons.OK);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            #region 加载一个只含有文本的Word文档
            Document document = new Document();
            document.LoadFromFile(Application.StartupPath + "\\TT.docx");
            #endregion

            #region 创建文本框
            Spire.Doc.Fields.TextBox TB = document.Sections[0].Paragraphs[0].AppendTextBox(150, 300);
            TB.Format.HorizontalOrigin = HorizontalOrigin.Page;
            TB.Format.HorizontalPosition = 370;
            TB.Format.VerticalOrigin = VerticalOrigin.Page;
            TB.Format.VerticalPosition = 155;

            TB.Format.TextWrappingStyle = TextWrappingStyle.Tight;
            TB.Format.TextWrappingType = TextWrappingType.Both;
            #endregion

            #region 设置文本框框的颜色，内部边距，图片填充。
            TB.Format.LineStyle = TextBoxLineStyle.Simple;
            TB.Format.LineColor = Color.Transparent;
            TB.Format.LineDashing = LineDashing.Solid;
            TB.Format.LineWidth = 3;
            TB.Format.FillEfects.Type = BackgroundType.Picture;
            //TB.Format.FillEfects.Picture = Image.FromFile(Application.StartupPath + "\\2.jpg");
            #endregion

            #region 在文本框内添加段落文本，图片，设置字体，字体颜色，行间距，段后距，对齐方式等。然后保存文档，打开查看效果。
            Paragraph para1 = TB.Body.AddParagraph();
            para1.Format.AfterSpacing = 6;
            para1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
            TextRange TR1 = para1.AppendText("标题");
            TR1.CharacterFormat.FontName = "华文新魏";
            TR1.CharacterFormat.FontSize = 16;
            TR1.CharacterFormat.Bold = true;

            Paragraph para2 = TB.Body.AddParagraph();
            Image image = Image.FromFile(Application.StartupPath + "\\李白.jpg");
            DocPicture picture = para2.AppendPicture(image);
            picture.Width = 120;
            picture.Height = 160;
            para2.Format.AfterSpacing = 8;
            para2.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

            Paragraph para3 = TB.Body.AddParagraph();
            TextRange TR2 = para3.AppendText("描述--盛唐最杰出的诗人，中国历史最伟大的浪漫主义诗人杜甫赞其文章“笔落惊风雨，诗成泣鬼神”");
            TR2.CharacterFormat.FontName = "华文新魏";
            TR2.CharacterFormat.FontSize = 11;
            para3.Format.LineSpacing = 15;
            para3.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
            para3.Format.SuppressAutoHyphens = true;

            document.SaveToFile(Application.StartupPath + "\\Testt.docx");
            System.Diagnostics.Process.Start(Application.StartupPath + "\\Testt.docx");
            #endregion
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            docConfig.DocName = textBox2.Text;
            docConfig.SetNodeValue("DocName", docConfig.DocName);
        }
    }
}
