using Aspose.Pdf;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Common
{
    public class AsposeHelp
    {
        public static int tighten, min, max, page;
        public static Dictionary<int, int> keyValues = new Dictionary<int, int>() { { 500, 28 }, { 700, 24 }, { 1000, 20 }, { 1300, 18 }, { 1600, 16 }, { 2100, 14 } };

        public bool Help(string oldstr, string newStr)
        {
            try
            {
                using (Presentation presentation = new Presentation(oldstr))
                {
                    ISlideCollection slds = presentation.Slides;
                    for (int i = page; i < slds.Count; i++)
                    {
                        ReplaceTags(slds[i]);
                    }
                    presentation.Save(newStr, Aspose.Slides.Export.SaveFormat.Pptx);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.WriteLog(typeof(AsposeHelp), ex.ToString());
                return false;
            }
            return true;
        }
        public void ReplaceTags(ISlide pSlide)
        {
            foreach (IShape curShape in pSlide.Shapes)
            {
                try
                {
                    if (curShape is IAutoShape)
                    {
                        IAutoShape shape = curShape as IAutoShape;
                        if (shape.TextFrame == null) continue;
                        int textlang = shape.TextFrame.Text.Length;
                        if (textlang == 0) continue;
                        int key = keyValues.Keys.First(a => a > textlang);
                        int fontSize = 0;
                        keyValues.TryGetValue(key, out fontSize);
                        foreach (IParagraph para in shape.TextFrame.Paragraphs)
                        {
                            if (!string.IsNullOrEmpty(para.Text))
                            {
                                List<FontProperty> list = new List<FontProperty>();
                                foreach (IPortion range in para.Portions)
                                {
                                    list.Add(new FontProperty() { text = range.Text.Trim(), portion = range, size = range.PortionFormat.FontHeight });
                                }
                                if (list.Where(a => a.size.Equals(float.NaN)).Count() > 0)
                                {
                                    float size = list.First(a => a.size.Equals(float.NaN)).size;
                                    foreach (var item in list)
                                    {
                                        if (float.IsNaN(item.size)) item.size = size;
                                        if (float.IsNaN(item.size)) item.size = fontSize;
                                    }
                                }

                                para.Portions.Clear();
                                foreach (var item in list)
                                {
                                    Random r = new Random();
                                    int number = r.Next(min, max);
                                    List<string> strList = subStringByCount(item.text, number);
                                    IPortion portion = Voluation(item.portion);
                                    if (float.IsNaN(portion.PortionFormat.FontHeight)) portion.PortionFormat.FontHeight = item.size;
                                    if (strList.Count == 0)
                                        portion.Text = "";
                                    else
                                        portion.Text = strList[0];
                                    para.Portions.Add(portion);
                                    for (int i = 1; i < strList.Count; i++)
                                    {
                                        string str = strList[i].Substring(0, 1);
                                        if (CheckStringChinese(str))
                                        {
                                            portion = Voluation(item.portion);
                                            if (float.IsNaN(portion.PortionFormat.FontHeight)) portion.PortionFormat.FontHeight = item.size;
                                            portion.Text = str;
                                            portion.PortionFormat.Spacing = -item.size;
                                            para.Portions.Add(portion);
                                        }
                                        portion = Voluation(item.portion);
                                        if (float.IsNaN(portion.PortionFormat.FontHeight)) portion.PortionFormat.FontHeight = item.size;
                                        portion.Text = strList[i];
                                        para.Portions.Add(portion);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.WriteLog(typeof(AsposeHelp), ex.ToString());
                }
            }
        }

        public Portion Voluation(IPortion portion)
        {
            Portion item = new Portion();
            item.PortionFormat.LanguageId = portion.PortionFormat.LanguageId;
            item.PortionFormat.KerningMinimalSize = portion.PortionFormat.KerningMinimalSize;
            item.PortionFormat.Escapement = portion.PortionFormat.Escapement;
            item.PortionFormat.SymbolFont = portion.PortionFormat.SymbolFont;
            item.PortionFormat.ComplexScriptFont = portion.PortionFormat.ComplexScriptFont;
            item.PortionFormat.EastAsianFont = portion.PortionFormat.EastAsianFont;
            item.PortionFormat.LatinFont = portion.PortionFormat.LatinFont;
            item.PortionFormat.FontHeight = portion.PortionFormat.FontHeight;
            item.PortionFormat.IsHardUnderlineFill = portion.PortionFormat.IsHardUnderlineFill;
            item.PortionFormat.IsHardUnderlineLine = portion.PortionFormat.IsHardUnderlineLine;
            item.PortionFormat.StrikethroughType = portion.PortionFormat.StrikethroughType;
            item.PortionFormat.TextCapType = portion.PortionFormat.TextCapType;
            item.PortionFormat.FontUnderline = portion.PortionFormat.FontUnderline;
            item.PortionFormat.ProofDisabled = portion.PortionFormat.ProofDisabled;
            item.PortionFormat.NormaliseHeight = portion.PortionFormat.NormaliseHeight;
            item.PortionFormat.Kumimoji = portion.PortionFormat.Kumimoji;
            item.PortionFormat.FontItalic = portion.PortionFormat.FontItalic;
            item.PortionFormat.FontBold = portion.PortionFormat.FontBold;
            item.PortionFormat.AlternativeLanguageId = portion.PortionFormat.AlternativeLanguageId;
            item.PortionFormat.FillFormat.SolidFillColor.Color = portion.PortionFormat.FillFormat.SolidFillColor.Color;
            return item;
        }

        /// <summary>
        /// 判断字符串是否是数字
        /// </summary>
        public static bool IsNumber(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            const string pattern = "^[0-9a-zA-Z]*$";
            Regex rx = new Regex(pattern);
            //var patrn = "^[`~!@#$%^&*()_\\-+=<>?:\"{ }|,.\\/;'\\[\\]·~！@#￥%……&*（）——\\-+={}|《》？：“”【】、；‘'，。、]$";
            return rx.IsMatch(s);
        }
        /// <summary>
        /// 用 ASCII 码范围判断字符是不是汉字
        /// </summary>
        /// <param name="text">待判断字符或字符串</param>
        /// <returns>真：是汉字；假：不是</returns>
        public bool CheckStringChinese(string text)
        {
            bool res = false;
            foreach (char t in text)
            {
                if ((int)t > 127)
                    res = true;
            }
            return res;
        }
        public static List<string> subStringByCount(string text, int count)
        {
            int start_index = 0;//开始索引
            int end_index = count - 1;//结束索引
            double count_value = 1.0 * text.Length / count;
            double newCount = Math.Ceiling(count_value);//向上取整，只有有小数就取整，比如3.14，结果4
            List<string> list = new List<string>();
            for (int i = 0; i < newCount; i++)
            {
                //如果end_index大于字符长度，则添加剩下字符串
                if (end_index > text.Length - 1)
                {
                    list.Add(text.Substring(start_index).Trim());
                    break;
                }
                else
                {
                    list.Add(text.Substring(start_index, count).Trim());

                    start_index += count;
                    end_index += count;
                }
            }

            return list.Where(a => !string.IsNullOrEmpty(a)).ToList();
        }
    
        /// <summary>
        /// pdf转doc
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        public static void PdfOrDoc(string sourcePath, string targetPath)
        {
            try
            {
                using (Document pdfDocument = new Document(sourcePath))
                {
                    DocSaveOptions saveOptions = new DocSaveOptions();
                    saveOptions.Format = DocSaveOptions.DocFormat.Doc;
                    pdfDocument.Save(targetPath, saveOptions);
                    pdfDocument.Dispose();
                }
            }
            catch (Exception ex)
            {
            }
        }
    }
}
