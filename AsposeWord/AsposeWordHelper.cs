﻿using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeWord
{
    public class AsposeWordHelper
    {
        #region word导出帮助类
        /// <summary>
        /// 设置段落格式
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="alignment">段落对齐方式</param>
        /// <param name="lineSpacing">行距</param>
        public static void SetParagraph(DocumentBuilder builder, ParagraphAlignment alignment, int lineSpacing)
        {
            var ph = builder.ParagraphFormat;
            ph.Alignment = alignment;
            // One line equals 12 points. so 1.5 lines = 18 points
            ph.LineSpacing = lineSpacing;
        }

        /// <summary>
        /// 设置段落格式
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="alignment">段落对齐方式</param>
        /// <param name="lineSpacing">行距</param>
        /// <param name="firstLineIndent">首行缩进长度</param>
        public static void SetParagraph(DocumentBuilder builder, ParagraphAlignment alignment, int lineSpacing, 
            int firstLineIndent)
        {
            var ph = builder.ParagraphFormat;
            ph.Alignment = alignment;
            // One line equals 12 points. so 1.5 lines = 18 points
            ph.LineSpacing = lineSpacing;
            ph.FirstLineIndent = firstLineIndent;
        }


        public static void SetHeaderText(DocumentBuilder builder, string mainTitle, string subTitle = "")
        {
            builder.Font.Size = 12;
            builder.Font.Bold = true;
            builder.Writeln(mainTitle);
            if (!string.IsNullOrEmpty(subTitle))
            {
                builder.Writeln(subTitle);
            }
            builder.Writeln("");
        }

        /// <summary>
        /// 添加文字
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="text"></param>
        public static void SetLabelText(DocumentBuilder builder, string text)
        {
            builder.Write(text);
        }

        /// <summary>
        /// 添加文字（带回车）
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="text"></param>
        public static void SetLabelTextln(DocumentBuilder builder, string text)
        {
            builder.Writeln(text);
        }

        /// <summary>
        /// 添加回车
        /// </summary>
        /// <param name="builder"></param>
        public static void SetTextLn(DocumentBuilder builder)
        {
            builder.Writeln();
        }

        /// <summary>
        /// 填写文字，如果为空则填充空格
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="text"></param>
        public static void SetValueText(DocumentBuilder builder, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                builder.Write(text + "    ");
            }
            else
            {
                builder.Write("        ");
            }
        }

        public static void SetValueTextln(DocumentBuilder builder, string text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                builder.Writeln(text);
            }
            else
            {
                builder.Writeln("        ");
            }
        }

        /// <summary>
        /// 设置文字格式
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="size">字体大小</param>
        /// <param name="isBold">是否粗体</param>
        /// <param name="under">是否有下划线</param>
        public static void SetFont(DocumentBuilder builder, int size, bool isBold, Underline under = Underline.None)
        {
            Font font = builder.Font;
            font.Size = size;
            font.Bold = isBold;
            font.Underline = under;
        }

        /// <summary>
        /// 开始添加表格
        /// </summary>
        /// <param name="builder"></param>
        public static void StartTable(DocumentBuilder builder)
        {
            builder.StartTable();
            builder.Font.Size = 11;
            builder.Font.Bold = false;
        }

        /// <summary>
        /// 结束表格，设置自适应及居中显示
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="doc"></param>
        public static void EndTable(DocumentBuilder builder, Document doc)
        {
            builder.EndTable();
            //表格宽度自适应页面
            doc.FirstSection.Body.Tables[0].PreferredWidth = PreferredWidth.Auto;
            //表格在页面中居中
            doc.FirstSection.Body.Tables[0].Alignment = TableAlignment.Center;
        }

        /// <summary>
        /// 设置表格行高
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="rowHeight"></param>
        public static void SetTableRow(DocumentBuilder builder, int rowHeight)
        {
            RowFormat rowf = builder.RowFormat;
            rowf.Height = rowHeight;
        }

        /// <summary>
        /// 结束表格行
        /// </summary>
        /// <param name="builder"></param>
        public static void EndRow(DocumentBuilder builder)
        {
            builder.EndRow();
        }

        /// <summary>
        /// 添加合并单元格
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellWidth"></param>
        /// <param name="hMerge">水平方向合并</param>
        /// <param name="vMerge">垂直方向合并</param>
        /// <param name="cellText"></param>
        public static void SetMergeCellText(DocumentBuilder builder, int cellWidth, CellMerge hMerge, CellMerge vMerge, string cellText = "")
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = hMerge;
            builder.CellFormat.VerticalMerge = vMerge;
            builder.CellFormat.Width = cellWidth;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            //单元格水平对齐方向
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.FitText = true;//单元格内文字为多行（默认为单行，会影响单元格宽）
            if (!string.IsNullOrEmpty(cellText))
            {
                builder.Write(cellText);
            }
        }

        /// <summary>
        /// 添加合并单元格，可设置文字方向
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellWidth"></param>
        /// <param name="hMerge">水平方向合并</param>
        /// <param name="vMerge">垂直方向合并</param>
        /// <param name="textOri">文字方向</param>
        /// <param name="cellText"></param>
        public static void SetMergeCellText(DocumentBuilder builder, int cellWidth, CellMerge hMerge, CellMerge vMerge, TextOrientation textOri, string cellText = "")
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = hMerge;
            builder.CellFormat.VerticalMerge = vMerge;
            builder.CellFormat.Width = cellWidth;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            //单元格水平对齐方向
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Orientation = textOri;
            builder.CellFormat.FitText = true;//单元格内文字为多行（默认为单行，会影响单元格宽）
            if (!string.IsNullOrEmpty(cellText))
            {
                builder.Write(cellText);
            }
        }

        /// <summary>
        /// 添加单元格并添加文字
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellWidth"></param>
        /// <param name="cellText"></param>
        public static void SetNormalCellText(DocumentBuilder builder, int cellWidth, string cellText = "")
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.Width = cellWidth;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            //单元格水平对齐方向
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.FitText = true;//单元格内文字为多行（默认为单行，会影响单元格宽）
            if (!string.IsNullOrEmpty(cellText))
            {
                builder.Write(cellText);
            }
        }

        /// <summary>
        /// 保存doc文件
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="fileName">自定义文件名</param>
        /// <param name="name">自定义文件名</param>
        /// <returns>完整文件名</returns>
        public static string SaveDoc(Document doc, string fileName, string name)
        {
            string filepath = fileName + DateTime.Now.ToString("yyyy-MM-dd") + name + ".doc";
            doc.Save(filepath);
            return filepath;
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="imagePath">图片路径</param>
        public static void SetImage(DocumentBuilder builder, string imagePath)
        {
            builder.InsertImage(imagePath);
        }

        /// <summary>
        /// 可设置文字在单元格内是否水平居中
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellWidth"></param>
        /// <param name="cellAlign">文字水平对齐方向</param>
        /// <param name="cellText"></param>
        public static void SetValueCellText(DocumentBuilder builder, int cellWidth, ParagraphAlignment cellAlign, string cellText = "")
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.Width = cellWidth;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            //单元格水平对齐方向
            builder.ParagraphFormat.Alignment = cellAlign;

            builder.CellFormat.FitText = true;//单元格内文字为多行（默认为单行，会影响单元格宽）
            if (!string.IsNullOrEmpty(cellText))
            {
                builder.Write(cellText);
            }
        }

        /// <summary>
        /// 填写单元格文字（日期格式）
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="cellWidth"></param>
        /// <param name="cellAlign">文字水平对齐方向</param>
        /// <param name="cellText"></param>
        public static void SetDateCellText(DocumentBuilder builder, int cellWidth, ParagraphAlignment cellAlign, DateTime dt)
        {
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.CellFormat.Width = cellWidth;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            //单元格水平对齐方向
            builder.ParagraphFormat.Alignment = cellAlign;

            builder.CellFormat.FitText = true;//单元格内文字为多行（默认为单行，会影响单元格宽）

            builder.Write(dt.Year.ToString());
            builder.Write("年");
            builder.Write(dt.Month.ToString());
            builder.Write("月");
            builder.Write(dt.Day.ToString());
            builder.Write("日");
        }

        /// <summary>
        /// 写入时间，精确到日
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="dt"></param>
        public static void SetDate(DocumentBuilder builder, DateTime dt)
        {
            builder.Write(dt.Year.ToString());
            builder.Write("年");
            builder.Write(dt.Month.ToString());
            builder.Write("月");
            builder.Write(dt.Day.ToString());
            builder.Write("日");
        }

        /// <summary>
        /// 写入时间，精确到分
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="dt"></param>
        public static void SetDateTime(DocumentBuilder builder, DateTime dt)
        {
            builder.Write(dt.Year.ToString());
            builder.Write("年");
            builder.Write(dt.Month.ToString());
            builder.Write("月");
            builder.Write(dt.Day.ToString());
            builder.Write("日");
            builder.Write(dt.Hour.ToString());
            builder.Write("时");
            builder.Write(dt.Minute.ToString());
            builder.Write("分");
        }

        /// <summary>
        /// 写入时间段，精确到日
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="dt"></param>
        public static void SetDate(DocumentBuilder builder, DateTime dtStart, DateTime dtEnd)
        {
            builder.Write(dtStart.Year.ToString());
            builder.Write("年");
            builder.Write(dtStart.Month.ToString());
            builder.Write("月");
            builder.Write(dtStart.Day.ToString());
            builder.Write("日至");
            builder.Write(dtEnd.Month.ToString());
            builder.Write("月");
            builder.Write(dtEnd.Day.ToString());
            builder.Write("日");
        }

        /// <summary>
        /// 写入时间段，精确到分
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="dt"></param>
        public static void SetDateTime(DocumentBuilder builder, DateTime dtStart, DateTime dtEnd)
        {
            builder.Write(dtStart.Year.ToString());
            builder.Write("年");
            builder.Write(dtStart.Month.ToString());
            builder.Write("月");
            builder.Write(dtStart.Day.ToString());
            builder.Write("日");
            builder.Write(dtStart.Hour.ToString());
            builder.Write("时");
            builder.Write(dtStart.Minute.ToString());
            builder.Write("分至");
            builder.Write(dtEnd.Month.ToString());
            builder.Write("月");
            builder.Write(dtEnd.Day.ToString());
            builder.Write("日");
            builder.Write(dtEnd.Hour.ToString());
            builder.Write("时");
            builder.Write(dtEnd.Minute.ToString());
            builder.Write("分");
        }

        #endregion
    }
}
