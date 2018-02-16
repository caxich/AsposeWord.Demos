using Aspose.Words;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsposeWord.Demos
{
    public class WordTableDemo
    {
        /// <summary>
        /// 导出word表格文件
        /// </summary>
        /// <returns>文件名</returns>
        public static string Export()
        {
            string mainTitle = "江州市环境保护局";
            string subTitle = "立案登记表";

            //创建文件对象
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //添加标题
            AsposeWordHelper.SetParagraph(builder, ParagraphAlignment.Center, 18);
            AsposeWordHelper.SetHeaderText(builder, mainTitle, subTitle);

            //添加正文（表格）
            AsposeWordHelper.SetParagraph(builder, ParagraphAlignment.Center, 12);
            AsposeWordHelper.StartTable(builder);

            //添加表格行，每一行的单元格宽度相加的总和必须相等，保证表格每一行对齐
            AsposeWordHelper.SetTableRow(builder, 40);
            AsposeWordHelper.SetNormalCellText(builder, 100, "案件来源");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None, "立案号");
            AsposeWordHelper.SetNormalCellText(builder, 60);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 40);
            AsposeWordHelper.SetNormalCellText(builder, 100, "案    由");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 30);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.First, TextOrientation.VerticalFarEast, "当    事    人");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, TextOrientation.Horizontal, "名称或姓名");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.First, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 30);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.Previous);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "地址（住址）");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None, "邮政编码");
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.None, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 30);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.Previous);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "营业执照注册号（公民身份号码）");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None, "组织机构代码");
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.None, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 30);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.Previous);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "社会信用代码");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.First, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 30);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.Previous);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "法定代表人\r\n（负责人）");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.None, CellMerge.None, "职    务");
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.None, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 70);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "案情简介及\r\n立案理由");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None, "\r\n\r\n\r\n          承办人：\r\n                                      年        月        日");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 70);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "承办机构负责人\r\n意见");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None, "\r\n\r\n\r\n          签    名：\r\n                                      年        月        日");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 70);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "环保部门负责人\r\n审批意见");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None, "\r\n\r\n\r\n          签    名：\r\n                                      年        月        日");
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            AsposeWordHelper.SetTableRow(builder, 70);
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.None, CellMerge.None, "备    注");
            AsposeWordHelper.SetMergeCellText(builder, 100, CellMerge.First, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 80, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.SetMergeCellText(builder, 60, CellMerge.Previous, CellMerge.None);
            AsposeWordHelper.EndRow(builder);

            return AsposeWordHelper.SaveDoc(doc, subTitle, "企业名");
        }
    }
}
