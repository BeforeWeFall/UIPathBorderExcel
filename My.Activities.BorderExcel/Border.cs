using System;
using System.Activities;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace My.Activities.BorderExcel
{
    public class Border : CodeActivity
    {


        private IXLWorkbook book;
        private IXLWorksheet worksheet;

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> PathExcel { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> SheetName { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<String> Cell { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public XLBorderStyleValues StyleBorder { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public System.Drawing.Color Color { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<bool> InsideBorder { get; set; }
        [Category("Input")]
        [RequiredArgument]
        public InArgument<bool> OutsideBorder { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            string cell = Cell.Get(context);
            string path = Regex.Replace(PathExcel.Get(context), @"[^\P{C}\n]+", "");
            string sheetName = SheetName.Get(context);
            bool outside = OutsideBorder.Get(context);
            bool inside = InsideBorder.Get(context);
            XLColor color = XLColor.FromColor(Color);
            SetTip(path, sheetName, cell, outside, inside, color, StyleBorder);

        }
        public void SetTip(string path, string sheetName, string cell, bool outB, bool insideB, XLColor color, XLBorderStyleValues styleB)
        {
            if (File.Exists(path))
            {
                book = new XLWorkbook(path);

                try
                {
                    worksheet = book.Worksheet(sheetName);
                }
                catch
                {
                    worksheet = book.AddWorksheet(sheetName);
                }
            }
            else
            {
                book = new XLWorkbook();
                worksheet = book.AddWorksheet(sheetName);
            }

            if (!cell.Contains(":"))
                TipCell(cell, outB, insideB, color, styleB);
            else
                TipRange(cell, outB, insideB, color, styleB);
            book.Save();
            //wb.SaveAs(filePath);
        }

        private void TipCell(string target, bool outB, bool insideB, XLColor color, XLBorderStyleValues styleB)
        {
            worksheet.Cell(target).Value = Convert.ToDecimal(worksheet.Cell(target).Value);
            if (outB)
            {
                worksheet.Cell(target).Style.Border.OutsideBorderColor = color;
                worksheet.Cell(target).Style.Border.OutsideBorder = styleB;
            }
            if (insideB)
            {
                worksheet.Cell(target).Style.Border.InsideBorderColor = color;
                worksheet.Cell(target).Style.Border.InsideBorder = styleB;
            }

        }
        private void TipRange(string target, bool outB, bool insideB, XLColor color, XLBorderStyleValues styleB)
        {
            string[] range = target.Split(':');

            IXLRange rangeXL;
            if (string.IsNullOrWhiteSpace(range[1]))
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), GetAlfb(worksheet.RangeUsed().FirstRowUsed().CellCount() - 1) + (worksheet.RangeUsed().RowCount()));
            }
            else
            {
                rangeXL = worksheet.Range(range[0].ToUpper(), range[1].ToUpper());
            }

            if (outB)
            {
                rangeXL.Style.Border.OutsideBorderColor = color;
                rangeXL.Style.Border.OutsideBorder = styleB;
            }
            if(insideB)
            {
                rangeXL.Style.Border.InsideBorderColor = color;
                rangeXL.Style.Border.InsideBorder = styleB;
            }
            
        }
        private string GetAlfb(int num)
        {
            return (065 + num) > 90 ? ((char)Math.Floor(64 + (64.0 + num) / 90)).ToString() + ((char)(num % 90)).ToString() : ((char)(065 + num)).ToString();
        }
    }
}