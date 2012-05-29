using System;
using System.Collections.Generic;
using System.Linq;
using System.Collections;

namespace ExcelExport.Models
{
    public class DescribeTheWorkbook
    {
        public DescribeTheWorkbook()
        {
            this.Worksheets = new List<DescribeTheWorkSheet>();
        }

        public IList<DescribeTheWorkSheet> Worksheets { get; private set; }

        public IEnumerable<dynamic> SheetModels { get; private set; }

        public DescribeTheWorkbook AddSheets<TSheetModels>(
            IEnumerable<TSheetModels> sheetModels, 
            Func<TSheetModels, DescribeTheWorkSheet> describeWorksheet)
            where TSheetModels : class
        {
            this.SheetModels = sheetModels;

           

            foreach (var sheetModel in sheetModels)
            {
                var sheet = describeWorksheet(sheetModel);
                
                this.AddSheet(sheet);
            }

            return this;
        }

        public DescribeTheWorkbook AddSheet(DescribeTheWorkSheet worksheet)
        {
            this.Worksheets.Add(worksheet);

            return this;
        }
    }
}