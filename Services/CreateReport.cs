using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using NPOI.SS.UserModel;
using Orchard;
using NPOI.HSSF.UserModel;
using ExcelExport.Models;
using System.IO;
using NPOI.HSSF.Util;

namespace ExcelExport.Services
{
    public interface IReportService : IDependency 
    {
        void Render(DescribeTheWorkbook workbook);

        byte[] Respond();
        byte[] RenderAndRespond(DescribeTheWorkbook workbook);
    }

    public class ReportService : IReportService
    {
        private CellStyle dateStyle;

        private CellStyle hyperlinkStyle;

        public ReportService() 
        {
            this.Workbook = new HSSFWorkbook();
            this.Init();
        }

        private void Init()
        {
            #region Set Styles 
            
            //date style 
            this.dateStyle = this.Workbook.CreateCellStyle();
            dateStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("d/m/yy h:mm");

            //hyperlink style
            this.hyperlinkStyle = this.Workbook.CreateCellStyle();
            var font = this.Workbook.CreateFont();

            font.Underline = (byte)FontUnderlineType.SINGLE;
            font.Color = HSSFColor.BLUE.index;
            hyperlinkStyle.SetFont(font);

            #endregion
        }

        private HSSFWorkbook Workbook { get; set; }

        private void AddSheets(HSSFWorkbook workbook, IEnumerable<DescribeTheWorkSheet> worksheets) 
        {
            int count = 1;

            foreach (var item in worksheets)
            { 
                var sameName = worksheets.Where(sheet => item.Title.Equals(sheet.Title, StringComparison.CurrentCultureIgnoreCase)).ToList();

                if (sameName.Count == 1)
                    continue;

                foreach (var same in sameName)
                {
                    var title = same.Title;
                    same.SetTitle(count.ToString() + title);
                    count++;
                }

                count = 1;
            }

            foreach (var item in worksheets) 
            {
                if (item.SheetModel.Count() == 0)
                    continue;
                this.AddWorkSheet(workbook, item);
            }
        }
  
        private void AddWorkSheet(HSSFWorkbook workbook, DescribeTheWorkSheet item)
        {
            var worksheet = workbook.CreateSheet(item.Title);

            this.AddRows(worksheet, item.ColumnRowDescriptor, item.RowDescriptors);
        }

        private void AddRows(Sheet worksheet, IDescribeRow columnRow, IEnumerable<IDescribeRow> rows)
        {
            var columns = columnRow.ColumnDescriptor.OrderBy(e => e.Order).ToList();
            var buildRows = rows.ToList(); // = //columnRow.WorkSheetDescription.SheetModel.ToList(); 

            int startRow = 0;

            var rowPane = worksheet.CreateRow(startRow);

            this.AddColumnHeaders(worksheet, rowPane, columns);

            for (int i = 0; i < buildRows.Count; i++) 
            {
                IDescribeRow describeRow = buildRows[i];

                columns = describeRow.ColumnDescriptor.OrderBy(e => e.Order).ToList();
                rowPane = worksheet.CreateRow(startRow + 1);

                var value = describeRow.Model;
                
                AddColumnValues(rowPane, columns, value);
                startRow++;
            }
        }

        private void SetCellStyle(IDescribeColumn description, Cell cell) 
        {
            
        }

        private void AddColumnHeaders(Sheet worksheet, Row rowPane, List<IDescribeColumn> columns)
        {
            for (int i = 0; i < columns.Count; i++) 
            {
                var cell = rowPane.CreateCell(i);
                var column = columns[i];
                cell.SetCellValue(column.Title);

                if(column.Width > 0)
                    worksheet.SetColumnWidth(i, column.Width * 256);
            }
        }

        private void AddColumnValues(Row rowPane, List<IDescribeColumn> columns, dynamic model)
        {
            for (int i = 0; i < columns.Count; i++) 
            {
                var cell = rowPane.CreateCell(i);

                string value = columns[i].Value(model);

                cell.SetCellValue(columns[i].Value(model));
                
                switch (columns[i].CellType)
                {
                    case ColumnType.String: 
                        { 
                            cell.SetCellType(CellType.STRING); 
                            break; 
                        }
                    case ColumnType.Number: cell.SetCellType(CellType.NUMERIC); break;
                    case ColumnType.Date: cell.CellStyle = dateStyle; break;
                    case ColumnType.Link: 
                        { 
                            cell.CellStyle = hyperlinkStyle;
                            cell.Hyperlink = new HSSFHyperlink(HyperlinkType.URL);
                            cell.Hyperlink.Address = value;
                            break; 
                        }
                    case ColumnType.Email:
                        {
                            cell.CellStyle = hyperlinkStyle;
                            string link = 
                                value.IndexOf("mailto:", StringComparison.CurrentCultureIgnoreCase) >= 0 ? 
                                    value : 
                                    string.Format("mailto:{0}", value ?? String.Empty);

                            cell.Hyperlink = new HSSFHyperlink(HyperlinkType.EMAIL);
                            cell.Hyperlink.Address = link;

                            break;
                        }
                }
            }
        }

        public void Render(DescribeTheWorkbook workbookDescription)
        {
            var workbook = this.Workbook ?? (this.Workbook = new HSSFWorkbook());

            this.AddSheets(workbook, workbookDescription.Worksheets);
        }

        public byte[] Respond()
        {
            using (var memoryStream = new MemoryStream())
            {
                this.Workbook.Write(memoryStream);
                return memoryStream.ToArray();
            }
        }

        public byte[] RenderAndRespond(DescribeTheWorkbook workbook) 
        {
            Render(workbook);

            return Respond();
        }


    }
}