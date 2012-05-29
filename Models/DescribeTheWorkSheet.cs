using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExport.Models
{
    public class DescribeTheWorkSheet
    {
        public DescribeTheWorkSheet()
        {
            this.RowDescriptors = new List<IDescribeRow>();
        }

        public string Title { get; private set; }

        public bool KeepHeadersAtTop { get; private set; }
        
        public IEnumerable<dynamic> SheetModel { get; private set; }

        public IDescribeRow ColumnRowDescriptor { get; set; }

        public IList<IDescribeRow> RowDescriptors { get; private set; }

        public DescribeTheWorkSheet SetTitle(string title)
        {
            this.Title = title.CleanTitle(20);

            return this;
        }

        //public DescribeTheWorkSheet SetData<TModel>(IEnumerable<TModel> items)
        //    where TModel : class
        //{
        //    this.Model = items;
        //    //Func<TModel, dynamic> change = c => c;

        //    return this;
        //}

        public DescribeTheWorkSheet SetFrozenPane(bool state)
        {
            this.KeepHeadersAtTop = state;

            return this;
        }
        
        public DescribeTheWorkSheet RowDefinition<TModel>(IEnumerable<TModel> items, Action<TModel, IDescribeRow> describeRow) 
            where TModel : class
        {
            this.SheetModel = items;

            var columnRowDescription = new DescribeRow(this);
            var header = default(TModel);
            describeRow(header, columnRowDescription);

            this.ColumnRowDescriptor = columnRowDescription;

            foreach(var item in items)
            {
                var rowDescription = new DescribeRow(this);
                describeRow(item, rowDescription);
                this.RowDescriptors.Add( rowDescription );
            }

            return this;
        }


       
    }
}