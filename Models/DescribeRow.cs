using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExport.Models
{
    public class DescribeRow : IDescribeRow
    {
        public DescribeRow(DescribeTheWorkSheet worksheetDescription)
        {
            this.ColumnDescriptor = new List<IDescribeColumn>();
            this.WorkSheetDescription = worksheetDescription;
        }

        public IList<IDescribeColumn> ColumnDescriptor { get; private set; }

        public DescribeTheWorkSheet WorkSheetDescription { get; private set; }


        public IDescribeRow Columns<TRowModel>(TRowModel model, Func<TRowModel, IEnumerable<Action<DescribeColumn>>> describeAction)
        {
            this.Model = model;

            var itemDescriptiors = describeAction(model);
            foreach (var columnDescription in itemDescriptiors) {
                var describeColumn = new DescribeColumn();
                columnDescription(describeColumn);
                this.ColumnDescriptor.Add(describeColumn);
            }

            return this;
        }

        public dynamic Model
        {
            get;
            private set;
        }


        #region IDescribeRow Members


        public IDescribeRow Column<TRowModel>(TRowModel model, Action<IDescribeColumn> describeAction)
        {
            this.Model = model;

            var describeColumn = new DescribeColumn();
            describeAction(describeColumn);

            this.ColumnDescriptor.Add(describeColumn);

            return this;
        }

        #endregion
    }
}