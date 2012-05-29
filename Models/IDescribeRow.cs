using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelExport.Models
{
    public interface IDescribeRow
    {
        dynamic Model { get; }

        DescribeTheWorkSheet WorkSheetDescription { get; }

        IList<IDescribeColumn> ColumnDescriptor { get; }

        IDescribeRow Column<TRowModel>(TRowModel model, Action<IDescribeColumn> describeAction);
        IDescribeRow Columns<TRowModel>(TRowModel model, Func<TRowModel, IEnumerable<Action<DescribeColumn>>> describeAction);
        //IDescribeRow Columns<TRowModel>(TRowModel model, Action<IEnumerable<IDescribeColumn>> describe);
    }
}