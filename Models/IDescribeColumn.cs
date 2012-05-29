using System;
using System.Linq;
using System.Drawing;

namespace ExcelExport.Models
{
    public interface IDescribeColumn
    {
        string Title { get; }
        float Order { get; }

        /// <summary>
        /// How many characters to show
        /// </summary>
        int Width { get; }

        ColumnType CellType { get; }

        string Value<TModel>(TModel model);

        CellSettings HeaderCellSettings { get; }
        CellSettings ValueCellSettings { get; }
        CellSettings AlternativeCellSettings { get; }

        IDescribeColumn SetOrder(float order);

        /// <summary>
        /// How many characters to show
        /// </summary>
        /// <param name="width"></param>
        /// <returns></returns>
        IDescribeColumn SetWidth(int width);

        IDescribeColumn FormatHeader(Color fontColour, Color cellColour);
        IDescribeColumn FormatCell(Color fontColour, Color cellColour);

        IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property);

        IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property, float order, int width);

        IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property, float order, int width, ColumnType type);
    }

    public class CellSettings 
    {
        public Color BackGroundColour { get; internal set; }
        public Color FontColour { get; internal set; }
    }

    
}