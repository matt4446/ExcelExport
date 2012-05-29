using System;
using System.Linq;
using System.Text;

namespace ExcelExport.Models
{
    public class DescribeColumn : IDescribeColumn
    {
        public DescribeColumn()
        {
            this.CellType = ColumnType.String;

            this.HeaderCellSettings = new CellSettings();
            this.ValueCellSettings = new CellSettings();
            this.AlternativeCellSettings = new CellSettings();
        }

        //Func<dynamic, string> exportValue;
        dynamic exportValue; 

        public string Title 
        { 
            get; 
            private set; 
        }

        public ColumnType CellType
        {
            get;
            private set;
        }

        public string Value<TModel>(TModel model)
        {
            return exportValue(model).ToString();
        }

        public IDescribeColumn Describe(string title)
        {
            this.Title = title.CleanTitle(150);
            return this;
        }

        public IDescribeColumn SetOrder(float order)
        {
            this.Order = order;

            return this;
        }

        public IDescribeColumn SetWidth(int width)
        {
            this.Width = width;

            return this;
        }

        public IDescribeColumn SetColumnType(ColumnType type)
        {
            this.CellType = type;

            return this;
        }

        public IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property)
            //where TModel : class
        {
            this.Title = title;
            this.exportValue = property;
            //Func<dynamic, string> thing = (a) => property(a);

            //this.exportValue = thing;

            return this;
        }

        public IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property, float order, int width)
            //where TModel : class
        {
            return this
                       .Describe(model, title, property)
                       .SetOrder(order)
                       .SetWidth(width);
        }

        #region IDescribeColumn Members

        public float Order { get; private set; }

        public int Width { get; private set; }
        #endregion

        public IDescribeColumn Describe<TModel>(TModel model, string title, Func<TModel, string> property, float order, int width, ColumnType type)
            //where TModel : class
        {
            this.SetColumnType(type);
            return this.Describe(model, title, property, order, width);
        }

        public IDescribeColumn FormatHeader(System.Drawing.Color fontColour, System.Drawing.Color cellColour)
        {
            this.HeaderCellSettings.FontColour = fontColour;
            this.HeaderCellSettings.BackGroundColour = cellColour;
            return this;
        }

        public CellSettings HeaderCellSettings
        {
            get;
            private set;
        }

        public CellSettings ValueCellSettings
        {
            get;
            private set;
        }

        public CellSettings AlternativeCellSettings
        {
            get;
            private set;    
        }

        public IDescribeColumn FormatCell(System.Drawing.Color fontColour, System.Drawing.Color cellColour)
        {
            this.ValueCellSettings.FontColour = fontColour;
            this.ValueCellSettings.BackGroundColour = cellColour;

            return this;
        }

    }

}