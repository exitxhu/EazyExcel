using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EazyExcel
{
	[AttributeUsage(AttributeTargets.Property)]
    public sealed class ColumnNameOrderAttribute:Attribute
    {
		private string _displayName;

		public string DisplayName
		{
			get { return _displayName; }
			set { _displayName = value; }
		}
		private int _order;

		public int Order
		{
			get { return _order; }
			set { _order = value; }
		}

		private Type _resourceType;

		public Type ResourceType
		{
			get { return _resourceType; }
			set { _resourceType = value; }
		}

		public ColumnNameOrderAttribute(string displayName,int order)
		{
			_displayName = displayName;
			_order = order;
		}

        public ColumnNameOrderAttribute(string displayName)
        {
            _displayName = displayName;
        }

        public ColumnNameOrderAttribute(string displayName, int order,Type resourceType)
        {
            _displayName = displayName;
            _order = order;
			_resourceType = resourceType;
        }
    }
}
