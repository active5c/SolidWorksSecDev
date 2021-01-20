using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SolidWorksSecDev
{
    public class CustomPropertyObject
    {
        public string Name { get;set; }
        public string Value { get;set; }
        public bool Delete { get; set; }
        public string NewValue { get; set; }
        public CustomPropertyObject() { }

        public CustomPropertyObject(string name, string value,string newValue)
        {
            Name = name;
            Value = value;
            NewValue = newValue;
        }
    }
}
