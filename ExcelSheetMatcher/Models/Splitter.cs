using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelSheetMatcher.Models
{
    public class Splitter
    {
        public string Value { get; set; }

        public static implicit operator string(Splitter obj)
        {
            return obj.Value;
        }
        public static implicit operator Splitter(string str)
        {
            return new Splitter(str);
        }
        public Splitter(string value)
        {
            this.Value = value;
        }
        public Splitter() { }
    }
}
