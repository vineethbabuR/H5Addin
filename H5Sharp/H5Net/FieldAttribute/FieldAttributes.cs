using System;

namespace H5Net
{
    [AttributeUsage(AttributeTargets.Property)]
    public class FieldAttribute : Attribute
    {
        public string FieldName { get; set; }
        public bool Mandatory { get; set; }

        public FieldAttribute(string fieldName, bool mandatory = false)
        {
            this.FieldName = fieldName;
            this.Mandatory = mandatory;
        }
    }
}