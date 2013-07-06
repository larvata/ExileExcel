namespace ExileExcel.Attribute
{
    using System;
    using System.ComponentModel;

    /// <summary>
    /// Mark class can be find by ExileParser
    /// </summary>
    [Description("Exiliable Class Attribute")]
    [AttributeUsage(AttributeTargets.Class)]
    public class ExiliableAttribute : Attribute
    {
        // descirption of class
        private readonly string _description;

        public string Description 
        { 
            get { return _description; } 
        }

        public ExiliableAttribute(string description="")
        {
            _description = description;
        }
    }
}
