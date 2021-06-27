using System;
using System.Collections.Generic;
using System.Text;

namespace FluentEpplus.Common.Processor
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    public class ExcelProcessorAttribute : Attribute
    {
        public Type SimpleProcessorType { get; }
        public Type ComplexProcessorType { get; }

        public ExcelProcessorAttribute(Type simpleProcessorType, Type complexProcessorType)
        {
            SimpleProcessorType = simpleProcessorType;
            ComplexProcessorType = complexProcessorType;
        }
    }
}
