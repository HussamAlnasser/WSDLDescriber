using System;
using System.Collections.Generic;
using System.Text;

namespace WSDLDescriber
{
    public class XSElement
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public bool IsComplexType { get; set; }
        public XSElement()
        {
            IsComplexType = false;
        }
        public string MinOccurs { get; set; }
        public string MaxOccurs { get; set; }
    }
}
