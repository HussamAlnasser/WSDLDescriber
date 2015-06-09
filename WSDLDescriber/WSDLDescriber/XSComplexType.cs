using System;
using System.Collections.Generic;
using System.Text;

namespace WSDLDescriber
{
    public class XSComplexType
    {
        private bool isAbstract = false;
        public string Name { get; set; }
        public string baseType { get; set; }
        public bool IsAbtract
        {
            get
            {
                return isAbstract;
            }
            set
            {
                isAbstract = value;
            }
        }
        public XSComplexType Base { get; set; }
        public List<XSElement> Sequence { get; set; }
    }
}
