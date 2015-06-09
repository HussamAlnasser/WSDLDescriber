using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace WSDLDescriber
{
    public class WebMethodInfo
    {
        public string PortType { get; set; }
        public WsdlOperation OperationInfo { get; set; }
        public List<XSElement> ElementInputs { get; set; }
        public List<XSComplexType> ComplexTypeInputs { get; set; }
        public XSElement ElementOutput { get; set; }
        public XSComplexType ComplexTypeOutput { get; set; }
    }
}
