using System;
using System.Collections.Generic;
using System.Text;

namespace WSDLDescriber
{
    public class WsdlPortType
    {
        public string Name { get; set; }
        public List<WsdlOperation> Operations { get; set; }
    }
}
