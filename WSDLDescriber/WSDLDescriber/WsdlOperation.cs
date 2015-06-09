using System;
using System.Collections.Generic;
using System.Text;

namespace WSDLDescriber
{
    public class WsdlOperation
    {
        public string Name { get; set; }
        public string InputName { get; set; }
        public string OutputName { get; set; }
        public List<string> FaultNames { get; set; }
    }
}
