using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WSDLDescriber
{
    public class XSSimpleType
    {
        public string Name { get; set; }
        public string Base { get; set; }
        public List<string> enumerations { get; set; }
    }
}
