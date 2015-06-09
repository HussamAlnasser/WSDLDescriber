using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WSDLDescriber
{
    public class XSSchema
    {
        private List<XSComplexType> complexTypes;
        private List<XSElement> elements;
        private List<XSSimpleType> simpleTypes;

        public XSSchema()
        {
            complexTypes = new List<XSComplexType>();
            elements = new List<XSElement>();
            simpleTypes = new List<XSSimpleType>();
        }

        public List<XSComplexType> ComplexTypes
        {
            get
            {
                return complexTypes;
            }
        }
        public List<XSElement> Elements 
        { 
            get
            {
                return elements;
            }
        }
        public List<XSSimpleType> SimpleTypes 
        { 
            get
            {
                return simpleTypes;
            }
        }
    }
}
