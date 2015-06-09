using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Web.Services.Description;
using System.Net;


namespace WSDLDescriber
{
    public class ApplicationManager
    {

        private List<WsdlPortType> wsdlPortTypeList;
        private XSSchema xsSchema;
        private List<WebMethodInfo> webMethodInfoList;
        private string xmlString;

        private void ExtractWSDLInfo(string wsdlUrl)
        {
            XmlTextReader reader = new XmlTextReader(wsdlUrl.Trim());
            ServiceDescription serviceDescription = ServiceDescription.Read(reader);
            xmlString = "";

            using (var webClient = new WebClient())
            {
                xmlString = webClient.DownloadString(wsdlUrl.Trim());
            }

            #region Getting WSDLPortType and their Operations
            wsdlPortTypeList = new List<WsdlPortType>();
            foreach (PortType portType in serviceDescription.PortTypes)
            {
                WsdlPortType wsdlPortType = new WsdlPortType();
                wsdlPortType.Name = portType.Name;
                wsdlPortType.Operations = new List<WsdlOperation>();
                foreach (Operation op in portType.Operations)
                {
                    WsdlOperation operation = new WsdlOperation();
                    operation.Name = op.Name;
                    operation.InputName = op.Messages.Input.Name;
                    operation.OutputName = op.Messages.Output.Name;
                    if (op.Faults.Count > 0)
                    {
                        operation.FaultNames = new List<string>();

                        foreach (OperationFault of in op.Faults)
                        {
                            operation.FaultNames.Add(of.Name);
                        }
                    }
                    wsdlPortType.Operations.Add(operation);
                }
                wsdlPortTypeList.Add(wsdlPortType);
            }
            #endregion

            #region Getting XSSchema
            Types types = serviceDescription.Types;
            xsSchema = new XSSchema();
            foreach (XmlSchema xmlSchema in types.Schemas)
            {
                foreach (object xmlItem in xmlSchema.Items)
                {
                    XmlSchemaElement schemaElement = xmlItem as XmlSchemaElement;
                    XmlSchemaComplexType complexType = xmlItem as XmlSchemaComplexType;
                    XmlSchemaSimpleType simpleType = xmlItem as XmlSchemaSimpleType;

                    if (schemaElement != null)
                    {
                        XSElement xsElement = new XSElement();
                        xsElement.Name = schemaElement.Name;
                        xsElement.Type = schemaElement.SchemaTypeName.Name;
                        xsElement.MinOccurs = schemaElement.MinOccursString;
                        //if (schemaElement.SchemaType)
                        xsSchema.Elements.Add(xsElement);
                    }
                    else if (complexType != null)
                    {
                        XSComplexType xsComplexType = new XSComplexType();
                        xsComplexType.Name = complexType.Name;
                        XmlSchemaSequence sequence = complexType.Particle as XmlSchemaSequence;
                        XmlSchemaComplexContent complexContent = complexType.ContentModel as XmlSchemaComplexContent;
                        if (complexType.IsAbstract)
                        {
                            xsComplexType.IsAbtract = true;
                        }
                        if (sequence != null)
                        {
                            xsComplexType.Sequence = new List<XSElement>();
                            foreach (XmlSchemaElement childElement in sequence.Items)
                            {
                                XSElement xsElement = new XSElement();
                                xsElement.Name = childElement.Name;
                                xsElement.Type = childElement.SchemaTypeName.Name;
                                xsElement.MinOccurs = childElement.MinOccursString;
                                xsComplexType.Sequence.Add(xsElement);
                            }
                        }
                        if (complexContent != null)
                        {
                            XmlSchemaComplexContentExtension extension = complexContent.Content as XmlSchemaComplexContentExtension;
                            xsComplexType.baseType = extension.BaseTypeName.Name;
                            sequence = extension.Particle as XmlSchemaSequence;
                            if (sequence != null)
                            {
                                xsComplexType.Sequence = new List<XSElement>();
                                foreach (XmlSchemaElement childElement in sequence.Items)
                                {
                                    XSElement xsElement = new XSElement();
                                    xsElement.Name = childElement.Name;
                                    xsElement.Type = childElement.SchemaTypeName.Name;
                                    xsElement.MinOccurs = childElement.MinOccursString;
                                    xsComplexType.Sequence.Add(xsElement);
                                }
                            }
                        }
                        xsSchema.ComplexTypes.Add(xsComplexType);
                    }
                    else if (simpleType != null)
                    {
                        XSSimpleType xsSimpleType = new XSSimpleType();

                        xsSimpleType.Name = simpleType.Name;
                        XmlSchemaSimpleTypeRestriction restriction = simpleType.Content as XmlSchemaSimpleTypeRestriction;
                        if (restriction != null)
                        {
                            List<string> enumerations = new List<string>();
                            //xsSimpleType.Base = simpleType.BaseXmlSchemaType.Name;
                            foreach (XmlSchemaObject facet in restriction.Facets)
                            {
                                if (facet is XmlSchemaEnumerationFacet)
                                    enumerations.Add(((XmlSchemaEnumerationFacet)facet).Value);
                            }
                            xsSimpleType.enumerations = enumerations;
                        }
                        xsSchema.SimpleTypes.Add(xsSimpleType);
                    }
                }

                #region Find Base Types in Complex Types

                foreach (XSComplexType xsComplexType in xsSchema.ComplexTypes)
                {
                    if (xsComplexType.baseType != null)
                    {
                        xsComplexType.Base = xsSchema.ComplexTypes.Where(ct => ct.Name == xsComplexType.baseType).First();
                    }
                }

                #endregion

            }

            #endregion
        }

        public void FillInWebMethodInfo()
        {
            WsdlPortType wsdlPortType = wsdlPortTypeList.ElementAt(0);
            webMethodInfoList = new List<WebMethodInfo>();
            foreach (WsdlOperation op in wsdlPortType.Operations)
            {
                int countElement = xsSchema.Elements.Where(emt => emt.Name == op.InputName).Count();
                int countComplexType = -1;
                int countSimpleType = -1;
                WebMethodInfo webMethodInfo = new WebMethodInfo();
                webMethodInfo.OperationInfo = op;
                XSElement xsElement = null;
                if (countElement == 1)
                {
                    xsElement = new XSElement();
                    xsElement = xsSchema.Elements.Where(emt => emt.Name == op.InputName).First();
                    countElement = xsSchema.Elements.Where(emt => emt.Name == xsElement.Type).Count();
                    countComplexType = xsSchema.ComplexTypes.Where(emt => emt.Name == xsElement.Type).Count();
                    countSimpleType = xsSchema.SimpleTypes.Where(emt => emt.Name == xsElement.Type).Count();
                    if (countComplexType > 0)
                    {
                        webMethodInfo.ComplexTypeInputs = new List<XSComplexType>();
                        if (countComplexType == 1)
                        {
                            List<XSElement> sequence = xsSchema.ComplexTypes.Where(emt => emt.Name == xsElement.Type).First().Sequence;
                            foreach (XSElement element in sequence)
                            {
                                countComplexType = xsSchema.ComplexTypes.Where(emt => emt.Name == element.Type).Count();
                                if (countComplexType > 0)
                                {
                                    webMethodInfo.ComplexTypeInputs.Add(xsSchema.ComplexTypes.Where(emt => emt.Name == element.Type).First());
                                }
                                else
                                {
                                    if (webMethodInfo.ElementInputs == null)
                                    {
                                        webMethodInfo.ElementInputs = new List<XSElement>();
                                    }
                                    webMethodInfo.ElementInputs.Add(element);
                                }

                            }

                        }
                        
                    }
                }

                countElement = xsSchema.Elements.Where(emt => emt.Name == op.OutputName).Count();
                countComplexType = -1;
                countSimpleType = -1;
                if (countElement == 1)
                {
                    xsElement = new XSElement();
                    xsElement = xsSchema.Elements.Where(emt => emt.Name == op.OutputName).First();
                    countElement = xsSchema.Elements.Where(emt => emt.Name == xsElement.Type).Count();
                    countComplexType = xsSchema.ComplexTypes.Where(emt => emt.Name == xsElement.Type).Count();
                    countSimpleType = xsSchema.SimpleTypes.Where(emt => emt.Name == xsElement.Type).Count();
                    if (countComplexType > 0)
                    {
                        List<XSElement> sequence = xsSchema.ComplexTypes.Where(emt => emt.Name == xsElement.Type).First().Sequence;
                        foreach (XSElement element in sequence)
                        {
                            countComplexType = xsSchema.ComplexTypes.Where(emt => emt.Name == element.Type).Count();
                            if (countComplexType > 0)
                            {
                                webMethodInfo.ComplexTypeOutput = xsSchema.ComplexTypes.Where(emt => emt.Name == element.Type).First();
                            }
                            else
                            {
                                webMethodInfo.ElementOutput = element;
                            }
                        }
                    }
                }

                webMethodInfoList.Add(webMethodInfo);

            }
        }

        public void DescribeWSDLInDocument(string wsdlUrl, string saveFileLocation, string authorName,ref int statusInt, ref string statusString)
        {
            statusInt += 10;
            statusString = "Extracting info from WSDL";
            ExtractWSDLInfo(wsdlUrl);
            statusString = "Organizing info extracted from WSDL";
            statusInt += 10;
            FillInWebMethodInfo();
            statusInt += 10;
            statusString = "Beginning to write document";
            WordWriter.WriteDocument(webMethodInfoList, xsSchema, wsdlPortTypeList[0], saveFileLocation, xmlString, authorName, ref statusInt, ref statusString);
            statusString = "Done";
        }
    }
}
