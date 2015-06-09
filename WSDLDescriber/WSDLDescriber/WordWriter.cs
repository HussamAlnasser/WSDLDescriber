using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

namespace WSDLDescriber
{
    public class WordWriter
    {
        private static Document wordDoc;
        private static Application wordApp;
        private static Paragraph newParagraph;
        private static ListTemplate listTemplate;
        private static XSSchema xsSchema;
        private static object oMissing = System.Reflection.Missing.Value;  

        private static void CreateListTemplate()
        {
            listTemplate = wordDoc.ListTemplates.Add(true, "WSDLDescriberIndexTemplate");
            for (int i = 1; i <= 6; i++)
            {
                listTemplate.ListLevels[i].NumberFormat = "";
                for (int j = 1; j <= i; j++)
                    listTemplate.ListLevels[i].NumberFormat = listTemplate.ListLevels[i].NumberFormat + "%" + j.ToString() + ".";
                listTemplate.ListLevels[i].NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            }


        }

        private static void CreateCoverPage(string authorName, string portTypeName)
        {
            #region Service Name Title
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "";
            newParagraph.Range.Font.Color = WdColor.wdColorBlack;
            newParagraph.Range.Font.Bold = 1;
            newParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorBlue;
            newParagraph.Range.Font.Size = 40;
            newParagraph.Range.Font.Name = "Cambria";
            newParagraph.Range.Text = "\n\n\n"+portTypeName;
            newParagraph.Range.InsertParagraphAfter();
            #endregion
            #region WSDL Description Subtitle
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "";
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].Color = WdColor.wdColorBlack;
            newParagraph.Range.Font.Bold = 1;
            newParagraph.Range.Font.Size = 22;
            newParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            newParagraph.Range.Font.Name = "Cambria";
            newParagraph.Range.Text = "WSDL Description";
            newParagraph.Range.InsertParagraphAfter();
            #endregion
            #region Author Name
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Font.Bold = 1;
            newParagraph.Range.Font.Size = 11;
            newParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            newParagraph.Range.Font.Name = "Calibri";
            newParagraph.Range.Text = authorName;
            newParagraph.Range.InsertParagraphAfter();
            #endregion
            #region Date
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            newParagraph.Range.Font.Bold = 1;
            newParagraph.Range.Font.Size = 11;
            newParagraph.Range.Font.Name = "Calibri";
            newParagraph.Range.Text = DateTime.Today.ToString("MMMM d, yyyy");
            newParagraph.Range.Text += "";
            newParagraph.Range.InsertParagraphAfter();
            #endregion
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "";
            newParagraph.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            object oPageBreak = WdBreakType.wdPageBreak;
            newParagraph.Range.InsertBreak(oPageBreak);
        }

        private static void CreateWSDLAppendix(string wsdlXmlString)
        {
            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "Appendix-WSDL";
            newParagraph.Range.Font.Color = WdColor.wdColorBlue;

            newParagraph.Range.Font.Size = 14;

            //newParagraph.Range.Font.Bold = 1;

            newParagraph.Range.Font.Name = "Times New Roman";
            newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
            newParagraph.Range.SetListLevel(1);

            newParagraph.Range.InsertParagraphAfter();

            newParagraph.Range.ListFormat.RemoveNumbers();

            #region WSDL XML
            newParagraph = wordDoc.Content.Paragraphs.Add();


            newParagraph.Range.Text = "";

            int rgbColor = Information.RGB(79, 129, 189);

            WdColor wdColorBackgroundColor = (WdColor)rgbColor;
            newParagraph.Range.Shading.BackgroundPatternColor = wdColorBackgroundColor;

            newParagraph.Range.Font.Name = "Arial";

            newParagraph.Range.Font.Size = 11;

            newParagraph.Range.Font.Color = WdColor.wdColorWhite;
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;

            
            newParagraph.Range.Text = wsdlXmlString;
            newParagraph.Range.Text += "";
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
            newParagraph.Range.Font.Color = WdColor.wdColorBlue;
            newParagraph.Range.Font.Name = "Times New Roman";
            newParagraph.Range.Font.Size = 12;
            newParagraph.Range.InsertParagraphAfter();
            #endregion
        }

        private static string CreateXSDString(XSComplexType complexType)
        {
            
            string complexTypeXSD = "";
            if (complexType.Base != null)
            {
                #region Inhereted
                if (complexType.IsAbtract)
                {
                    complexTypeXSD += "<xs:complexType abstract=\"true\" name=\"" + complexType.Name + "\">\n\t<xs:complexContent>\n\t\t<xs:extension base=\"tns:" + complexType.Base.Name + "\">\t\t\t<xs:sequence>";
                }
                else
                {
                    complexTypeXSD += "<xs:complexType name=\"" + complexType.Name + "\">\n\t<xs:complexContent>\n\t\t<xs:extension base=\"tns:" + complexType.Base.Name + "\">\n\t\t\t<xs:sequence>\n";
                }

                foreach (XSElement sequenceElement in complexType.Sequence)
                {
                    complexTypeXSD += "\t\t\t\t<xs:element ";
                    if (sequenceElement.MinOccurs != null)
                    {
                        complexTypeXSD += "minOccurs=\"" + sequenceElement.MinOccurs + "\" ";
                    }

                    complexTypeXSD += "name=\"" + sequenceElement.Name + "\" ";
                    if (xsSchema.SimpleTypes.Where(st => st.Name == sequenceElement.Type).Count() > 0 || xsSchema.ComplexTypes.Where(xsct => xsct.Name == sequenceElement.Type).Count() > 0)
                    {
                        complexTypeXSD += "type=\"tns:" + sequenceElement.Type + "\"/>\n";
                    }
                    else
                    {
                        complexTypeXSD += "type=\"xs:" + sequenceElement.Type + "\"/>\n";
                    }
                }
                complexTypeXSD += "\t\t\t</xs:sequence>\n\t\t</xs:extension>\n\t</xs:complexContent>\n</xs:complexType>\n\n";
                complexTypeXSD += CreateXSDString(complexType.Base);
#endregion
            }
            else
            {
                #region Not Inhereted
                if (complexType.IsAbtract)
                {
                    complexTypeXSD += "<xs:complexType abstract=\"true\" name=\"" + complexType.Name + "\">\n\t<xs:sequence>\n";
                }
                else
                {
                    complexTypeXSD += "<xs:complexType name=\"" + complexType.Name + "\">\n\t<xs:sequence>\n";
                }
                foreach (XSElement sequenceElement in complexType.Sequence)
                {
                    complexTypeXSD += "\t\t<xs:element ";
                    if (sequenceElement.MinOccurs != null)
                    {
                        complexTypeXSD += "minOccurs=\"" + sequenceElement.MinOccurs + "\" ";
                    }

                    complexTypeXSD += "name=\"" + sequenceElement.Name + "\" ";
                    if (xsSchema.SimpleTypes.Where(st => st.Name == sequenceElement.Type).Count() > 0 || xsSchema.ComplexTypes.Where(xsct => xsct.Name == sequenceElement.Type).Count() > 0)
                    {
                        complexTypeXSD += "type=\"tns:" + sequenceElement.Type + "\"/>\n";
                    }
                    else
                    {
                        complexTypeXSD += "type=\"xs:" + sequenceElement.Type + "\"/>\n";
                    }
                }
                complexTypeXSD += "\t</xs:sequence>\n</xs:complexType>";
                #endregion
            }
            return complexTypeXSD;
        }

        private static void CreateIntroduction()
        {
            #region Introducion

            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "Introduction";
            newParagraph.Range.Font.Color = WdColor.wdColorBlue;

            newParagraph.Range.Font.Size = 14;

            newParagraph.Range.Font.Bold = 0;
            
            newParagraph.Range.Font.Name = "Times New Roman";
            newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
            newParagraph.Range.SetListLevel(1);

            newParagraph.Range.InsertParagraphAfter();

            newParagraph.Range.ListFormat.RemoveNumbers();

            newParagraph = wordDoc.Content.Paragraphs.Add();
            newParagraph.Range.Text = "";

            newParagraph.Indent();
            newParagraph.Indent();

            newParagraph.Range.Font.Size = 12;

            newParagraph.Range.Font.Color = WdColor.wdColorBlack;
            string purposeString =
                "The purpose of this document is to clearly describe the XML of a WSDL of Web Services developed in Java. It analyzes each web method by getting their input, output and fault parameters, as well as describing and analyzing each parameter.";

            newParagraph.Range.Text = purposeString;

            newParagraph.Range.InsertParagraphAfter();

            #endregion
        }

        private static void CreateXSD(XSComplexType complexType)
        {
            #region Complex Type XSD

            newParagraph = wordDoc.Content.Paragraphs.Add();


            newParagraph.Range.Text = "";

            int rgbColor = Information.RGB(79, 129, 189);

            WdColor wdColorBackgroundColor = (WdColor)rgbColor;
            newParagraph.Range.Shading.BackgroundPatternColor = wdColorBackgroundColor;

            newParagraph.Range.Font.Name = "Arial";

            newParagraph.Range.Font.Size = 11;

            newParagraph.Range.Font.Color = WdColor.wdColorWhite;
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;

            string complexTypeXSD = "";
            complexTypeXSD += CreateXSDString(complexType);
            newParagraph.Range.Text = complexTypeXSD;
            newParagraph.Range.Text += "";
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
            newParagraph.Range.Font.Color = WdColor.wdColorBlue;
            newParagraph.Range.Font.Name = "Times New Roman";
            newParagraph.Range.Font.Size = 12;
            newParagraph.Range.InsertParagraphAfter();
            #endregion

        }

        private static void AddDataToFieldDecription(XSComplexType complexType, bool additionalRows, ref int rowNumber, ref Table fieldDescTable)
        {
            if (complexType.Base != null)
            {
                AddDataToFieldDecription(complexType.Base, true, ref rowNumber, ref fieldDescTable);
            }
            foreach (XSElement element in complexType.Sequence)
            {
                if (additionalRows)
                {
                    fieldDescTable.Rows.Add(ref oMissing);
                }
                fieldDescTable.Cell(rowNumber, 1).Range.Font.Name = "Arial";
                fieldDescTable.Cell(rowNumber, 2).Range.Font.Name = "Arial";
                fieldDescTable.Cell(rowNumber, 3).Range.Font.Name = "Arial";
                fieldDescTable.Cell(rowNumber, 4).Range.Font.Name = "Arial";

                fieldDescTable.Cell(rowNumber, 1).Range.Text = element.Name;

                var r = new Regex(@"
            (?<=[A-Z])(?=[A-Z][a-z]) |
                (?<=[^A-Z])(?=[A-Z]) |
                (?<=[A-Za-z])(?=[^A-Za-z])", RegexOptions.IgnorePatternWhitespace);

                string elementDescription = r.Replace(element.Name, " ");

                char[] elementDescriptionCharacters = elementDescription.ToCharArray();

                elementDescriptionCharacters[0] = char.ToUpper(elementDescription[0]);

                elementDescription = new string(elementDescriptionCharacters);

                fieldDescTable.Cell(rowNumber, 2).Range.Text = elementDescription;
                fieldDescTable.Cell(rowNumber, 3).Range.Text = element.Type;

                XSSimpleType simpleType = xsSchema.SimpleTypes.FirstOrDefault(st => st.Name == element.Type);
                XSComplexType otherComplexType = xsSchema.ComplexTypes.FirstOrDefault(st => st.Name == element.Type);
                if (simpleType != null)
                {
                    string comment = "This type is an enum. Here are the possible values for this enum:\n";
                    foreach (string value in simpleType.enumerations)
                    {
                        comment += value + "\n";
                    }
                    fieldDescTable.Cell(rowNumber, 4).Range.Text = comment;
                }
                else if (otherComplexType != null)
                {
                    string comment = "This type is a complex type. Here are the properties in this type:\n";
                    foreach (XSElement value in otherComplexType.Sequence)
                    {
                        comment += value.Name + "\n";
                    }
                    fieldDescTable.Cell(rowNumber, 4).Range.Text = comment;
                }
                rowNumber++;
            }
            
        }

        private static void CreateFieldDescription(XSComplexType complexType)
        {
            #region Field Description
                newParagraph = wordDoc.Content.Paragraphs.Add();

                Table fieldDescTable = wordDoc.Content.Tables.Add(newParagraph.Range, complexType.Sequence.Count + 1, 4);

                for (int i = 1; i <= fieldDescTable.Columns.Count; i++)
                {

                    for (int j = 1; j <= fieldDescTable.Rows.Count; j++)
                    {
                        fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    }

                }

                fieldDescTable.Range.Font.Color = WdColor.wdColorBlack;

                int rgbColor = Information.RGB(79, 129, 189);

                WdColor wdColorCellColor = (WdColor)rgbColor;

                fieldDescTable.Cell(1, 1).Range.Shading.BackgroundPatternColor = wdColorCellColor;
                fieldDescTable.Cell(1, 2).Range.Shading.BackgroundPatternColor = wdColorCellColor;

                fieldDescTable.Cell(1, 3).Range.Shading.BackgroundPatternColor = wdColorCellColor;
                fieldDescTable.Cell(1, 4).Range.Shading.BackgroundPatternColor = wdColorCellColor;

                fieldDescTable.Cell(1, 1).Range.Text = "Element Name";
                fieldDescTable.Cell(1, 2).Range.Text = "Description";
                fieldDescTable.Cell(1, 3).Range.Text = "Type";
                fieldDescTable.Cell(1, 4).Range.Text = "Comment";

                fieldDescTable.Cell(1, 1).Range.Font.Size = 14;
                fieldDescTable.Cell(1, 2).Range.Font.Size = 14;

                fieldDescTable.Cell(1, 1).Range.Font.Name = "Times New Roman";
                fieldDescTable.Cell(1, 2).Range.Font.Name = "Times New Roman";

                fieldDescTable.Cell(1, 1).Range.Font.Color = WdColor.wdColorWhite;
                fieldDescTable.Cell(1, 2).Range.Font.Color = WdColor.wdColorWhite;

                fieldDescTable.Cell(1, 3).Range.Font.Name = "Times New Roman";
                fieldDescTable.Cell(1, 4).Range.Font.Name = "Times New Roman";

                fieldDescTable.Cell(1, 3).Range.Font.Color = WdColor.wdColorWhite;
                fieldDescTable.Cell(1, 4).Range.Font.Color = WdColor.wdColorWhite;

                int rowNumber = 2;
                StringBuilder stringBuilder = new StringBuilder();
                AddDataToFieldDecription(complexType, false, ref rowNumber, ref fieldDescTable);
                newParagraph.Range.InsertParagraphAfter();

            #endregion
        }

        private static void CreateXSD(XSElement element)
        {
            #region Element XSD

            newParagraph = wordDoc.Content.Paragraphs.Add();


            newParagraph.Range.Text = "";

            int rgbColor = Information.RGB(79, 129, 189);

            WdColor wdColorBackgroundColor = (WdColor)rgbColor;
            newParagraph.Range.Shading.BackgroundPatternColor = wdColorBackgroundColor;

            newParagraph.Range.Font.Name = "Arial";

            newParagraph.Range.Font.Size = 11;

            newParagraph.Range.Font.Color = WdColor.wdColorWhite;
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;

            string elementXSD = "";
            //complexTypeXSD += "<xs:complexType name=\"" + complexType.Name + "\">\n\t<xs:sequence>\n";
            //foreach (XSElement sequenceElement in complexType.Sequence)
            //{
            elementXSD += "\t\t<xs:element ";
            if (element.MinOccurs != null)
            {
                elementXSD += "minOccurs=\"" + element.MinOccurs + "\" ";
            }

            elementXSD += "name=\"" + element.Name + "\" ";
            if (xsSchema.SimpleTypes.Where(st => st.Name == element.Type).Count() > 0 || xsSchema.ComplexTypes.Where(xsct => xsct.Name == element.Type).Count() > 0)
            {
                elementXSD += "type=\"tns:" + element.Type + "\"/>\n";
            }
            else
            {
                elementXSD += "type=\"xs:" + element.Type + "\"/>\n";
            }
            //}
            //elementXSD += "\t</xs:sequence>\n</xs:complexType>";
            newParagraph.Range.Text = elementXSD;
            newParagraph.Range.Text += "";
            newParagraph.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleNone;
            newParagraph.Range.Shading.BackgroundPatternColor = WdColor.wdColorWhite;
            newParagraph.Range.Font.Color = WdColor.wdColorBlue;
            newParagraph.Range.Font.Name = "Times New Roman";
            newParagraph.Range.Font.Size = 12;
            newParagraph.Range.InsertParagraphAfter();
            #endregion

        }

        private static void CreateFieldDescription(XSElement element)
        {
            #region Field Description

            newParagraph = wordDoc.Content.Paragraphs.Add();

            Table fieldDescTable = wordDoc.Content.Tables.Add(newParagraph.Range, 2, 4);

            for (int i = 1; i <= fieldDescTable.Columns.Count; i++)
            {

                for (int j = 1; j <= fieldDescTable.Rows.Count; j++)
                {
                    fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                    fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                    fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    fieldDescTable.Cell(j, i).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                }

            }

            fieldDescTable.Range.Font.Color = WdColor.wdColorBlack;

            int rgbColor = Information.RGB(79, 129, 189);

            WdColor wdColorCellColor = (WdColor)rgbColor;

            fieldDescTable.Cell(1, 1).Range.Shading.BackgroundPatternColor = wdColorCellColor;
            fieldDescTable.Cell(1, 2).Range.Shading.BackgroundPatternColor = wdColorCellColor;

            fieldDescTable.Cell(1, 3).Range.Shading.BackgroundPatternColor = wdColorCellColor;
            fieldDescTable.Cell(1, 4).Range.Shading.BackgroundPatternColor = wdColorCellColor;

            fieldDescTable.Cell(1, 1).Range.Text = "Element Name";
            fieldDescTable.Cell(1, 2).Range.Text = "Description";
            fieldDescTable.Cell(1, 3).Range.Text = "Type";
            fieldDescTable.Cell(1, 4).Range.Text = "Comment";

            fieldDescTable.Cell(1, 1).Range.Font.Size = 14;
            fieldDescTable.Cell(1, 2).Range.Font.Size = 14;

            fieldDescTable.Cell(1, 1).Range.Font.Name = "Times New Roman";
            fieldDescTable.Cell(1, 2).Range.Font.Name = "Times New Roman";

            fieldDescTable.Cell(1, 1).Range.Font.Color = WdColor.wdColorWhite;
            fieldDescTable.Cell(1, 2).Range.Font.Color = WdColor.wdColorWhite;

            fieldDescTable.Cell(1, 3).Range.Font.Name = "Times New Roman";
            fieldDescTable.Cell(1, 4).Range.Font.Name = "Times New Roman";

            fieldDescTable.Cell(1, 3).Range.Font.Color = WdColor.wdColorWhite;
            fieldDescTable.Cell(1, 4).Range.Font.Color = WdColor.wdColorWhite;

            int rowNumber = 2;
            StringBuilder stringBuilder = new StringBuilder();
            fieldDescTable.Cell(rowNumber, 1).Range.Font.Name = "Arial";
            fieldDescTable.Cell(rowNumber, 2).Range.Font.Name = "Arial";
            fieldDescTable.Cell(rowNumber, 3).Range.Font.Name = "Arial";
            fieldDescTable.Cell(rowNumber, 4).Range.Font.Name = "Arial";

            fieldDescTable.Cell(rowNumber, 1).Range.Text = element.Name;

            var r = new Regex(@"
        (?<=[A-Z])(?=[A-Z][a-z]) |
            (?<=[^A-Z])(?=[A-Z]) |
            (?<=[A-Za-z])(?=[^A-Za-z])", RegexOptions.IgnorePatternWhitespace);

            string elementDescription = r.Replace(element.Name, " ");

            char[] elementDescriptionCharacters = elementDescription.ToCharArray();

            elementDescriptionCharacters[0] = char.ToUpper(elementDescription[0]);

            elementDescription = new string(elementDescriptionCharacters);

            fieldDescTable.Cell(rowNumber, 2).Range.Text = elementDescription;
            fieldDescTable.Cell(rowNumber, 3).Range.Text = element.Type;

            XSSimpleType simpleType = xsSchema.SimpleTypes.FirstOrDefault(st => st.Name == element.Type);

            if (simpleType != null)
            {
                string comment = "";
                foreach (string value in simpleType.enumerations)
                {
                    comment += value + "\n";
                }
                fieldDescTable.Cell(rowNumber, 4).Range.Text = comment;
            }
            rowNumber++;

            newParagraph.Range.InsertParagraphAfter();

            #endregion
        }

        public static void WriteDocument(List<WebMethodInfo> webMethodInfoList, XSSchema xsSchemaParameter, WsdlPortType portTypes, string saveFileLocation, string wsdlXmlString, string authorName, ref int statusInt, ref string statusString)
        {
            xsSchema = xsSchemaParameter;
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordDoc = wordApp.Documents.Add();
            wordDoc.Activate();
            CreateListTemplate();
            statusInt += 10;
            statusString = "Writing cover page";
            CreateCoverPage(authorName, portTypes.Name);
            statusInt += 10;
            statusString = "Writing introduction";
            CreateIntroduction();
            statusInt += 10;
            statusString = "Writing each web method";
            foreach (WebMethodInfo webMethodInfo in webMethodInfoList)
            {
                newParagraph = wordDoc.Content.Paragraphs.Add();
                newParagraph.Range.Text = webMethodInfo.OperationInfo.Name;
                newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                newParagraph.Range.Font.Size = 14;

                //newParagraph.Range.Font.Bold = 1;

                newParagraph.Range.Font.Name = "Times New Roman";
                newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                newParagraph.Range.SetListLevel(1);

                newParagraph.Range.InsertParagraphAfter();
                #region Method Parameters

                newParagraph = wordDoc.Content.Paragraphs.Add();
                newParagraph.Range.Text = "Method Parameters";
                newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                newParagraph.Range.Font.Size = 12;

                //newParagraph.Range.Font.Bold = 1;

                newParagraph.Range.Font.Name = "Times New Roman";
                newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                newParagraph.Range.SetListLevel(2);
                newParagraph.Range.InsertParagraphAfter();
                newParagraph.Range.ListFormat.RemoveNumbers();
                newParagraph = wordDoc.Content.Paragraphs.Add();

                Table table = wordDoc.Content.Tables.Add(newParagraph.Range, 5, 2);

                table.Range.Font.Color = WdColor.wdColorBlack;
                int rgbColor = Information.RGB(79, 129, 189);

                WdColor wdColorCellColor = (WdColor)rgbColor;
                table.Cell(1, 1).Range.Shading.BackgroundPatternColor = wdColorCellColor;
                table.Cell(1, 2).Range.Shading.BackgroundPatternColor = wdColorCellColor;

                table.Cell(1, 1).Range.Text = "Parameter Name";
                table.Cell(1, 2).Range.Text = "Value";

                table.Cell(1, 1).Range.Font.Size = 14;
                table.Cell(1, 2).Range.Font.Size = 14;

                table.Cell(1, 1).Range.Font.Name = "Times New Roman";
                table.Cell(1, 2).Range.Font.Name = "Times New Roman";

                table.Cell(1, 1).Range.Font.Color = WdColor.wdColorWhite;
                table.Cell(1, 2).Range.Font.Color = WdColor.wdColorWhite;

                table.Cell(2, 1).Range.Font.Bold = 1;
                table.Cell(3, 1).Range.Font.Bold = 1;
                table.Cell(4, 1).Range.Font.Bold = 1;
                table.Cell(5, 1).Range.Font.Bold = 1;

                table.Cell(2, 1).Range.Font.Name = "Arial";
                table.Cell(3, 1).Range.Font.Name = "Arial";
                table.Cell(4, 1).Range.Font.Name = "Arial";
                table.Cell(5, 1).Range.Font.Name = "Arial";

                table.Cell(2, 2).Range.Font.Name = "Arial";
                table.Cell(3, 2).Range.Font.Name = "Arial";
                table.Cell(4, 2).Range.Font.Name = "Arial";
                table.Cell(5, 2).Range.Font.Name = "Arial";

                table.Cell(1, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Cell(1, 2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                table.Cell(2, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Cell(3, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Cell(4, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                table.Cell(5, 1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                table.Cell(2, 1).Range.Text = "Method Name";
                table.Cell(3, 1).Range.Text = "Request Parameter";
                table.Cell(4, 1).Range.Text = "Response Parameter";
                table.Cell(5, 1).Range.Text = "Fault Parameter";

                table.Cell(2, 2).Range.Font.Bold = 0;
                table.Cell(3, 2).Range.Font.Bold = 0;
                table.Cell(4, 2).Range.Font.Bold = 0;
                table.Cell(5, 2).Range.Font.Bold = 0;

                table.Cell(2, 2).Range.Text = webMethodInfo.OperationInfo.Name;
                table.Cell(3, 2).Range.Text = "";
                string stringToInput = "";
                int countInputs = 0;

                if (webMethodInfo.ElementInputs != null)
                {
                    if (webMethodInfo.ElementInputs.Count > 0)
                    {
                        foreach (XSElement element in webMethodInfo.ElementInputs)
                        {
                            //table.Cell(3, 2).Range.ListFormat.ApplyBulletDefault();
                            //table.Cell(3, 2).Range.Text += element.Name + "\n";
                            if (countInputs > 0)
                            {
                                stringToInput += "\n" + element.Name;
                            }
                            else
                            {
                                stringToInput += element.Name;
                            }

                            countInputs++;
                        }

                    }
                }
                if (webMethodInfo.ComplexTypeInputs != null)
                {
                    if (webMethodInfo.ComplexTypeInputs.Count > 0)
                    {
                        foreach (XSComplexType complexType in webMethodInfo.ComplexTypeInputs)
                        {
                            //table.Cell(3, 2).Range.ListFormat.ApplyBulletDefault();
                            //table.Cell(3, 2).Range.Text += complexType.Name + "\n";
                            if (countInputs > 0)
                            {
                                stringToInput += "\n" + complexType.Name;
                            }
                            else
                            {
                                stringToInput += complexType.Name;
                            }
                            countInputs++;
                        }

                    }
                }
                //table.Cell(3, 2).Range.Delete(WdUnits.wdCharacter, 1);
                table.Cell(3, 2).Range.Text = stringToInput;
                if (countInputs > 1)
                {
                    table.Cell(3, 2).Range.ListFormat.ApplyBulletDefault();
                }

                //table.Cell(3, 2).Range.Text = operation.InputName;
                if (webMethodInfo.ComplexTypeOutput != null)
                {
                    table.Cell(4, 2).Range.Text = webMethodInfo.ComplexTypeOutput.Name;
                }
                else if (webMethodInfo.ElementOutput != null)
                {
                    table.Cell(4, 2).Range.Text = webMethodInfo.ElementOutput.Name;
                }

                //table.Cell(5, 2).Range.Text = webMethodInfo.OperationInfo.FaultNames[0];
                int countFaultName = 0;
                string stringToFault = "";
                if (webMethodInfo.OperationInfo.FaultNames != null)
                {
                    if (webMethodInfo.OperationInfo.FaultNames.Count > 0)
                    {
                        foreach (string faultNames in webMethodInfo.OperationInfo.FaultNames)
                        {
                            //table.Cell(3, 2).Range.ListFormat.ApplyBulletDefault();
                            //table.Cell(3, 2).Range.Text += element.Name + "\n";
                            if (countFaultName > 0)
                            {
                                stringToFault += "\n" + faultNames;
                            }
                            else
                            {
                                stringToFault += faultNames;
                            }

                            countFaultName++;
                        }

                    }
                }
                table.Cell(5, 2).Range.Text = stringToFault;
                if (countFaultName > 1)
                {
                    table.Cell(5, 2).Range.ListFormat.ApplyBulletDefault();
                }
                for (int i = 1; i <= table.Columns.Count; i++)
                {

                    for (int j = 1; j <= table.Rows.Count; j++)
                    {
                        table.Cell(j, i).Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleSingle;
                        table.Cell(j, i).Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleSingle;
                        table.Cell(j, i).Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        table.Cell(j, i).Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleSingle;
                    }

                }
                newParagraph.Range.InsertParagraphAfter();
                #endregion

                #region Request Message

                newParagraph = wordDoc.Content.Paragraphs.Add();
                newParagraph.Range.Text = "Request Message";
                newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                newParagraph.Range.Font.Size = 12;

                //newParagraph.Range.Font.Bold = 1;

                newParagraph.Range.Font.Name = "Times New Roman";
                newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                newParagraph.Range.SetListLevel(2);
                newParagraph.Range.InsertParagraphAfter();
                //newParagraph.Range.ListFormat.RemoveNumbers();
                #region ComplexTypes Inputs
                if (webMethodInfo.ComplexTypeInputs != null)
                {
                    foreach (XSComplexType ct in webMethodInfo.ComplexTypeInputs)
                    {
                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = ct.Name;
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                        newParagraph.Range.SetListLevel(3);
                        newParagraph.Range.InsertParagraphAfter();

                        #region XML Schema Definition

                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = "XML Schema Definition";
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                        newParagraph.Range.SetListLevel(4);
                        newParagraph.Range.InsertParagraphAfter();
                        newParagraph.Range.ListFormat.RemoveNumbers();

                        CreateXSD(ct);
                        #endregion
                        #region Field Description

                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = "Field Description";
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                        newParagraph.Range.SetListLevel(4);
                        newParagraph.Range.InsertParagraphAfter();
                        newParagraph.Range.ListFormat.RemoveNumbers();
                        CreateFieldDescription(ct);

                        #endregion

                    }
                }
                #endregion

                #region Element Inputs
                if (webMethodInfo.ElementInputs != null)
                {
                    foreach (XSElement element in webMethodInfo.ElementInputs)
                    {
                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = element.Name;
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                        newParagraph.Range.SetListLevel(3);
                        newParagraph.Range.InsertParagraphAfter();

                        #region XML Schema Definition

                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = "XML Schema Definition";
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                        newParagraph.Range.SetListLevel(4);
                        newParagraph.Range.InsertParagraphAfter();
                        newParagraph.Range.ListFormat.RemoveNumbers();

                        CreateXSD(element);
                        #endregion
                        #region Field Description

                        newParagraph = wordDoc.Content.Paragraphs.Add();
                        newParagraph.Range.Text = "Field Description";
                        newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                        newParagraph.Range.Font.Size = 12;

                        //newParagraph.Range.Font.Bold = 1;

                        newParagraph.Range.Font.Name = "Times New Roman";
                        newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                        newParagraph.Range.SetListLevel(4);
                        newParagraph.Range.InsertParagraphAfter();
                        newParagraph.Range.ListFormat.RemoveNumbers();
                        CreateFieldDescription(element);

                        #endregion
                    }
                }
                #endregion

                #endregion

                #region Response Message

                newParagraph = wordDoc.Content.Paragraphs.Add();
                newParagraph.Range.Text = "Response Message";
                newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                newParagraph.Range.Font.Size = 12;

                //newParagraph.Range.Font.Bold = 1;

                newParagraph.Range.Font.Name = "Times New Roman";
                newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                newParagraph.Range.SetListLevel(2);
                newParagraph.Range.InsertParagraphAfter();
                #region ComplexTypes Output
                if (webMethodInfo.ComplexTypeOutput != null)
                {
                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = webMethodInfo.ComplexTypeOutput.Name;
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                    newParagraph.Range.SetListLevel(3);
                    newParagraph.Range.InsertParagraphAfter();

                    #region XML Schema Definition

                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = "XML Schema Definition";
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                    newParagraph.Range.SetListLevel(4);
                    newParagraph.Range.InsertParagraphAfter();
                    newParagraph.Range.ListFormat.RemoveNumbers();

                    CreateXSD(webMethodInfo.ComplexTypeOutput);

                    #endregion
                    #region Field Description

                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = "Field Description";
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                    newParagraph.Range.SetListLevel(4);
                    newParagraph.Range.InsertParagraphAfter();
                    newParagraph.Range.ListFormat.RemoveNumbers();
                    CreateFieldDescription(webMethodInfo.ComplexTypeOutput);
                    #endregion
                }
                #endregion

                #region Element Output
                if (webMethodInfo.ElementOutput != null)
                {
                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = webMethodInfo.ElementOutput.Name;
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                    newParagraph.Range.SetListLevel(3);
                    newParagraph.Range.InsertParagraphAfter();

                    #region XML Schema Definition

                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = "XML Schema Definition";
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.CanContinuePreviousList(listTemplate);
                    newParagraph.Range.SetListLevel(4);
                    newParagraph.Range.InsertParagraphAfter();
                    newParagraph.Range.ListFormat.RemoveNumbers();

                    CreateXSD(webMethodInfo.ElementOutput);
                    #endregion
                    #region Field Description

                    newParagraph = wordDoc.Content.Paragraphs.Add();
                    newParagraph.Range.Text = "Field Description";
                    newParagraph.Range.Font.Color = WdColor.wdColorBlue;

                    newParagraph.Range.Font.Size = 12;

                    //newParagraph.Range.Font.Bold = 1;

                    newParagraph.Range.Font.Name = "Times New Roman";
                    newParagraph.Range.ListFormat.ApplyListTemplateWithLevel(listTemplate, ContinuePreviousList: true, ApplyTo: WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior: WdDefaultListBehavior.wdWord10ListBehavior);
                    newParagraph.Range.SetListLevel(4);
                    newParagraph.Range.InsertParagraphAfter();
                    newParagraph.Range.ListFormat.RemoveNumbers();
                    CreateFieldDescription(webMethodInfo.ElementOutput);

                    #endregion
                }
                #endregion
                #endregion
            }
            statusInt += 20;
            statusString = "Writing appendix";
            CreateWSDLAppendix(wsdlXmlString);
            statusInt += 20;
            statusString = "Saving file......";
            wordDoc.SaveAs(saveFileLocation);

            wordApp.Application.Quit();
        }
    }
}
