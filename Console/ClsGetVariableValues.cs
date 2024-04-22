using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace TestConsole
{
    public static class ClsGetVariableValues
    {

        public static void ReadVariableValue(string filepath)
        {
            try
            {
                using (WordprocessingDocument wordDocument =
              WordprocessingDocument.Open(filepath, false))
                {
                    // Assign a reference to the existing document body.  
                    var mainDocument = wordDocument.MainDocumentPart.Document;
                    var contentControls = mainDocument.Descendants<SimpleField>().ToList();
                    var contentControls1 = mainDocument.Descendants<FieldCode>().ToList();

                    foreach (var item in contentControls1)
                    {
                        var property = item.InnerText;
                        OpenXmlElement node = item.Parent.Parent;
                        var match = node.Descendants<Text>().First();
                    }

                    var Header = wordDocument.MainDocumentPart.HeaderParts;
                    foreach (var item in Header)
                    {
                        var c1 = item.Header.Descendants<SimpleField>().ToList();
                        var c2 = item.Header.Descendants<FieldCode>().ToList();

                        foreach (var i in c1)
                        {
                            var property = i.InnerText;
                            OpenXmlElement node = i.Parent.Parent;
                            var match = node.Descendants<Text>().First();
                        }
                        foreach (var i in c2)
                        {
                            var property = i.InnerText;
                            OpenXmlElement node = i.Parent.Parent;
                            var match = node.Descendants<Text>().First();
                        }
                    }
                    var Footer = wordDocument.MainDocumentPart.FooterParts;
                    foreach (var item in Footer)
                    {
                        var c1 = item.Footer.Descendants<SimpleField>().ToList();
                        var c2 = item.Footer.Descendants<FieldCode>().ToList();

                        foreach (var i in c1)
                        {
                            var property = i.InnerText;
                            OpenXmlElement node = i.Parent.Parent;
                            var match = node.Descendants<Text>().First();
                        }
                        foreach (var i in c2)
                        {
                            var property = i.InnerText;
                            OpenXmlElement node = i.Parent.Parent;
                            var match = node.Descendants<Text>().First();
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
