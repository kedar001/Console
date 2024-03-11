using System.Collections.Generic;
using System.IO.Packaging;
using System.Xml.Linq;

namespace eDocsDN_DocX_IS
{
  public class Section : Container
  {

    public SectionBreakType SectionBreakType;

    internal Section(DocX document, XElement xml) : base(document, xml)
    {
    }

    public List<Paragraph> SectionParagraphs { get; set; }
  }
}