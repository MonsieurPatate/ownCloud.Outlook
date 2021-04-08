using System.Xml.Serialization;

namespace ownCloud.Outlook
{
    internal class SharedItem
    {
        [XmlAttribute("url")]
        public string Url { get; set; }
    }
    
}