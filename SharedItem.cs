using System.Xml.Serialization;

namespace ownCloud.Outlook
{
    /// <summary>
    /// Shared file
    /// </summary>
    internal class SharedItem
    {
        [XmlAttribute("url")]
        public string Url { get; set; }
    }
    
}