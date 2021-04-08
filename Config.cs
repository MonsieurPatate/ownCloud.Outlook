using Newtonsoft.Json;

namespace ownCloud.Outlook
{
    public class Config
    {
        [JsonProperty("ownCloudUrl")]
        public string OwnCloudUrl { get; set; }

        [JsonProperty("maxAttachmentSize")]
        public int MaxAttachmentSize { get; set; }


    }
}