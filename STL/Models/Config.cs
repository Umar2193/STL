using Newtonsoft.Json;

namespace STL.Models {
    internal class Config {

        public Config() {
            CustomerMaterial = new Dictionary<string, string>();
            BladeSources = new List<Source>();
            TowerSources = new List<Source>();
            MainSources = new List<Source>();
        }

        [JsonProperty("CustomerMaterial")]
        public Dictionary<string, string> CustomerMaterial { get; set; }

        [JsonProperty("bladeSources")]
        public List<Source> BladeSources { get; set; }

        [JsonProperty("towerSources")]
        public List<Source> TowerSources { get; set; }

        [JsonProperty("mainSources")]
        public List<Source> MainSources { get; set; }
    }
}
