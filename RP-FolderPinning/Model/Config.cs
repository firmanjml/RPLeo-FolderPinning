using System.Collections.Generic;

namespace RP_FolderPinning.Model
{
    class Config
    {
        public List<Calendar> calendar { get; set; }
        public string folder_dir { get; set; }
        public string prev_dir { get; set; }
        public bool verbose { get; set; }
        public int verbose_sleep_time { get; set; }
    }
}
