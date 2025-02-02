using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPT.ReviewTab.Code.Model
{
    public class ItemGroup
    {
        public string Name { get; set; }
        public List<Item> Items { get; set; }
        public GroupShape Shape { get; set; }
    }
}
