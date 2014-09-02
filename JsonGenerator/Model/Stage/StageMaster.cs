using JsonGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Model.Stage
{
    [ExcelData]
    public class StageMaster
    {
        public int StageId { get; set; }

        public string Name { get; set; }
    }
}
