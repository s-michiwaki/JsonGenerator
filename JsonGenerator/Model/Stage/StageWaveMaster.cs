using JsonGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Model.Stage
{
    [ExcelData]
    public class StageWaveMaster
    {
        public int StageId { get; set; }

        public int CharacterId { get; set; }

        public int Rank { get; set; }

        public float AppearTime { get; set; }
    }
}
