using JsonGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Model.Character
{
    [ExcelData]
    public class RankMaster
    {
        public int Rank { get; set; }

        public float HpUpRate { get; set; }

        public float OffenseUpRate { get; set; }

        public float DefenseUpRate { get; set; }

        public float SpeedUpRate { get; set; }
    }
}
