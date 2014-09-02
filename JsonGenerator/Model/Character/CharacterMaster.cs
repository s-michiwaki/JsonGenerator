using JsonGenerator.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonGenerator.Model.Character
{
    [ExcelData]
    public class CharacterMaster
    {
        public uint CharacterId { get; set; }

        public string Name { get; set; }

        public float AttackSpeed { get; set; }

        public float AttackRange { get; set; }

        public float AttackInterval { get; set; }

        public int HP { get; set; }

        public int Offence { get; set; }

        public int Defense { get; set; }

        public float Speed { get; set; }

        public int Cost { get; set; }

        public float MakeInterval { get; set; }

        public int DeadResource { get; set; }

        public string ImageName { get; set; }
    }
}
