using System;

namespace AutoTrader_Scraper
{
    [Serializable]
    public class Car
    {
        public string Name { get; set; }

        public int Year { get; set; }
        public int Mileage { get; set; }

        public int HorsePower { get; set; }
        public int Price { get; set; }

        public string Fuel { get; set; }
        public string Transmission { get; set; }
        public string Trim { get; set; }
        public decimal Litre { get; set; }

        public string MiscDetails { get; set; }

        public string URL { get; set; }

        public string Distance { get; set; }
        public string Reg { get; internal set; }

        public override string ToString()
        {
            return $"{Name},{Year},{Mileage},{HorsePower},{Price},{Fuel},{Transmission},{Trim},{Litre},{(MiscDetails ?? "").Replace(",", "")},{URL},{Distance},{Reg}";
        }
    }
}
