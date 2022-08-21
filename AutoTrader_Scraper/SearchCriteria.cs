using System.Text;

namespace AutoTrader_Scraper
{
    public class SearchCriteria
    {
        public string postcode { get; set; }
        public int radius { get; set; }

        public string make { get; set; }
        public string model { get; set; }

        public int price_from { get; set; }
        public int price_to { get; set; }

        public int year_from { get; set; }

        public int year_to { get; set; }

        public string zero_to_60 { get; set; }
        public string drivetrain { get; set; }
        public string annual_tax_cars { get; set; }
        public string insuranceGroup { get; set; }
        public string co2_emissions_cars { get; set; }
        public string fuel_consumption { get; set; }
        public string exclude_writeoff_categories { get; set; }


        public string colour { get; set; }

        public string maximum_mileage { get; set; }

        public string minimum_badge_engine_size { get; set; }
        public string maximum_badge_engine_size { get; set; }

        public string min_engine_power { get; set; }
        public string max_engine_power { get; set; }

        public int quantity_of_doors { get; set; }

        public int minimum_seats { get; set; }
        public int maximum_seats { get; set; }

        public string aggregatedTrim { get; set; }


        public string body_type { get; set; }

        public string transmission { get; set; }

        public string fuel_type { get; set; }

        public string seller_type { get; set; }

        public string GetSearchURL()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(@"https://www.autotrader.co.uk/car-search?");

            sb.Append(string.IsNullOrEmpty(max_engine_power) ? "" : $"&{nameof(max_engine_power).Replace('_', '-')}={max_engine_power}");
            sb.Append(string.IsNullOrEmpty(min_engine_power) ? "" : $"&{nameof(min_engine_power).Replace('_', '-')}={min_engine_power}");
            sb.Append(string.IsNullOrEmpty(maximum_badge_engine_size) ? "" : $"&{nameof(maximum_badge_engine_size).Replace('_', '-')}={maximum_badge_engine_size}");
            sb.Append(string.IsNullOrEmpty(minimum_badge_engine_size) ? "" : $"&{nameof(minimum_badge_engine_size).Replace('_', '-')}={minimum_badge_engine_size}");
            sb.Append(maximum_mileage == default ? "" : $"&{nameof(maximum_mileage).Replace('_', '-')}={maximum_mileage}");
            sb.Append(string.IsNullOrEmpty(colour) ? "" : $"&{nameof(colour).Replace('_', '-')}={colour}");
            sb.Append(string.IsNullOrEmpty(fuel_type) ? "" : $"&{nameof(fuel_type).Replace('_', '-')}={fuel_type}");
            sb.Append(string.IsNullOrEmpty(transmission) ? "" : $"&{nameof(transmission).Replace('_', '-')}={transmission}");

            sb.Append(postcode == default ? "" : $"&{nameof(postcode).Replace('_', '-')}={postcode.Replace(" ", "")}");
            sb.Append(string.IsNullOrEmpty(make) ? "" : $"&{nameof(make).Replace('_', '-')}={make}");
            sb.Append(string.IsNullOrEmpty(model) ? "" : $"&{nameof(model).Replace('_', '-')}={model}");
            sb.Append(radius == default ? "" : $"&{nameof(radius).Replace('_', '-')}={radius}");
            sb.Append(price_from == default ? "" : $"&{nameof(price_from).Replace('_', '-')}={price_from}");
            sb.Append(price_to == default ? "" : $"&{nameof(price_to).Replace('_', '-')}={price_to}");
            sb.Append(year_from == default ? "" : $"&{nameof(year_from).Replace('_', '-')}={year_from}");
            sb.Append(year_to == default ? "" : $"&{nameof(year_to).Replace('_', '-')}={year_to}");
            sb.Append(string.IsNullOrEmpty(body_type) ? "" : $"&{nameof(body_type).Replace('_', '-')}={body_type}");
            sb.Append(string.IsNullOrEmpty(seller_type) ? "" : $"&{nameof(seller_type).Replace('_', '-')}={seller_type}");
            sb.Append(string.IsNullOrEmpty(exclude_writeoff_categories) ? "" : $"&{nameof(exclude_writeoff_categories).Replace('_', '-')}={exclude_writeoff_categories}");
            sb.Append(string.IsNullOrEmpty(fuel_consumption) ? "" : $"&{nameof(fuel_consumption).Replace('_', '-')}={fuel_consumption}");
            sb.Append(string.IsNullOrEmpty(co2_emissions_cars) ? "" : $"&{nameof(co2_emissions_cars).Replace('_', '-')}={co2_emissions_cars}");
            sb.Append(string.IsNullOrEmpty(insuranceGroup) ? "" : $"&{nameof(insuranceGroup).Replace('_', '-')}={insuranceGroup}");
            sb.Append(string.IsNullOrEmpty(annual_tax_cars) ? "" : $"&{nameof(annual_tax_cars).Replace('_', '-')}={annual_tax_cars}");
            sb.Append(string.IsNullOrEmpty(drivetrain) ? "" : $"&{nameof(drivetrain).Replace('_', '-')}={drivetrain}");
            sb.Append(string.IsNullOrEmpty(zero_to_60) ? "" : $"&{nameof(zero_to_60).Replace('_', '-')}={zero_to_60}");
            sb.Append(string.IsNullOrEmpty(aggregatedTrim) ? "" : $"&{nameof(aggregatedTrim).Replace('_', '-')}={aggregatedTrim}");
            sb.Append(maximum_seats == default ? "" : $"&{nameof(maximum_seats).Replace('_', '-')}={maximum_seats}");
            sb.Append(minimum_seats == default ? "" : $"&{nameof(minimum_seats).Replace('_', '-')}={minimum_seats}");
            sb.Append(quantity_of_doors == default ? "" : $"&{nameof(quantity_of_doors).Replace('_', '-')}={quantity_of_doors}");

            sb.Append("&include-delivery-option=on&advertising-location=at_cars&page=");

            return sb.ToString().Replace("?&", "?");
        }
    }
}
