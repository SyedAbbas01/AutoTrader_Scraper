using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.IO;
using HtmlAgilityPack;
using System.Xml.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using OfficeOpenXml;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Windows.Media;

namespace AutoTrader_Scraper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public ObservableCollection<Car> Cars { get; set; }
        public bool ComboBoxesPopulated { get; private set; } = false;

        public MainWindow()
        {
            InitializeComponent();
            Cars = new ObservableCollection<Car>();

            dgResults.ItemsSource = Cars;

            InitializeOptions();
        }

        public T DeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;

                return (T)formatter.Deserialize(ms);
            }
        }

        async void InitializeOptions()
        {
            try
            {
                //URL to pull option details from
                var url = "https://www.autotrader.co.uk/search-form?moreOptions=visible";

                var httpClient = new HttpClient();

                var html = await httpClient.GetStringAsync(url);

                var doc = new HtmlDocument();
                doc.LoadHtml(html);

                StringBuilder sb = new StringBuilder();

                var dropDowns = doc.DocumentNode.Descendants("select");

                List<Options> bodyTypes = doc.DocumentNode.Descendants("ul")
                    .Where(x => x.GetAttributes().Any(x => x.Name == "class" && x.Value == "body-type-selector__container"))
                    .SelectMany(y => y.Descendants("input"))
                    .Select(z => new Options() { Header = z.Id, Value = z.Id })
                    .ToList();

                bodyTypes.Insert(0, new Options() { Header = "", Value = "dummy" });

                cboBodyType.ItemsSource = bodyTypes;
                cboBodyType.DisplayMemberPath = "Header";
                cboBodyType.SelectedValuePath = "Value";

                if (!string.IsNullOrEmpty(Settings1.Default.cboBodyType))
                {
                    cboBodyType.SelectedValue = Settings1.Default.cboBodyType;
                }

                if (Settings1.Default.postcodes == null)
                {
                    Settings1.Default.postcodes = new System.Collections.Specialized.StringCollection();
                    Settings1.Default.Save();
                }

                if (Settings1.Default.postcodes.Count > 0)
                {
                    cboPostCode.ItemsSource = Settings1.Default.postcodes;
                    cboPostCode.SelectedIndex = cboPostCode.Items.Count;
                }

                PopulateCombobox(null, cboYearFrom, "year");
                if (!string.IsNullOrEmpty(Settings1.Default.cboYearFrom))
                {
                    cboYearFrom.SelectedValue = Settings1.Default.cboYearFrom;
                }

                PopulateCombobox(null, cboYearTo, "year");
                if (!string.IsNullOrEmpty(Settings1.Default.cboYearTo))
                {
                    cboYearTo.SelectedValue = Settings1.Default.cboYearTo;
                }

                if (!string.IsNullOrEmpty(Settings1.Default.postcode))
                {
                    cboPostCode.ItemsSource = Settings1.Default.postcodes;
                    cboPostCode.SelectedValue = Settings1.Default.postcode;
                }

                if (!string.IsNullOrEmpty(Settings1.Default.txtPages))
                {
                    txtPages.Text = Settings1.Default.txtPages;
                }

                foreach (var options in dropDowns)
                {
                    switch (options.Id)
                    {
                        case "minPrice":
                            PopulateCombobox(options, cboPriceFrom);

                            if (!string.IsNullOrEmpty(Settings1.Default.cboPriceFrom))
                            {
                                cboPriceFrom.SelectedValue = Settings1.Default.cboPriceFrom;
                            }
                            break;

                        case "bodyType":
                            PopulateCombobox(options, cboBodyType);

                            if (!string.IsNullOrEmpty(Settings1.Default.cboBodyType))
                            {
                                cboBodyType.SelectedValue = Settings1.Default.cboBodyType;
                            }
                            break;

                        case "maxPrice":
                            PopulateCombobox(options, cboPriceTo);

                            if (!string.IsNullOrEmpty(Settings1.Default.cboPriceTo))
                            {
                                cboPriceTo.SelectedValue = Settings1.Default.cboPriceTo;
                            }
                            break;
                        case "distance":
                            PopulateCombobox(options, cboRadius);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboRadius))
                            {
                                cboRadius.SelectedValue = Settings1.Default.cboRadius;
                            }
                            break;
                        case "make":
                            PopulateCombobox(options, cboMake);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMake))
                            {
                                cboMake.SelectedValue = Settings1.Default.cboMake;
                            }
                            break;
                        case "model":
                            if (!string.IsNullOrEmpty(Settings1.Default.model))
                            {
                                cboModel.ItemsSource = Settings1.Default.models;
                                cboModel.SelectedValue = Settings1.Default.model;
                            }
                            break;

                        case "maxMileage":
                            PopulateCombobox(options, cboMaxMileage);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMaxMileage))
                            {
                                cboMaxMileage.SelectedValue = Settings1.Default.cboMaxMileage;
                            }
                            break;
                        case "transmission":
                            PopulateCombobox(options, cboTransmission);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboTransmission))
                            {
                                cboTransmission.SelectedValue = Settings1.Default.cboTransmission;
                            }
                            break;
                        case "doorsValue":
                            PopulateCombobox(options, cboDoorsValue);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboDoorsValue))
                            {
                                cboDoorsValue.SelectedValue = Settings1.Default.cboDoorsValue;
                            }
                            break;
                        case "minSeats":
                            PopulateCombobox(options, cboMinSeats);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMinSeats))
                            {
                                cboMinSeats.SelectedValue = Settings1.Default.cboMinSeats;
                            }
                            break;
                        case "maxSeats":
                            PopulateCombobox(options, cboMaxSeats);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMaxSeats))
                            {
                                cboMaxSeats.SelectedValue = Settings1.Default.cboMaxSeats;
                            }
                            break;
                        case "minEngineSizeLitres":
                            PopulateCombobox(options, cboMinEngineSizeLitres);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMinEngineSizeLitres))
                            {
                                cboMinEngineSizeLitres.SelectedValue = Settings1.Default.cboMinEngineSizeLitres;
                            }
                            break;
                        case "maxEngineSizeLitres":
                            PopulateCombobox(options, cboMaxEngineSizeLitres);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMaxEngineSizeLitres))
                            {
                                cboMaxEngineSizeLitres.SelectedValue = Settings1.Default.cboMaxEngineSizeLitres;
                            }
                            break;
                        case "minEnginePower":
                            PopulateCombobox(options, cboMinEnginePower);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMinEnginePower))
                            {
                                cboMinEnginePower.SelectedValue = Settings1.Default.cboMinEnginePower;
                            }
                            break;
                        case "maxEnginePower":
                            PopulateCombobox(options, cboMaxEnginePower);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMaxEnginePower))
                            {
                                cboMaxEnginePower.SelectedValue = Settings1.Default.cboMaxEnginePower;
                            }
                            break;
                        case "accelerationValue":
                            PopulateCombobox(options, cboAccelerationValue);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboAccelerationValue))
                            {
                                cboAccelerationValue.SelectedValue = Settings1.Default.cboAccelerationValue;
                            }
                            break;
                        case "drivetrain":
                            PopulateCombobox(options, cboDrivetrain);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboDrivetrain))
                            {
                                cboDrivetrain.SelectedValue = Settings1.Default.cboDrivetrain;
                            }
                            break;
                        case "annualTaxValue":
                            PopulateCombobox(options, cboAnnualTaxValue);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboAnnualTaxValue))
                            {
                                cboAnnualTaxValue.SelectedValue = Settings1.Default.cboAnnualTaxValue;
                            }
                            break;
                        case "maxInsuranceGroup":
                            PopulateCombobox(options, cboMaxInsuranceGroup);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboMaxInsuranceGroup))
                            {
                                cboMaxInsuranceGroup.SelectedValue = Settings1.Default.cboMaxInsuranceGroup;
                            }
                            break;
                        case "fuelConsumptionValue":
                            PopulateCombobox(options, cboFuelConsumptionValue);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboFuelConsumptionValue))
                            {
                                cboFuelConsumptionValue.SelectedValue = Settings1.Default.cboFuelConsumptionValue;
                            }
                            break;
                        case "co2EmissionValue":
                            PopulateCombobox(options, cboCo2EmissionValue);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboCo2EmissionValue))
                            {
                                cboCo2EmissionValue.SelectedValue = Settings1.Default.cboCo2EmissionValue;
                            }
                            break;
                        case "sellerType":
                            PopulateCombobox(options, cboSellerType);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboSellerType))
                            {
                                cboSellerType.SelectedValue = Settings1.Default.cboSellerType;
                            }
                            break;
                        case "showWriteOff":
                            PopulateCombobox(options, cboShowWriteOff);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboShowWriteOff))
                            {
                                cboShowWriteOff.SelectedValue = Settings1.Default.cboShowWriteOff;
                            }
                            break;
                        case "colour":
                            PopulateCombobox(options, cboColour);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboColour))
                            {
                                cboColour.SelectedValue = Settings1.Default.cboColour;
                            }
                            break;
                        case "fuelType":
                            PopulateCombobox(options, cboFuelType);
                            if (!string.IsNullOrEmpty(Settings1.Default.cboFuelType))
                            {
                                cboFuelType.SelectedValue = Settings1.Default.cboFuelType;
                            }
                            break;
                        default:
                            break;
                    }
                }

                ComboBoxesPopulated = true;
            }
            catch (Exception ex)
            {

                //throw;
            }
        }

        private int Extract(string url)
        {
            int counter = 0;

            try
            {
                HttpClient httpClient = new HttpClient();
                string html = httpClient.GetStringAsync(url).Result;

                MatchCollection? m = Regex.Matches(html, "(?<=\")(.*?)(?=\")");

                foreach (Match mc in m.Where(c => !string.IsNullOrEmpty(c.Value)))
                {
                    html = html.Replace(mc.Value, mc.Value.Replace("<", "d").Replace(">", "gt"));
                }

                var doc = new HtmlDocument();
                var doc2 = new HtmlDocument();
                doc.LoadHtml(html);

                var sections = doc.DocumentNode.Descendants("section").Where(x => x.GetAttributes()
                    .Any(x => x.Name == "class" && x.Value == "product-card-details"));

                var allsections = doc.DocumentNode.DescendantNodes();

                var car = new Car();

                foreach (var item in allsections)
                {
                    if (item.Name == "li" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "search-page__result"))
                    {
                        doc2.LoadHtml(item.OuterHtml);
                        car.URL = @"https://www.autotrader.co.uk/car-details/" + item.Id;

                        car.Distance = item.GetAttributeValue<string>("data-distance-value", "").ReplaceLineEndings("");
                    }

                    if (item.Name == "h3" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "product-card-details__title"))
                    {
                        car.Name = item.InnerText.ReplaceLineEndings("");
                    }

                    if (item.Name == "p" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "product-card-details__subtitle"))
                    {
                        car.Trim = item.InnerText.ReplaceLineEndings("");
                    }

                    if (item.Name == "p" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "product-card-details__attention-grabber"))
                    {
                        car.MiscDetails = item.InnerText.ReplaceLineEndings("");
                    }

                    if (item.Name == "div" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "product-card-pricing__price"))
                    {
                        car.Price = Int32.Parse(item.ChildNodes.FirstOrDefault(x => x.Name == "span").InnerHtml.Replace("£", "").Replace(",", ""));
                    }


                    if (item.Name == "ul" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "listing-key-specs"))
                    {
                        foreach (var detailItem in item.Descendants("li"))
                        {
                            var match = Regex.Match(detailItem.InnerText, "[0-9][0-9][0-9][0-9]");
                            if (match.Success)
                            {
                                car.Year = Int32.Parse(match.Value);
                                car.Reg = detailItem.InnerText.Replace(match.Value, "");
                            }

                            match = Regex.Match(detailItem.InnerText, "miles");
                            if (match.Success)
                            {
                                car.Mileage = Int32.Parse(detailItem.InnerText.Replace("miles", "").Replace(",", ""));
                            }

                            match = Regex.Match(detailItem.InnerText, "[0-9][PS|BHP]");
                            if (match.Success)
                            {
                                car.HorsePower = Int32.Parse(detailItem.InnerText.Replace("PS", "").Replace("BHP", ""));
                            }

                            match = Regex.Match(detailItem.InnerText, "^Petrol$|^Diesel$|^Electric$");
                            if (match.Success)
                            {
                                car.Fuel = detailItem.InnerText;
                            }

                            match = Regex.Match(detailItem.InnerText, "^Automatic$|^Manual$");
                            if (match.Success)
                            {
                                car.Transmission = detailItem.InnerText;
                            }

                            match = Regex.Match(detailItem.InnerText, "[0-9]L");
                            if (match.Success)
                            {
                                car.Litre = Decimal.Parse(detailItem.InnerText.Replace("L", ""));
                            }
                        }
                    }

                    if (item.Name == "div" && item.GetAttributes().Any(x => x.Name == "class" && x.Value == "product-card-seller-info"))
                    {
                        if (!Cars.Any(c => c.ToString() == car.ToString()))
                        {
                            Cars.Add(DeepClone(car));
                            counter++;
                        }

                        car = new Car();
                    }
                }

                return counter;
            }
            catch (Exception ex)
            {
                return counter;
            }
        }

        private static void PopulateCombobox(HtmlNode options, ComboBox comboBox, string bypass = null)
        {
            var list = new List<Options>();
            list.Add(new Options() { Header = "", Value = "Dummy" });

            if (bypass != null)
            {
                switch (bypass)
                {
                    case "year":
                        for (int i = 1975; i < DateTime.Now.Year + 1; i++)
                        {
                            var option = new Options();
                            option.Header = i.ToString();
                            option.Value = i.ToString();

                            list.Add(option);
                        }
                        break;
                    default:
                        break;
                }
            }
            else
                foreach (var choice in options.Descendants("option"))
                {
                    var element = XElement.Parse(choice.OuterHtml);
                    if (element.Attribute("value")?.Value != "" || element?.Value == "National")
                    {
                        var option = new Options();

                        option.Header = (element?.Value ?? "");
                        option.Value = element?.Value == "National" ? "1500" : element.Attribute("value")?.Value;

                        int index = option.Header.IndexOf(" (");
                        if (index > 0)
                            option.Header = option.Header.Substring(0, index);

                        list.Add(option);
                    }
                }

            if (list.Count > 1)
            {
                comboBox.ItemsSource = list;
                comboBox.DisplayMemberPath = "Header";
                comboBox.SelectedValuePath = "Value";

                int index = list.FindIndex(x => x.Value == DateTime.Now.Year.ToString());

                if (bypass != null && index > -1)
                {
                    comboBox.SelectedIndex = index;
                }
                else
                    comboBox.SelectedIndex = 0;
            }
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnExcelExport.IsEnabled = false;
                Cars.Clear();

                if (string.IsNullOrEmpty(cboPostCode.Text))
                {
                    MessageBox.Show("Post code cannot be blank");
                    return;
                }
                else
                {
                    if (!Settings1.Default.postcodes.Contains(cboPostCode.Text.ToUpper()))
                    {
                        Settings1.Default.postcodes.Add(cboPostCode.Text.ToUpper());
                        Settings1.Default.Save();
                    }
                }

                SearchCriteria criteria = new SearchCriteria();

                criteria.make = GetComboBoxStringValue(cboMake);
                criteria.model = GetComboBoxStringValue(cboModel);

                criteria.postcode = GetComboBoxStringValue(cboPostCode);
                criteria.distance = GetComboBoxIntValue(cboRadius);

                criteria.price_from = GetComboBoxIntValue(cboPriceFrom);
                criteria.price_to = GetComboBoxIntValue(cboPriceTo);

                criteria.year_from = GetComboBoxIntValue(cboYearFrom);
                criteria.year_to = GetComboBoxIntValue(cboYearTo);

                criteria.zero_to_60 = GetComboBoxStringValue(cboAccelerationValue);
                criteria.minimum_badge_engine_size = GetComboBoxStringValue(cboMinEngineSizeLitres);
                criteria.maximum_badge_engine_size = GetComboBoxStringValue(cboMaxEngineSizeLitres);

                // criteria.aggregatedTrim = todo

                criteria.body_type = GetComboBoxStringValue(cboBodyType);
                criteria.drivetrain = GetComboBoxStringValue(cboDrivetrain);
                criteria.colour = GetComboBoxStringValue(cboColour);
                criteria.quantity_of_doors = GetComboBoxIntValue(cboDoorsValue);

                criteria.fuel_type = GetComboBoxStringValue(cboFuelType);
                criteria.fuel_consumption = GetComboBoxStringValue(cboFuelConsumptionValue);

                criteria.maximum_seats = GetComboBoxIntValue(cboMaxSeats);
                criteria.minimum_seats = GetComboBoxIntValue(cboMinSeats);

                criteria.min_engine_power = GetComboBoxStringValue(cboMinEnginePower);
                criteria.max_engine_power = GetComboBoxStringValue(cboMaxEnginePower);

                criteria.insuranceGroup = GetComboBoxStringValue(cboMaxInsuranceGroup);
                criteria.exclude_writeoff_categories = GetComboBoxStringValue(cboShowWriteOff);
                criteria.seller_type = GetComboBoxStringValue(cboSellerType);

                criteria.annual_tax_cars = GetComboBoxStringValue(cboAnnualTaxValue);
                criteria.co2_emissions_cars = GetComboBoxStringValue(cboCo2EmissionValue);
                criteria.maximum_mileage = GetComboBoxStringValue(cboMaxMileage);

                string url = criteria.GetSearchURL();

                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

                for (int i = 1; i <= Convert.ToInt32(txtPages.Text); i++)
                {
                    int counter = Extract($"{url}{i}");

                    if (counter == 0)
                    {
                        MessageBox.Show($"No results for page {i}.{Environment.NewLine}Returning.");
                        break;
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (Cars.Count > 0)
                    btnExcelExport.IsEnabled = true;

                Mouse.OverrideCursor = null;
            }
        }

        private string GetComboBoxStringValue(ComboBox comboBox)
        {
            if (comboBox.SelectedIndex > 0) //skip dummy item
            {
                return comboBox.SelectedValue.ToString();
            }

            if (!string.IsNullOrEmpty(comboBox.Text))
            {
                return comboBox.Text;
            }

            return "";
        }

        private int GetComboBoxIntValue(ComboBox comboBox)
        {
            if (comboBox.SelectedIndex > 0 && int.TryParse(comboBox.Text, out int parsedvalue)) //skip dummy item
            {
                return parsedvalue;
            }

            if (!string.IsNullOrEmpty(comboBox.Text) && int.TryParse(comboBox.Text, out int parsedtext))
            {
                return parsedtext;
            }

            return default;
        }

        private void txtPages_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");
            // e.Handled =if
            var match = Regex.Match(e.Text, "[^0-9]+");

            if (!match.Success)
                Settings1.Default.txtPages = e.Text;

            Settings1.Default.Save();

            e.Handled = match.Success;
        }

        private void SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;

            if (ComboBoxesPopulated && (cb.SelectedIndex > -1))
            {
                switch (cb.Tag)
                {
                    case "make":
                        Settings1.Default.cboMake = cb.SelectedValue.ToString();
                        break;
                    case "yearfrom":
                        Settings1.Default.cboYearFrom = cb.SelectedValue.ToString();
                        break;
                    case "yearto":
                        Settings1.Default.cboYearTo = cb.SelectedValue.ToString();
                        break;
                    case "pricefrom":
                        Settings1.Default.cboPriceFrom = cb.SelectedValue.ToString();
                        break;
                    case "priceto":
                        Settings1.Default.cboPriceTo = cb.SelectedValue.ToString();
                        break;
                    case "radius":
                        Settings1.Default.cboRadius = cb.SelectedValue.ToString();
                        break;
                    case "bodytype":
                        Settings1.Default.cboBodyType = cb.SelectedValue.ToString();
                        break;
                    case "fuelType":
                        Settings1.Default.cboFuelType = cb.SelectedValue.ToString();
                        break;
                    case "transmission":
                        Settings1.Default.cboTransmission = cb.SelectedValue.ToString();
                        break;
                    case "colour":
                        Settings1.Default.cboColour = cb.SelectedValue.ToString();
                        break;
                    case "minSeats":
                        Settings1.Default.cboMinSeats = cb.SelectedValue.ToString();
                        break;
                    case "maxSeats":
                        Settings1.Default.cboMaxSeats = cb.SelectedValue.ToString();
                        break;
                    case "maxMileage":
                        Settings1.Default.cboMaxMileage = cb.SelectedValue.ToString();
                        break;
                    case "doorsValue":
                        Settings1.Default.cboDoorsValue = cb.SelectedValue.ToString();
                        break;
                    case "minEngineSizeLitres":
                        Settings1.Default.cboMinEngineSizeLitres = cb.SelectedValue.ToString();
                        break;
                    case "maxEngineSizeLitres":
                        Settings1.Default.cboMaxEngineSizeLitres = cb.SelectedValue.ToString();
                        break;
                    case "minEnginePower":
                        Settings1.Default.cboMinEnginePower = cb.SelectedValue.ToString();
                        break;
                    case "maxEnginePower":
                        Settings1.Default.cboMaxEnginePower = cb.SelectedValue.ToString();
                        break;
                    case "accelerationValue":
                        Settings1.Default.cboAccelerationValue = cb.SelectedValue.ToString();
                        break;
                    case "drivetrain":
                        Settings1.Default.cboDrivetrain = cb.SelectedValue.ToString();
                        break;
                    case "annualTaxValue":
                        Settings1.Default.cboAnnualTaxValue = cb.SelectedValue.ToString();
                        break;
                    case "maxInsuranceGroup":
                        Settings1.Default.cboMaxInsuranceGroup = cb.SelectedValue.ToString();
                        break;
                    case "fuelConsumptionValue":
                        Settings1.Default.cboFuelConsumptionValue = cb.SelectedValue.ToString();
                        break;
                    case "co2EmissionValue":
                        Settings1.Default.cboCo2EmissionValue = cb.SelectedValue.ToString();
                        break;
                    case "sellerType":
                        Settings1.Default.cboSellerType = cb.SelectedValue.ToString();
                        break;
                    case "showWriteOff":
                        Settings1.Default.cboShowWriteOff = cb.SelectedValue.ToString();
                        break;
                    default:
                        break;
                }

                Settings1.Default.Save();
            }
        }

        private void cboKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                var cb = (ComboBox)sender;
                var option = cb.Text;

                if (!string.IsNullOrWhiteSpace(option))
                {
                    switch (cb.Tag)
                    {
                        case "postcode":
                            if (Settings1.Default.postcodes == null)
                                Settings1.Default.postcodes = new System.Collections.Specialized.StringCollection();

                            var postcodes = Settings1.Default.postcodes;
                            var postcode = cb.Text.ToString().ToUpper().Trim().Replace(" ", "");

                            if (!postcodes.Contains(postcode))
                            {
                                Settings1.Default.postcodes.Add(postcode);
                                Settings1.Default.postcode = postcode;

                                cb.ItemsSource = null;
                                cb.ItemsSource = Settings1.Default.postcodes;
                                cb.SelectedValue = postcode;
                            }

                            break;
                        case "model":

                            if (Settings1.Default.models == null)
                                Settings1.Default.models = new System.Collections.Specialized.StringCollection();

                            var models = Settings1.Default.models;
                            var model = cb.Text.ToString().ToUpper().Trim().Replace(" ", "");

                            if (!models.Contains(model))
                            {
                                Settings1.Default.models.Add(model);
                                Settings1.Default.model = model;

                                cb.ItemsSource = null;
                                cb.ItemsSource = Settings1.Default.models;
                                cb.SelectedValue = model;
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void btnExcelExport_Click(object sender, RoutedEventArgs e)
        {
            if (Cars.Count > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel file (*.xlsx)|*.xlsx";

                if (string.IsNullOrEmpty(Settings1.Default.saveDirectory) || !Directory.Exists(Settings1.Default.saveDirectory))
                {
                    saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                else
                    saveFileDialog.InitialDirectory = Settings1.Default.saveDirectory;

                if (saveFileDialog.ShowDialog() == true)
                {
                    Settings1.Default.saveDirectory = Path.GetDirectoryName(saveFileDialog.FileName);
                    Settings1.Default.Save();

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (var package = new ExcelPackage(new FileInfo(saveFileDialog.FileName)))
                    {
                        var sheet = package.Workbook.Worksheets.Add("AutoTrader results");

                        sheet.Cells["A1"].LoadFromCollection(Cars, true, OfficeOpenXml.Table.TableStyles.Medium9);

                        sheet.View.FreezePanes(2, 1);
                        sheet.Column(1).AutoFit();
                        sheet.Column(2).Style.Numberformat.Format = "#,##0";
                        sheet.Column(5).Style.Numberformat.Format = "£#,##0";

                        var start = sheet.Dimension.Columns;
                        var end = sheet.Dimension.End;

                        // Save to file
                        package.Save();

                    }

                    new Process { StartInfo = new ProcessStartInfo(saveFileDialog.FileName) { UseShellExecute = true } }.Start();
                }
            }
        }

        public IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj == null) yield return (T)Enumerable.Empty<T>();
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                DependencyObject ithChild = VisualTreeHelper.GetChild(depObj, i);
                if (ithChild == null) continue;
                if (ithChild is T t) yield return t;
                foreach (T childOfChild in FindVisualChildren<T>(ithChild)) yield return childOfChild;
            }
        }

        private void btnClearFilters_Click(object sender, RoutedEventArgs e)
        {
            foreach (ComboBox cb in FindVisualChildren<ComboBox>(this))
            {
                if (cb.IsEditable)
                    cb.Text = "";
                else
                    if (cb.SelectedIndex > 0)
                    cb.SelectedIndex = 0;
            }
        }
    }
}
