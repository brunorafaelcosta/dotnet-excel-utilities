using System.Collections.Generic;

namespace dotnet_excel_utilities
{
    public static class DataSet
    {
        public static IEnumerable<Region> GetData()
        {
            var data = new List<Region>();

            data.Add(new Region()
            {
                Name = "ASIA (EX. NEAR EAST)",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Afghanistan ", Population = "31056997", Area = "647500" },
                    new Country() { Name = "Bangladesh ", Population = "147365352", Area = "144000" },
                    new Country() { Name = "Bhutan ", Population = "2279723", Area = "47000" },
                    new Country() { Name = "Brunei ", Population = "379444", Area = "5770" },
                    new Country() { Name = "Burma ", Population = "47382633", Area = "678500" },
                    new Country() { Name = "Cambodia ", Population = "13881427", Area = "181040" },
                    new Country() { Name = "China ", Population = "1313973713", Area = "9596960" },
                    new Country() { Name = "East Timor ", Population = "1062777", Area = "15007" },
                    new Country() { Name = "Hong Kong ", Population = "6940432", Area = "1092" },
                    new Country() { Name = "India ", Population = "1095351995", Area = "3287590" },
                    new Country() { Name = "Indonesia ", Population = "245452739", Area = "1919440" },
                    new Country() { Name = "Iran ", Population = "68688433", Area = "1648000" },
                    new Country() { Name = "Japan ", Population = "127463611", Area = "377835" },
                    new Country() { Name = "Korea, North ", Population = "23113019", Area = "120540" },
                    new Country() { Name = "Korea, South ", Population = "48846823", Area = "98480" },
                    new Country() { Name = "Laos ", Population = "6368481", Area = "236800" },
                    new Country() { Name = "Macau ", Population = "453125", Area = "28" },
                    new Country() { Name = "Malaysia ", Population = "24385858", Area = "329750" },
                    new Country() { Name = "Maldives ", Population = "359008", Area = "300" },
                    new Country() { Name = "Mongolia ", Population = "2832224", Area = "1564116" },
                    new Country() { Name = "Nepal ", Population = "28287147", Area = "147181" },
                    new Country() { Name = "Pakistan ", Population = "165803560", Area = "803940" },
                    new Country() { Name = "Philippines ", Population = "89468677", Area = "300000" },
                    new Country() { Name = "Singapore ", Population = "4492150", Area = "693" },
                    new Country() { Name = "Sri Lanka ", Population = "20222240", Area = "65610" },
                    new Country() { Name = "Taiwan ", Population = "23036087", Area = "35980" },
                    new Country() { Name = "Thailand ", Population = "64631595", Area = "514000" },
                    new Country() { Name = "Vietnam ", Population = "84402966", Area = "329560" }
                }
            });

            data.Add(new Region()
            {
                Name = "BALTICS",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Estonia ", Population = "1324333", Area = "45226" },
                    new Country() { Name = "Latvia ", Population = "2274735", Area = "64589" },
                    new Country() { Name = "Lithuania ", Population = "3585906", Area = "65200" }
                }
            });

            data.Add(new Region()
            {
                Name = "C.W. OF IND. STATES",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Armenia ", Population = "2976372", Area = "29800" },
                    new Country() { Name = "Azerbaijan ", Population = "7961619", Area = "86600" },
                    new Country() { Name = "Belarus ", Population = "10293011", Area = "207600" },
                    new Country() { Name = "Georgia ", Population = "4661473", Area = "69700" },
                    new Country() { Name = "Kazakhstan ", Population = "15233244", Area = "2717300" },
                    new Country() { Name = "Kyrgyzstan ", Population = "5213898", Area = "198500" },
                    new Country() { Name = "Moldova ", Population = "4466706", Area = "33843" },
                    new Country() { Name = "Russia ", Population = "142893540", Area = "17075200" },
                    new Country() { Name = "Tajikistan ", Population = "7320815", Area = "143100" },
                    new Country() { Name = "Turkmenistan ", Population = "5042920", Area = "488100" },
                    new Country() { Name = "Ukraine ", Population = "46710816", Area = "603700" },
                    new Country() { Name = "Uzbekistan ", Population = "27307134", Area = "447400" }
                }
            });

            data.Add(new Region()
            {
                Name = "EASTERN EUROPE",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Albania ", Population = "3581655", Area = "28748" },
                    new Country() { Name = "Bosnia & Herzegovina ", Population = "4498976", Area = "51129" },
                    new Country() { Name = "Bulgaria ", Population = "7385367", Area = "110910" },
                    new Country() { Name = "Croatia ", Population = "4494749", Area = "56542" },
                    new Country() { Name = "Czech Republic ", Population = "10235455", Area = "78866" },
                    new Country() { Name = "Hungary ", Population = "9981334", Area = "93030" },
                    new Country() { Name = "Macedonia ", Population = "2050554", Area = "25333" },
                    new Country() { Name = "Poland ", Population = "38536869", Area = "312685" },
                    new Country() { Name = "Romania ", Population = "22303552", Area = "237500" },
                    new Country() { Name = "Serbia ", Population = "9396411", Area = "88361" },
                    new Country() { Name = "Slovakia ", Population = "5439448", Area = "48845" },
                    new Country() { Name = "Slovenia ", Population = "2010347", Area = "20273" }
                }
            });

            data.Add(new Region()
            {
                Name = "LATIN AMER. & CARIB",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Anguilla ", Population = "13477", Area = "102" },
                    new Country() { Name = "Antigua & Barbuda ", Population = "69108", Area = "443" },
                    new Country() { Name = "Argentina ", Population = "39921833", Area = "2766890" },
                    new Country() { Name = "Aruba ", Population = "71891", Area = "193" },
                    new Country() { Name = "Bahamas, The ", Population = "303770", Area = "13940" },
                    new Country() { Name = "Barbados ", Population = "279912", Area = "431" },
                    new Country() { Name = "Belize ", Population = "287730", Area = "22966" },
                    new Country() { Name = "Bolivia ", Population = "8989046", Area = "1098580" },
                    new Country() { Name = "Brazil ", Population = "188078227", Area = "8511965" },
                    new Country() { Name = "British Virgin Is. ", Population = "23098", Area = "153" },
                    new Country() { Name = "Cayman Islands ", Population = "45436", Area = "262" },
                    new Country() { Name = "Chile ", Population = "16134219", Area = "756950" },
                    new Country() { Name = "Colombia ", Population = "43593035", Area = "1138910" },
                    new Country() { Name = "Costa Rica ", Population = "4075261", Area = "51100" },
                    new Country() { Name = "Cuba ", Population = "11382820", Area = "110860" },
                    new Country() { Name = "Dominica ", Population = "68910", Area = "754" },
                    new Country() { Name = "Dominican Republic ", Population = "9183984", Area = "48730" },
                    new Country() { Name = "Ecuador ", Population = "13547510", Area = "283560" },
                    new Country() { Name = "El Salvador ", Population = "6822378", Area = "21040" },
                    new Country() { Name = "French Guiana ", Population = "199509", Area = "91000" },
                    new Country() { Name = "Grenada ", Population = "89703", Area = "344" },
                    new Country() { Name = "Guadeloupe ", Population = "452776", Area = "1780" },
                    new Country() { Name = "Guatemala ", Population = "12293545", Area = "108890" },
                    new Country() { Name = "Guyana ", Population = "767245", Area = "214970" },
                    new Country() { Name = "Haiti ", Population = "8308504", Area = "27750" },
                    new Country() { Name = "Honduras ", Population = "7326496", Area = "112090" },
                    new Country() { Name = "Jamaica ", Population = "2758124", Area = "10991" },
                    new Country() { Name = "Martinique ", Population = "436131", Area = "1100" },
                    new Country() { Name = "Mexico ", Population = "107449525", Area = "1972550" },
                    new Country() { Name = "Montserrat ", Population = "9439", Area = "102" },
                    new Country() { Name = "Netherlands Antilles ", Population = "221736", Area = "960" },
                    new Country() { Name = "Nicaragua ", Population = "5570129", Area = "129494" },
                    new Country() { Name = "Panama ", Population = "3191319", Area = "78200" },
                    new Country() { Name = "Paraguay ", Population = "6506464", Area = "406750" },
                    new Country() { Name = "Peru ", Population = "28302603", Area = "1285220" },
                    new Country() { Name = "Puerto Rico ", Population = "3927188", Area = "13790" },
                    new Country() { Name = "Saint Kitts & Nevis ", Population = "39129", Area = "261" },
                    new Country() { Name = "Saint Lucia ", Population = "168458", Area = "616" },
                    new Country() { Name = "Saint Vincent and the Grenadines ", Population = "117848", Area = "389" },
                    new Country() { Name = "Suriname ", Population = "439117", Area = "163270" },
                    new Country() { Name = "Trinidad & Tobago ", Population = "1065842", Area = "5128" },
                    new Country() { Name = "Turks & Caicos Is ", Population = "21152", Area = "430" },
                    new Country() { Name = "Uruguay ", Population = "3431932", Area = "176220" },
                    new Country() { Name = "Venezuela ", Population = "25730435", Area = "912050" },
                    new Country() { Name = "Virgin Islands ", Population = "108605", Area = "1910" }
                }
            });
            
            data.Add(new Region()
            {
                Name = "NEAR EAST",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Bahrain ", Population = "698585", Area = "665" },
                    new Country() { Name = "Cyprus ", Population = "784301", Area = "9250" },
                    new Country() { Name = "Gaza Strip ", Population = "1428757", Area = "360" },
                    new Country() { Name = "Iraq ", Population = "26783383", Area = "437072" },
                    new Country() { Name = "Israel ", Population = "6352117", Area = "20770" },
                    new Country() { Name = "Jordan ", Population = "5906760", Area = "92300" },
                    new Country() { Name = "Kuwait ", Population = "2418393", Area = "17820" },
                    new Country() { Name = "Lebanon ", Population = "3874050", Area = "10400" },
                    new Country() { Name = "Oman ", Population = "3102229", Area = "212460" },
                    new Country() { Name = "Qatar ", Population = "885359", Area = "11437" },
                    new Country() { Name = "Saudi Arabia ", Population = "27019731", Area = "1960582" },
                    new Country() { Name = "Syria ", Population = "18881361", Area = "185180" },
                    new Country() { Name = "Turkey ", Population = "70413958", Area = "780580" },
                    new Country() { Name = "United Arab Emirates ", Population = "2602713", Area = "82880" },
                    new Country() { Name = "West Bank ", Population = "2460492", Area = "5860" },
                    new Country() { Name = "Yemen ", Population = "21456188", Area = "527970" }
                }
            });
            
            data.Add(new Region()
            {
                Name = "NORTHERN AFRICA",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Algeria ", Population = "32930091", Area = "2381740" },
                    new Country() { Name = "Egypt ", Population = "78887007", Area = "1001450" },
                    new Country() { Name = "Libya ", Population = "5900754", Area = "1759540" },
                    new Country() { Name = "Morocco ", Population = "33241259", Area = "446550" },
                    new Country() { Name = "Tunisia ", Population = "10175014", Area = "163610" },
                    new Country() { Name = "Western Sahara ", Population = "273008", Area = "266000" }
                }
            });
            
            data.Add(new Region()
            {
                Name = "NORTHERN AMERICA",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Bermuda ", Population = "65773", Area = "53" },
                    new Country() { Name = "Canada ", Population = "33098932", Area = "9984670" },
                    new Country() { Name = "Greenland ", Population = "56361", Area = "2166086" },
                    new Country() { Name = "St Pierre & Miquelon ", Population = "7026", Area = "242" },
                    new Country() { Name = "United States ", Population = "298444215", Area = "9631420" }
                }
            });

            data.Add(new Region()
            {
                Name = "OCEANIA",
                Countries = new List<Country>()
                {
                    new Country() { Name = "American Samoa ", Population = "57794", Area = "199" },
                    new Country() { Name = "Australia ", Population = "20264082", Area = "7686850" },
                    new Country() { Name = "Cook Islands ", Population = "21388", Area = "240" },
                    new Country() { Name = "Fiji ", Population = "905949", Area = "18270" },
                    new Country() { Name = "French Polynesia ", Population = "274578", Area = "4167" },
                    new Country() { Name = "Guam ", Population = "171019", Area = "541" },
                    new Country() { Name = "Kiribati ", Population = "105432", Area = "811" },
                    new Country() { Name = "Marshall Islands ", Population = "60422", Area = "11854" },
                    new Country() { Name = "Micronesia, Fed. St. ", Population = "108004", Area = "702" },
                    new Country() { Name = "Nauru ", Population = "13287", Area = "21" },
                    new Country() { Name = "New Caledonia ", Population = "219246", Area = "19060" },
                    new Country() { Name = "New Zealand ", Population = "4076140", Area = "268680" },
                    new Country() { Name = "N. Mariana Islands ", Population = "82459", Area = "477" },
                    new Country() { Name = "Palau ", Population = "20579", Area = "458" },
                    new Country() { Name = "Papua New Guinea ", Population = "5670544", Area = "462840" },
                    new Country() { Name = "Samoa ", Population = "176908", Area = "2944" },
                    new Country() { Name = "Solomon Islands ", Population = "552438", Area = "28450" },
                    new Country() { Name = "Tonga ", Population = "114689", Area = "748" },
                    new Country() { Name = "Tuvalu ", Population = "11810", Area = "26" },
                    new Country() { Name = "Vanuatu ", Population = "208869", Area = "12200" },
                    new Country() { Name = "Wallis and Futuna ", Population = "16025", Area = "274" }
                }
            });

            data.Add(new Region()
            {
                Name = "NEAR EASTSUB-SAHARAN AFRICA",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Angola ", Population = "12127071", Area = "1246700" },
                    new Country() { Name = "Benin ", Population = "7862944", Area = "112620" },
                    new Country() { Name = "Botswana ", Population = "1639833", Area = "600370" },
                    new Country() { Name = "Burkina Faso ", Population = "13902972", Area = "274200" },
                    new Country() { Name = "Burundi ", Population = "8090068", Area = "27830" },
                    new Country() { Name = "Cameroon ", Population = "17340702", Area = "475440" },
                    new Country() { Name = "Cape Verde ", Population = "420979", Area = "4033" },
                    new Country() { Name = "Central African Rep. ", Population = "4303356", Area = "622984" },
                    new Country() { Name = "Chad ", Population = "9944201", Area = "1284000" },
                    new Country() { Name = "Comoros ", Population = "690948", Area = "2170" },
                    new Country() { Name = "Congo, Dem. Rep. ", Population = "62660551", Area = "2345410" },
                    new Country() { Name = "Congo, Repub. of the ", Population = "3702314", Area = "342000" },
                    new Country() { Name = @"Cote d""Ivoire ", Population = "17654843", Area = "322460" },
                    new Country() { Name = "Djibouti ", Population = "486530", Area = "23000" },
                    new Country() { Name = "Equatorial Guinea ", Population = "540109", Area = "28051" },
                    new Country() { Name = "Eritrea ", Population = "4786994", Area = "121320" },
                    new Country() { Name = "Ethiopia ", Population = "74777981", Area = "1127127" },
                    new Country() { Name = "Gabon ", Population = "1424906", Area = "267667" },
                    new Country() { Name = "Gambia, The ", Population = "1641564", Area = "11300" },
                    new Country() { Name = "Ghana ", Population = "22409572", Area = "239460" },
                    new Country() { Name = "Guinea ", Population = "9690222", Area = "245857" },
                    new Country() { Name = "Guinea-Bissau ", Population = "1442029", Area = "36120" },
                    new Country() { Name = "Kenya ", Population = "34707817", Area = "582650" },
                    new Country() { Name = "Lesotho ", Population = "2022331", Area = "30355" },
                    new Country() { Name = "Liberia ", Population = "3042004", Area = "111370" },
                    new Country() { Name = "Madagascar ", Population = "18595469", Area = "587040" },
                    new Country() { Name = "Malawi ", Population = "13013926", Area = "118480" },
                    new Country() { Name = "Mali ", Population = "11716829", Area = "1240000" },
                    new Country() { Name = "Mauritania ", Population = "3177388", Area = "1030700" },
                    new Country() { Name = "Mauritius ", Population = "1240827", Area = "2040" },
                    new Country() { Name = "Mayotte ", Population = "201234", Area = "374" },
                    new Country() { Name = "Mozambique ", Population = "19686505", Area = "801590" },
                    new Country() { Name = "Namibia ", Population = "2044147", Area = "825418" },
                    new Country() { Name = "Niger ", Population = "12525094", Area = "1267000" },
                    new Country() { Name = "Nigeria ", Population = "131859731", Area = "923768" },
                    new Country() { Name = "Reunion ", Population = "787584", Area = "2517" },
                    new Country() { Name = "Rwanda ", Population = "8648248", Area = "26338" },
                    new Country() { Name = "Saint Helena ", Population = "7502", Area = "413" },
                    new Country() { Name = "Sao Tome & Principe ", Population = "193413", Area = "1001" },
                    new Country() { Name = "Senegal ", Population = "11987121", Area = "196190" },
                    new Country() { Name = "Seychelles ", Population = "81541", Area = "455" },
                    new Country() { Name = "Sierra Leone ", Population = "6005250", Area = "71740" },
                    new Country() { Name = "Somalia ", Population = "8863338", Area = "637657" },
                    new Country() { Name = "South Africa ", Population = "44187637", Area = "1219912" },
                    new Country() { Name = "Sudan ", Population = "41236378", Area = "2505810" },
                    new Country() { Name = "Swaziland ", Population = "1136334", Area = "17363" },
                    new Country() { Name = "Tanzania ", Population = "37445392", Area = "945087" },
                    new Country() { Name = "Togo ", Population = "5548702", Area = "56785" },
                    new Country() { Name = "Uganda ", Population = "28195754", Area = "236040" },
                    new Country() { Name = "Zambia ", Population = "11502010", Area = "752614" },
                    new Country() { Name = "Zimbabwe ", Population = "12236805", Area = "390580" }
                }
            });

            data.Add(new Region()
            {
                Name = "WESTERN EUROPE",
                Countries = new List<Country>()
                {
                    new Country() { Name = "Andorra ", Population = "71201", Area = "468" },
                    new Country() { Name = "Austria ", Population = "8192880", Area = "83870" },
                    new Country() { Name = "Belgium ", Population = "10379067", Area = "30528" },
                    new Country() { Name = "Denmark ", Population = "5450661", Area = "43094" },
                    new Country() { Name = "Faroe Islands ", Population = "47246", Area = "1399" },
                    new Country() { Name = "Finland ", Population = "5231372", Area = "338145" },
                    new Country() { Name = "France ", Population = "60876136", Area = "547030" },
                    new Country() { Name = "Germany ", Population = "82422299", Area = "357021" },
                    new Country() { Name = "Gibraltar ", Population = "27928", Area = "7" },
                    new Country() { Name = "Greece ", Population = "10688058", Area = "131940" },
                    new Country() { Name = "Guernsey ", Population = "65409", Area = "78" },
                    new Country() { Name = "Iceland ", Population = "299388", Area = "103000" },
                    new Country() { Name = "Ireland ", Population = "4062235", Area = "70280" },
                    new Country() { Name = "Isle of Man ", Population = "75441", Area = "572" },
                    new Country() { Name = "Italy ", Population = "58133509", Area = "301230" },
                    new Country() { Name = "Jersey ", Population = "91084", Area = "116" },
                    new Country() { Name = "Liechtenstein ", Population = "33987", Area = "160" },
                    new Country() { Name = "Luxembourg ", Population = "474413", Area = "2586" },
                    new Country() { Name = "Malta ", Population = "400214", Area = "316" },
                    new Country() { Name = "Monaco ", Population = "32543", Area = "2" },
                    new Country() { Name = "Netherlands ", Population = "16491461", Area = "41526" },
                    new Country() { Name = "Norway ", Population = "4610820", Area = "323802" },
                    new Country() { Name = "Portugal ", Population = "10605870", Area = "92391" },
                    new Country() { Name = "San Marino ", Population = "29251", Area = "61" },
                    new Country() { Name = "Spain ", Population = "40397842", Area = "504782" },
                    new Country() { Name = "Sweden ", Population = "9016596", Area = "449964" },
                    new Country() { Name = "Switzerland ", Population = "7523934", Area = "41290" },
                    new Country() { Name = "United Kingdom ", Population = "60609153", Area = "244820" }
                }
            });

            return data;
        }

        [ExcelUtilities.ExportTable("Regions", HasChildren = true)]
        public class Region : ExcelUtilities.IData
        {
            [ExcelUtilities.ExportColumn()]
            public string Name { get; set; }

            [ExcelUtilities.ExportColumnTable(IsCollapsed = true)]
            public IEnumerable<Country> Countries { get; set; }
        }

        [ExcelUtilities.ExportTable("Countries", HasChildren = false)]
        public class Country : ExcelUtilities.IData
        {
            [ExcelUtilities.ExportColumn()]
            public string Name { get; set; }

            [ExcelUtilities.ExportColumn()]
            public string Population { get; set; }

            [ExcelUtilities.ExportColumn(Title = "Area (sq. mi.)")]
            public string Area { get; set; }
        }
    }
}
