using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace GetInfoPBS
{
    class Program
    {
        private static readonly (string Cyrillic, string Latin)[] Mapping = new[]
        {
            ("А", "A"), ("Б", "B"), ("В", "V"), ("Г", "G"), ("Д", "D"), ("Ђ", "Đ"), ("Е", "E"), ("Ж", "Ž"), ("З", "Z"), ("И", "I"),
            ("Ј", "J"), ("К", "K"), ("Л", "L"), ("М", "M"), ("Н", "N"), ("О", "O"), ("П", "P"), ("Р", "R"), ("С", "S"), ("Т", "T"),
            ("Ћ", "Ć"), ("У", "U"), ("Ф", "F"), ("Х", "H"), ("Ц", "C"), ("Ч", "Č"), ("Ш", "Š"), ("Џ", "Dž"), ("Љ", "Lj"), ("Њ", "Nj"),
            ("а", "a"), ("б", "b"), ("в", "v"), ("г", "g"), ("д", "d"), ("ђ", "đ"), ("е", "e"), ("ж", "ž"), ("з", "z"), ("и", "i"),
            ("ј", "j"), ("к", "k"), ("л", "l"), ("м", "m"), ("н", "n"), ("о", "o"), ("п", "p"), ("р", "r"), ("с", "s"), ("т", "t"),
            ("ћ", "ć"), ("у", "u"), ("ф", "f"), ("х", "h"), ("ц", "c"), ("ч", "č"), ("ш", "š"), ("џ", "dž"), ("љ", "lj"), ("њ", "nj")
        };

        public static string toLatin(string input)
        {
            var sb = new StringBuilder(input);

            foreach (var (Cyrillic, Latin) in Mapping)
            {
                sb.Replace(Cyrillic, Latin);
            }

            return sb.ToString();
        }

        static void Main(string[] args)
        {
            // Prikupljanje informacija sa sajta
            Console.WriteLine("Unesite link kategorije PoslovneBazeSrbije");
            string readline = Console.ReadLine();
            Console.Write("---------------------------------------------------------------------------------"); Console.WriteLine(); Console.WriteLine();
            char ch = readline.Last();
            char ch1 = readline.Skip(readline.Length - 2).First();
            bool pom = Char.IsDigit(ch1);
            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load(readline);
            string baseUrl;
            if (pom)
            {
                baseUrl = "https://www.poslovnabazasrbije.rs/Search?query=&page={0}&category=" + ch1 + ch;
            }
            else
            {
                baseUrl = "https://www.poslovnabazasrbije.rs/Search?query=&page={0}&category=" + ch;
            }

            StringBuilder urlBuilder = new StringBuilder();
            var countPages = doc.DocumentNode.SelectNodes("//*[@id=\"divContent1\"]/div/div[22]/ul/li[7]/a").First().InnerText;
            string[,,] matrix = new string[Convert.ToInt32(countPages), 20, 4];
            int count = 2;

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Podaci");

                worksheet.Cells[1, 1].Value = "Naziv Kompanije";
                worksheet.Cells[1, 2].Value = "Kategorija";
                worksheet.Cells[1, 3].Value = "Adresa";
                worksheet.Cells[1, 4].Value = "Email";

                int rowIndex = 2;

                for (int i = 2; i <= Convert.ToInt32(countPages); i++)
                {
                    for (int j = 0; j < 20; j++)
                    {
                        var companyName = doc.DocumentNode.SelectNodes("//*[@id=\"divContent1\"]/div/div[" + count.ToString() + "]/p[1]").First().InnerText;
                        var companyCategory = doc.DocumentNode.SelectNodes("//*[@id=\"divContent1\"]/div/div[" + count.ToString() + "]/div[1]/div[1]").First().InnerText;
                        var companyAddress = doc.DocumentNode.SelectNodes("//*[@id=\"divContent1\"]/div/div[" + count.ToString() + "]/div[1]/div[2]").First().InnerText;
                        var companyEmail = doc.DocumentNode.SelectNodes("//*[@id=\"divContent1\"]/div/div[" + count.ToString() + "]/div[2]/div/p").First().InnerText;
                        companyEmail = companyEmail.Replace("-", "");
                        matrix[i - 2, j, 0] = toLatin(companyName).Trim();
                        matrix[i - 2, j, 1] = toLatin(companyCategory).Trim();
                        matrix[i - 2, j, 2] = toLatin(companyAddress).Trim();
                        matrix[i - 2, j, 3] = companyEmail.Trim();

                        worksheet.Cells[rowIndex, 1].Value = matrix[i - 2, j, 0];
                        worksheet.Cells[rowIndex, 2].Value = matrix[i - 2, j, 1];
                        worksheet.Cells[rowIndex, 3].Value = matrix[i - 2, j, 2];
                        worksheet.Cells[rowIndex, 4].Value = matrix[i - 2, j, 3];
                        rowIndex++;

                        Console.WriteLine(matrix[i - 2, j, 0]);
                        Console.WriteLine(matrix[i - 2, j, 1]);
                        Console.WriteLine(matrix[i - 2, j, 2]);
                        Console.WriteLine(matrix[i - 2, j, 3] + "\n" + "\n");

                        count++;
                    }
                    string url = string.Format(baseUrl, i);
                    urlBuilder.AppendLine(url);
                    doc = web.Load(url);
                    count = 2;
                }

                string filePath = "";
                if (pom) { filePath = Path.Combine(Directory.GetCurrentDirectory(), "PoslovnaBazaSrbije - Kategorija " + ch1 + ch + ".xlsx"); }
                else { filePath = Path.Combine(Directory.GetCurrentDirectory(), "PoslovnaBazaSrbije - Kategorija " + ch + ".xlsx"); }
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
                Console.WriteLine($"Podaci su snimljeni u {filePath}");
            }
        }
    }
}
