using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using OfficeOpenXml;

class Program
{
    static async System.Threading.Tasks.Task Main(string[] args)
    {
        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Компании");

            // Add column headers
            string[] headers = { "Название", "Описание", "Адрес", "Телефон", "Почта", "Сайт" };
            worksheet.Cells["A1"].LoadFromArrays(new List<string[]>() { headers });

            // Read links from the file
            List<string> links = new List<string>();
            using (StreamReader reader = new StreamReader(@"C:\Users\koval\RiderProjects\PARSER GOVNA\PARSER GOVNA\links.txt"))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    links.Add(line);
                }
            }

            // Base URL for company pages
            string baseUrl = "https://manufacturers.ru";

            using (HttpClient httpClient = new HttpClient())
            {
                foreach (string link in links)
                {
                    string fullLink = baseUrl + link;
                    HttpResponseMessage response = await httpClient.GetAsync(fullLink);
                    string html = await response.Content.ReadAsStringAsync();

                    HtmlDocument doc = new HtmlDocument();
                    doc.LoadHtml(html);

                    // Extract company information
                    string company_name = doc.DocumentNode.SelectSingleNode("//h1[@class='like-h1']").InnerText;
                    string description = doc.DocumentNode.SelectSingleNode("//p").InnerText;

                    // Extract contact information
                    HtmlNode contact_info = doc.DocumentNode.SelectSingleNode("//div[@id='contact-list']");

                    // Extract address information using regular expression
                    HtmlNode addressNode = contact_info.SelectSingleNode("td[contains(text(), 'Адрес')]/following-sibling::td/p");
                    string address = addressNode != null ? addressNode.InnerText.Trim() : "Адрес не найден";


                    string phonePattern = @"tel:(\+?[0-9() -]+)";
                    Match phoneMatch = Regex.Match(html, phonePattern);
                    string phone = phoneMatch.Success ? phoneMatch.Groups[1].Value : "Телефон не найден";

                    string emailPattern = @"mailto:([\w\.-]+@[\w\.-]+)";
                    Match emailMatch = Regex.Match(html, emailPattern);

                    string email = emailMatch.Success ? emailMatch.Groups[1].Value : "Почта не найдена";
                    // Extract website information using regular expression
                    string websitePattern = @"<a href=""(https?://[^""]+)""";
                    Match websiteMatch = Regex.Match(html, websitePattern);

                    string website = websiteMatch.Success ? websiteMatch.Groups[1].Value : "Сайт не найден";


                    // Create company info array
                    string[] company_info = { company_name, description, address, phone, email, website };

                    // Add the company info to the worksheet
                    worksheet.Cells[worksheet.Dimension.End.Row + 1, 1].LoadFromArrays(new List<string[]>() { company_info });

                    // Output debug information
                    Console.WriteLine($"Ссылка: {baseUrl}{link}");
                    Console.WriteLine($"Добавлена компания: {company_name}");
                    Console.WriteLine($"Адрес: {address}");
                    Console.WriteLine($"Телефон: {phone}");
                    Console.WriteLine($"Почта: {email}");
                    Console.WriteLine($"Сайт: {website}");
                }
            }

            // Save the Excel file
            FileInfo excelFile = new FileInfo("companies.xlsx");
            package.SaveAs(excelFile);
        }
    }
}
