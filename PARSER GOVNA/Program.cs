using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using OfficeOpenXml;

class Program
{
    static async Task Main(string[] args)
    {
        string excelFilePath = @"C:\Users\koval\RiderProjects\PARSER GOVNA\PARSER GOVNA\Penis.xlsx";
        string linksFilePath = @"C:\Users\koval\RiderProjects\PARSER GOVNA\PARSER GOVNA\links.txt";

        await ProcessCompaniesAsync(excelFilePath, linksFilePath);
    }

    static async Task ProcessCompaniesAsync(string excelFilePath, string linksFilePath)
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Компании");

            string[] headers = { "Название", "Описание", "Адрес", "Телефон", "Почта", "Сайт" };
            worksheet.Cells["A1"].LoadFromArrays(new List<string[]>() { headers });

            List<string> links = ReadLinksFromFile(linksFilePath);
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

                    string company_name = doc.DocumentNode.SelectSingleNode("//h1[@class='like-h1']").InnerText;
                    string description = doc.DocumentNode.SelectSingleNode("//p").InnerText;

                    HtmlNode contact_info = doc.DocumentNode.SelectSingleNode("//div[@id='contact-list']");

                    string address = ExtractAddress(contact_info, doc);
                    string phone = ExtractPhone(html);
                    string email = ExtractEmail(html);
                    string website = ExtractWebsite(html);

                    string[] company_info = { company_name, description, address, phone, email, website };

                    worksheet.Cells[worksheet.Dimension.End.Row + 1, 1].LoadFromArrays(new List<string[]>() { company_info });

                    Console.WriteLine($"Ссылка: {baseUrl}{link}");
                    Console.WriteLine($"Добавлена компания: {company_name}");
                    Console.WriteLine($"Адрес: {address}");
                    Console.WriteLine($"Телефон: {phone}");
                    Console.WriteLine($"Почта: {email}");
                    Console.WriteLine($"Сайт: {website}");
                }
            }

            FileInfo excelFile = new FileInfo(excelFilePath);
            await package.SaveAsAsync(excelFile);
        }
    }

    static List<string> ReadLinksFromFile(string filePath)
    {
        List<string> links = new List<string>();

        using (StreamReader reader = new StreamReader(filePath))
        {
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                links.Add(line);
            }
        }

        return links;
    }

    static string ExtractAddress(HtmlNode contactInfoNode, HtmlDocument doc)
    {
        HtmlNode addressNode = contactInfoNode.SelectSingleNode("//td[contains(text(), 'Адрес')]/following-sibling::td/p");
        return addressNode != null ? addressNode.InnerText.Trim() : "Адрес не найден";
    }

    static string ExtractPhone(string html)
    {
        string phonePattern = @"tel:([\d\s()+-]+)";
        MatchCollection phoneMatches = Regex.Matches(html, phonePattern);
        List<string> phones = new List<string>();

        foreach (Match match in phoneMatches)
        {
            phones.Add(match.Groups[1].Value);
        }

        return phones.Count > 0 ? string.Join(", ", phones) : "Телефон не найден";
    }

    static string ExtractEmail(string html)
    {
        string emailPattern = @"mailto:([\w\.-]+@[\w\.-]+)";
        Match emailMatch = Regex.Match(html, emailPattern);
        return emailMatch.Success ? emailMatch.Groups[1].Value : "Почта не найдена";
    }

    static string ExtractWebsite(string html)
    {
        string websitePattern = @"<a href=""(https?://[^""]+)""";
        Match websiteMatch = Regex.Match(html, websitePattern);
        return websiteMatch.Success ? websiteMatch.Groups[1].Value : "Сайт не найден";
    }
}
