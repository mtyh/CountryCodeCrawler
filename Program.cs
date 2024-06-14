// See https://aka.ms/new-console-template for more information

using HtmlAgilityPack;
using OfficeOpenXml;
using System.Reflection;

#region 实体
public class Country
{
    /// <summary>
    /// ISO 二字代码
    /// </summary>
    public string ISO2Code { get; set; }
    /// <summary>
    /// ISO 三字代码
    /// </summary>
    public string ISO3Code { get; set; }
    /// <summary>
    /// ISO 数字代码
    /// </summary>
    public string ISONumericCode { get; set; }
    /// <summary>
    /// 国家/地区
    /// </summary>
    public string CountryOrRegion { get; set; }
    /// <summary>
    /// 国家/地区(英文)
    /// </summary>
    public string CountryOrRegionEn { get; set; }
    /// <summary>
    /// 首都/省会
    /// </summary>
    public string CapitalOrState { get; set; }
    /// <summary>
    /// 面积 (km²)
    /// </summary>
    public string Area_km2 { get; set; }
    /// <summary>
    /// 人口
    /// </summary>
    public string Population { get; set; }
    /// <summary>
    /// 洲
    /// </summary>
    public string ContinentISO2Code { get; set; }
}
#endregion

namespace CountryCodeCrawler
{
    public static class Program
    {
        static async Task Main()
        {
            HttpClient httpClient = new HttpClient();
            var html = await httpClient.GetStringAsync("https://www.nowmsg.com/iso/country_code.asp");
            if (!string.IsNullOrEmpty(html))
            {
                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);
                var nodes = htmlDocument.DocumentNode.SelectNodes(".//table");
                var tableNode = nodes[1];
                List<Country> countries = new List<Country>();
                var trNodes = tableNode.SelectNodes("./tbody/tr");
                foreach (var trNode in trNodes)
                {
                    var tdNodes = trNode.SelectNodes("./td");
                    var country = new Country
                    {
                        ISO2Code = tdNodes[0].InnerHtml,
                        ISO3Code = tdNodes[1].InnerHtml,
                        ISONumericCode = tdNodes[2].InnerHtml,
                        CapitalOrState = tdNodes[4].InnerHtml,
                        Area_km2 = tdNodes[5].InnerHtml,
                        Population = tdNodes[6].InnerHtml,
                        ContinentISO2Code = tdNodes[7].InnerHtml
                    };
                    var countryOrRegionStr = tdNodes[3].InnerHtml;
                    if (countryOrRegionStr.Contains("<br>"))
                    {
                        var countryOrRegions = tdNodes[3].InnerHtml.Split("<br>");
                        country.CountryOrRegionEn = countryOrRegions[0];
                        country.CountryOrRegion = countryOrRegions[1];
                        
                    }
                    else
                    {
                        country.CountryOrRegion = countryOrRegionStr;
                        country.CountryOrRegionEn = countryOrRegionStr;
                    }
                    countries.Add(country);
                }

                // 导出List到Excel
                ExportToExcel(countries, "output.xlsx");

                Console.WriteLine("抓取完成，按任意键结束");
                Console.Read();
            }
        }

        static void ExportToExcel<T>(List<T> dataList, string filePath)
        {
            // 设置许可上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                // 添加工作表
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                if (dataList == null || !dataList.Any())
                {
                    throw new ArgumentException("The dataList cannot be null or empty.");
                }

                // 获取类型和属性
                Type type = typeof(T);
                PropertyInfo[] properties = type.GetProperties();

                // 添加表头
                for (int i = 0; i < properties.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = properties[i].Name;
                }

                // 添加数据
                for (int i = 0; i < dataList.Count; i++)
                {
                    for (int j = 0; j < properties.Length; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = properties[j].GetValue(dataList[i]);
                    }
                }

                // 保存Excel文件
                File.WriteAllBytes(filePath, package.GetAsByteArray());

                Console.WriteLine($"Excel文件已保存到 {filePath}");
            }
        }
    }
    
}






