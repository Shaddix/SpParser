using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using HtmlAgilityPack;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.PhantomJS;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace SpParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private const int PageCount = 422;
        //private const string PageUrl = "http://sp.tomica.ru/forum/phpBB3/viewtopic.php?keys_=243&f=244&t=431244"; //Dec
        //private const string PageUrl = "http://sp.tomica.ru/forum/phpBB3/viewtopic.php?keys_=243&f=244&t=436615"; //jan
        private const string PageUrl = "http://sp.tomica.ru/forum/phpBB3/viewtopic.php?keys_=135&f=244&t=441821"; //feb

        private async void Form1_Load(object sender, EventArgs e)
        {
            await DoWork();
            WriteToExcel();
        }

        private void WriteToExcel()
        {
            var result = new List<LineData>();
            for (int i = 0; i < PageCount; i++)
            {
                var data = JsonConvert.DeserializeObject<List<LineData>>(File.ReadAllText(i + ".txt"));
                result.AddRange(data);
            }

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");
            var line = 1;
            foreach (LineData lineData in result)
            {
                ws.Cell(line, 1).Value = lineData.Category;
                ws.Cell(line, 2).Value = lineData.SubCategory;
                ws.Cell(line, 3).Value = lineData.Title;
                ws.Cell(line, 4).Value = lineData.User;
                ws.Cell(line, 5).Value = lineData.ItemTitle;
                ws.Cell(line, 6).Value = lineData.ItemLink;
                ws.Cell(line, 7).Value = lineData.Parameter;
                ws.Cell(line, 8).Value = lineData.Price;
                ws.Cell(line, 9).Value = lineData.State;
                ws.Cell(line, 10).Value = lineData.Order;
                line++;
            }
            try
            {
                wb.SaveAs("HelloWorld_feb.xlsx");
            }
            catch (Exception)
            {

                throw;
            }

        }

        private async Task DoWork()
        {
            try
            {

                var driver = GetDriver();

                driver.Navigate().GoToUrl("http://sp.tomica.ru/forum/phpBB3/ucp.php?b_=320&mode=login");
                var loginPage = new LoginPage(driver);
                loginPage.EnterLogin("linsp", "123qweASD");

                var cookies = driver.Manage().Cookies.AllCookies;
                var link = PageUrl;

                var handler = new HttpClientHandler();
                foreach (var cookie in cookies)
                {
                    handler.CookieContainer.Add(new System.Net.Cookie(cookie.Name, cookie.Value, cookie.Path, cookie.Domain));
                }
//driver.Manage().
                driver.Quit();

                //handler.CookieContainer.Add(new System.Net.Cookie("phpbb3_oaauj_u", "179106")
                //{
                //    Domain = "sp.tomica.ru",
                //});
                //handler.CookieContainer.Add(new System.Net.Cookie("phpbb3_oaauj_sid", "af175b8af7bd69f9f95a70ab2f8b68e7")
                //{
                //    Domain = "sp.tomica.ru",
                //});

                var client = new HttpClient(handler);
                client.DefaultRequestHeaders.Add("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36");


                for (int i = 0; i < PageCount; i++)
                {
                    Console.WriteLine("Page " + i);
                    var result = await ParseListOfZakupka(client, PageUrl, i * 10);
                    var serialized = JsonConvert.SerializeObject(result);
                    File.WriteAllText(i + ".txt", serialized);
                }

            }
            catch (Exception)
            {
                MessageBox.Show("asd"); //check that UserAgent is the same as in Chrome!
            }
            finally
            {
                WriteLog();
            }

        }

        private void WriteLog()
        {
            File.WriteAllText("zakupkaWithNoTable" + Guid.NewGuid().ToString() + ".txt", JsonConvert.SerializeObject(ZakupkaWithNoTable));
            File.WriteAllText("PostWithNoLink" + Guid.NewGuid().ToString() + ".txt", JsonConvert.SerializeObject(PostWithNoLink));
        }

        private async Task<List<LineData>> ParseListOfZakupka(HttpClient client, string link, int start)
        {
            var pageContent = await (await client.GetAsync(link + "&start=" + start)).Content.ReadAsStringAsync();

            var doc = new HtmlDocument();
            doc.LoadHtml(pageContent);

            var result = new List<LineData>();
            var posts = doc.DocumentNode.SelectNodes("//div[@id='page-body']/div[contains(@class,'post')]");
            foreach (var post in posts)
            {

                var linkToZakupka = post.SelectSingleNode(".//div[@class='content']//a");
                if (linkToZakupka != null && (linkToZakupka.GetAttributeValue("href", "") != "#"))
                {
                    var href = linkToZakupka.GetAttributeValue("href", "").Replace("hhttp", "http");
                    href = href.Replace("&amp;", "&");
                    try
                    {
                        var data = await OpenZakupkaAndParseTable(client, href);
                        result.AddRange(data);
                    }
                    catch (Exception)
                    {
                        PostWithNoLink.Add(post.InnerHtml);
                        continue;
                    }
                }
                else
                {
                    PostWithNoLink.Add(post.InnerHtml);
                }
            }
            return result;
        }

        public class LineData
        {
            public string Category { get; set; }
            public string SubCategory { get; set; }
            public string Title { get; set; }

            public string User { get; set; }
            public string ItemTitle { get; set; }
            public string ItemLink { get; set; }
            public decimal Price { get; set; }
            public string Parameter { get; set; }
            public string State { get; set; }
            public string Order { get; set; }

        }

        public List<string> ZakupkaWithNoTable = new List<string>();
        public List<string> PostWithNoLink = new List<string>();
        private async Task<List<LineData>> OpenZakupkaAndParseTable(HttpClient client, string link)
        {
            var result = new List<LineData>();
            var pageContent = await (await client.GetAsync(link)).Content.ReadAsStringAsync();

            var doc = new HtmlDocument();
            doc.LoadHtml(pageContent);

            var tableComment = doc.DocumentNode.SelectSingleNode("//comment()[contains(., 'Отчет по закупке')]");
            if (tableComment == null)
            {
                ZakupkaWithNoTable.Add(link);
                return result;
            }

            var categoryBlock = doc.DocumentNode.SelectNodes("//li[@class='icon-home']/a").Skip(1).ToList();
            var mainCategory = categoryBlock[0].InnerText;
            var subCategory = categoryBlock[1].InnerText;
            var title = doc.DocumentNode.SelectSingleNode("//div[@id='page-body']//h2").InnerText;

            var lines = tableComment.ParentNode.SelectNodes(".//table//tr");
            foreach (HtmlNode line in lines.Skip(1))
            {
                try
                {
                    var lineData = new LineData()
                    {
                        Category = mainCategory,
                        SubCategory = subCategory,
                        Title = title,

                        User = line.ChildNodes[0].InnerText,
                        ItemTitle = line.ChildNodes[1].InnerText,
                        ItemLink = line.ChildNodes[1].SelectSingleNode("a")?.GetAttributeValue("href", ""),
                        Price = decimal.Parse(line.ChildNodes[2].InnerText),
                        Parameter = line.ChildNodes[3].InnerText,
                        State = line.ChildNodes[4].InnerText,
                        Order = line.ChildNodes[5].InnerText,
                    };
                    result.Add(lineData);
                }
                catch (Exception)
                {

                    throw;
                }

            }
            return result;
        }

        private IWebDriver _driver;
        private IWebDriver GetDriver()
        {
            if (_driver != null)
                return _driver;
            _driver = new ChromeDriver();
            {

            };

            _driver.Manage().Window.Size = new Size(1360, 768);
            return _driver;
        }
    }
}
