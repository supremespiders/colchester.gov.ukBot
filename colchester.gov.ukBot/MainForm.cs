using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using colchester.gov.ukBot.Models;
using ExcelHelperExe;
using MetroFramework.Controls;
using MetroFramework.Forms;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace colchester.gov.ukBot
{
    public partial class MainForm : MetroForm
    {
        public bool LogToUi = true;
        public bool LogToFile = true;

        Regex _regex = new Regex("[^a-zA-Z0-9]");
        private readonly string _path = Application.StartupPath;
        private int _maxConcurrency;
        private Dictionary<string, string> _config;
        public HttpCaller HttpCaller = new HttpCaller();
        private ChromeDriver _driver;
        public MainForm()
        {
            InitializeComponent();
        }


        private async Task MainWork()
        {
            SuccessLog("Work started");
            NormalLog($"Starting chrome driver");
            var chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            var options = new ChromeOptions();
            options.AddArgument("--window-position=-32000,-32000");
            _driver = new ChromeDriver(chromeDriverService, options);
            var allLinks = new List<string>();
            for (int i = 0; i < 26; i++)
            {
                Display($"Collecting links from {i + 1} / 26 , links collected : {allLinks.Count}");
                try
                {
                    var links = await GetLinks($"https://www.colchester.gov.uk/wam-received-list/?id={i}");
                    allLinks.AddRange(links);
                }
                catch (Exception e)
                {
                    ErrorLog(e.ToString());
                }
            }
            File.WriteAllLines("links", allLinks);
            allLinks = File.ReadAllLines("links").ToList();
            var items = new List<Item>();
            for (var i = 0; i < allLinks.Count; i++)
            {
                Display($"working on {i + 1} / {allLinks.Count}");
                SetProgress((i + 1) * 100 / allLinks.Count);
                var link = allLinks[i];
                var item = await GetDetails(link);
                items.Add(item);
            }

            File.WriteAllText("items", JsonConvert.SerializeObject(items));
            await items.SaveToExcel(outputI.Text);
            SuccessLog("Work completed");
        }

        private async Task<Item> GetDetails(string href)
        {
            // _driver.Navigate().GoToUrl(href);
            var doc = await HttpCaller.GetDoc(href);
            var nodes = doc.DocumentNode.SelectNodes("//div[@class='control']");
            var textInfo = new CultureInfo("en-US", false).TextInfo;
            var propertyInfo = typeof(Item).GetProperties().ToDictionary(x => x.Name);
            var item = new Item();
            foreach (var node in nodes)
            {
                var info = node.SelectSingleNode("./preceding-sibling::div[@class='info']")?.InnerText.Trim();
                if (info == null) continue;
                var value = node.SelectSingleNode("./input")?.GetAttributeValue("value", "") ?? node.InnerText.Replace("\r\n", "");
                info = _regex.Replace(textInfo.ToTitleCase(info), string.Empty);
                var property = propertyInfo[info];
                property.SetValue(item, value);

                //Console.WriteLine($"public string {info} {{ get; set; }}");
            }

            item.Url = href;

            return item;
            //doc.Save("doc.html");
            //Process.Start("doc.html");
        }

        private async Task<List<string>> GetLinks(string baseLink)
        {
            _driver.Navigate().GoToUrl(baseLink);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            var firstHref = "";
            var links = new HashSet<string>();
            do
            {
                await Task.Delay(10000);
                do
                {
                    var fh = _driver.FindElementByXPath("//a[@class='details-link has-tooltip']").GetAttribute("href");
                    if (!fh.Equals(firstHref))
                    {
                        firstHref = fh;
                        break;
                    }
                    await Task.Delay(1000);
                } while (true);
                var nodes = _driver.FindElementsByXPath("//a[@class='details-link has-tooltip']");
                foreach (var a in nodes)
                {
                    var href = a.GetAttribute("href");
                    links.Add(href);
                }

                if (_driver.FindElementsByXPath("//a[@aria-label='Next page']").Count == 0) break;
                var nextPage = _driver.FindElementByXPath("//a[@aria-label='Next page']");
                if (nextPage.Displayed == false)
                    break;
                var isDisabled = nextPage.FindElement(By.XPath("./..")).GetAttribute("class").Equals("disabled");
                if (isDisabled) break;
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
                if (_driver.FindElementsByXPath("//button[@class='button-small cookie-notice-button']").Count > 0)
                    try
                    {
                        _driver.FindElementByXPath("//button[@class='button-small cookie-notice-button']").Click();
                    }
                    catch (Exception)
                    {//
                    }
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                try
                {
                    nextPage.Click();
                }
                catch (Exception e)
                {
                    Console.WriteLine($"couldn't click on next page {baseLink} " + e.Message);
                    break;
                }
            } while (true);

            return links.ToList();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ServicePointManager.DefaultConnectionLimit = 65000;
            Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            Directory.CreateDirectory("data");
            outputI.Text = _path + @"\output.xlsx";
            LoadConfig();
        }

        void InitControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    try
                    {
                        if (x.Name.EndsWith("I"))
                        {
                            switch (x)
                            {
                                case MetroCheckBox _:
                                case CheckBox _:
                                    ((CheckBox)x).Checked = bool.Parse(_config[((CheckBox)x).Name]);
                                    break;
                                case RadioButton radioButton:
                                    radioButton.Checked = bool.Parse(_config[radioButton.Name]);
                                    break;
                                case TextBox _:
                                case RichTextBox _:
                                case MetroTextBox _:
                                    x.Text = _config[x.Name];
                                    break;
                                case NumericUpDown numericUpDown:
                                    numericUpDown.Value = int.Parse(_config[numericUpDown.Name]);
                                    break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    InitControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public void SaveControls(Control parent)
        {
            try
            {
                foreach (Control x in parent.Controls)
                {
                    #region Add key value to disctionarry

                    if (x.Name.EndsWith("I"))
                    {
                        switch (x)
                        {
                            case MetroCheckBox _:
                            case RadioButton _:
                            case CheckBox _:
                                _config.Add(x.Name, ((CheckBox)x).Checked + "");
                                break;
                            case TextBox _:
                            case RichTextBox _:
                            case MetroTextBox _:
                                _config.Add(x.Name, x.Text);
                                break;
                            case NumericUpDown _:
                                _config.Add(x.Name, ((NumericUpDown)x).Value + "");
                                break;
                            default:
                                Console.WriteLine(@"could not find a type for " + x.Name);
                                break;
                        }
                    }
                    #endregion
                    SaveControls(x);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private void SaveConfig()
        {
            _config = new Dictionary<string, string>();
            SaveControls(this);
            try
            {
                File.WriteAllText("config.txt", JsonConvert.SerializeObject(_config, Formatting.Indented));
            }
            catch (Exception e)
            {
                ErrorLog(e.ToString());
            }
        }
        private void LoadConfig()
        {
            try
            {
                _config = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText("config.txt"));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return;
            }
            InitControls(this);
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            MessageBox.Show(e.Exception.ToString(), @"Unhandled Thread Exception");
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            MessageBox.Show((e.ExceptionObject as Exception)?.ToString(), @"Unhandled UI Exception");
        }
        #region UIFunctions
        public delegate void WriteToLogD(string s, Color c);
        public void WriteToLog(string s, Color c)
        {
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new WriteToLogD(WriteToLog), s, c);
                    return;
                }
                if (LogToUi)
                {
                    if (DebugT.Lines.Length > 5000)
                    {
                        DebugT.Text = "";
                    }
                    DebugT.SelectionStart = DebugT.Text.Length;
                    DebugT.SelectionColor = c;
                    DebugT.AppendText(DateTime.Now.ToString(Utility.SimpleDateFormat) + " : " + s + Environment.NewLine);
                }
                Console.WriteLine(DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s);
                if (LogToFile)
                {
                    File.AppendAllText(_path + "/data/log.txt", DateTime.Now.ToString(Utility.SimpleDateFormat) + @" : " + s + Environment.NewLine);
                }
                Display(s);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        public void NormalLog(string s)
        {
            WriteToLog(s, Color.Black);
        }
        public void ErrorLog(string s)
        {
            WriteToLog(s, Color.Red);
        }
        public void SuccessLog(string s)
        {
            WriteToLog(s, Color.Green);
        }
        public void CommandLog(string s)
        {
            WriteToLog(s, Color.Blue);
        }

        public delegate void SetProgressD(int x);
        public void SetProgress(int x)
        {
            if (InvokeRequired)
            {
                Invoke(new SetProgressD(SetProgress), x);
                return;
            }
            if ((x <= 100))
            {
                ProgressB.Value = x;
            }
        }
        public delegate void DisplayD(string s);
        public void Display(string s)
        {
            if (InvokeRequired)
            {
                Invoke(new DisplayD(Display), s);
                return;
            }
            displayT.Text = s;
        }

        #endregion
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveConfig();
            _driver?.Quit();
        }
        private void openOutputB_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(outputI.Text);
            }
            catch (Exception ex)
            {
                ErrorLog(ex.ToString());
            }
        }
        private void loadOutputB_Click_1(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog
            {
                Filter = @"xlsx file|*.xlsx",
                Title = @"Select the output location"
            };
            saveFileDialog1.ShowDialog();
            if (saveFileDialog1.FileName != "")
            {
                outputI.Text = saveFileDialog1.FileName;
            }
        }

        private async void startB_Click_1(object sender, EventArgs e)
        {
            SaveConfig();
            LogToUi = logToUII.Checked;
            LogToFile = logToFileI.Checked;
            await Task.Run(MainWork);
        }
    }
}
