using Microsoft.Practices.EnterpriseLibrary.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace USD_Price_Download
{
    class Program
    {

        private static void Main(string[] args)
        {
            try
            {
                SqlParameter[] Params = new SqlParameter[2]
                {
          new SqlParameter("@WJN_JobId", (object) 1001),
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                if (Program.ExecuteDatasetProc("SPROC_WerpJobStatus", Params).Tables[0].Rows.Count != 0)
                    return;
                Program.DollarRate();
            }
            catch (Exception ex)
            {
                SqlParameter[] Params = new SqlParameter[4]
                {
          new SqlParameter("@WJN_JobId", (object) 1001),
          null,
          null,
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                Params[2] = new SqlParameter("@WJM_IsMailSent", (object)int.Parse("0"));
                Params[2].DbType = DbType.Int16;
                Params[3] = new SqlParameter("@WJM_ErrorInMailSent", (object)ex.Message);
                Params[3].DbType = DbType.String;
                Program.ExecuteStoredProc("SPROC_UpdateEmailSentDetails", Params);
            }
        }

        public static void BSEIndexPriceDownLoad()
        {
            Hashtable hashtable1 = new Hashtable();
            IWebDriver webDriver = (IWebDriver)new FirefoxDriver(new FirefoxBinary("C:\\Program Files (x86)\\Mozilla Firefox\\Firefox.exe"), new FirefoxProfile(), TimeSpan.FromMinutes(10.0));
            webDriver.Manage().Window.Maximize();
            webDriver.Navigate().GoToUrl("http://www.bseindia.com/sensexview/IndexHighlight.aspx?expandable=2&type=1");
            Thread.Sleep(5000);
            IList<IWebElement> source = (IList<IWebElement>)null;
            try
            {
                if (webDriver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_divRT']/table")).Displayed)
                    source = (IList<IWebElement>)webDriver.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr[1]/th"));
                DataTable dataTable = new DataTable();
                int num1 = source.Count<IWebElement>();
                for (int index = 0; index < num1; ++index)
                    dataTable.Columns.Add(source.ElementAt<IWebElement>(index).Text, typeof(string));
                for (int index1 = 1; index1 <= 9; ++index1)
                {
                    try
                    {
                        webDriver.Navigate().GoToUrl("http://www.bseindia.com/sensexview/IndexHighlight.aspx?expandable=2&type=" + (object)index1);
                        if (webDriver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_divRT']/table")).Displayed)
                        {
                            int num2 = webDriver.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr")).Count<IWebElement>();
                            for (int index2 = 2; index2 <= num2; ++index2)
                            {
                                DataRow row = dataTable.NewRow();
                                if (webDriver.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr[" + (object)index2 + "]/td")).Count == num1)
                                {
                                    for (int index3 = 1; index3 <= num1; ++index3)
                                    {
                                        IList<IWebElement> elements = (IList<IWebElement>)webDriver.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr[" + (object)index2 + "]/td[" + (object)index3 + "]"));
                                        int num3 = Decimal.TryParse(elements.ElementAt<IWebElement>(0).Text, out Decimal _) ? 0 : (index3 != 1 ? 1 : 0);
                                        row[index3 - 1] = num3 != 0 ? (object)0 : (object)elements.ElementAt<IWebElement>(0).Text;
                                    }
                                }
                                dataTable.Rows.Add(row);
                            }
                        }
                        Thread.Sleep(5000);
                    }
                    catch (NoSuchElementException ex)
                    {
                        hashtable1.Add((object)("http://www.bseindia.com/sensexview/IndexHighlight.aspx?expandable=2&type=" + (object)index1), (object)ex.Message);
                    }
                }
                webDriver.Close();
                try
                {
                }
                catch (SqlException ex)
                {
                    Hashtable hashtable2=null;
                    hashtable2.Add((object)"Issue with upload to SQL DB", (object)ex.Message);
                }
            }
            catch (NoSuchElementException ex)
            {
                hashtable1.Add((object)"http://www.bseindia.com/sensexview/IndexHighlight.aspx?expandable=2&type=1", (object)ex.Message);
            }
            if (hashtable1.Count <= 0)
                return;
            string str = string.Empty;
            foreach (string key in (IEnumerable)hashtable1.Keys)
                str = str + key + ":   :" + hashtable1[(object)key].ToString() + "<br/>";
        }

        public static void DollarRate()
        {
            // IWebDriver webDriver = (IWebDriver)new FirefoxDriver();

            //IWebDriver webDriver = new FirefoxDriver(new FirefoxBinary("C:\Program Files\Mozilla Firefox\firefox.exe"), new FirefoxProfile(), TimeSpan.FromMinutes(10));
            //IWebDriver webDriver = new ChromeDriver(@"E:\Win-Werp_Upgraded-Jobs-2021\Chromedriver");
            IWebDriver webDriver = new ChromeDriver(@"C:\Users\aanja\Desktop\New folder (4)\NSEEquityIndexPriceDownload\NSEINXD-Chrome_Driver");


            webDriver.Manage().Window.Maximize();
            Hashtable hashtable = new Hashtable();
            Decimal result = 0M;
            webDriver.Navigate().GoToUrl("https://www.oanda.com/currency-converter/en/?from=USD&to=INR&amount=1");
            Thread.Sleep(10000);
            try
            {
                if (webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete")).Displayed && webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete")).Displayed && (webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[1]/div[2]/div[1]/div/input")).Displayed && webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[3]/div[2]/div[1]/div/input")).Displayed) && webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[5]/div[1]/div[2]/div/div/input")).Displayed)

                {
                //    webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete")).Click();
                //    webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete")).Clear();
                //    char ch1;
                //    foreach (char ch2 in "US Dollar")
                //    {
                //        IWebElement element = webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete"));
                //        ch1 = ch2;
                //        string text = ch1.ToString();
                //        element.SendKeys(text);
                //        Thread.Sleep(700);
                //    }
                //    webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete")).SendKeys(Keys.Tab);
                //    webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete")).Click();
                //    webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete")).Clear();
                //    foreach (char ch2 in "Indian Rupee")
                //    {
                //        IWebElement element = webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete"));
                //        ch1 = ch2;
                //        string text = ch1.ToString();
                //        element.SendKeys(text);
                //         Thread.Sleep(700);
                //    }
                //    webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete")).SendKeys(Keys.Tab);
                    //webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[1]/div[2]/div[1]/div/input")).Clear();
                    //webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[1]/div[2]/div[1]/div/input")).SendKeys("1");
                    //webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[1]/div[2]/div[1]/div/input")).SendKeys(Keys.Tab);
                    //Thread.Sleep(2000);
                    if (webDriver.FindElement(By.Id("baseCurrency_currency_autocomplete")).GetAttribute("value") == "US Dollar" && webDriver.FindElement(By.Id("quoteCurrency_currency_autocomplete")).GetAttribute("value") == "Indian Rupee")
                    {
                        string attribute1 = webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[2]/div[3]/div[2]/div[1]/div/input")).GetAttribute("value");
                        string attribute2 = webDriver.FindElement(By.XPath("//*[@id='cc-main-conversion-block']/div/div[5]/div[1]/div[2]/div/div/input")).GetAttribute("value");
                        if (Decimal.TryParse(attribute1, out result) && DateTime.TryParse(attribute2, out DateTime _))
                        {
                            try
                            {
                                SqlParameter[] Params = new SqlParameter[2]
                                {
                  new SqlParameter("@WCDV_Value", (object) Decimal.Parse(attribute1)),
                  null
                                };
                                Params[0].DbType = DbType.Decimal;
                                Params[1] = new SqlParameter("@WCDV_Date", (object)DateTime.Now.Date);
                                Params[1].DbType = DbType.Date;
                                Program.USAStockPriceInsert("SPROC_USDPriceInsert", Params);
                            }
                            catch (SqlException ex)
                            {
                                hashtable.Add((object)"Issue while uploading USD price to SQL DB", (object)ex.Message);
                            }
                        }
                        else
                            hashtable.Add((object)"Please Check Arrow 4 and 5", (object)"Possible reasons:<br/>1.The Id of the element is not same.<br/>2.The Value is not readable.<br/>3.The element is not visible.");
                    }
                    else
                        hashtable.Add((object)"Please Check Arrow 1 and 2", (object)"Please make sure that those fields are still accepting the input shown in the image.");
                }
            }
            catch (NoSuchElementException ex)
            {
                hashtable.Add((object)"Please Check the website ", (object)"Possible reasons:<br/>1.The website design is not same.<br/>2.The internet connection was too slow(Try restarting the job again).");
                webDriver.Close();
            }
            if (hashtable.Count > 0)
            {
                string str = string.Empty;
                foreach (string key in (IEnumerable)hashtable.Keys)
                    str = str + key + ":   :" + hashtable[(object)key].ToString() + "<br/>";
                SqlParameter[] Params = new SqlParameter[4]
                {
          new SqlParameter("@WJN_JobId", (object) 1001),
          null,
          null,
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                Params[2] = new SqlParameter("@WJM_ErrorMsg", (object)str);
                Params[2].DbType = DbType.String;
                Params[3] = new SqlParameter("@WJM_IsSuccess", (object)int.Parse("0"));
                Params[3].DbType = DbType.Int16;
                Program.ExecuteStoredProc("SPROC_WerpJobStatusProcessingState", Params);
            }
            else
            {
                SqlParameter[] Params = new SqlParameter[4]
                {
          new SqlParameter("@WJN_JobId", (object) 1001),
          null,
          null,
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                Params[2] = new SqlParameter("@WJM_ErrorMsg", (object)string.Empty);
                Params[2].DbType = DbType.String;
                Params[3] = new SqlParameter("@WJM_IsSuccess", (object)1);
                Params[3].DbType = DbType.Int16;
                Program.ExecuteStoredProc("SPROC_WerpJobStatusProcessingState", Params);
            }
            webDriver.Close();
        }

        public static void USAStockPriceDownload()
        {
            IWebDriver webDriver = (IWebDriver)new FirefoxDriver();
            Hashtable hashtable1 = new Hashtable();
            Hashtable hashtable2 = new Hashtable();
            hashtable2.Add((object)"CTSH", (object)"16834");
            hashtable2.Add((object)"IBN", (object)"16833");
            hashtable2.Add((object)"DEO", (object)"16835");
            Regex regex = new Regex("(\\d+)[/](\\d+)[/](\\d+)");
            foreach (string key in (IEnumerable)hashtable2.Keys)
            {
                try
                {
                    webDriver.Navigate().GoToUrl("http://www.bloomberg.com/quote/" + key + ":US");
                    Thread.Sleep(5000);
                    if (webDriver.FindElement(By.ClassName("price")).Displayed && webDriver.FindElement(By.ClassName("price-datetime")).Displayed)
                    {
                        string text = webDriver.FindElement(By.ClassName("price")).Text;
                        MatchCollection matchCollection = regex.Matches(webDriver.FindElement(By.ClassName("price-datetime")).Text);
                        if (matchCollection.Count > 0)
                        {
                            string s = matchCollection[0].ToString();
                            try
                            {
                                SqlParameter[] Params = new SqlParameter[4]
                                {
                  new SqlParameter("@PEM_ScripCode", (object) int.Parse(hashtable2[(object) key].ToString())),
                  null,
                  null,
                  null
                                };
                                Params[0].DbType = DbType.Int32;
                                Params[1] = new SqlParameter("@PESPH_ClosePrice", (object)Decimal.Parse(text));
                                Params[1].DbType = DbType.Decimal;
                                Params[2] = new SqlParameter("@PESPH_PreviousClose", SqlDbType.BigInt);
                                Params[2].DbType = DbType.Decimal;
                                Params[3] = new SqlParameter("@PESPH_Date", (object)DateTime.Parse(s));
                                Params[3].DbType = DbType.Date;
                                if (Program.USAStockPriceInsert("SPROC_USAStockEquityPriceInsert", Params) == 0)
                                    hashtable1.Add((object)("Issue while uploading Stock Price for the code " + key + " to SQL DB"), (object)("USD Price for" + s + " not available"));
                            }
                            catch (SqlException ex)
                            {
                                hashtable1.Add((object)("Issue while uploading Stock Price for the code " + key + " to SQL DB"), (object)ex.Message);
                            }
                        }
                        else
                            hashtable1.Add((object)"Please Check Closing Date Arrow", (object)"Possible reasons:<br/>1.The Id of the element is not same.<br/>2.The Value is not readable.<br/>3.The element is not visible.");
                    }
                }
                catch (NoSuchElementException ex)
                {
                    hashtable1.Add((object)("Please Check the website for stock code" + key + "  "), (object)"Possible reasons:<br/>1.The website design is not same.<br/>2.The internet connection was too slow(Try restarting the job again).");
                }
                catch (Exception ex)
                {
                    hashtable1.Add((object)("Something Unexcepted happend while downloading the stock price for" + key), (object)"Possible reasons:<br/>1.The website design is not same.<br/>2.The internet connection was too slow(Try restarting the job again).");
                }
            }
            if (hashtable1.Count > 0)
            {
                string str = string.Empty;
                foreach (string key in (IEnumerable)hashtable1.Keys)
                    str = str + key + ": :" + hashtable1[(object)key].ToString() + "<br/>";
                SqlParameter[] Params = new SqlParameter[4]
                {
          new SqlParameter("@WJN_JobId", (object) 1002),
          null,
          null,
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                Params[2] = new SqlParameter("@WJM_ErrorMsg", (object)str);
                Params[2].DbType = DbType.String;
                Params[3] = new SqlParameter("@WJM_IsSuccess", (object)int.Parse("0"));
                Params[3].DbType = DbType.Int16;
                Program.ExecuteStoredProc("SPROC_WerpJobStatusProcessingState", Params);
            }
            else
            {
                SqlParameter[] Params = new SqlParameter[4]
                {
          new SqlParameter("@WJN_JobId", (object) 1002),
          null,
          null,
          null
                };
                Params[0].DbType = DbType.Int32;
                Params[1] = new SqlParameter("@WJM_Date", (object)DateTime.Now);
                Params[1].DbType = DbType.Date;
                Params[2] = new SqlParameter("@WJM_ErrorMsg", (object)string.Empty);
                Params[2].DbType = DbType.String;
                Params[3] = new SqlParameter("@WJM_IsSuccess", (object)1);
                Params[3].DbType = DbType.Int16;
                Program.ExecuteStoredProc("SPROC_WerpJobStatusProcessingState", Params);
            }
            webDriver.Close();
        }

        public static void Test()
        {
            IWebDriver webDriver = (IWebDriver)new FirefoxDriver();
            try
            {
                webDriver.Navigate().GoToUrl("https://app.wealtherp.net");
                IWebElement element = webDriver.FindElement(By.Id("mainframe"));
                webDriver.SwitchTo().Frame(element);
                webDriver.FindElement(By.Id("ctrl_Userlogin_txtLoginId")).SendKeys("ssahu");
                webDriver.FindElement(By.Id("ctrl_Userlogin_txtPassword")).SendKeys("samir@3010");
                webDriver.FindElement(By.Id("ctrl_Userlogin_btnSignIn")).Click();
            }
            catch (Exception ex)
            {
            }
        }

        public static void UploadToDataBaseTable(DataTable dt)
        {
            Database database = DatabaseFactory.CreateDatabase("wealtherp");
            DbCommand storedProcCommand = database.GetStoredProcCommand("SPROC_BSEIndexPriceInsert");
            SqlParameter sqlParameter = new SqlParameter();
            sqlParameter.ParameterName = "@ProductEquityIndexHistory";
            sqlParameter.SqlDbType = SqlDbType.Structured;
            sqlParameter.Value = (object)dt;
            storedProcCommand.Parameters.Add((object)sqlParameter);
            database.ExecuteNonQuery(storedProcCommand);
        }

        public static int USAStockPriceInsert(string CommandText, SqlParameter[] Params)
        {
            Database database = DatabaseFactory.CreateDatabase("wealtherp");
            DbCommand storedProcCommand = database.GetStoredProcCommand(CommandText);
            foreach (SqlParameter sqlParameter in Params)
                database.AddInParameter(storedProcCommand, sqlParameter.ParameterName, sqlParameter.DbType, sqlParameter.Value);
            database.AddOutParameter(storedProcCommand, "@returnValue", DbType.Int16, 5000);
            database.ExecuteNonQuery(storedProcCommand);
            return int.Parse(storedProcCommand.Parameters["@returnValue"].Value.ToString());
        }

        public static DataSet ExecuteDatasetProc(string CommandText, SqlParameter[] Params)
        {
            Database database = DatabaseFactory.CreateDatabase("wealtherp");
            DbCommand storedProcCommand = database.GetStoredProcCommand(CommandText);
            foreach (SqlParameter sqlParameter in Params)
                database.AddInParameter(storedProcCommand, sqlParameter.ParameterName, sqlParameter.DbType, sqlParameter.Value);
            return database.ExecuteDataSet(storedProcCommand);
        }

        public static void ExecuteStoredProc(string CommandText, SqlParameter[] Params)
        {
            Database database = DatabaseFactory.CreateDatabase("wealtherp");
            DbCommand storedProcCommand = database.GetStoredProcCommand(CommandText);
            foreach (SqlParameter sqlParameter in Params)
                database.AddInParameter(storedProcCommand, sqlParameter.ParameterName, sqlParameter.DbType, sqlParameter.Value);
            database.ExecuteNonQuery(storedProcCommand);
        }
    }
}
