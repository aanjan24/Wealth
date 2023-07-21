using Microsoft.SqlServer.Dts.Runtime;
//using Microsoft.SqlServer.Dts.Runtime.Wrapper;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.IO.Compression;
using SeleniumWebScraping;

namespace USD_Price_Download
{
    public class SeleniumJobs
    {
        //private static string refLink;
        //private static string srcPath = "C:\\Users\\smishra\\Downloads\\Zip\\";
        //private static string destPath = "C:\\Users\\smishra\\Downloads\\UnZip\\";
        //private static string fileName;
        //private static string fileNameAfterUnZip;
        //private static string FilePath = "C:\\Users\\smishra\\Downloads\\UnZip";
        //private static string PackagePath = "C:\\Users\\smishra\\Desktop\\Console Applications\\BhavCopy_Daily_Upload\\BhavCopy_Daily_Upload\\Package.dtsx";
        //private static string DtsConfigPath = "C:\\Users\\smishra\\Desktop\\Console Applications\\BhavCopy_Daily_Upload\\BhavCopy_Daily_Upload\\PackageConfig.dtsConfig";
        //private static string uploadType = "BhavCopy";


        private static string refLink;
        private static string srcPath = @"C:\Users\Admin\Downloads\Zip\";
        private static string destPath = @"C:\Users\Admin\Downloads\UnZip\";
        private static string fileName;
        private static string fileNameAfterUnZip;
        private static string FilePath = @"C:\Users\Admin\Downloads\UnZip\";
        private static string PackagePath = @"D:\BhavCopy_Daily_Upload\BhavCopy_Daily_Upload\Package.dtsx";
        private static string DtsConfigPath = @"D:\BhavCopy_Daily_Upload\BhavCopy_Daily_Upload\PackageConfig.dtsConfig";
        private static string uploadType = "BhavCopy";


        private static DateTime CurrentDate = DateTime.Now;

        public static void EquityDailyPrice()
        {
            try
            {
                IWebDriver webDriver = (IWebDriver)new FirefoxDriver();
                webDriver.Navigate().GoToUrl("http://www.bseindia.com/markets/equity/EQReports/Equitydebcopy.aspx");
                IWebElement element = webDriver.FindElement(By.Id("btnhylZip"));
                SeleniumJobs.refLink = element.FindElement(By.Id("btnhylZip")).Text;
                MatchCollection matchCollection = new Regex("(\\d+)[/](\\d+)[/](\\d+)").Matches(SeleniumJobs.refLink);
                IFormatProvider provider = (IFormatProvider)new CultureInfo("fr-FR", true);
                if (!(DateTime.Parse(matchCollection[0].ToString(), provider).Date == SeleniumJobs.CurrentDate.Date))
                    return;
                SeleniumJobs.refLink = element.FindElement(By.Id("btnhylZip")).GetAttribute("href").ToString();
                string[] strArray = SeleniumJobs.refLink.Split('/');
                SeleniumJobs.fileName = strArray[strArray.Length - 1].ToString();
                string refLink = SeleniumJobs.refLink;
                string str = Path.Combine(SeleniumJobs.srcPath, SeleniumJobs.fileName);
                WebClient webClient = new WebClient();
                if (System.IO.File.Exists(str))
                    System.IO.File.Delete(str);
                webClient.DownloadFile(SeleniumJobs.refLink, str);
                if (SeleniumJobs.Extract())
                {
                    Console.Write("Extraction Successful");
                    webDriver.Close();
                    SeleniumJobs.FilePath = Path.Combine(SeleniumJobs.FilePath, SeleniumJobs.fileNameAfterUnZip);
                    if (SeleniumJobs.RunPackage(SeleniumJobs.PackagePath, SeleniumJobs.DtsConfigPath, SeleniumJobs.uploadType, SeleniumJobs.FilePath, SeleniumJobs.CurrentDate))
                        System.IO.File.Delete(Path.Combine(SeleniumJobs.destPath, SeleniumJobs.fileNameAfterUnZip));
                }
                else
                    Console.Write("Extraction Was not Successful");
            }
            catch (Exception ex)
            {
                Console.Write(ex.ToString());
            }
        }

        public static void BseIndex()
        {
            IWebDriver webDriver = (IWebDriver)new FirefoxDriver();
            webDriver.Navigate().GoToUrl("http://www.bseindia.com/sensexview/IndexHighlight.aspx?expandable=2");
            if (webDriver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr")).Displayed)
            {
                IWebElement element = webDriver.FindElement(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr"));
                int count1 = element.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr")).Count;
                IList<IWebElement> elements1 = (IList<IWebElement>)element.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr[1]/th"));
                int count2 = elements1.Count;
                DataTable dataTable = new DataTable();
                for (int index = 0; index < count2; ++index)
                    dataTable.Columns.Add(elements1.ElementAt<IWebElement>(index).Text, typeof(string));
                for (int index1 = 2; index1 <= count1; ++index1)
                {
                    DataRow row = dataTable.NewRow();
                    for (int index2 = 1; index2 <= count2; ++index2)
                    {
                        IList<IWebElement> elements2 = (IList<IWebElement>)element.FindElements(By.XPath("//*[@id='ctl00_ContentPlaceHolder1_gvRealTime']/tbody/tr[" + (object)index1 + "]/td[" + (object)index2 + "]"));
                        row[index2 - 1] = (object)elements2.ElementAt<IWebElement>(0).Text;
                    }
                    dataTable.Rows.Add(row);
                }
            }
            webDriver.Close();
        }

        public static bool Extract()
        {

           
            ZipStorer zipStorer = ZipStorer.Open(Path.Combine(SeleniumJobs.srcPath, SeleniumJobs.fileName), FileAccess.Read);
            List<ZipStorer.ZipFileEntry> zipFileEntryList = zipStorer.ReadCentralDir();
            bool flag = false;
            foreach (ZipStorer.ZipFileEntry _zfe in zipFileEntryList)
            {
                string str = Path.Combine(SeleniumJobs.destPath, Path.GetFileName(_zfe.FilenameInZip));
                SeleniumJobs.fileNameAfterUnZip = _zfe.FilenameInZip.ToString();
                if (System.IO.File.Exists(str))
                    System.IO.File.Delete(str);
                flag = zipStorer.ExtractFile(_zfe, str);
            }
            zipStorer.Close();
            return flag;
        }

        public static bool RunPackage(
          string PackagePath,
          string DtsConfigPath,
          string uploadType,
          string FilePath,
          DateTime currentDate)
        {
            bool flag = false;


            bool returnType = false;
            Application app = new Application();
            Package package = null;

            //  Package package = new Application().LoadPackage(PackagePath, null);

            package = app.LoadPackage(PackagePath, null);
            package.Variables[(object)"User::FilePath"].Value = (object)FilePath;
            package.Variables[(object)"User::UploadType"].Value = (object)uploadType;
            package.Variables[(object)"User::BhavCopyDate"].Value = (object)currentDate;
            package.Configurations[(object)0].ConfigurationString = DtsConfigPath;
            DTSExecResult dtsExecResult = package.Execute();
            if (dtsExecResult == DTSExecResult.Failure)
            {
                string str1 = "";
                foreach (DtsError error in package.Errors)
                {
                    string str2 = error.Description.ToString();
                    str1 += str2;
                }
            }
            if (dtsExecResult == DTSExecResult.Success)
                flag = true;
            return flag;
        }

        public static void CreateExcelFile(DataTable dt)
        {
            StreamWriter streamWriter = new StreamWriter("C:\\Users\\smishra\\Downloads\\Indices.xls");
            try
            {
                for (int index = 0; index < dt.Columns.Count; ++index)
                    streamWriter.Write(dt.Columns[index].ToString().ToUpper() + "\t");
                streamWriter.WriteLine();
                for (int index = 0; index < dt.Rows.Count; ++index)
                {
                    for (int columnIndex = 0; columnIndex < dt.Columns.Count; ++columnIndex)
                    {
                        if (dt.Rows[index][columnIndex] != null)
                            streamWriter.Write(Convert.ToString(dt.Rows[index][columnIndex]) + "\t");
                        else
                            streamWriter.Write("\t");
                    }
                    streamWriter.WriteLine();
                }
                streamWriter.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
