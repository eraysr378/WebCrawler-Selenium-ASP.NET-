using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using WebApp1.Models;
using WebApp1.ResponseMangament;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using HtmlAgilityPack;
using NUnit.Framework;
using OpenQA.Selenium.Edge;
using System.Threading;
using System.Collections.ObjectModel;
using System.Dynamic;
using ClosedXML;
using ClosedXML.Excel;
using System.Data;
using System.Reflection;
using System.Text.Json;
using Newtonsoft.Json;

namespace WebApp1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        public static List<SearchedItem> SearchedItems = new List<SearchedItem>();
        public static List<SearchedItem> ReadItems = new List<SearchedItem>();
        public static List<PriceModel> PriceModels = new List<PriceModel>();
        public static List<ExpandoModel> ExpandoModels = new List<ExpandoModel>();
        public static List<List<SearchedItem>> SavedItemsList = new List<List<SearchedItem>>();
        public static string flowName = "Undefined flow";
        public static List<Flow> FlowList = new List<Flow>();
        public static BigViewModel bigViewModel = new BigViewModel();
        static string connectionString = "Data Source=(localdb)\\MSSQLLocalDB;Initial Catalog=\"Test Database\";Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            FlowList = new List<Flow>();
            SearchedItems = new List<SearchedItem>();
            ExpandoModels = new List<ExpandoModel>();
            return View();
        }
        public IActionResult ItemSelection()
        {
            return View();
        }
        [HttpPost]
        public IActionResult ItemSelection(Flow flow)
        {
            flowName = flow.Name;
            return View();
        }
        public IActionResult CreateNewFlowList()
        {
            return View(FlowList);
        }
        // Get the flow name from the user
        public IActionResult GetFlowName()
        {
            return View();
        }
        public IActionResult SelectFromFlowList(int id)
        {
            SearchedItems = FlowList.ElementAt(id).ItemList;
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult SelectFromSaved(int id)
        {
            SearchedItems = SavedItemsList.ElementAt(id);
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult SaveSelectedItems()
        {
            SavedItemsList.Add(SearchedItems);
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult ListSavedItems()
        {
            return View(SavedItemsList);
        }
        public IActionResult GetExcelDataInput()
        {
            ExcelDataModel excelData = new ExcelDataModel();
            return View(excelData);
        }

        public IActionResult AddAttributeName()
        {
            return View();
        }
        public IActionResult FlowIsDone()
        {
            FlowList.Add(new Flow
            {
                ItemList = SearchedItems,
                priceModels = PriceModels,
                Name = flowName
            });
            SearchedItems = new List<SearchedItem>();
            PriceModels = new List<PriceModel>();
            flowName = "undefined flow";
            return View("CreateNewFlowList", FlowList);
        }
        public IActionResult CreateGivenAttribute(PriceModel priceModel)
        {
            PriceModels.Add(priceModel);
            return View("SelectRead");
        }
        public IActionResult SelectGoToPage()
        {
            return View();
        }
        public IActionResult ShowAttributeNames()
        {
            return View(PriceModels);
        }
        public IActionResult SelectClick()
        {
            return View();
        }
        public IActionResult ShowFlowLists()
        {
            return View();
        }
        public IActionResult PrintReadItems(BigViewModel model)
        {
            return View(model);
        }
        public IActionResult SortBy(BigViewModel model)
        {
            bool WillBeSorted = false;
            foreach (var item in bigViewModel.ExpandoModels)
            {
                foreach (var obj in item.Expando)
                {
                    if (obj.Key.Equals(model.SortName))
                    {
                        WillBeSorted = true;
                    }
                }
            }
            if (WillBeSorted)
            {
                // selection sort
                for (int i = 0; i < bigViewModel.ExpandoModels.Count - 1; i++)
                {
                    int min_index = i;
                    for (int j = i + 1; j < bigViewModel.ExpandoModels.Count; j++)
                    {
                        if (((string)((IDictionary<String, Object>)(bigViewModel.ExpandoModels.ElementAt(j).Expando))[model.SortName]).CompareTo
                            ((string)((IDictionary<String, Object>)(bigViewModel.ExpandoModels.ElementAt(i).Expando))[model.SortName]) < 0)
                        {
                            min_index = j;
                        }

                    }
                    ExpandoObject temp = bigViewModel.ExpandoModels.ElementAt(min_index).Expando;
                    bigViewModel.ExpandoModels.ElementAt(min_index).Expando = bigViewModel.ExpandoModels.ElementAt(i).Expando;
                    bigViewModel.ExpandoModels.ElementAt(i).Expando = temp;
                }
                model.ExpandoModels = bigViewModel.ExpandoModels;
                model.SortName = "";
            }
            else
            {
                model.ExpandoModels = new List<ExpandoModel>();
            }
            return View("PrintReadItems", model);
        }
        public IActionResult SaveFlowListJSON()
        {
            string fileName = "Products.json";
            string jString = JsonConvert.SerializeObject(FlowList, Formatting.Indented);
            System.IO.File.WriteAllText(fileName, jString);
            return View("CreateNewFlowList", FlowList);
        }
        public IActionResult ReadFlowListJSON()
        {
            string jString = System.IO.File.ReadAllText("Products.json");
            //string jString = System.IO.File.ReadAllText("Youtube.json");
            FlowList = JsonConvert.DeserializeObject<List<Flow>>(jString);
            return View("CreateNewFlowList", FlowList);
        }
        public IActionResult SelectSetInputVal()
        {
            return View();
        }
        public IActionResult ExportToExcel()
        {
            DataTable dt = ToDataTable(bigViewModel.ExpandoModels);
            ToExcelFile(dt, @"D:\ExcelSaved\ProductList.xlsx");

            return View("PrintReadItems", bigViewModel);
        }
        public IActionResult SelectRead()
        {
            return View();
        }
        public IActionResult SelectedItemsList(SearchedItem searchedItem)
        {
            return View(SearchedItems);
        }
        public IActionResult ReadElementsList(SearchedItem searchedItem)
        {
            return View(ReadItems);
        }
        public IActionResult AddEvent(SearchedItem searchedItem)
        {
            if (!(searchedItem.EventType == EventTypes.GoToPage && (searchedItem.Text == null || searchedItem.Text == "")))
            {
                if (searchedItem.EventType == EventTypes.ReadInnerElements)
                {
                    ReadItems.Add(searchedItem);
                }

                SearchedItems.Add(searchedItem);
            }
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult DeleteSearched(int id)
        {
            foreach (var item in SearchedItems)
            {
                if (item.ItemId == id)
                {
                    SearchedItems.Remove(SearchedItems.Single(s => s.ItemId == item.ItemId));
                    if (item.EventType == EventTypes.ReadInnerElements)
                    {
                        return DeleteRead(item.ItemId);
                    }
                    break;
                }
            }
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult DeleteRead(int id)
        {
            foreach (var item in ReadItems)
            {
                if (item.ItemId == id)
                {
                    SearchedItems.Remove(SearchedItems.Single(s => s.ItemId == item.ItemId));
                    ReadItems.Remove(ReadItems.Single(s => s.ItemId == item.ItemId));
                    break;
                }
            }
            return View("SelectedItemsList", SearchedItems);
        }
        public IActionResult Done(ExcelDataModel excelData)
        {
            DataTable dt = ReadExcel(excelData.FilePath);
            List<string> columns = new List<string>();
            // store all the names in a list from the given column
            List<string> excelProductNames = ConvertDataTable(dt, excelData.Column, columns);
            EdgeOptions edgeOptions = new EdgeOptions() { UseWebView = false };
            /*******************/
            // data table in which found data will be stored
            DataTable newDt = new DataTable();
            newDt.Clear();
            foreach (var item in columns)
            {
                newDt.Columns.Add(item);
            }
            newDt.Columns.Add("Flow");

            /****************/
            //  edgeOptions.AddArgument("headless");
            //  edgeOptions.AddArgument("disable-gpu");


            using (IWebDriver driver = new EdgeDriver(AppDomain.CurrentDomain.BaseDirectory, edgeOptions))
            {
                
                driver.Manage().Window.Maximize();
                //string lastReplacedUrl = "@productName";
                List<string> lastReplacedUrls = new List<string>();
                for (int i = 0; i < FlowList.Count; i++)
                {
                    lastReplacedUrls.Add("@productName");
                }
                foreach (var productName in excelProductNames)
                {
                    for (int j = 0; j < FlowList.Count; j++)
                    {
                        dynamic expando = new ExpandoObject();
                        var currentFlow = FlowList.ElementAt(j);
                        for (int i = 0; i < currentFlow.ItemList.Count; i++)
                        {
                            List<PriceModel> priceModels = currentFlow.priceModels;
                            SearchedItem searchedItem = currentFlow.ItemList[i];
                            if (searchedItem.ReadItemsId != null)
                            {
                                foreach (var itm in ReadItems)
                                {
                                    if (itm.ItemId.ToString() == searchedItem.ReadItemsId)
                                    {
                                        searchedItem.Text = itm.Text;
                                        break;
                                    }
                                }
                            }
                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
                            IWebElement item = null;
                            switch (searchedItem.EventType)
                            {
                                case EventTypes.GoToPage:
                                    try
                                    {
                                        searchedItem.Text = searchedItem.Text.Replace(lastReplacedUrls.ElementAt(j), productName);
                                        lastReplacedUrls[j] = productName;
                                        driver.Url = searchedItem.Text;
                                        driver.Navigate();
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }
                                    break;
                                case EventTypes.Click:
                                    IWebElement btnItem = null;
                                    try
                                    {
                                        btnItem = driver.FindElement(By.XPath(searchedItem.XPath));
                                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);

                                        btnItem.Click();
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }


                                    break;
                                case EventTypes.SetInputValue:
                                    IWebElement inputElement = null;
                                    try
                                    {
                                        inputElement = driver.FindElement(By.XPath(searchedItem.XPath));
                                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
                                        inputElement.Clear();
                                        inputElement.SendKeys(searchedItem.Text);
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }

                                    break;
                                case EventTypes.ReadInnerElements:
                                    IWebElement readItem = null;
                                    ReadOnlyCollection<IWebElement> readItems = null;
                                    try
                                    {
                                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(5000);
                                        readItem = driver.FindElement(By.XPath(searchedItem.XPath));
                                        expando = new ExpandoObject();
                                        foreach (var itm in priceModels)
                                        {
                                            // bunların xpathini verirken başına . koymak lazım mesela .//div[@class="s"]
                                            string str = readItem.FindElement(By.XPath(itm.XPath)).GetAttribute("innerHTML");
                                            AddProperty(expando, itm.Name, str);
                                        }
                                        List<string> values = new List<string>();
                                        List<string> keys = new List<string>();
                                        foreach (var itm in (ExpandoObject)expando)
                                        {
                                            values.Add(itm.Value.ToString());
                                            keys.Add(itm.Key.ToString());
                                        }
                                        DataRow row = newDt.NewRow();
                                        row[0] = productName;
                                        // k < newDt.Columns.Count-1 because flow added at the end of the columns
                                        for (int k = 1; k < newDt.Columns.Count-1; k++)
                                        {
                                            if (keys.Count >= k &&  columns[k] == keys[k - 1].ToString())
                                            {
                                                row[k] = values[k - 1].ToString();
                                            }
                                            else
                                            {
                                                row[k] = "NOT FOUND!";
                                            }
                                        }
                                        row[newDt.Columns.Count-1] = currentFlow.Name;
                                        newDt.Rows.Add(row);
                                        bigViewModel.ExpandoModels.Add(new ExpandoModel
                                        {
                                            Expando = expando
                                        });
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }

                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
                Thread.Sleep(5000);
                driver.Quit();
            }
            // hardcoded filename for now
            ToExcelFile(newDt, @"D:\ExcelRead\WriteProduct.xlsx");
            return View("PrintReadItems", bigViewModel);
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public static void AddProperty(ExpandoObject expando, string propertyName, object propertyValue)
        {
            // ExpandoObject supports IDictionary so we can extend it like this
            var expandoDict = expando as IDictionary<string, object>;
            if (expandoDict.ContainsKey(propertyName))
                expandoDict[propertyName] = propertyValue;
            else
                expandoDict.Add(propertyName, propertyValue);
        }
        public static DataTable ToDataTable(List<ExpandoModel> list)
        {
            DataTable MethodResult = null;

            DataTable dt = new DataTable();
            ExpandoModel expandoModel = list.ElementAt(0);
            if (expandoModel.Expando == null)
            {
                return null;
            }
            // Add a column for each key of the expando object
            foreach (var item in expandoModel.Expando)
            {
                dt.Columns.Add(item.Key);
            }
            // Add the rows with values in the expando Objects
            foreach (var item in list)
            {
                int i = 0;
                DataRow dr = dt.NewRow();
                foreach (var obj in item.Expando)
                {
                    dr[i] = obj.Value;
                    i++;
                }
                dt.Rows.Add(dr);
            }
            dt.AcceptChanges();
            MethodResult = dt;
            return MethodResult;
        }
        public static bool ToExcelFile(DataTable dt, string filename)
        {
            bool Success = false;
            try
            {
                XLWorkbook wb = new XLWorkbook();

                wb.Worksheets.Add(dt, "Sheet 1");

                if (filename.Contains("."))
                {
                    int IndexOfLastFullStop = filename.LastIndexOf('.');

                    filename = filename.Substring(0, IndexOfLastFullStop) + ".xlsx";
                }
                filename = filename + ".xlsx";
                wb.SaveAs(filename);
                Success = true;
            }
            catch (Exception ex)
            {
                return false;
            }
            return Success;
        }

        public DataTable ReadExcel(string filePath)
        {

            filePath = @"D:\ExcelRead\ProductNames.xlsx";

            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
                return dt;

            }

        }
        private List<string> ConvertDataTable(DataTable dt, string column, List<string> cols)
        {
            if (column.Length != 1)
            {
                return null;
            }
            column = column.ToLower();
            int colInt = (int)column[0] - 'a';
            List<string> data = new List<string>();
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string value = dt.Columns[i].ToString();
                cols.Add(value);
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (dt.Rows[i][colInt] == DBNull.Value)
                {
                    //data.Add("INVALID DATA");
                }
                else
                {
                    string value = (string)(dt.Rows[i][colInt]);
                    data.Add(value);
                }
            }
            return data;
        }
    }

}
