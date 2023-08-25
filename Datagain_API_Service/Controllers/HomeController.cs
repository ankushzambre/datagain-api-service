using System;
using System.IO;
using Datagain_API_Service.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Data.SqlClient;
using Bytescout.Spreadsheet;

namespace Datagain_API_Service.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        //[HttpGet]
        //public IActionResult GenrateReport(string tax_year, string company_code)
        //{
        //    List<Report> reportList = new List<Report>();
        //    String connectionString = "Data Source=LAPTOP;Initial Catalog=datagain;Integrated Security=True";
        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();
        //        String query = "EXEC Report @tax_year=" + tax_year + ",@company_code=" + company_code;
        //        using (SqlCommand command = new SqlCommand(query, connection))
        //        {
        //            using (
        //                SqlDataReader reader = command.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    Report report = new Report();
        //                    report.year_1 = reader["year_1"] != DBNull.Value ? reader.GetDouble(0) : 0;
        //                    report.year_2 = reader["year_2"] != DBNull.Value ? reader.GetDouble(1) : 0;
        //                    report.description = reader.GetString(2) == null ? "" : reader.GetString(2); ;
        //                    reportList.Add(report);
        //                }
        //                if(reportList.Count>0)
        //                {

        //                    Spreadsheet document = new Spreadsheet();
        //                    document.LoadFromFile("./Final_Output.xlsx");

        //                    //Get worksheet by name
        //                    Worksheet worksheet = document.Workbook.Worksheets.ByName("Page 3");
        //                    for (var i = 0; i < reportList.Count; i++)
        //                    {
        //                        //Total Operating Revenue
        //                        if (reportList[i].description.ToString().ToLower() == "operating revenues (600)") 
        //                        {
        //                            worksheet.Cell("E9").Value = reportList[i].year_1 == null? 0 : reportList[i].year_1;
        //                            worksheet.Cell("F9").Value = reportList[i].year_2 == null ? 0 : reportList[i].year_2;
        //                        }
        //                        // Depreciation Expense
        //                        if (reportList[i].description.ToString().ToLower() == "depreciation and amortization (540)")
        //                        {
        //                            worksheet.Cell("E10").Value = reportList[i].year_1 == null ? 0 : reportList[i].year_1 * -1;
        //                            worksheet.Cell("F10").Value = reportList[i].year_2 == null ? 0 : reportList[i].year_2 * -1;
        //                        }

        //                        //Other Operating Expenses
        //                        if (reportList[i].description.ToString().ToLower() == "(less) operating expenses (610)")
        //                        {
        //                            worksheet.Cell("E13").Value = reportList[i].year_1 == null ? 0 : "=" + (reportList[i].year_1 * -1).ToString() + "-E10";
        //                            worksheet.Cell("F13").Value = reportList[i].year_2 == null ? 0 : "=" + (reportList[i].year_2 * -1).ToString() + "-E10";
        //                        }

        //                        //(Less) Interest Expense (650)
        //                        if (reportList[i].description.ToString().ToLower() == "(less) interest expense (650)")
        //                        {
        //                            worksheet.Cell("E21").Value = reportList[i].year_1 == null ? 0 : (reportList[i].year_1);
        //                            worksheet.Cell("F21").Value = reportList[i].year_2 == null ? 0 : (reportList[i].year_2);
        //                        }

        //                        //Capital Expenditures
        //                        if (reportList[i].description.ToString().ToLower() == "total (lines 37 thru 45)")
        //                        {
        //                            worksheet.Cell("E28").Value = reportList[i].year_1 == null ? 0 : (reportList[i].year_1 * -1);
        //                            worksheet.Cell("F28").Value = reportList[i].year_2 == null ? 0 : (reportList[i].year_2 * -1);
        //                        }

        //                        //Net Owned Operating Property*	
        //                        if (reportList[i].description.ToString().ToLower() == "total tangible property (total of lines 31, 32, and 35)")
        //                        {
        //                            worksheet.Cell("E50").Value = reportList[i].year_1 == null ? 0 : (reportList[i].year_1 * -1);
        //                            worksheet.Cell("F50").Value = reportList[i].year_2 == null ? 0 : (reportList[i].year_2 * -1);
        //                        }

        //                        //Construction Work in Progress- Real and Personal	
        //                        if (reportList[i].description.ToString().ToLower() == "total tangible property (total of lines 31, 32, and 35)")
        //                        {
        //                            worksheet.Cell("E50").Value = reportList[i].year_1 == null ? 0 : (reportList[i].year_1 * -1);
        //                            worksheet.Cell("F50").Value = reportList[i].year_2 == null ? 0 : (reportList[i].year_2 * -1);
        //                        }
        //                    }

        //                    // Save document
        //                    var file_name = DateTime.Now.ToString("yyyyMMddTHHmmss");
        //                    document.SaveAs("report_"+file_name+".xls");

        //                    // Close Spreadsheet
        //                    document.Close();
        //                    return Ok("Report Genrated");
        //                }
        //                else
        //                {
        //                    return BadRequest("Unable to genrate report");
        //                }
        //            }
        //        }
        //    }
        //}

        //[HttpGet]
        //public IActionResult GenrateCompnyList()
        //{
        //    List<CompanyList> company_list = new List<CompanyList>();
        //    String connectionString = "Data Source=LAPTOP;Initial Catalog=datagain;Integrated Security=True";
        //    using (SqlConnection connection = new SqlConnection(connectionString))
        //    {
        //        connection.Open();
        //        String query = "EXEC Get_Company_List_With_Tax_Year";
        //        using (SqlCommand command = new SqlCommand(query, connection))
        //        {
        //            using (
        //                SqlDataReader reader = command.ExecuteReader())
        //            {
        //                while (reader.Read())
        //                {
        //                    CompanyList company_obj = new CompanyList();
        //                    company_obj.tax_year = reader.GetDouble(0);
        //                    company_obj.company_name = "" + reader.GetString(1).ToString();
        //                    company_obj.company_code = "" + reader.GetString(2).ToString();
        //                    company_obj.state = "" + reader.GetString(3).ToString();
        //                    company_obj.country = "" + reader.GetString(4).ToString();
        //                    company_list.Add(company_obj);
        //                }
        //                if (company_list.Count > 0)
        //                {
        //                    return Ok(company_list);
        //                }
        //                else
        //                {
        //                    return NotFound("No record found");
        //                }
        //            }
        //        }
        //    }
        //}
    }
}