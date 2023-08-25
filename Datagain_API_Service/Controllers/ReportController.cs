using System;
using System.IO;
using Datagain_API_Service.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Data.SqlClient;
using Bytescout.Spreadsheet;
using System.Diagnostics.Metrics;

namespace Datagain_API_Service.Controllers
{
    public class ReportController : Controller
    {
        // GET: ReportController
        public ActionResult Index()
        {
            return View();
        }

        // GET: ReportController/Details/5
        public ActionResult Details(int id)
        {
            return View();
        }

        // GET: ReportController/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: ReportController/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ReportController/Edit/5
        public ActionResult Edit(int id)
        {
            return View();
        }

        // POST: ReportController/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }

        // GET: ReportController/Delete/5
        public ActionResult Delete(int id)
        {
            return View();
        }

        // POST: ReportController/Delete/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id, IFormCollection collection)
        {
            try
            {
                return RedirectToAction(nameof(Index));
            }
            catch
            {
                return View();
            }
        }


        [HttpGet]
        public IActionResult GenrateReport(string tax_year, string company_code)
        {
            List<Report> reportList = new List<Report>();
            String connectionString = "Data Source=LAPTOP;Initial Catalog=datagain;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                String query = "EXEC Report @tax_year=" + tax_year + ",@company_code=" + company_code;
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (
                        SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            Report report = new Report();
                            report.name = reader.GetString(0) == null ? "" : reader.GetString(0);
                            report.year_1 = reader["year_1"] != DBNull.Value ? reader.GetDouble(1) : 0;
                            report.year_2 = reader["year_2"] != DBNull.Value ? reader.GetDouble(2) : 0;
                            reportList.Add(report);
                        }
                        if (reportList.Count > 0)
                        {

                            Spreadsheet document = new Spreadsheet();
                            document.LoadFromFile("./Final_Output.xlsx");

                            // ====================   Setting up Page 4 sheet data start    ====================
                            //Get worksheet Page 4
                            Worksheet worksheet_page_4 = document.Workbook.Worksheets.ByName("Page 4");

                            //Construction work in progress - Total
                            var construction_work_in_progress_total = reportList.Find(x => x.name.ToLower().Trim() == "construction work in progress - total");
                            worksheet_page_4.Cell("C11").Value = construction_work_in_progress_total.year_2 == null ? 0 : construction_work_in_progress_total.year_2;
                            worksheet_page_4.Cell("D11").Value = construction_work_in_progress_total.year_1 == null ? 0 : construction_work_in_progress_total.year_1;

                            //Inventories, materials and supplies*
                            var inventories_materials_and_supplies= reportList.Find(x => x.name.ToLower().Trim() == "inventories, materials and supplies*");
                            worksheet_page_4.Cell("C16").Value = inventories_materials_and_supplies.year_2 == null ? 0 : inventories_materials_and_supplies.year_2;
                            worksheet_page_4.Cell("D16").Value = inventories_materials_and_supplies.year_1 == null ? 0 : inventories_materials_and_supplies.year_1;

                            //Accumulated depreciation
                            var accumulated_depreciation = reportList.Find(x => x.name.ToLower().Trim() == "accumulated depreciation");
                            worksheet_page_4.Cell("C18").Value = accumulated_depreciation.year_2 == null ? 0 : accumulated_depreciation.year_2;
                            worksheet_page_4.Cell("D18").Value = accumulated_depreciation.year_1 == null ? 0 : accumulated_depreciation.year_1;

                            //Current assets (less materials and supplies)
                            var current_assets = reportList.Find(x => x.name.ToLower().Trim() == "current assets (less materials and supplies)");
                            worksheet_page_4.Cell("C21").Value = current_assets.year_2 == null ? "=0 - C16" : "=" + current_assets.year_2.ToString() + "-C16";
                            worksheet_page_4.Cell("D21").Value = current_assets.year_1 == null ? "=0 - D16" : "=" + current_assets.year_1.ToString() + "-C16";

                            //Investments and other assets
                            var investments_other_assets= reportList.Find(x => x.name.ToLower().Trim() == "investments and other assets");
                            worksheet_page_4.Cell("C22").Value = investments_other_assets.year_2 == null ? 0 : investments_other_assets.year_2;
                            worksheet_page_4.Cell("D22").Value = investments_other_assets.year_1 == null ? 0 : investments_other_assets.year_1;

                            //Common stock and paid-in capital
                            var common_stock_and_paid_in_capital = reportList.Find(x => x.name.ToLower().Trim() == "common stock and paid-in capital");
                            worksheet_page_4.Cell("C36").Value = common_stock_and_paid_in_capital.year_2 == null ? 0 : common_stock_and_paid_in_capital.year_2;
                            worksheet_page_4.Cell("D36").Value = common_stock_and_paid_in_capital.year_1 == null ? 0 : common_stock_and_paid_in_capital.year_1;

                            //Retained earnings
                            var retained_earnings = reportList.Find(x => x.name.ToLower().Trim() == "retained earnings");
                            worksheet_page_4.Cell("C24").Value = retained_earnings.year_2 == null ? 0 : retained_earnings.year_2;
                            worksheet_page_4.Cell("D24").Value = retained_earnings.year_1 == null ? 0 : retained_earnings.year_1;

                            //Long-term debt due after one year
                            var long_term_debt_due_after_one_year = reportList.Find(x => x.name.ToLower().Trim() == "long-term debt due after one year");
                            worksheet_page_4.Cell("C26").Value = long_term_debt_due_after_one_year.year_2 == null ? 0 : long_term_debt_due_after_one_year.year_2;
                            worksheet_page_4.Cell("D26").Value = long_term_debt_due_after_one_year.year_1 == null ? 0 : long_term_debt_due_after_one_year.year_1;

                            //Current liabilities
                            var current_liabilities = reportList.Find(x => x.name.ToLower().Trim() == "current liabilities");
                            worksheet_page_4.Cell("C42").Value = current_liabilities.year_2 == null ? 0 : current_liabilities.year_2;
                            worksheet_page_4.Cell("D42").Value = current_liabilities.year_1 == null ? 0 : current_liabilities.year_1;

                            // ====================   Setting up Page 4 sheet data end    ====================

                            // ====================   Setting up Page 3 sheet data start  ====================
                            //Get worksheet Page 3
                            Worksheet worksheet = document.Workbook.Worksheets.ByName("Page 3");

                            //Total Operating Revenue
                            var total_operating_revenue = reportList.Find(x => x.name.ToLower() == "total operating revenue");
                            worksheet.Cell("E9").Value = total_operating_revenue.year_1 == null ? 0 : total_operating_revenue.year_1;
                            worksheet.Cell("F9").Value = total_operating_revenue.year_2 == null ? 0 : total_operating_revenue.year_2;
                            
                            //Depreciation Expense
                            var depreciation_expense = reportList.Find(x => x.name.ToLower() == "depreciation expense");
                            worksheet.Cell("E10").Value = depreciation_expense.year_1 == null ? 0 : depreciation_expense.year_1 * -1;
                            worksheet.Cell("F10").Value = depreciation_expense.year_2 == null ? 0 : depreciation_expense.year_2 * -1;

                            //Other Operating Expenses
                            var other_operating_expenses = reportList.Find(x => x.name.ToLower() == "other operating expenses");
                            worksheet.Cell("E13").Value = other_operating_expenses.year_1 == null ? 0 : "=" + (other_operating_expenses.year_1 * -1).ToString() + "-E10";
                            worksheet.Cell("F13").Value = other_operating_expenses.year_2 == null ? 0 : "=" + (other_operating_expenses.year_2 * -1).ToString() + "-F10";

                            //Non - operating Income(Loss)
                            var account_640 = reportList.Find(x => x.name.ToLower() == "640");
                            var account_660 = reportList.Find(x => x.name.ToLower() == "660");
                            //var year_1 = account_640.year_1 != null & account_660.year_1 != null ? account_640.year_1 - account_660.year_1 :( account_640.year_1 == null ? (account_660.year_1 == null ? 0 : account_660.year_2) : (account_660.year_1 == null ? 0 : account_660.year_2));

                            var account_640_year_1 = account_640.year_1 == null ? 0 : account_640.year_1;
                            var account_660_year_1 = account_660.year_1 == null ? 0 : account_660.year_1;
                            var account_640_year_2 = account_640.year_2 == null ? 0 : account_640.year_2;
                            var account_660_year_2 = account_660.year_2 == null ? 0 : account_660.year_2;
                            worksheet.Cell("E19").Value = "=" + account_640_year_1.ToString() +" - "+account_660_year_1.ToString();
                            worksheet.Cell("F19").Value = "=" + account_640_year_2.ToString() +" - "+ account_660_year_2.ToString();


                            //Interest Expense
                            var interest_expense = reportList.Find(x => x.name.ToLower() == "interest expense");
                            worksheet.Cell("E21").Value = interest_expense.year_1 == null ? 0 : interest_expense.year_1;
                            worksheet.Cell("F21").Value = interest_expense.year_2 == null ? 0 : interest_expense.year_2;

                            //Capital Expenditures
                            var capital_expenditures = reportList.Find(x => x.name.ToLower() == "capital expenditures");
                            worksheet.Cell("E28").Value = capital_expenditures.year_1 == null ? 0 : capital_expenditures.year_1;
                            worksheet.Cell("F28").Value = capital_expenditures.year_2 == null ? 0 : capital_expenditures.year_2;

                            //Net Owned Operating Property*	
                            var net_owned_operating_property = reportList.Find(x => x.name.ToLower() == "net owned operating property*");
                            worksheet.Cell("E50").Value = net_owned_operating_property.year_1 == null ? 0 : net_owned_operating_property.year_1 ;
                            worksheet.Cell("F50").Value = net_owned_operating_property.year_2 == null ? 0 : net_owned_operating_property.year_2;
                            
                            //Construction Work in Progress- Real and Personal	
                            var construction_work_in_progress = reportList.Find(x => x.name.ToLower() == "construction work in progress- real and personal");
                            worksheet.Cell("E53").Value = construction_work_in_progress.year_1 == null ? 0 : construction_work_in_progress.year_1;
                            worksheet.Cell("F53").Value = construction_work_in_progress.year_2 == null ? 0 : construction_work_in_progress.year_2;


                            //Get worksheet Page 4
                            //Worksheet worksheet_page_4 = document.Workbook.Worksheets.ByName("Page 4");
                            worksheet.Cell("E51").Value = worksheet_page_4.Cell("C29").Value;
                            worksheet.Cell("F51").Value = worksheet_page_4.Cell("D29").Value;

                            // ====================   Setting up Page 3 sheet data end    ====================

                            // Save new report
                            var file_name = DateTime.Now.ToString("yyyyMMddTHHmmss");
                            document.SaveAs("report_" + file_name + ".xls");

                            // Close Spreadsheet
                            document.Close();
                            return Ok("Report Genrated");
                        }
                        else
                        {
                            return BadRequest("Unable to genrate report");
                        }
                    }
                }
            }
        }

        [HttpGet]
        public IActionResult GenrateCompnyList()
        {
            List<CompanyList> company_list = new List<CompanyList>();
            String connectionString = "Data Source=LAPTOP;Initial Catalog=datagain;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                String query = "EXEC Get_Company_List_With_Tax_Year";
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (
                        SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            CompanyList company_obj = new CompanyList();
                            company_obj.tax_year = reader.GetDouble(0);
                            company_obj.company_name = "" + reader.GetString(1).ToString();
                            company_obj.company_code = "" + reader.GetString(2).ToString();
                            company_obj.state = "" + reader.GetString(3).ToString();
                            company_obj.country = "" + reader.GetString(4).ToString();
                            company_list.Add(company_obj);
                        }
                        if (company_list.Count > 0)
                        {
                            return Ok(company_list);
                        }
                        else
                        {
                            return NotFound("No record found");
                        }
                    }
                }
            }
        }
    }
}
