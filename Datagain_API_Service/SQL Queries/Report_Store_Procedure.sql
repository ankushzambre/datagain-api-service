--===================================== Report Procedure ============================

USE [datagain]
GO
 
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:     Ankush Zambre
-- Create date: 21-8-23
-- Description: Generate Report
-- =============================================
CREATE PROCEDURE [dbo].[Report] 
@tax_year INT= NULL,
@company_code INT= NULL
     
AS
BEGIN
     
    SET NOCOUNT ON;
    BEGIN
		SELECT  'Total Operating Revenue' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no' FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account = '600' AND Page = 114 AND line_no = 1
	UNION
		SELECT 'Depreciation Expense' AS name,  balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no' FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account = '540' AND Page = 302 AND line_no = 13
	UNION
		SELECT 'Other Operating Expenses' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='GT' AND Page = 302 AND line_no = 23
	UNION
		SELECT '640' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='640' AND Page = 114 AND line_no = 6
	UNION
		SELECT '660' AS name,  balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='660' AND Page = 114 AND line_no = 9
	UNION
		SELECT 'Interest Expense' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='650' AND Page = 114 AND line_no = 8
	UNION
		SELECT 'Capital Expenditures' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='187' AND Page = 212 AND line_no = 45
	UNION
		SELECT 'Net Owned Operating Property*' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='NCP' AND Page = 111 AND line_no = 31
	UNION
		SELECT 'Construction Work in Progress- Real and Personal' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 3' AS 'report_page_no'  FROM FERC$ WHERE company_code = @company_code AND tax_year = @tax_year AND Account='187' AND Page = 212 AND line_no = 45
	UNION
		SELECT 'Historical cost of plant in service' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='30' AND Page = 110 AND line_no = 28
	UNION
		SELECT 'Construction work in progress - Total' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='187' AND Page = 212 AND line_no = 45
	UNION
		SELECT  'Inventories, materials and supplies*' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='17' AND Page = 110 AND line_no = 10
	UNION
		SELECT  'Accumulated depreciation' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='31' AND Page = 111 AND line_no = 29
	UNION
		SELECT  'Current assets (less materials and supplies)' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='TCA' AND Page = 110 AND line_no = 14
	UNION
		SELECT  'Investments and other assets' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='43' AND Page = 111 AND line_no = 40
	UNION
		SELECT  'Common stock and paid-in capital' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='73' AND Page = 113 AND line_no = 71
	UNION 
		SELECT  'Retained earnings' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='75' AND Page = 113 AND line_no = 73
	UNION
		SELECT  'Long-term debt due after one year' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='60' AND Page = 113 AND line_no = 58
	UNION
		SELECT  'Current liabilities' AS name, balance_12_31_2021 AS 'year_1', balance_12_31_2020 AS 'year_2', 'Page 4' AS 'report_page_no' FROM FERC$ WHERE company_code = 19 AND tax_year = 2022 AND Account='TCL' AND Page = 113 AND line_no = 57;
    END   
END

-- To RUN Procedure
-- EXEC Report 2022, 19