--===================================== Get_Company_List_With_Tax_Year Procedure ============================


USE [datagain]
GO
/****** Object:  StoredProcedure [dbo].[Get_Company_List_With_Tax_Year]    Script Date: 25-08-2023 12:05:55 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:     Ankush Zambre
-- Create date: 22-8-23
-- Description: Get company list with tax year and state wise
-- =============================================
CREATE PROCEDURE [dbo].[Get_Company_List_With_Tax_Year] 

AS
BEGIN
     
    SET NOCOUNT ON;
    BEGIN

		SELECT TaxYear, company, [Company code] as company_code, state, county as country 
		FROM dbo.GIS$ 
		WHERE TaxYear IS NOT NULL AND [Company code] IS NOT NULL AND state IS NOT NULL AND county IS NOT NULL
		GROUP BY county,company,[Company code], state, TaxYear
		ORDER BY company_code, state, county;

    END   
END