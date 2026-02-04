from BackEnd.Database.General.get_connection import DatabaseConnection


class SelectAdaptData:

    def __init__(self, cursor, wo):
        self.wo = wo
        self.cursor = cursor
        pass

    def select_adapt_data(self):

        qry = """
            SELECT 
                st.ClientSampleID, 
                st.LabAnalysisRefMethodID, 
                st.LabSampleID, 
                st.LabID, 
                st.ClientAnalyteID, 
                st.AnalyteName, 
                st.Result, 
                st.Error, 
                st.ResultUnits, 
                st.LabQualifiers, 
                st.DetectionLimit, 
                st.AnalyteType, 
                st.Dilution, 
                st.PercentMoisture, 
                st.PercentRecovery, 
                st.RelativePercentDifference, 
                st.ReportingLimit, 
                st.ProjectNumber, 
                st.ProjectName, 
                st.DateCollected, 
                st.MatrixID, 
                st.QCType, 
                st.ShippingBatchID, 
                st.Temperature, 
                st.PreparationType, 
                st.AnalysisType, 
                CASE WHEN st.ReportableResult = 1 THEN 'Yes' ELSE 'No' END AS Reportable,
                FORMAT(TRY_CONVERT(datetime, st.DatePrepared), 'MM/dd/yy HH:mm') AS Dateprep,
                FORMAT(TRY_CONVERT(datetime, st.DateAnalyzed), 'MM/dd/yy HH:mm') AS DateAnalyz,
                st.TotalOrDissolved, 
                st.PrepBatchID, 
                st.MethodBatchID, 
                CASE WHEN st.PreservationIntact = 1 THEN 'Yes' ELSE 'No' END AS Preservation,
                st.QCSpikeAdded, 
                st.ResultComments, 
                st.LabReportingBatchID,
                st.GroupLongName,
                st.High_Limit,
                st.Low_Limit,
                st.[Order],
                st.Analyst,
                st.Limits
            FROM Sample_Tests st
            INNER JOIN Samples s ON st.LabSampleID = s.LabSampleID
            WHERE st.LabReportingBatchID = ?
            ORDER BY st.LabAnalysisRefMethodID, st.LabSampleID, st.[Order];
        """

        self.cursor.execute(qry, (self.wo,))
        data = self.cursor.fetchall()
        return [tuple(row) for row in data]


    def select_epp_data(self):
       
        qry = """
                SELECT DISTINCT 
                    fdd.WACS_Testsite_ID, 
                    fdd.WACS_Testsite_Name, 
                    fdd.WACS_Facility_ID, 
                    fdd.WACS_Facility_Name, 
                    CASE 
                        WHEN fdd.Matrix LIKE '%S%' THEN 'E'
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN 'E'
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN 'E'
                        ELSE ''
                    END AS Sample_Type,
                    fdd.Matrix, 
                    CASE 
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN ''
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN ''
                        ELSE fmp.Field_Measurement_Method
                    END AS Field_Measurement_Method,
                    CASE 
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN ''
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN ''
                        ELSE fmp.Field_Parameter_NameAnalyteName
                    END AS Field_Parameter_Name,
                    fdd.Result, 
                    CASE 
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN ''
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN ''
                        ELSE fmp.Result_Units
                    END AS Result_Units,
                    CASE 
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN ''
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN ''
                        ELSE fmp.Field_Parameter_Qualifier_Code
                    END AS Field_Parameter_Qualifier_Code,
                    fdd.Field_Comments, 
                    fdd.Sampler, 
                    fdd.CollectionAgency, 
                    fdd.Date_Sampled, 
                    fdd.Shipping_Batch_ID, 
                    fdd.Well_Purged_Flag, 
                    fdd.WACS_Report_Type
                FROM (
                    -- Esta subconsulta replica *_FDD_Elements_From_LDD
                    SELECT DISTINCT 
                        '' AS WACS_Testsite_ID,
                        t.ClientSampleID AS WACS_Testsite_Name,
                        t.ProjectNumber AS WACS_Facility_ID,
                        t.ProjectName AS WACS_Facility_Name,
                        'E' AS Sample_Type,
                        CASE 
                            WHEN t.MatrixID LIKE 'AQ%' THEN 'W'
                            ELSE 'S'
                        END AS Matrix,
                        '' AS Field_Measurement_Method,
                        '' AS Field_Parameter_NameAnalyteName,
                        '' AS Result,
                        '' AS Result_Units,
                        '' AS Field_Parameter_Qualifier_Code,
                        '' AS Field_Comments,
                        s.Sampler,
                        s.CollectionAgency,
                        t.DateCollected AS Date_Sampled,
                        t.ShippingBatchID AS Shipping_Batch_ID,
                        CASE 
                            WHEN t.MatrixID LIKE 'AQ%' THEN 'Y'
                            ELSE ''
                        END AS Well_Purged_Flag,
                        CASE 
                            WHEN t.LabAnalysisRefMethodID LIKE '%1311%' THEN 'TCLP'
                            WHEN t.LabAnalysisRefMethodID LIKE '%1312%' THEN 'SPLP'
                            WHEN t.MatrixID LIKE 'AQ%' THEN 'SEMGW'
                            ELSE 'ASMNT'
                        END AS WACS_Report_Type
                    FROM tblHoldDepEdd t
                    INNER JOIN Samples s ON t.LabSampleID = s.LabSampleID
                    WHERE t.AnalyteType = 'TRG' 
                        AND t.QCType = 'N'
                ) fdd
                INNER JOIN Field_Measured_Parameters fmp 
                    ON fdd.Matrix = fmp.Matrix
                ORDER BY 
                    fdd.WACS_Testsite_Name, 
                    CASE 
                        WHEN fdd.WACS_Report_Type LIKE '%LP%' THEN ''
                        WHEN UPPER(fdd.WACS_Testsite_Name) LIKE '%BLANK%' THEN ''
                        ELSE fmp.Field_Parameter_NameAnalyteName
                    END;
            """

        self.cursor.execute(qry)
        data = self.cursor.fetchall()
        return [tuple(row) for row in data]
    
    def clear_temp_table(self):
        
        """
        
        Limpia la tabla temporal tblHoldDepEdd
        Equivalente a: DoCmd.RunSQL ("DELETE* FROM tblHoldDepEdd;")
        
        """
        qry = "DELETE FROM tblHoldDepEdd;"
        self.cursor.execute(qry)
        self.cursor.connection.commit()
    
    def insert_adapt_data(self):

        qry = """
            INSERT INTO tblHoldDepEdd (
                ClientSampleID, LabAnalysisRefMethodID, LabSampleID, LabID, 
                ClientAnalyteID, AnalyteName, Result, Error, ResultUnits, 
                LabQualifiers, DetectionLimit, AnalyteType, Dilution, 
                PercentMoisture, PercentRecovery, RelativePercentDifference, 
                ReportingLimit, ProjectNumber, ProjectName, DateCollected, 
                MatrixID, QCType, ShippingBatchID, Temperature, PreparationType, 
                AnalysisType, ReportableResult, DatePrepared, DateAnalyzed, 
                TotalOrDissolved, PrepBatchID, MethodBatchID, PreservationIntact, 
                QCSpikeAdded, ResultComments, LabReportingBatchID, [Order],
                GroupLongName, High_Limit, Low_Limit, Analyst, Limits
            )
            SELECT 
                st.ClientSampleID, 
                st.LabAnalysisRefMethodID, 
                st.LabSampleID, 
                st.LabID, 
                st.ClientAnalyteID, 
                st.AnalyteName, 
                st.Result, 
                st.Error, 
                st.ResultUnits, 
                st.LabQualifiers, 
                st.DetectionLimit, 
                st.AnalyteType, 
                st.Dilution, 
                st.PercentMoisture, 
                st.PercentRecovery, 
                st.RelativePercentDifference, 
                st.ReportingLimit, 
                st.ProjectNumber, 
                st.ProjectName, 
                TRY_CONVERT(datetime, st.DateCollected),
                st.MatrixID, 
                st.QCType, 
                st.ShippingBatchID, 
                st.Temperature, 
                st.PreparationType, 
                st.AnalysisType,
                st.ReportableResult,
                TRY_CONVERT(datetime, st.DatePrepared),
                TRY_CONVERT(datetime, st.DateAnalyzed),
                st.TotalOrDissolved, 
                st.PrepBatchID, 
                st.MethodBatchID,
                st.PreservationIntact,
                st.QCSpikeAdded, 
                st.ResultComments, 
                st.LabReportingBatchID,
                st.[Order],
                st.GroupLongName,
                st.High_Limit,
                st.Low_Limit,
                st.Analyst,
                st.Limits
            FROM Sample_Tests st
            WHERE st.LabReportingBatchID = ?
            ORDER BY st.LabAnalysisRefMethodID, st.LabSampleID, st.[Order];
        """

        self.cursor.execute(qry, (self.wo,))
        self.cursor.connection.commit()

        self.cursor.execute("SELECT COUNT(*) FROM tblHoldDepEdd")
        return self.cursor.fetchone()[0]
