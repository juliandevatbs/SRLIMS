from BackEnd.Database.General.get_connection import DatabaseConnection


class SelectDataByWo:


    def __init__(self):

        self.connection = None
        self.cursor = None


        self.lab_sample_ids = None
        self.login_data = None
        self.samples = None
        self.sample_tests = None


        self.query_lab_sample_ids = """
 
                                    SELECT DISTINCT LabSampleID FROM Samples WHERE LabReportingBatchID = ?;

                                    """

        self.query_login_data =     """
        
                                SELECT TOP 1
                                    SL.LabReportingBatchID,
                                    S.LabSampleID,               
                                    SL.ProjectName,
                                    SL.ProjectLocation,
                                    SL.Date_Due,
                                    SL.LabReceiptDate,
                                    SL.Contact,
                                    SL.Email,
                                    SL.Address_1,
                                    SL.City,
                                    SL.Postal_Code,
                                    SL.Phone,
                                    SL.State_Prov,
                                    SL.ClientProjectLocation,
                                    S.MatrixID

                                FROM Samples S 
                                LEFT JOIN Sample_Login SL ON S.LabReportingBatchID = SL.LabReportingBatchID 
                                WHERE S.LabReportingBatchID = ?
                                ORDER BY S.LabReportingBatchID DESC;
                            """

        self.samples_query = """
        
                                SELECT  
		
                                            ItemID,
                                            LabReportingBatchID,
                                            LabSampleID,
                                            ClientSampleID,
                                            Sampler,
                                            DateCollected,
                                            MatrixID,
                                            Temperature,
                                            ShippingBatchID,
                                            CollectionMethod,
                                            CollectionAgency,
                                            AdaptMatrixID,
                                            LabID
                                    
                                    FROM Samples 
                                    WHERE LabReportingBatchID = ?
                                    ORDER BY LabReportingBatchID DESC; 

        
        
        
        
        """

        self.sample_tests_query = """
        
        
                            SELECT 
                                ST.ClientSampleID,
                                ST.LabAnalysisRefMethodID,
                                ST.LabSampleID,
                                ST.AnalyteName,
                                ST.Result,
                                ST.ResultUnits,
                                ST.DetectionLimit,
                                ST.Dilution,
                                ST.ReportingLimit,
                                ST.ProjectName,
                                ST.DateCollected,
                                ST.MatrixID,
                                ST.AnalyteType,
                                ST.LabReportingBatchID,
                                ST.Notes,
                                S.Sampler,
                                ST.Analyst,
        
                            FROM Sample_Tests ST
                            LEFT JOIN Samples S
                            ON ST.LabReportingBatchID = S.LabReportingBatchID
                            WHERE ST.LabReportingBatchID = ? ;

        """

        return

    def load_connection(self):

        try:

            instance_db = DatabaseConnection()
            self.connection = instance_db.get_conn()

            if self.connection:

                self.cursor = self.connection.cursor()

            else:

                print("No database connection")

        except Exception as e:
            print(f"Error in the database connection: {e}")

    def close_connection(self):

        if self.cursor:
            self.cursor.close()

        if self.connection:
            self.connection.close()

    def select_lab_sample_ids_by_wo(self, wo: str):

        self.cursor.execute(self.query_lab_sample_ids, (wo, ))

        results = self.cursor.fetchall()

        self.lab_sample_ids = results

    def select_login_data(self, wo: str):

        self.cursor.execute(self.query_login_data, (wo, ))

        results = self.cursor.fetchall()

        self.login_data = results

        print(f"SELECTED LOGIN DATA {results}")

        return results

    def select_samples_by_wo(self, wo: str):

        self.cursor.execute(self.samples_query, (wo, ))

        results = self.cursor.fethall()

        self.samples = results

    def select_sample_tests_by_wo(self, wo: str):

        self.cursor.execute(self.sample_tests_query, (wo, ))

        results = self.cursor.fetchall()

        self.sample_tests = results














