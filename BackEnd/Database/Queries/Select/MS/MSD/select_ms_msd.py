from BackEnd.Database.General.get_connection import DatabaseConnection


class SelectMSMSD:

    def __init__(self, wo, client_sample_id):
        
        self.wo = wo
        self.client_sample_id = client_sample_id
        self.cursor = None
        
    def get_conn(self):
        
        instance_db = DatabaseConnection()
        connection = DatabaseConnection.get_conn(instance_db) 
        self.cursor = connection.cursor()

    def select_ms_msd(self):

        qry = """
        
            SELECT *
            FROM Sample_Tests WHERE
            LabReportingBatchID = ?
            AND ClientSampleID = ?
            AND TagParentSample = 1;
            
            """

        self.cursor.execute(qry, (self.wo, self.client_sample_id))
        data = self.cursor.fetchall()
        return [tuple(row) for row in data]

