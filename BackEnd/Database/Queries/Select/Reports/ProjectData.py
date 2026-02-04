from BackEnd.Database.General.get_connection import DatabaseConnection


class ProjectData:



    def __init__(self):

        return



    def project_data_query(self, lab_reporting_batch_id, cursor):


        qry = """
            SELECT
            SL.Address_1,
            SL.City,
            C.Client,
            SL.ClientProjectNumber,
            SL.Contact,
            SL.LabReceiptDate,
            SL.LabReportingBatchID,
            SL.Phone,
            SL.Postal_Code,
            SL.ProjectLocation,
            SL.ProjectName,
            SL.State_Prov                        
            FROM Sample_Login SL
            JOIN Clients C ON SL.Client_ID = C.Client_ID
            WHERE LabReportingBatchID = ?

            """

        cursor.execute(qry, (lab_reporting_batch_id, ))


        data = cursor.fetchall()

        return data







