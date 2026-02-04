from BackEnd.Database.General.get_connection import DatabaseConnection


class FirstPageQuery:



    def __init__(self):

        pass



    def first_page_query(self, lab_reporting_batch_id, cursor):


        qry = """
            SELECT
                S.ItemID,
                S.LabSampleID,
                S.ClientSampleID,
                S.DateCollected,
                S.MatrixID,
                Methods = STUFF(
                    (
                        SELECT DISTINCT ', ' + ST.LabAnalysisRefMethodID
                        FROM Sample_Tests ST
                        WHERE ST.LabReportingBatchID = S.LabReportingBatchID
                        FOR XML PATH(''), TYPE
                    ).value('.', 'NVARCHAR(MAX)')
                , 1, 2, '')
            FROM Samples S
            WHERE S.LabReportingBatchID = ?
            AND QCSample = 0
            AND ClientSampleID NOT LIKE '%Trip Blank%'
            ORDER BY S.ItemID;

            """

        cursor.execute(qry, (lab_reporting_batch_id, ))


        data = cursor.fetchall()

        return data







