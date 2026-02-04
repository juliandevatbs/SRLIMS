from BackEnd.Database.General.get_connection import DatabaseConnection


class AnalyticalResultsQuery:



    def __init__(self):

        return



    def analytical_results_query(self, lab_reporting_batch_id, cursor):


        qry = """
        
              SELECT
                s.ClientSampleID,
                s.LabSampleID,
                s.DateCollected,
                s.Sampler,
                s.MatrixID,

                t.AnalyteName,
                t.Result,
                t.ResultUnits,
                t.Dilution,
                t.DetectionLimit,
                t.ReportingLimit,
                t.LabAnalysisRefMethodID,
                t.Analyst,
                t.MethodBatchID,
                t.Notes,
                t.GroupLongName,
                t.LabQualifiers

                FROM Sample_Tests t
                JOIN Samples s
                    ON s.LabSampleID = t.LabSampleID
                AND s.LabReportingBatchID = t.LabReportingBatchID
                AND s.QCSample = 0
                WHERE t.LabReportingBatchID = ?
                ORDER BY
                    s.LabSampleID,
                    t.GroupLongName,
                    --t.Test_Group
                    t.AnalyteName;

            """

        cursor.execute(qry, (lab_reporting_batch_id, ))


        data = cursor.fetchall()

        return data







