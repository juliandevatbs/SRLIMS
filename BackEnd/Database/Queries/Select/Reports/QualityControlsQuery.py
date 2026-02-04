from BackEnd.Database.General.get_connection import DatabaseConnection


class QualityControlsQuery:

    def __init__(self):
        return

    def quality_controls_query(self, lab_reporting_batch_id, cursor):
        qry = """

            SELECT
                s.ClientSampleID,
                s.LabSampleID,
                COALESCE(s.DateCollected, t.DatePrepared) AS DateCollected,
                t.DatePrepared,
                s.DateAnalyzed,
                s.Sampler,
                s.MatrixID,
                t.AnalyteName,
                t.Result,
                t.QCSpikeAdded,
                t.LabQualifiers,
                t.ResultUnits,
                t.Dilution,
                t.DetectionLimit,
                t.ReportingLimit,
                t.LabAnalysisRefMethodID,
                t.Analyst,
                t.MethodBatchID,
                t.Notes,
                t.RelativePercentDifference,
                t.PercentRecovery,
                t.Limits as src,
                CASE 
                    WHEN t.Low_Limit IS NOT NULL AND t.High_Limit IS NOT NULL
                    THEN CAST(t.Low_Limit AS VARCHAR(10)) + '-' + CAST(t.High_Limit AS VARCHAR(10))
                    ELSE NULL
                END AS Limites
            FROM Sample_Tests t
            JOIN Samples s ON s.LabSampleID = t.LabSampleID
                AND s.LabReportingBatchID = t.LabReportingBatchID
                AND s.QCSample = 1
            WHERE t.LabReportingBatchID = ?

            """

        cursor.execute(qry, (lab_reporting_batch_id,))

        data = cursor.fetchall()

        return data







