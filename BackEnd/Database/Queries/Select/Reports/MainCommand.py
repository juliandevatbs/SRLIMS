from BackEnd.Database.General.get_connection import DatabaseConnection
from BackEnd.Database.Queries.Select.Reports.AnalyticalResultsQuery import AnalyticalResultsQuery
from BackEnd.Database.Queries.Select.Reports.FirstPageQuery import FirstPageQuery
from BackEnd.Database.Queries.Select.Reports.ProjectData import ProjectData
from BackEnd.Database.Queries.Select.Reports.QualityControlsQuery import QualityControlsQuery
from BackEnd.Processes.DataTypes.Report.build_samples_structure import build_samples_structure
from BackEnd.Processes.DataTypes.Report.convert_row_to_dict import row_to_dicts


class MainCommand:


    def __init__(self):

        self.instance_db = DatabaseConnection()

        self.conn = None

        self.cursor = None

        self.first_data_instance = FirstPageQuery()

        self.analytical_results_instance = AnalyticalResultsQuery()

        self.quality_controls_instance = QualityControlsQuery()

        self.project_data_instance = ProjectData()
        
        
        self.first_data = None

        self.analytical_results = None

        self.quality_controls = None

        self.project_data = None

    def load_connection(self):

        self.conn = DatabaseConnection.get_conn(self.instance_db)

        self.cursor = self.conn.cursor()

    def close_connection(self):

        if self.cursor:

            self.cursor.close()

        if self.conn:

            self.conn.close()


    def caller(self, lab_reporting_batch_id):

        self.load_connection()

        try:
            
            self.project_data = self.project_data_instance.project_data_query(lab_reporting_batch_id, self.cursor)

            self.first_data = self.first_data_instance.first_page_query(lab_reporting_batch_id, self.cursor)

            self.analytical_results = self.analytical_results_instance.analytical_results_query(lab_reporting_batch_id, self.cursor)

            self.quality_controls = self.quality_controls_instance.quality_controls_query(lab_reporting_batch_id, self.cursor)

            self.project_data = self.project_data_instance.project_data_query(lab_reporting_batch_id, self.cursor)

        except Exception as ex:

            print(f"Error loading report data {ex}")

            raise

        finally:

            self.close_connection()
            
            project_data_dict = row_to_dicts(self.project_data, [
                    "Address_1",
                    "City",
                    "Client",
                    "ClientProjectNumber",
                    "Contact",
                    "LabReceiptDate",
                    "LabReportingBatchID",
                    "Phone",
                    "Postal_Code",
                    "ProjectLocation",
                    "ProjectName",
                    "State_Prov"
            ])
        
            samples_data_dict = row_to_dicts(self.first_data, [
                
                'itemID','labSampleId', 'clientSampleId', 'collectedDateTime',
                'sampleMatrix', 'analysisRequested'
            ])
            
            samples_tw = build_samples_structure(self.analytical_results)
            
            
            
            qc_data_dict = row_to_dicts(self.quality_controls, [
            'ClientSampleID',
            'LabSampleID',
            'DateCollected',
            'DatePrepared',
            'DateAnalyzed',
            'Sampler',
            'MatrixID',
            'AnalyteName',
            'Result',
            'QCSpikeAdded',
            'LabQualifiers',              
            'ResultUnits',
            'Dilution',
            'DetectionLimit',
            'ReportingLimit',
            'LabAnalysisRefMethodID',
            'Analyst',
            'MethodBatchID',
            'Notes',
            'RelativePercentDifference', 
            'PercentRecovery',
            'src',
            'Limites'
        ])

        for qc in qc_data_dict:
            lab_qual = qc.get('LabQualifiers', '')
            if lab_qual and lab_qual.strip():
                qc['Result'] = f"{qc['Result']}{lab_qual}"
            
            qc.pop('LabQualifiers', None)  


            return {
                
                'project_data': project_data_dict,
                'samples_data': samples_data_dict,
                'samples_tw': samples_tw,
                'quality_controls': qc_data_dict
                
            }





