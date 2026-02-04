from pathlib import Path
import pandas as pd
from datetime import datetime
from BackEnd.Database.General.get_connection import DatabaseConnection
from BackEnd.Database.Queries.Select.Adapt.select_adapt_data import SelectAdaptData


class GenerateAdaptExport:
    
    def __init__(self, wo, project_number, project_name, date_collected, collection_agency):
        self.wo = wo
        self.project_number = project_number
        self.project_name = project_name
        self.date_collected = date_collected
        self.collection_agency = collection_agency
        self.instance_db = DatabaseConnection()  
        self.conn = None
        self.cursor = None
        self.adapt_data = None
        self.epp_data = None
        
    def load_connection(self):
        self.conn = self.instance_db.get_conn() 
        self.cursor = self.conn.cursor()

    def close_connection(self):
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()
        
    def call_adapt_data(self):
        
        try:
            
            self.load_connection()
            
            instance_command = SelectAdaptData(self.cursor, self.wo)
            
            self.clean_result = instance_command.clear_temp_table()
            
            self.insert_temp = instance_command.insert_adapt_data()
            
            self.adapt_data = instance_command.select_adapt_data()
            
            self.epp_data = instance_command.select_epp_data()
            
            return True
        
        except Exception as e:
            print(f"Error getting adapt data: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            self.close_connection()
            
    def format_date(self, date_value):
        if date_value is None:
            return ""
        try:
            if isinstance(date_value, str):
                date_value = pd.to_datetime(date_value)
            return date_value.strftime("%m/%d/%y %H:%M")
        except:
            return str(date_value)
            
    def export_to_excel(self):
        if self.adapt_data and self.epp_data:
            try:
                if isinstance(self.date_collected, str):
                    date_obj = pd.to_datetime(self.date_collected)
                else:
                    date_obj = self.date_collected
                date_formatted = date_obj.strftime("%Y%m%d")
                
                base_dir = Path("C:/1. Techical/2008-2022  FDEP HW Format Reports-EDD Files/2024-2025")
                client_dir = base_dir / self.collection_agency / f"{self.wo} {self.project_name}"
                zip_dir = client_dir / f"{self.project_number}_{date_formatted}_PRzdd"
                
                client_dir.mkdir(parents=True, exist_ok=True)
                zip_dir.mkdir(exist_ok=True)
                
                real_columns = [
                    'ClientSampleID', 'LabAnalysisRefMethodID', 'LabSampleID', 
                    'LabID', 'ClientAnalyteID', 'AnalyteName', 'Result', 
                    'Error', 'ResultUnits', 'LabQualifiers', 'DetectionLimit', 
                    'AnalyteType', 'Dilution', 'PercentMoisture', 'PercentRecovery', 
                    'RelativePercentDifference', 'ReportingLimit', 'ProjectNumber', 
                    'ProjectName', 'DateCollected', 'MatrixID', 'QCType', 
                    'ShippingBatchID', 'Temperature', 'PreparationType', 
                    'AnalysisType', 'Reportable', 'DatePrepared', 'DateAnalyzed', 
                    'TotalOrDissolved', 'PrepBatchID', 'MethodBatchID', 
                    'Preservation', 'QCSpikeAdded', 'ResultComments', 
                    'LabReportingBatchID', 'GroupLongName', 'High_Limit', 
                    'Low_Limit', 'Order', 'Analyst', 'Limits'
                ]
                
                df_adapt = pd.DataFrame(self.adapt_data, columns=real_columns)
                
                if 'DatePrepared' in df_adapt.columns:
                    df_adapt['DatePrepared'] = df_adapt['DatePrepared'].apply(self.format_date)
                if 'DateAnalyzed' in df_adapt.columns:
                    df_adapt['DateAnalyzed'] = df_adapt['DateAnalyzed'].apply(self.format_date)
                
                missing_columns = [
                    'FQcType', 'FQcBatchEB', 'FQcBatchFB', 'FQcBatchTB', 'FQcBatchFD',
                    'Filename', 'DVModifiedConcentration', 'DVQualTemperature',
                    'DVQualPreservation', 'DVQualHTSamplingToAnalysis',
                    'DVQualHTSamplingToExtraction', 'DVQualHTExtractionToAnalysis',
                    'DVQualHoldingTime', 'DVQualMethodBlanks', 'DVQualSurrogateRecovery',
                    'DVQualMS', 'DVQualMSRecovery', 'DVQualMSRPD', 'DVQualLCS',
                    'DVQualLCSRecovery', 'DVQualLCSRPD', 'DVQualRepLimits',
                    'DVQualReportingLimits', 'DVQualFieldQC', 'DVQualFieldBlank',
                    'DVQualEquipmentBlank', 'DVQualTripBlank', 'DVQualFieldDuplicate',
                    'DVQualIC', 'DVQualInitialCalibrationRRF', 'DVQualInitialCalibrationRSD',
                    'DVQualInitialCalibrationCC', 'DVQualICV',
                    'DVQualInitialCalibrationVerificationRRF',
                    'DVQualInitialCalibrationVerificationPD', 'DVQualCCV',
                    'DVQualContinuingCalibrationVerificationRRF',
                    'DVQualContinuingCalibrationVerificationPD', 'DVQualOverall',
                    'TagLabSampleID', 'TagDetQual01', 'TagNonDetQual01', 'TagDetQual02',
                    'TagNonDetQual02', 'surDVQualDet', 'surDVQualNonDet',
                    'DVQualInstrumentPerformanceCheckRunBatch',
                    'DVQualInstrumentPerformanceCheckAnaBatch', 'DVQualIPC',
                    'DVQualLabDup', 'DVQualCode', 'FieldDupRPD', 'DVQualMergedQualifier',
                    'DVQualMergedResult', 'AdaptMatrixID', 'TagArchivedFqc',
                    'ReplicateRSD', 'DVQualReplicateRSD', 'DVQualFieldDupReplicateRSD',
                    'StoretCode', 'WACS_Facility', 'WACS_Well', 'Samp_Method',
                    'StoretCodeMatrix', 'ParentSampleID', 'RecordId', 
                    'Sampler', 'Sampling_Personnel', 'Collection_Agency'
                ]
                
                for col in missing_columns:
                    df_adapt[col] = ''
                
                adapt_file = client_dir / f"{self.project_number}_{date_formatted}_PRldd.xlsx"
                
                with pd.ExcelWriter(adapt_file, engine='openpyxl') as writer:
                    df_adapt.to_excel(writer, sheet_name='AdaptLabEDD', index=False)
                
                epp_columns = [
                    'WACS_Testsite_ID', 'WACS_Testsite_Name', 'WACS_Facility_ID', 
                    'WACS_Facility_Name', 'Sample_Type', 'Matrix', 
                    'Field_Measurement_Method', 'Field_Parameter_Name', 'Result', 
                    'Result_Units', 'Field_Parameter_Qualifier_Code', 'Field_Comments', 
                    'Sampler', 'CollectionAgency', 'Date_Sampled', 'Shipping_Batch_ID', 
                    'Well_Purged_Flag', 'WACS_Report_Type'
                ]
                
                df_epp = pd.DataFrame(self.epp_data, columns=epp_columns)
                
                epp_file = client_dir / f"{self.project_number}_{date_formatted}_PRfdd.xlsx"
                
                with pd.ExcelWriter(epp_file, engine='openpyxl') as writer:
                    df_epp.to_excel(writer, sheet_name="AdaptFieldEDD", index=False)
                
                import os
                import time
                
                os.startfile(str(adapt_file.absolute()))
                time.sleep(0.5)
                os.startfile(str(epp_file.absolute()))
                
                print(f"Archivos exportados exitosamente:")
                print(f"  Lab EDD: {adapt_file}")
                print(f"  Field EDD: {epp_file}")
                    
                return True
            except Exception as e:
                print(f"Error exporting adapt data to excel: {e}")
                import traceback
                traceback.print_exc()
                return False