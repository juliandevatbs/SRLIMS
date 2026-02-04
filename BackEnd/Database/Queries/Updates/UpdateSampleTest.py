from BackEnd.Database.General.get_connection import DatabaseConnection


class UpdateSampleTest:
    """
    Clase para actualizar registros en la tabla SampleTests
    Usa SampleTestsID como clave única
    """
    def __init__(self):
        self.instance_db = DatabaseConnection()
        self.conn = None
        self.cursor = None
        
    def load_connection(self):
        """Carga la conexión a la base de datos"""
        self.conn = DatabaseConnection.get_conn(self.instance_db)
        self.cursor = self.conn.cursor()
    
    def close_connection(self):
        """Cierra la conexión a la base de datos"""
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()
    
    def update_field(self, sample_tests_id, field_name, new_value):
        """
        Actualiza un campo específico de un registro
        
        Args:
            sample_tests_id (int): SampleTestsID del registro a actualizar
            field_name (str): Nombre del campo a actualizar
            new_value: Nuevo valor del campo
            
        Returns:
            bool: True si la actualización fue exitosa, False en caso contrario
        """
        try:
            # Validar que el campo sea válido
            valid_fields = [
                'ItemID', 'ClientSampleID', 'LabAnalysisRefMethodID', 
                'LabSampleID', 'LabID', 'ClientAnalyteID', 'AnalyteName',
                'Result', 'Error', 'ResultUnits', 'LabQualifiers',
                'DetectionLimit', 'AnalyteType', 'Dilution', 'PercentMoisture',
                'PercentRecovery', 'RelativePercentDifference', 'ReportingLimit',
                'ProjectNumber', 'ProjectName', 'DateCollected', 'MatrixID',
                'QCType', 'ShippingBatchID', 'Temperature', 'PreparationType',
                'AnalysisType', 'ReportableResult', 'DatePrepared', 'DateAnalyzed',
                'TotalOrDissolved', 'PrepBatchID', 'MethodBatchID',
                'PreservationIntact', 'QCSpikeAdded', 'ResultComments',
                'LabReportingBatchID', 'GroupLongName', 'Low_Limit', 'High_Limit',
                'Limits', 'Test_Group', 'Prep_Due', 'Test_Due', 'Work_Area',
                'Format_Type', 'Test_Cost', 'Result_Text', 'Analyst',
                'Group_Cost', 'Include', 'Cont_ID', 'Result_ID', 'Order',
                'SamplingPersonnel', 'CollectionAgency', 'CustodyIntactSeal',
                'ReceiptComments', 'CoolerNumber', 'LocationCode', 'LabReceiptDate',
                'AdaptMatrixID', 'ProgramType', 'CollectionMethod', 'SamplingDepth',
                'LoctionCode', 'ProjectLocation', 'ProjectCode', 'Test_ID',
                'TotalContainers', 'Prep_Hold', 'Prep_Hold_Units',
                'HoldDectionLimit', 'HoldReportingLimit', 'TagMb', 'tagLcs',
                'TagLCSD', 'tagMs', 'TagLabDup', 'TagSurr', 'TagMdlPql',
                'TagParentSample', 'LibraryFlag', 'SampleWgt', 'Run_Code',
                'Field_Filtered', 'Lab_Filtered', 'Iced', 'Preservative',
                'Notes', 'RP1', 'RP2', 'RP3', 'ReportingBasis', 'HoldCal',
                'QCSpikeLow', 'QCSpikeHigh', 'QCSpikeAmount', 'SortSRLReport',
                'SortSRLQC', 'SortOthers', 'Parent_Sample', 'Sample_Receipt_Date'
            ]
            
            if field_name not in valid_fields:
                print(f"Campo inválido: {field_name}")
                return False
            
            # Manejar valores None para campos numéricos
            if new_value == '' or new_value == 'None':
                new_value = None
            
            # Construir la query de forma segura
            query = f"""
                UPDATE Sample_Tests 
                SET [{field_name}] = ? 
                WHERE SampleTestsID = ?
            """
            
            # Ejecutar la query
            self.load_connection()
            self.cursor.execute(query, (new_value, sample_tests_id))
            self.conn.commit()
            
            rows_affected = self.cursor.rowcount
            
            if rows_affected > 0:
                print(f"[DB] Updated SampleTest ID={sample_tests_id}, {field_name}={new_value}")
                return True
            else:
                print(f"[DB] No rows updated for SampleTestsID={sample_tests_id}")
                return False
            
        except Exception as e:
            print(f"Error updating SampleTest: {e}")
            if self.conn:
                self.conn.rollback()
            return False
            
        finally:
            self.close_connection()
    
    def update_multiple_fields(self, sample_tests_id, updates_dict):
        """
        Actualiza múltiples campos de un registro en una sola transacción
        
        Args:
            sample_tests_id (int): SampleTestsID del registro a actualizar
            updates_dict (dict): Diccionario {campo: nuevo_valor}
            
        Returns:
            bool: True si la actualización fue exitosa
        """
        try:
            if not updates_dict:
                return False
            
            # Construir la query con múltiples SET
            set_clauses = []
            values = []
            
            for field_name, new_value in updates_dict.items():
                set_clauses.append(f"[{field_name}] = ?")
                values.append(new_value)
            
            values.append(sample_tests_id)  # Para el WHERE
            
            query = f"""
                UPDATE Sample_Tests 
                SET {', '.join(set_clauses)}
                WHERE SampleTestsID = ?
            """
            
            self.load_connection()
            self.cursor.execute(query, tuple(values))
            self.conn.commit()
            
            print(f"[DB] Updated SampleTest ID={sample_tests_id}, fields={list(updates_dict.keys())}")
            return True
            
        except Exception as e:
            print(f"Error updating multiple fields: {e}")
            if self.conn:
                self.conn.rollback()
            return False
            
        finally:
            self.close_connection()