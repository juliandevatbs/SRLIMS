from BackEnd.Database.General.get_connection import DatabaseConnection


class InsertSampleTestsFromGroup:
    """
    Inserta múltiples Sample_Tests basándose en los analitos de un Test Group
    """

    def __init__(self):
        self.db = DatabaseConnection()
        self.cursor = None
        self.conn = None

    def load_connection(self):
        """Establecer conexión a la base de datos"""
        self.conn = self.db.get_conn()
        self.cursor = self.conn.cursor()

    def close_connection(self):
        """Cerrar la conexión"""
        if self.cursor:
            self.cursor.close()
        if self.conn:
            self.conn.close()

    def insert_tests_from_group(self, work_order, lab_sample_id, test_group_data):
        """
        Inserta múltiples Sample_Tests basándose en los datos de un Test Group
        
        Args:
            work_order: Work Order number (ej: 2501011)
            lab_sample_id: Lab Sample ID (ej: "2501011-001")
            test_group_data: Lista de diccionarios con datos de Tests obtenidos de SelectTestsByGroup
            
        Returns:
            int: Número de registros insertados
        """
        
        if not test_group_data:
            return 0
        
        # Query de inserción - SOLO columnas que EXISTEN en Sample_Tests
        insert_query = """
            INSERT INTO Sample_Tests (
                LabReportingBatchID,
                LabSampleID,
                Test_Group,
                LabAnalysisRefMethodID,
                ClientAnalyteID,
                AnalyteName,
                Result,
                ResultUnits,
                DetectionLimit,
                ReportingLimit,
                Low_Limit,
                High_Limit,
                AnalyteType,
                Dilution,
                MatrixID,
                LabQualifiers,
                PercentMoisture,
                PercentRecovery,
                Prep_Hold,
                Test_Due,
                Work_Area,
                MS_RPD,
                MS_Limit_Low,
                MS_Limit_Upper,
                LCS_RPD,
                LCS_Limit_Low,
                LCS_LimitUpper,
                TagMb,
                tagLcs,
                TagLCSD,
                tagMs,
                TagLabDup,
                TagSurr,
                TagMdlPql,
                [Order]
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        
        inserted_count = 0
        
        try:
            for test_data in test_group_data:
                # Preparar valores para inserción
                values = (
                    work_order,                                    # LabReportingBatchID
                    lab_sample_id,                                 # LabSampleID
                    test_data.get('Test_Group'),                   # Test_Group
                    test_data.get('LabAnalysisRefMethodID'),       # LabAnalysisRefMethodID
                    test_data.get('ClientAnalyteID'),              # ClientAnalyteID
                    test_data.get('AnalyteName'),                  # AnalyteName
                    None,                                          # Result (NULL - se llenará después)
                    test_data.get('ResultUnits'),                  # ResultUnits
                    test_data.get('DetectionLimit'),               # DetectionLimit
                    test_data.get('ReportingLimit'),               # ReportingLimit
                    test_data.get('Low_Limit'),                    # Low_Limit
                    test_data.get('High_Limit'),                   # High_Limit
                    test_data.get('AnalyteType'),                  # AnalyteType (TRG, SURR, etc.)
                    test_data.get('Dilution', 1),                  # Dilution (default 1)
                    test_data.get('MatrixID'),                     # MatrixID
                    None,                                          # LabQualifiers (NULL - se llenará después)
                    None,                                          # PercentMoisture
                    None,                                          # PercentRecovery
                    test_data.get('Prep_Hold'),                    # Prep_Hold
                    test_data.get('Test_Hold'),                    # Test_Due (en Tests se llama Test_Hold)
                    test_data.get('Work_Area'),                    # Work_Area
                    test_data.get('MS_RPD'),                       # MS_RPD
                    test_data.get('MS_Limit_Low'),                 # MS_Limit_Low
                    test_data.get('MS_Limit_Upper'),               # MS_Limit_Upper
                    test_data.get('LCS_RPD'),                      # LCS_RPD
                    test_data.get('LCS_Limit_Low'),                # LCS_Limit_Low
                    test_data.get('LCS_LimitUpper'),               # LCS_LimitUpper
                    test_data.get('TagMb', 0),                     # TagMb
                    test_data.get('tagLcs', 0),                    # tagLcs (nota: minúscula en Sample_Tests)
                    test_data.get('TagLcsd', 0),                   # TagLCSD (nota: mayúsculas en Sample_Tests)
                    test_data.get('TagMs', 0),                     # tagMs (nota: minúscula en Sample_Tests)
                    test_data.get('TagLabDup', 0),                 # TagLabDup
                    test_data.get('TagSurr', 0),                   # TagSurr
                    test_data.get('TagMdlPql', 0),                 # TagMdlPql
                    test_data.get('Order')                         # Order
                )
                
                self.cursor.execute(insert_query, values)
                inserted_count += 1
            
            # Commit todas las inserciones
            self.conn.commit()
            
            return inserted_count
            
        except Exception as e:
            self.conn.rollback()
            print(f"Error inserting sample tests from group: {e}")
            import traceback
            traceback.print_exc()
            raise

    def get_sample_info(self, lab_sample_id):
        """
        Obtiene información del sample para validación
        
        Args:
            lab_sample_id: Lab Sample ID
            
        Returns:
            dict: Información del sample o None
        """
        
        query = """
            SELECT 
                LabReportingBatchID,
                LabSampleID,
                ClientSampleID,
                MatrixID,
                DateCollected,
                Sampler
            FROM Samples
            WHERE LabSampleID = ?
        """
        
        self.cursor.execute(query, (lab_sample_id,))
        row = self.cursor.fetchone()
        
        if row:
            columns = [column[0] for column in self.cursor.description]
            return dict(zip(columns, row))
        
        return None