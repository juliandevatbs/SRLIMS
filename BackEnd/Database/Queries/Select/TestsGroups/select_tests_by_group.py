from BackEnd.Database.General.get_connection import DatabaseConnection


class SelectTestsByGroup:
    """
    Query para obtener todos los analitos (Tests) de un Test Group específico
    Estos datos se usarán como plantilla para crear Sample_Tests
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

    def get_tests_by_group(self, test_group_name, matrix_id=None):
        """
        Obtiene todos los Tests (analitos) de un Test Group
        
        Args:
            test_group_name: Nombre del grupo (ej: "EPA 8260 (BTEX+MTBE) SOILS")
            matrix_id: Opcional, filtrar por matriz (ej: "SOILS")
        
        Returns:
            list: Lista de diccionarios con los datos de cada analito
        """
        
        # SOLO seleccionar columnas que EXISTEN en la tabla Tests
        query = """
            SELECT 
                t.Test_Group,
                t.LabAnalysisRefMethodID,
                t.ClientAnalyteID,
                t.AnalyteName,
                t.ResultUnits,
                t.DetectionLimit,
                t.ReportingLimit,
                t.Low_Limit,
                t.High_Limit,
                t.AnalyteType,
                t.Dilution,
                t.MatrixID,
                t.Prep_Hold,
                t.Test_Hold,
                t.Work_Area,
                t.MS_RPD,
                t.MS_Limit_Low,
                t.MS_Limit_Upper,
                t.LCS_RPD,
                t.LCS_Limit_Low,
                t.LCS_LimitUpper,
                t.TagMb,
                t.tagLcs,
                t.TagLcsd,
                t.TagMs,
                t.TagLabDup,
                t.TagSurr,
                t.TagMdlPql,
                t.[Order]
            FROM Tests t
            WHERE t.Test_Group = ?
        """
        
        params = [test_group_name]
        
        # Si se proporciona matriz, filtrar por ella
        if matrix_id:
            query += " AND t.MatrixID = ?"
            params.append(matrix_id)
        
        # Ordenar por el campo Order para mantener la secuencia correcta
        query += " ORDER BY t.[Order]"
        
        self.cursor.execute(query, params)
        
        # Obtener nombres de columnas
        columns = [column[0] for column in self.cursor.description]
        
        # Convertir resultados a lista de diccionarios
        results = []
        for row in self.cursor.fetchall():
            results.append(dict(zip(columns, row)))
        
        return results

    def get_available_test_groups(self, matrix_id=None):
        """
        Obtiene la lista de Test Groups disponibles
        
        Args:
            matrix_id: Opcional, filtrar por matriz
            
        Returns:
            list: Lista de nombres de Test Groups
        """
        
        query = """
            SELECT DISTINCT 
                tg.Test_Group,
                tg.Test_Group_Title,
                tg.MatrixID
            FROM Test_Groups tg
            WHERE tg.Active = 1
        """
        
        if matrix_id:
            query += " AND tg.MatrixID = ?"
            self.cursor.execute(query, (matrix_id,))
        else:
            self.cursor.execute(query)
        
        results = self.cursor.fetchall()
        
        # Retornar solo los nombres de grupos
        return [row[0] for row in results]