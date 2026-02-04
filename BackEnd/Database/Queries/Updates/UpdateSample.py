from BackEnd.Database.General.get_connection import DatabaseConnection


class UpdateSample:
    """
    Clase para actualizar registros en la tabla Samples (Sample_Login)
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
    
    def update_field(self, item_id, field_name, new_value):
        """
        Actualiza un campo específico de un registro
        
        Args:
            item_id (int): LabSampleID del registro a actualizar
            field_name (str): Nombre del campo a actualizar
            new_value: Nuevo valor del campo
            
        Returns:
            bool: True si la actualización fue exitosa, False en caso contrario
        """
        try:
            # Validar que el campo sea válido (seguridad contra SQL injection)
            valid_fields = [
                'LabReportingBatchID', 'LabSampleID', 'ClientSampleID', 
                'ResultComments', 'Temperature', 'ShippingBatchID', 
                'CollectMethod', 'MatrixID', 'DateCollected', 'Sampler',
                'TotalContainers', 'CoolerNumber', 'PreservationIntact',
                'SamplingPersonnel', 'CollectionAgency', 'CustodyIntactSeal',
                'ReceiptComments', 'LocationCode', 'LabReceiptDate',
                'AdaptMatrixID', 'ProgramType', 'CollectionMethod',
                'SamplingDepth', 'LoctionCode', 'ProjectNumber', 'LabID',
                'Include', 'TagParentSample', 'TagMB', 'tagLcs', 'TagLCSD',
                'TagMS', 'TagLabDup', 'PercentMositure', 'HoldCal',
                'Sample Wgt', 'DateAnalyzed', 'QCSample', 'Test_Group', 'QCType', 'ItemID'
            ]
            
            if field_name not in valid_fields:
                print(f"Campo inválido: {field_name}")
                return False
            
            # Construir la query de forma segura
            query = f"""
                UPDATE Samples
                SET [{field_name}] = ? 
                WHERE LabSampleID = ?
            """
            
            # Ejecutar la query
            self.load_connection()
            self.cursor.execute(query, (new_value, item_id))
            self.conn.commit()
            
            print(f"[DB] Updated Samples LabSampleID={item_id}, {field_name}={new_value}")
            return True
            
        except Exception as e:
            print(f"Error updating Sample: {e}")
            if self.conn:
                self.conn.rollback()
            return False
            
        finally:
            self.close_connection()
    
    def update_multiple_fields(self, item_id, updates_dict):
        """
        Actualiza múltiples campos de un registro en una sola transacción
        
        Args:
            item_id (int): LabSampleID del registro a actualizar
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
            
            values.append(item_id)  # Para el WHERE
            
            query = f"""
                UPDATE Samples 
                SET {', '.join(set_clauses)}
                WHERE LabSampleID = ?
            """
            
            self.load_connection()
            self.cursor.execute(query, tuple(values))
            self.conn.commit()
            
            print(f"[DB] Updated Samples LabSampleID={item_id}, fields={list(updates_dict.keys())}")
            return True
            
        except Exception as e:
            print(f"Error updating multiple fields: {e}")
            if self.conn:
                self.conn.rollback()
            return False
            
        finally:
            self.close_connection()