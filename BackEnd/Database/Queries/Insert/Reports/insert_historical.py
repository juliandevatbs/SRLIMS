from BackEnd.Database.General.get_connection import DatabaseConnection


class InsertHistorical:
    
    
    
    def __init__(self):
        
        self.instance_db = DatabaseConnection()
        
        self.conn = None
        
        self.cursor = None
        
    def load_connection(self):
        
        self.conn = DatabaseConnection.get_conn(self.instance_db)
        
        self.cursor = self.conn.cursor()
    
    def close_conn(self):
        
        if self.cursor:
            
            self.cursor.close()
        
        if self.conn:
            
            self.conn.close()
       
    
    def insert_historical_record(self,  historical_data):
        
        
        
        query = """
        
        
                INSERT INTO dbo.Historical_Reports 
                    (
                        
                        created_by,
                        work_order
                        
                    )
                
                VALUES (
                    
                    
                    ?,
                    ?
                )
            
                
                
                """
                
        created_by = historical_data.get('created_by')
        work_order = historical_data.get('work_order')
        
        try:
        
            self.cursor.execute(query, (created_by, work_order))
            self.conn.commit()
        
        except Exception as e:
            
            self.conn.rollback()
            raise e
        