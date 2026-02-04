from BackEnd.Database.General.get_connection import DatabaseConnection


class SelectHistoricals:
    
    
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
            
    def select_historical_reports(self):
        
        
        qry =   """
                
                
                SELECT 
	
                    HR.id,
                    HR.creation_date,
                    HR.created_by,
                    HR.work_order,
                    CL.Client,
                    SL.Email

                    
                FROM dbo.Historical_Reports HR
                JOIN dbo.Sample_Login SL ON HR.work_order = SL.LabReportingBatchID
                JOIN dbo.Clients CL ON SL.Client_ID = CL.Client_ID
                ORDER BY HR.creation_date DESC
                ;
                
                
                
                """
        
        try:
            
            self.cursor.execute(qry)
            
            results = self.cursor.fetchall()
            
            return results
        
        except Exception as e:
            
            raise e
        
        
        