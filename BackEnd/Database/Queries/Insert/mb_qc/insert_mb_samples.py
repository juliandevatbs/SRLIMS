from BackEnd.Database.General.get_connection import DatabaseConnection


def insert_mb_samples(data_to_create: dict) -> bool:
    connection = None
    cursor = None
    
    try:
        instance_db = DatabaseConnection()
        connection = DatabaseConnection.get_conn(instance_db)
        cursor = connection.cursor()
        
        query = """
        INSERT INTO Samples (
            ItemID, LabReportingBatchID, LabSampleID, ClientSampleID,
            ResultComments, Temperature, ShippingBatchID,
            CollectMethod, MatrixID, DateCollected, Sampler,
            TotalContainers, CoolerNumber, PreservationIntact,
            CollectionAgency, CustodyIntactSeal, AdaptMatrixID,
            ProgramType, CollectionMethod, SamplingDepth, LocationCode,
            ProjectNumber, LabID, 
            QCType, TagMB, QCSample,  -- ← AGREGADOS
            PercentMositure, DateAnalyzed
        )
        SELECT 
            (SELECT MAX(ItemID) + 1 FROM Samples 
             WHERE LabReportingBatchID = ?),
            s.LabReportingBatchID, 
            ?,  -- inc_lab_sample_id
            ?,  -- client_sample_id
            s.ResultComments, s.Temperature, s.ShippingBatchID,
            s.CollectMethod, s.MatrixID, s.DateCollected, 
            'Lab QC',  -- Sampler
            s.TotalContainers, s.CoolerNumber, s.PreservationIntact,
            s.CollectionAgency, s.CustodyIntactSeal, s.AdaptMatrixID,
            s.ProgramType, s.CollectionMethod, s.SamplingDepth, s.LocationCode,
            s.ProjectNumber, s.LabID,
            'MB',  -- ← QCType
            1,     -- ← TagMB
            1,     -- ← QCSample
            s.PercentMositure, s.DateAnalyzed
        FROM Samples s
        WHERE s.LabSampleID = ?
        """
        
        cursor.execute(query, (
            data_to_create["work_order"],
            data_to_create["inc_lab_sample_id"],
            data_to_create["client_sample_id"],
            data_to_create["lab_sample_id_orig"]
        ))
        
        connection.commit()
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        if connection:
            connection.rollback()
        return False
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()