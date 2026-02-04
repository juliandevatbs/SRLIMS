from BackEnd.Database.General.get_connection import DatabaseConnection


def insert_mb_tests(data_to_create: dict) -> bool:
    """
    COPIA TODOS los analitos de la muestra original como MB
    """
    connection = None
    cursor = None
    
    try:
        instance_db = DatabaseConnection()
        connection = DatabaseConnection.get_conn(instance_db)
        cursor = connection.cursor()
        
        query = """
        INSERT INTO Sample_Tests (
            ItemID, ClientSampleID, LabAnalysisRefMethodID, LabSampleID,
            LabID, AnalyteName, AnalyteType, ResultUnits, 
            DetectionLimit, ReportingLimit, LabReportingBatchID,
            GroupLongName, PreservationType,
            QCType, TagMb, QCSpikeAdded, Result, PercentRecovery
        )
        SELECT 
            (SELECT MAX(ItemID) + 1 FROM Sample_Tests 
             WHERE LabReportingBatchID = ?),
            ?,                          -- Nuevo ClientSampleID
            st.LabAnalysisRefMethodID,  -- Copiar
            ?,                          -- Nuevo LabSampleID
            st.LabID,
            st.AnalyteName,             -- ← Copia TODOS
            st.AnalyteType,
            st.ResultUnits,
            st.DetectionLimit,
            st.ReportingLimit,
            ?,                          -- LabReportingBatchID
            st.GroupLongName,
            st.PreservationType,
            'MB',   -- QCType
            1,      -- TagMb
            NULL,   -- QCSpikeAdded (MB no tiene)
            NULL,   -- Result (usuario lo llena)
            NULL    -- PercentRecovery (MB no calcula)
        FROM Sample_Tests st
        WHERE st.LabSampleID = ?  -- ← Muestra original
        """
        
        cursor.execute(query, (
            data_to_create["work_order"],
            data_to_create["ClientSampleID"],
            data_to_create["LabSampleID"],
            data_to_create["LabReportingBatchID"],
            data_to_create["lab_sample_id_orig"]
        ))
        
        connection.commit()
        print(f"✓ MB tests created (copied from {data_to_create['lab_sample_id_orig']})")
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