from BackEnd.Database.General.get_connection import DatabaseConnection


def insert_lcs_tests(data_to_create: dict) -> bool:
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
            QCType, tagLcs, QCSpikeAdded, Result, PercentRecovery
        )
        SELECT 
            (SELECT MAX(ItemID) + 1 FROM Sample_Tests 
             WHERE LabReportingBatchID = ?),
            ?,
            st.LabAnalysisRefMethodID,
            ?,
            st.LabID,
            st.AnalyteName,
            st.AnalyteType,
            st.ResultUnits,
            st.DetectionLimit,
            st.ReportingLimit,
            ?,
            st.GroupLongName,
            st.PreservationType,
            'LCS',  -- ← QCType
            1,      -- ← tagLcs
            CASE 
                WHEN st.AnalyteType = 'TRG' THEN 20
                WHEN st.AnalyteType = 'SURR' THEN 50
                ELSE NULL
            END,  -- ← QCSpikeAdded
            NULL,  -- Result
            NULL   -- PercentRecovery
        FROM Sample_Tests st
        WHERE st.LabSampleID = ?
        """
        
        cursor.execute(query, (
            data_to_create["work_order"],
            data_to_create["ClientSampleID"],
            data_to_create["LabSampleID"],
            data_to_create["LabReportingBatchID"],
            data_to_create["lab_sample_id_orig"]
        ))
        
        connection.commit()
        print(f"✓ LCS tests created with spike values")
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