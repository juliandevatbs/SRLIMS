from BackEnd.Database.General.get_connection import DatabaseConnection


class InsertQualityControl:
    
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

    def get_sample_info(self, lab_sample_id):
            query = """
                SELECT 
                    LabReportingBatchID,
                    MatrixID,
                    LabID,
                    ItemID,
                    ClientSampleID,
                    DateCollected,
                    Temperature,
                    ShippingBatchID,
                    PercentMositure,
                    LabSampleID
                FROM Samples
                WHERE LabSampleID = ?
            """
            self.cursor.execute(query, (lab_sample_id,))
            result = self.cursor.fetchone()
            if not result:
                raise ValueError(f"Sample {lab_sample_id} not found")
            return result
    
    def get_next_sample_number(self, work_order_id):
        query = """
            SELECT MAX(CAST(SUBSTRING(LabSampleID, LEN(LabReportingBatchID) + 2, 3) AS INT))
            FROM Samples 
            WHERE LabReportingBatchID = ?
        """
        self.cursor.execute(query, (work_order_id,))
        result = self.cursor.fetchone()
        max_number = result[0] if result and result[0] else 0
        return max_number + 1
    
    def get_analytes_for_sample(self, lab_sample_id):
        query = """
            SELECT DISTINCT
                LabAnalysisRefMethodID,
                ClientAnalyteID,
                AnalyteName,
                ResultUnits,
                DetectionLimit,
                ReportingLimit,
                AnalyteType
            FROM Sample_Tests
            WHERE LabSampleID = ?
            AND (QCType IS NULL OR QCType = '')
        """
        self.cursor.execute(query, (lab_sample_id,))
        return self.cursor.fetchall()
    
    def create_method_blank(self, reference_lab_sample_id, prep_batch_id=None):
        try:
            sample_info = self.get_sample_info(reference_lab_sample_id)
            work_order_id = sample_info[0]
            matrix_id = sample_info[1] if sample_info[1] else None
            lab_id = sample_info[2] if sample_info[2] else None
            
            next_number = self.get_next_sample_number(work_order_id)
            
            lab_sample_id = f"{work_order_id}-{next_number:03d}"
            client_sample_id = f"{lab_sample_id} MB"
            
            insert_sample = """
                INSERT INTO Samples (
                    LabReportingBatchID,
                    LabSampleID,
                    ClientSampleID,
                    TagMB,
                    TagLCS,
                    TagLCSD,
                    TagMS,
                    TagLabDup,
                    TagParentSample,
                    QCSample,
                    QCType,
                    MatrixID,
                    LabID,
                    CollectionAgency,
                    Sampler,
                    PreservationIntact,
                    CustodyIntactSeal
                )
                VALUES (?, ?, ?, 1, 0, 0, 0, 0, 0, 1, ?, ?, ?, ?, ?, 1, 1)
            """
            
            self.cursor.execute(insert_sample, (
                work_order_id,
                lab_sample_id,
                client_sample_id,
                'MB',
                matrix_id,
                lab_id,
                'Lab QC',
                'Lab QC'
            ))
            
            self.conn.commit()
            
            self._create_qc_tests(lab_sample_id, 'MB', reference_lab_sample_id, prep_batch_id)
            
            return lab_sample_id
            
        except Exception as e:
            self.conn.rollback()
            raise Exception(f"Error creating Method Blank: {e}")
    
    def create_lcs_pair(self, reference_lab_sample_id, prep_batch_id=None):
        try:
            sample_info = self.get_sample_info(reference_lab_sample_id)
            work_order_id = sample_info[0]
            matrix_id = sample_info[1] if sample_info[1] else None
            lab_id = sample_info[2] if sample_info[2] else None
            
            next_number = self.get_next_sample_number(work_order_id)
            lcs_number = next_number
            lcsd_number = next_number + 1
            
            lcs_lab_sample_id = f"{work_order_id}-{lcs_number:03d}"
            lcs_client_sample_id = f"{lcs_lab_sample_id} LCS"
            
            insert_sample = """
                INSERT INTO Samples (
                    LabReportingBatchID, 
                    LabSampleID, 
                    ClientSampleID,
                    TagMB, 
                    TagLCS, 
                    TagLCSD, 
                    TagMS, 
                    TagLabDup, 
                    TagParentSample,
                    QCSample, 
                    QCType, 
                    MatrixID, 
                    LabID,
                    CollectionAgency, 
                    Sampler, 
                    PreservationIntact, 
                    CustodyIntactSeal
                )
                VALUES (?, ?, ?, 0, 1, 0, 0, 0, 0, 1, ?, ?, ?, ?, ?, 1, 1)
            """
            
            self.cursor.execute(insert_sample, (
                work_order_id, 
                lcs_lab_sample_id, 
                lcs_client_sample_id,
                'LCS', 
                matrix_id, 
                lab_id, 
                'Lab QC', 
                'Lab QC'
            ))
            
            lcsd_lab_sample_id = f"{work_order_id}-{lcsd_number:03d}"
            lcsd_client_sample_id = f"{lcsd_lab_sample_id} LCSD"
            
            self.cursor.execute("""
                INSERT INTO Samples (
                    LabReportingBatchID, 
                    LabSampleID, 
                    ClientSampleID,
                    TagMB, 
                    TagLCS, 
                    TagLCSD, 
                    TagMS, 
                    TagLabDup, 
                    TagParentSample,
                    QCSample, 
                    QCType, 
                    MatrixID, 
                    LabID,
                    CollectionAgency, 
                    Sampler, 
                    PreservationIntact, 
                    CustodyIntactSeal
                )
                VALUES (?, ?, ?, 0, 0, 1, 0, 0, 1, 1, ?, ?, ?, ?, ?, 1, 1)
            """, (
                work_order_id, 
                lcsd_lab_sample_id, 
                lcsd_client_sample_id,
                'LCSD', 
                matrix_id, 
                lab_id, 
                'Lab QC', 
                'Lab QC'
            ))
            
            self.conn.commit()
            
            self._create_qc_tests(lcs_lab_sample_id, 'LCS', reference_lab_sample_id, prep_batch_id)
            self._create_qc_tests(lcsd_lab_sample_id, 'LCSD', reference_lab_sample_id, prep_batch_id, lcs_lab_sample_id)
            
            return lcs_lab_sample_id, lcsd_lab_sample_id
            
        except Exception as e:
            self.conn.rollback()
            raise Exception(f"Error creating LCS pair: {e}")
    
    def create_matrix_spike_pair(self, parent_lab_sample_id, prep_batch_id=None):
        try:
            parent_info = self.get_sample_info(parent_lab_sample_id)
            work_order_id = parent_info[0]
            matrix_id = parent_info[1] if parent_info[1] else None
            lab_id = parent_info[2] if parent_info[2] else None
            date_collected = parent_info[5] if parent_info[5] else None
            temperature = parent_info[6] if parent_info[6] else None
            shipping_batch = parent_info[7] if parent_info[7] else None
            percent_moisture = parent_info[8] if parent_info[8] else None
            
            next_number = self.get_next_sample_number(work_order_id)
            ms_number = next_number
            msd_number = next_number + 1
            
            ms_lab_sample_id = f"{work_order_id}-{ms_number:03d}"
            ms_client_sample_id = f"{ms_lab_sample_id} MS"
            
            insert_sample = """
                INSERT INTO Samples (
                    LabReportingBatchID, 
                    LabSampleID, 
                    ClientSampleID,
                    TagMB, 
                    TagLCS, 
                    TagLCSD, 
                    TagMS, 
                    TagLabDup, 
                    TagParentSample,
                    QCSample, 
                    QCType, 
                    MatrixID, 
                    LabID,
                    CollectionAgency, 
                    Sampler, 
                    DateCollected, 
                    Temperature, 
                    ShippingBatchID,
                    PercentMositure, 
                    PreservationIntact, 
                    CustodyIntactSeal
                )
                VALUES (?, ?, ?, 0, 0, 0, 1, 0, 1, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 1)
            """
            
            self.cursor.execute(insert_sample, (
                work_order_id, 
                ms_lab_sample_id, 
                ms_client_sample_id,
                'MS', 
                matrix_id, 
                lab_id, 
                'Lab QC', 
                'Lab QC',
                date_collected, 
                temperature, 
                shipping_batch, 
                percent_moisture
            ))
            
            msd_lab_sample_id = f"{work_order_id}-{msd_number:03d}"
            msd_client_sample_id = f"{msd_lab_sample_id} MSD"
            
            self.cursor.execute("""
                INSERT INTO Samples (
                    LabReportingBatchID, 
                    LabSampleID, 
                    ClientSampleID,
                    TagMB, 
                    TagLCS, 
                    TagLCSD, 
                    TagMS, 
                    TagLabDup, 
                    TagParentSample,
                    QCSample, 
                    QCType, 
                    MatrixID, 
                    LabID,
                    CollectionAgency, 
                    Sampler, 
                    DateCollected, 
                    Temperature, 
                    ShippingBatchID,
                    PercentMositure, 
                    PreservationIntact, 
                    CustodyIntactSeal
                )
                VALUES (?, ?, ?, 0, 0, 0, 0, 1, 1, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 1)
            """, (
                work_order_id, 
                msd_lab_sample_id, 
                msd_client_sample_id,
                'MSD', 
                matrix_id, 
                lab_id, 
                'Lab QC', 
                'Lab QC',
                date_collected, 
                temperature, 
                shipping_batch, 
                percent_moisture
            ))
            
            self.conn.commit()
            
            self._create_qc_tests(ms_lab_sample_id, 'MS', parent_lab_sample_id, prep_batch_id, parent_lab_sample_id)
            self._create_qc_tests(msd_lab_sample_id, 'MSD', parent_lab_sample_id, prep_batch_id, parent_lab_sample_id, ms_lab_sample_id)
            
            return ms_lab_sample_id, msd_lab_sample_id
            
        except Exception as e:
            self.conn.rollback()
            raise Exception(f"Error creating Matrix Spike pair: {e}")
    
    def _create_qc_tests(self, qc_lab_sample_id, qc_type, reference_lab_sample_id, 
                        prep_batch_id=None, parent_sample_id=None, parent_qc_id=None):
        try:
            analytes = self.get_analytes_for_sample(reference_lab_sample_id)
            
            if not analytes:
                raise ValueError(f"No analytes found for reference sample {reference_lab_sample_id}")
            
            sample_info = self.get_sample_info(qc_lab_sample_id)
            work_order_id = sample_info[0]
            qc_client_sample_id = sample_info[4]
            
            parent_item_id = None
            if parent_sample_id:
                parent_info = self.get_sample_info(parent_sample_id)
                parent_item_id = parent_info[3]
            
            for analyte in analytes:
                insert_test = """
                    INSERT INTO Sample_Tests (
                        ItemID,
                        ClientSampleID,
                        LabSampleID,
                        LabAnalysisRefMethodID,
                        ClientAnalyteID,
                        AnalyteName,
                        ResultUnits,
                        DetectionLimit,
                        ReportingLimit,
                        AnalyteType,
                        QCType,
                        LabReportingBatchID,
                        PrepBatchID,
                        TagParentSample
                    )
                    SELECT 
                        (SELECT ItemID FROM Samples WHERE LabSampleID = ?),
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?,
                        ?
                """
                
                self.cursor.execute(insert_test, (
                    qc_lab_sample_id,
                    qc_client_sample_id,
                    qc_lab_sample_id,
                    analyte[0],
                    analyte[1],
                    analyte[2],
                    analyte[3],
                    analyte[4],
                    analyte[5],
                    analyte[6],
                    qc_type,
                    work_order_id,
                    prep_batch_id,
                    parent_item_id
                ))
            
            self.conn.commit()
            print(f"✓ Created {len(analytes)} tests for {qc_type} {qc_lab_sample_id}")
            
        except Exception as e:
            self.conn.rollback()
            raise Exception(f"Error creating QC tests: {e}")