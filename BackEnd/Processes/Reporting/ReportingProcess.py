from BackEnd.Database.Queries.Insert.Reports.insert_historical import InsertHistorical
from BackEnd.Database.Queries.Select.Reports.MainCommand import MainCommand


class ReportingProcess:


    def __init__(self):

        self.getter_data_instance = MainCommand()


        self.initial_data = None

        self.analytic_data = None

        self.quality_data = None

    def get_data(self):

        try:

            self.getter_data_instance.load_connection()

            initial_data, analytical_data, quality_data = self.getter_data_instance.caller(self.lab_reporting_batch_id)

        except Exception as e:

            print(f"Error getting the report data {e}")

        finally:

            self.getter_data_instance.close_connection()


        return self.initial_data, self.analytic_data, self.quality_data

    def process_data(self):


        


        return
    
    def insert_historical_record(self, historical_data):

        try:

            instance_db = InsertHistorical()

            instance_db.load_connection()

            instance_db.insert_historical_record(historical_data)

        except Exception as e:

            print(f"Error inserting historical record: {e}")

        finally:

            instance_db.close_conn()









