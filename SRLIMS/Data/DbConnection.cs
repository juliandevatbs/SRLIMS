using Microsoft.Data.SqlClient;
using System.Data;
using System.Threading.Tasks;

namespace SRLIMS.Data
{
    public class DbConnection : IDisposable
    {
        private readonly SqlConnection _connection;

        public DbConnection()
        {
            var builder = new SqlConnectionStringBuilder
            {
                DataSource = "192.168.0.121",
                UserID = "SRLADMIN",
                Password = "FL4R1D42025*.*",
                InitialCatalog = "SRLSQL",
                TrustServerCertificate = true,
                ConnectTimeout = 30
            };
            _connection = new SqlConnection(builder.ConnectionString);
        }

        public async Task OpenAsync()
        {
            if (_connection.State != ConnectionState.Open)
            {
                await _connection.OpenAsync();
            }
        }

        public async Task<DataTable> ExecuteQueryAsync(string query)
        {
            var dataTable = new DataTable();

            using (var command = new SqlCommand(query, _connection))
            {
                using (var reader = await command.ExecuteReaderAsync())
                {
                    dataTable.Load(reader); 
                }
            }

            return dataTable;
        }

        public void Dispose()
        {
            if (_connection.State == ConnectionState.Open)
            {
                _connection.Close();
            }
            _connection.Dispose();
        }
    }
}