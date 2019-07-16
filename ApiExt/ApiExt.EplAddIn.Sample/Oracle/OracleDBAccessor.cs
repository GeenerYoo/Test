using ApiExt.EplAddIn.Sample.Abstracts;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace ApiExt.EplAddIn.Sample.Oracle
{
    public class OracleDBAccessor : IDBAccessor
    {
        #region Constructor

        private IConfigurationProvider _configurationProvider;
        private OracleConnection _connection;
        private string _connectionName;

        /// <summary>
        /// Constructor which takes ConfigurationProvider and uses the "OracleConnection"
        /// </summary>
        /// <param name="configurationProvider"></param>
        public OracleDBAccessor(IConfigurationProvider configurationProvider): this("OracleConnection", configurationProvider)
        {
        }

        /// <summary>
        /// Constructor which takes the connection name and ConfigurationProvider
        /// </summary>
        /// <param name="connectionName"></param>
        /// <param name="configurationProvider"></param>
        public OracleDBAccessor(string connectionName, IConfigurationProvider configurationProvider)
        {
            if (string.IsNullOrWhiteSpace(connectionName))
                throw new ArgumentNullException("connectionName");

            if (configurationProvider == null)
                throw new ArgumentNullException("configurationProvider");

            this._configurationProvider = configurationProvider;
            this._connectionName = connectionName;

            this._connection = new OracleConnection(_configurationProvider.GetConnectionString(_connectionName));
        }

        #endregion

        #region IDBAccessor Support

        public string DataSource
        {
            get { return this._connection.DataSource; }
        }

        public int ExecuteNonQuery(string commandText, Dictionary<string, object> parameters)
        {
            int result = 0;

            if (string.IsNullOrWhiteSpace(commandText))
                throw new ArgumentException("Command text cannot be null or empty.", "commandText");

            try
            {
                EnsureConnectionOpen();
                using (OracleCommand command = CreateCommand(commandText, parameters))
                {
                    result = command.ExecuteNonQuery();
                }
            }
            finally
            {
                EnsureConnectionClosed();
            }

            return result;
        }

        public object ExecuteScalar(string commandText, Dictionary<string, object> parameters)
        {
            object result = null;

            if (string.IsNullOrWhiteSpace(commandText))
                throw new ArgumentException("Command text cannot be null or empty.", "commandText");

            try
            {
                EnsureConnectionOpen();
                using (OracleCommand command = CreateCommand(commandText, parameters))
                {
                    result = command.ExecuteScalar();
                }
            }
            finally
            {
                EnsureConnectionClosed();
            }

            return result;
        }

        public string GetStrValue(string commandText, Dictionary<string, object> parameters)
        {
            object value = ExecuteScalar(commandText, parameters);

            return value == null ? null : value.ToString();
        }

        public DataTable ExecuteSelect(string commandText, Dictionary<string, object> parameters)
        {
            DataTable rows = new DataTable();
            OracleDataAdapter dataAdapter = null;

            if (string.IsNullOrWhiteSpace(commandText))
                throw new ArgumentException("Command text cannot be null or empty.", "commandText");

            try
            {
                EnsureConnectionOpen();
                using (OracleCommand command = CreateCommand(commandText, parameters))
                {
                    dataAdapter = new OracleDataAdapter(command);
                    dataAdapter.Fill(rows);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (dataAdapter != null)
                    dataAdapter.Dispose();

                EnsureConnectionClosed();
            }

            return rows;
        }

        #endregion

        #region IDisposable Support

        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (_connection != null)
                    {
                        _connection.Dispose();
                        _connection = null;
                    }
                }

                disposedValue = true;
            }
        }

        // This code added to correctly implement the disposable pattern.
        public void Dispose()
        {
            // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
            Dispose(true);
            // TODO: uncomment the following line if the finalizer is overridden above.
            // GC.SuppressFinalize(this);
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Opens a connection if not open
        /// </summary>
        private void EnsureConnectionOpen()
        {
            var retries = 3;

            if (_connection.State == ConnectionState.Open)
                return;
            else
                while (retries >= 0 && _connection.State != ConnectionState.Open)
                {
                    _connection.Open();
                    retries--;
                    Thread.Sleep(30);
                }
        }

        /// <summary>
        /// Closes a connection if open
        /// </summary>
        private void EnsureConnectionClosed()
        {
            if (_connection.State == ConnectionState.Open)
                _connection.Close();
        }

        /// <summary>
        /// Creates a OracleCommand with the given parameters
        /// </summary>
        /// <param name="commandText">The Oracle query to execute</param>
        /// <param name="parameters">Parameters to pass to the Oracle query</param>
        /// <returns></returns>
        private OracleCommand CreateCommand(string commandText, Dictionary<string, object> parameters)
        {
            OracleCommand command = _connection.CreateCommand();
            command.CommandType = CommandType.Text;
            command.BindByName = true;
            command.Parameters.Clear();

            command.CommandText = commandText;
            AddParameters(command, parameters);

            return command;
        }

        /// <summary>
        /// Adds the parameters to a Oracle command
        /// </summary>
        /// <param name="commandText">The Oracle query to execute</param>
        /// <param name="parameters">Parameters to pass to the Oracle query</param>
        private void AddParameters(OracleCommand command, Dictionary<string, object> parameters)
        {
            if (parameters == null)
                return;

            foreach (var param in parameters)
            {
                var parameter = command.CreateParameter();
                parameter.ParameterName = param.Key;
                parameter.Value = param.Value ?? DBNull.Value;
                command.Parameters.Add(parameter);
            }
        }

        #endregion
    }
}
