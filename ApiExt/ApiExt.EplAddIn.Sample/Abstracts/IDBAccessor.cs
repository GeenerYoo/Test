using System;
using System.Collections.Generic;
using System.Data;

namespace ApiExt.EplAddIn.Sample.Abstracts
{
    public interface IDBAccessor : IDisposable
    {
        string DataSource { get; }
        int ExecuteNonQuery(string commandText, Dictionary<string, object> parameters = null);

        object ExecuteScalar(string commandText, Dictionary<string, object> parameters = null);
        string GetStrValue(string commandText, Dictionary<string, object> parameters = null);

        DataTable ExecuteSelect(string commandText, Dictionary<string, object> parameters = null);
    }
}
