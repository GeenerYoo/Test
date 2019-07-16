using ApiExt.EplAddIn.Sample.Abstracts;
using System;
using System.Data;

namespace ApiExt.EplAddIn.Sample.Repositories
{
    public class SampleRepository : IRepository
    {
        private readonly IDBAccessor _db;

        #region Constructor

        public SampleRepository(IDBAccessor dbAccessor)
        {
            Exception argumentNullException = null;
 
            if (dbAccessor == null) {
                argumentNullException = new ArgumentNullException("dbAccessor");

            throw argumentNullException;
            }

            this._db = dbAccessor;
        }

        #endregion

        #region Public Properties

        public string DataSource
        {
            get { return this._db.DataSource; }
        }

        #endregion

        public DataTable GetEmpList()
        {
            try
            {
                string commandText = @"SELECT *
                                         FROM EMP T";

                var resultTable = this._db.ExecuteSelect(commandText);
                resultTable.TableName = "SAMPLE_EMP";

                return resultTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable GetDeptList()
        {
            try
            {
                string commandText = @"SELECT *
                                         FROM DEPT T";

                var resultTable = this._db.ExecuteSelect(commandText);
                resultTable.TableName = "SAMPLE_DEPT";

                return resultTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
