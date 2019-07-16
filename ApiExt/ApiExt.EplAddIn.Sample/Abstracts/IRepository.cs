using System.Data;

namespace ApiExt.EplAddIn.Sample.Abstracts
{
    public interface IRepository
    {
        #region Public Properties

        string DataSource { get; }

        #endregion

        #region Query Data

        DataTable GetEmpList();

        DataTable GetDeptList();

        #endregion
    }

}
