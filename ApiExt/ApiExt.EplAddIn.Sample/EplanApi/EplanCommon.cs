using Eplan.EplApi.DataModel;
using System;
using System.Linq;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanCommon
    {
        #region Public Methods

        public bool IsProjectOpen(string elkFullPath)
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();

                return projectManager.OpenProjects.Any(p => p.ProjectLinkFilePath.Equals(elkFullPath, StringComparison.OrdinalIgnoreCase)); 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "IsProjectOpen", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        #endregion
    }
}
