using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanReport
    {
        public void GenerateReport(Project targetProject)
        {
            try
            {
                Reports reportEngine = new Reports();
                reportEngine.GenerateProject(targetProject);

                MessageBox.Show("프로젝트 리포트가 성공적으로 생성되었습니다.", "Report 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Report 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }
    }
}
