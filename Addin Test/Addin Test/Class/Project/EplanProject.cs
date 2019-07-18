using Eplan.EplApi.DataModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Addin_Test.Class.Project
{
  public  class EplanProject
    {

        public void CurrentProjectFilePath()
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                Eplan.EplApi.DataModel.Project  project = projectManager.CurrentProject;//Project라는 이름이 중첩되는 문제가 있어서 이런 Namespace구조는 피해야 할듯
                                                                                        //일단 Eplan.EplApi.DataModel.Project 으로 테스트 진행


                if (project == null)
                    MessageBox.Show("현재 프로젝트가 없습니다.", "현재 프로젝트", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show(string.Format("현재 프로젝트 = [{0}]", project.ProjectLinkFilePath), "현재 프로젝트", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "현재 프로젝트 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }
    }
}
