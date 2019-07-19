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
        Eplan.EplApi.DataModel.Project project = null;

        public void CurrentProjectFilePath()
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                Eplan.EplApi.DataModel.Project project = projectManager.CurrentProject;//Project라는 이름이 중첩되는 문제가 있어서 이런 Namespace구조는 피해야 할듯(앞에 약자를 붙이는게 좋을까?)
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

        public void SetProjectProperties()
        {
            // 프로젝트 설정하기
            try
            {
                
                //속성을 찾으려면 메뉴얼 필수! 어떤게 어떤건지 번호로 표현되어 있어 확인이 어렵다~
                //그리고 현재 이렇게 단순 텍스트로 설정하면 에러나고 값적용 안됨!
                //MultiLangString로 다시반번 가공된 텍스트값 적용한걸로 예제소스에 나와 있는데 아직 이유는 파악못함.
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD04] = "어떤속성인가 확인하자";
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                throw ex;
            }




        }



    }
}
