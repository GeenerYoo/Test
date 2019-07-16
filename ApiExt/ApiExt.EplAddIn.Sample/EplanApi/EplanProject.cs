using ApiExt.EplAddIn.Sample.Extensions;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Specialized;
using System.IO;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanProject
    {
        public void RestoreProject(string zw1FullPath, string elkFullPath)
        {
            try
            {
                Restore restoreZw1 = new Restore();
                StringCollection targetZw1Path = new StringCollection();
                targetZw1Path.Add(zw1FullPath);

                string restorePath = Path.GetDirectoryName(elkFullPath);
                string restoredProjectName = Path.GetFileNameWithoutExtension(elkFullPath);

                restoreZw1.Project(targetZw1Path, restorePath, restoredProjectName, false, false);

                MessageBox.Show("프로젝트가 성공적으로 복구되었습니다.", "Restore 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Restore 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public Project OpenProject(string elkFullPath)
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                var toReturn = projectManager.OpenProject(elkFullPath, ProjectManager.OpenMode.Standard, false);

                MessageBox.Show("프로젝트가 성공적으로 열렸습니다.", "Open 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return toReturn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Open 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void CloseProject()
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                foreach (Project eplProject in projectManager.OpenProjects)
                {
                    eplProject.Close();
                }

                MessageBox.Show("프로젝트가 성공적으로 닫혔습니다.", "Close 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Close 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void SetProperties(Project project)
        {
            try
            {
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD04] = "EPLAN Sample 프로젝트".GetMultiLangString();

                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD07] = "담당자".GetMultiLangString();
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD08] = "담당자 연락처".GetMultiLangString();

                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD16] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss").GetMultiLangString();
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD17] = "대한민국".GetMultiLangString();
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD18] = "EPLAN KOREA".GetMultiLangString();

                #region 도면 표제란 정보

                // 작성자, 검토자, 승인자 정보
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD01] = "작성자".GetMultiLangString();
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD02] = "검토자".GetMultiLangString();
                project.Properties[Properties.Project.PROJ_CUSTOM_SUPPLEMENTARYFIELD03] = "승인자".GetMultiLangString();

                // 공사명(프로젝트 설명)
                project.Properties[Properties.Project.PROJ_INSTALLATIONNAME] = "EPLAN API 샘플 도면".GetMultiLangString();

                // 도면 생성 날짜(프로젝트 시작 날짜)
                project.Properties[Properties.Project.PROJ_PROJECTBEGIN] = DateTime.Now;

                #endregion

                MessageBox.Show("프로젝트 속성이 성공적으로 설정되었습니다.", "속성 설정 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "속성 설정 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void GetPages(Project project)
        {
            try
            {
                Page[] pageList = project.Pages;

                if (pageList.Length == 0) {
                    MessageBox.Show("선택하신 프로젝트에는 페이지가 없습니다.", "프로젝트 페이지", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    return;
                }

                foreach (Page page in pageList)
                {
                    string message = string.Format("페이지명: [{0}]", page.Name);
                    MessageBox.Show(message, "프로젝트 페이지", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                MessageBox.Show("프로젝트 페이지가 성공적으로 조회되었습니다.", "프로젝트 페이지", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "프로젝트 페이지", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void CurrentProject()
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                Project project = projectManager.CurrentProject;

                if(project == null)
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

        public Project GetProject(string elkFullPath)
        {
            try
            {
                ProjectManager projectManager = new ProjectManager();
                var toReturn = projectManager.GetProject(elkFullPath);

                MessageBox.Show(string.Format("프로젝트 [{0}]이(가) 선택되었습니다.", toReturn.ProjectFullName), "프로젝트 선택", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return toReturn;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "프로젝트 선택 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }
    }
}
