using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Collections.Specialized;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanExport
    {
        public void BackupProject(Project targetProject)
        {
            try
            {
                string elkFullPath = targetProject.ProjectLinkFilePath;
                string backupPath = Path.GetDirectoryName(elkFullPath);
                string backupFileName = Path.GetFileNameWithoutExtension(elkFullPath);

                Backup backupExport = new Backup();
                StringCollection projectPath = new StringCollection();
                StringCollection targetPath = new StringCollection();

                projectPath.Add(elkFullPath);
                targetPath.Add(string.Format("{0}.zw1", backupFileName));

                backupExport.Project(projectPath, string.Empty, backupPath, targetPath, Backup.Type.MakeBackup, Backup.Medium.Disk, 0.0, Backup.Amount.All, false, false, false);

                MessageBox.Show("프로젝트가 성공적으로 백업되었습니다.", "Backup 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Backup 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void PdfExport(Project targetProject)
        {
            try
            {
                string elkFullPath = targetProject.ProjectLinkFilePath;
                string exportFolder = Path.GetDirectoryName(elkFullPath);
                string pdfFileName  = string.Format("{0}.pdf", Path.GetFileNameWithoutExtension(elkFullPath));
                string pdfFullPath = Path.Combine(exportFolder, pdfFileName);

                Export pdfExport = new Export();
                pdfExport.PdfProject(elkFullPath, string.Empty, pdfFullPath, Export.DegreeOfColor.Color, false, string.Empty, true);

                MessageBox.Show("PDF 내보내기가 성공적으로 실행되었습니다.", "PDF 내보내기 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "PDF 내보내기 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void DwgDxfExport(Project targetProject)
        {
            try
            {
                string elkFullPath = targetProject.ProjectLinkFilePath;
                string dwgFilesExportFolder = Path.Combine(Path.GetDirectoryName(elkFullPath), DateTime.Now.ToString("MMdd_HHmmssfff"));
                string dwgExpotedZipFileName = string.Format("{0}.zip", Path.GetFileNameWithoutExtension(elkFullPath));
                string dwgExpotedZipFileFullPath = Path.Combine(Path.GetDirectoryName(elkFullPath), dwgExpotedZipFileName);

                Directory.CreateDirectory(dwgFilesExportFolder);

                Export dwgExport = new Export();
                dwgExport.DxfDwgProjectToDisk(targetProject, "Standard Settings", dwgFilesExportFolder);

                Thread.Sleep(100);

                ZipFile.CreateFromDirectory(dwgFilesExportFolder, dwgExpotedZipFileFullPath);

                MessageBox.Show("DWG/DXF 내보내기가 성공적으로 실행되었습니다.", "DWG/DXF 내보내기 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);

                DeleteTempWorkingFolder(dwgFilesExportFolder);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "DWG/DXF 내보내기 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        #region Private Methods

        private void DeleteTempWorkingFolder(string folderFullPath)
        {
            try
            {
                Directory.Delete(folderFullPath, true);
            }
            catch (Exception)
            {
                // 임시 파일을 제거하는 과정에서 오류(Exception)은  무시한다.
            }
        }

        #endregion
    }
}
