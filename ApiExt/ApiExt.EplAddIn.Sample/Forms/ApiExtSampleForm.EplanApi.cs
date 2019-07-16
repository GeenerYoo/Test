using Eplan.EplApi.DataModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.Forms
{
    partial class ApiExtSampleForm
    {
        Project selectedProject = null;
        Project secondProject   = null;
        Page    selectedPage    = null;

        #region Project

        private void btnProjectRestore_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanZW1.Text) || string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 백업 파일(*.zw1)을 선택하시기 바랍니다.", "Restore 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                this.eplanProject.RestoreProject(this.textEplanZW1.Text, this.textEplanProject.Text);
            }
            catch (Exception)
            {
                this.textEplanZW1.Text = string.Empty;
            }
        }

        private void btnProjectOpen_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "Open 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text)) {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "Open 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                this.selectedProject = this.eplanProject.OpenProject(this.textEplanProject.Text);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            try
            {
                this.eplanProject.CloseProject();
                //this.eplanProject.CurrentProject();//현재 활성화된 프로젝트?????

            }
            catch (Exception)
            {
                this.textEplanZW1.Text = string.Empty;
            }
            finally
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "Backup 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text)) {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "Backup 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                this.eplanExport.BackupProject(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnProperty_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "프로젝트 속성 설정 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = this.eplanProject.OpenProject(this.textEplanProject.Text);

                this.eplanProject.SetProperties(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnProjectPage_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text))
                {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "프로젝트 페이지 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = this.eplanProject.OpenProject(this.textEplanProject.Text);

                this.eplanProject.GetPages(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnProjectCurrent_Click(object sender, EventArgs e)
        {
            try
            {
                this.eplanProject.CurrentProject();
            }
            catch (Exception)
            {
            }
        }

        private void btnProjectSelect_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "프로젝트 선택 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text)) {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "프로젝트 선택 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                this.selectedProject = this.eplanProject.GetProject(this.textEplanProject.Text);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        #endregion

        #region Page

        private void btnPageCreate_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "페이지 생성 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                this.selectedPage = this.eplanPage.CreatePage(this.selectedProject, "KOREA", "SAMPLE", "Area01", "DRAWING", "ApiExt Sample Page");
                this.eplanPage.CreatePage(this.selectedProject, "KOREA", "SAMPLE", "Area01", string.Empty, "ApiExt Sample Page");
            }
            catch (Exception)
            {
                this.selectedPage = null;
            }
        }

        private void btnPageFilter_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textSecondProject.Text)) {
                    MessageBox.Show("먼저 두번째 프로젝트(*.elk)를 선택하시기 바랍니다.", "페이지 필터 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.secondProject == null)
                    this.secondProject = GetOrOpenProject(this.textSecondProject.Text);

                this.eplanPage.FilterPage(this.secondProject, string.Empty, "SAMPLE", "Area02", string.Empty);
            }
            catch (Exception)
            {
                this.textSecondProject.Text = string.Empty;
                this.secondProject = null;
            }
        }

        private void btnPageCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "페이지 복사 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                if (string.IsNullOrWhiteSpace(this.textSecondProject.Text)) {
                    MessageBox.Show("먼저 두번째 프로젝트(*.elk)를 선택하시기 바랍니다.", "페이지 복사 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.secondProject == null)
                    this.secondProject = GetOrOpenProject(this.textSecondProject.Text);

                IEnumerable<Page> sourcePages = this.eplanPage.FilterPage(this.secondProject, string.Empty, "SAMPLE", "Area02", string.Empty);

                if (sourcePages.Count() == 0) {
                    MessageBox.Show("두번째 프로젝트에 선택된 페이지가 없습니다.", "페이지 복사 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                foreach (Page sourcePage in sourcePages)
                {
                    this.eplanPage.CopyPage(sourcePage, this.selectedProject, "COPY", "TEST", "Area02", "DRAWING", "ApiExt Copied Page");
                }

            }
            catch (Exception)
            {
                this.textSecondProject.Text = string.Empty;
                this.secondProject = null;
            }
        }

        private void btnPageProperty_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "페이지 속성 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                IEnumerable<Page> selectedPages = this.eplanPage.FilterPage(this.selectedProject, "KOREA", "SAMPLE", "Area01", "DRAWING");

                if (selectedPages.Count() == 0) {
                    MessageBox.Show("프로젝트에서 선택된 페이지가 없습니다.", "페이지 속성 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                foreach (Page targetPage in selectedPages)
                {
                    this.eplanPage.SetProperties(targetPage);
                }
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        #endregion

        #region Macro

        private void btnMacroPageMacro_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<Page> selectedPages = GetPage();

                if (selectedPages == null || selectedPages.Count() == 0)
                    return;

                if (!ValidatePageMacro())
                    return;

                this.eplanMacro.InsertPageMacro(this.textEplanMacro.Text, selectedPages.LastOrDefault(), this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanMacro.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnMacroWindowMacro_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<Page> selectedPages = GetPage();

                if (selectedPages == null || selectedPages.Count() == 0)
                    return;

                if (!ValidateWindowMacro())
                    return;

                int positionX = 100;
                int positionY = 400;

                GetMacroPostion(ref positionX, ref positionY);


                this.eplanMacro.InsertWindowMacro(selectedPages.LastOrDefault(), this.textEplanMacro.Text, positionX, positionY, GetVariantIndex());
            }
            catch (Exception)
            {
                this.textEplanMacro.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnMacroProperty_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<Page> selectedPages = GetPage();

                if (selectedPages == null || selectedPages.Count() == 0)
                    return;

                if (!ValidateWindowMacro())
                    return;

                int positionX = 200;
                int positionY = 400;

                GetMacroPostion(ref positionX, ref positionY);


                this.eplanMacro.SetMacroProperties(selectedPages.LastOrDefault(), this.textEplanMacro.Text, positionX, positionY, GetVariantIndex());
            }
            catch (Exception)
            {
                this.textEplanMacro.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnMacroPlaceHolder_Click(object sender, EventArgs e)
        {
            try
            {
                IEnumerable<Page> selectedPages = GetPage();

                if (selectedPages == null || selectedPages.Count() == 0)
                    return;

                if (!ValidateWindowMacro())
                    return;

                int positionX = 300;
                int positionY = 400;

                GetMacroPostion(ref positionX, ref positionY);
                var placeHolderParameters = GetPlaceHolderParameters();

                this.eplanMacro.InsertPlaceHolderMacro(selectedPages.LastOrDefault(), this.textEplanMacro.Text, positionX, positionY, GetVariantIndex(), "EPLANApi", placeHolderParameters);
            }
            catch (Exception)
            {
                this.textEplanMacro.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        #endregion

        #region Report

        private void btnReport_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "레포트 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text))  {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "레포트 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                this.eplanReport.GenerateReport(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        #endregion

        #region Export

        private void btnExportPdf_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "PDF Export 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text)) {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "PDF Export 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                this.eplanExport.PdfExport(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnExportDwgDxf_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                    MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "DWG/DXF Export 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (!File.Exists(this.textEplanProject.Text)) {
                    MessageBox.Show("선택하신 프로젝트(*.elk)가 존재하지 않습니다.", "DWG/DXF Export 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return;
                }

                if (this.selectedProject == null)
                    this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

                this.eplanExport.DwgDxfExport(this.selectedProject);
            }
            catch (Exception)
            {
                this.textEplanProject.Text = string.Empty;
                this.selectedProject = null;
            }
        }

        private void btnExportZw1_Click(object sender, EventArgs e)
        {
            this.btnBackup_Click(sender, e);
        }

        #endregion

        #region Private Methods

        private Project GetOrOpenProject(string elkFullPath)
        {
            if (this.eplanCommon.IsProjectOpen(elkFullPath))
                return this.eplanProject.GetProject(elkFullPath);
            else
                return this.eplanProject.OpenProject(elkFullPath);
        }

        private int GetVariantIndex()
        {
            string mapVariantLetter = "ABCDEFGHIJKLMNOP";
            
            int toReturn = mapVariantLetter.IndexOf(this.cBoxEplanMacroType.SelectedText);

            if (toReturn < 0)
                toReturn = 0;

            return toReturn;
        }

        private IEnumerable<Page> GetPage()
        {
            if (string.IsNullOrWhiteSpace(this.textEplanProject.Text)) {
                MessageBox.Show("먼저 프로젝트(*.elk)를 선택하시기 바랍니다.", "매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return null;
            }

            if (this.selectedProject == null)
                this.selectedProject = GetOrOpenProject(this.textEplanProject.Text);

            IEnumerable<Page> selectedPages = this.eplanPage.FilterPage(this.selectedProject, "KOREA", "SAMPLE", "Area01", "DRAWING");

            if (selectedPages.Count() == 0) {
                MessageBox.Show("프로젝트에서 선택된 페이지가 없습니다.", "매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return null;
            }

            return selectedPages;
        }

        private bool ValidatePageMacro()
        {
            bool toReturn = ValidateMacroNameSelected();

            if (!toReturn)
                return toReturn;

            if (!this.textEplanMacro.Text.EndsWith(".emp", StringComparison.OrdinalIgnoreCase)) {
                MessageBox.Show("페이지 매크로를 선택하시기 바랍니다.", "페이지 매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                toReturn = false;
            }

            return toReturn;
        }

        private bool ValidateWindowMacro()
        {
            bool toReturn = ValidateMacroNameSelected();

            if (!toReturn)
                return toReturn;

            if (!this.textEplanMacro.Text.EndsWith(".ema", StringComparison.OrdinalIgnoreCase)) {
                MessageBox.Show("윈도우 매크로를 선택하시기 바랍니다.", "윈도우 매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                toReturn = false;
            }

            return toReturn;
        }

        private bool ValidateMacroNameSelected()
        {
            if (string.IsNullOrWhiteSpace(this.textEplanMacro.Text)) {
                MessageBox.Show("먼저 매크로 파일을 선택하시기 바랍니다.", "매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            if (!File.Exists(this.textEplanMacro.Text)) {
                MessageBox.Show("선택하신 매크로 파일이 존재하지 않습니다.", "매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return false;
            }

            return true;
        }

        private void GetMacroPostion(ref int positionX, ref int positionY)
        {
            positionX = GetPostion(this.textMacroPositionX.Text, positionX);
            positionY = GetPostion(this.textMacroPositionY.Text, positionY);
        }

        private int GetPostion(string postionText, int defaultValue)
        {
            int positionValue = 0;
            if (!Int32.TryParse(postionText, out positionValue))
                positionValue = defaultValue;

            if (positionValue <= 50)
                positionValue = defaultValue;

            return positionValue;
        }

        private Dictionary<string, string> GetPlaceHolderParameters()
        {
            const string PlaceHolderVariableFlowSummaryDataColumnSequenceNo = "SEQ_NO";      // 흐름 번호
            const string PlaceHolderVariableFlowSummaryDataColumnDescription = "DESCRIPTION"; // 흐름 명칭
            const string PlaceHolderVariableFlowSummaryDataColumnPhase = "PHASE";       // 상
            const string PlaceHolderVariableFlowSummaryDataColumnContents = "CONTENTS";    // 내용물
            const string PlaceHolderVariableFlowSummaryDataColumnOperationTemperature = "OPER_TEM";    // 운전 온도
            const string PlaceHolderVariableFlowSummaryDataColumnOperationPressure = "OPER_PRE";    // 운전 압력
            const string PlaceHolderVariableFlowSummaryDataColumnDesignPressure = "DES_PRE";     // 설계 압력
            const string PlaceHolderVariableSummaryDataColumnNote = "NOTE";        // 비고

            Dictionary<string, string> placeHolderParameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnSequenceNo, "API Ext 흐름 번호");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnDescription, "API Ext 흐름 명칭");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnPhase, "API Ext 흐름 상(PHASE)");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnContents, "API Ext 내용물");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnOperationTemperature, "API Ext 운전 온도");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnOperationPressure, "API Ext 운전 압력");
            placeHolderParameters.Add(PlaceHolderVariableFlowSummaryDataColumnDesignPressure, "API Ext 설계 압력");
            placeHolderParameters.Add(PlaceHolderVariableSummaryDataColumnNote, "API Ext 비고");

            return placeHolderParameters;
        }

        #endregion
    }
}
