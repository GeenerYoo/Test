using ApiExt.EplAddIn.Sample.EplanApi;
using System;
using System.IO;
using System.Windows.Forms;

namespace ApiExt.EplAddIn.Sample.Forms
{
    public partial class ApiExtSampleForm : Form
    {
        EplanProject eplanProject = null;
        EplanPage    eplanPage    = null;
        EplanMacro   eplanMacro   = null;
        EplanReport  eplanReport  = null;
        EplanExport  eplanExport  = null;
        EplanCommon eplanCommon = null;

        public ApiExtSampleForm()
        {
            InitializeComponent();
        }

        private void ApiExtSampleForm_Load(object sender, EventArgs e)
        {
            this.eplanProject = new EplanProject();
            this.eplanPage    = new EplanPage();
            this.eplanMacro   = new EplanMacro();
            this.eplanReport  = new EplanReport();
            this.eplanExport  = new EplanExport();
            this.eplanCommon  = new EplanCommon();

            this.cBoxEplanMacroType.SelectedIndex = 0;
        }

        private void btnZW1_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Title = "백업 프로젝트 선택";
            this.openFileDialog.Filter = "EPLAN 백업 프로젝트(*.zw1)|*.zw1";

            if (DialogResult.OK == this.openFileDialog.ShowDialog())
                this.textEplanZW1.Text = this.openFileDialog.FileName;
        }

        private void textEplanZW1_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.textEplanZW1.Text)) {
                this.textEplanProjectFolder.Text = string.Empty;

                return;
            }

            FileInfo zw1File = new FileInfo(this.textEplanZW1.Text);

            this.textEplanProjectFolder.Text = Path.Combine(zw1File.Directory.Parent.FullName, "Projects");
        }

        private void btnEplanProjectFolder_Click(object sender, EventArgs e)
        {
            this.folderBrowserDialog.SelectedPath = this.textEplanProjectFolder.Text;

            if (DialogResult.OK == this.folderBrowserDialog.ShowDialog())
                this.textEplanProjectFolder.Text = this.folderBrowserDialog.SelectedPath;
        }

        private void textEplanProjectFolder_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.textEplanZW1.Text) || string.IsNullOrWhiteSpace(this.textEplanProjectFolder.Text)) {
                this.textEplanProject.Text = string.Empty;

                return;
            }

            string zw1FileName = Path.GetFileNameWithoutExtension(this.textEplanZW1.Text);
            string dataTimeText = DateTime.Now.ToString("yyMMddHHmmssfff");

            this.textEplanProject.Text = Path.Combine(this.textEplanProjectFolder.Text, string.Format("{0}_{1}.elk", zw1FileName, dataTimeText));
        }

        private void btnEplanProject_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Title = "프로젝트 선택";
            this.openFileDialog.Filter = "EPLAN 프로젝트(*.elk)|*.elk";

            if (DialogResult.OK == this.openFileDialog.ShowDialog())
                this.textEplanProject.Text = this.openFileDialog.FileName;
        }

        private void btnSecondProject_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Title = "프로젝트 선택";
            this.openFileDialog.Filter = "EPLAN 프로젝트(*.elk)|*.elk";

            if (DialogResult.OK == this.openFileDialog.ShowDialog())
                this.textSecondProject.Text = this.openFileDialog.FileName;
        }

        private void btnEplanMacro_Click(object sender, EventArgs e)
        {
            this.openFileDialog.Title = "매크로 선택";
            this.openFileDialog.Filter = "EPLAN 페이지 매크로(*.emp)|*.emp|EPLAN 윈도우 매크로(*.ema)|*.ema";

            if (DialogResult.OK == this.openFileDialog.ShowDialog())
                this.textEplanMacro.Text = this.openFileDialog.FileName;
        }
    }
}
