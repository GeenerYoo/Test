using Eplan.EplApi.DataModel;
using ApiExt.EplAddIn.Sample.Extensions;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using Eplan.EplApi.DataModel.MasterData;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanPage
    {
        public Page CreatePage(Project targetProject, string functionalAssignment, string plant, string location, string docType, string pageDescription)
        {
            try
            {
                PagePropertyList oPagePropList = new PagePropertyList();

                oPagePropList.DESIGNATION_FUNCTIONALASSIGNMENT = functionalAssignment;
                oPagePropList.DESIGNATION_PLANT = plant;
                oPagePropList.DESIGNATION_LOCATION = location;
                oPagePropList.DESIGNATION_DOCTYPE = docType;
                oPagePropList.PAGE_NOMINATIOMN = pageDescription.GetMultiLangString();

                Page newPage = new Page();
                newPage.Create(targetProject, DocumentTypeManager.DocumentType.Circuit, oPagePropList);
                newPage.Properties.PAGE_NOMINATIOMN = pageDescription.GetMultiLangString();

                MessageBox.Show(string.Format("페이지 [{0}]이(가) 성공적으로 생성되었습니다.", newPage.Name), "페이지 생성 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return newPage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "페이지 생성 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public IEnumerable<Page> FilterPage(Project targetProject, string functionalAssignment, string plant, string location, string docType)
        {
            try
            {
                DMObjectsFinder pageFinder = new DMObjectsFinder(targetProject);
                PagesFilter pageFilter = new PagesFilter();
                PagePropertyList pageProperties = new PagePropertyList();

                pageProperties.DESIGNATION_FUNCTIONALASSIGNMENT = functionalAssignment;
                pageProperties.DESIGNATION_PLANT = plant;
                pageProperties.DESIGNATION_LOCATION = location;
                pageProperties.DESIGNATION_DOCTYPE = docType;

                pageFilter.SetFilteredPropertyList(pageProperties);

                Page[] filteredPages = pageFinder.GetPages(pageFilter);
                int totalPageCount = filteredPages.Length;

                MessageBox.Show("페이지 필터가 성공적으로 실행되었습니다.", "페이지 필터 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);

                for (int pageIndex = 0; pageIndex < totalPageCount; pageIndex++)
                {
                    MessageBox.Show(string.Format("페이지[{0}] = [{1}]", pageIndex, filteredPages[pageIndex].Name), "페이지 필터 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return filteredPages;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "페이지 필터 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public Page CopyPage(Page sourcePage, Project targetProject, string functionalAssignment, string plant, string location, string docType, string pageDescription)
        {
            try
            {
                PagePropertyList copyPageProperties = new PagePropertyList();
                sourcePage.Properties.CopyTo(copyPageProperties);

                copyPageProperties.DESIGNATION_FUNCTIONALASSIGNMENT = functionalAssignment;
                copyPageProperties.DESIGNATION_PLANT = plant;
                copyPageProperties.DESIGNATION_LOCATION = location;
                copyPageProperties.DESIGNATION_DOCTYPE = docType;

                Page copiedPage = sourcePage.CopyTo(targetProject, copyPageProperties, PageMacro.Enums.NumerationMode.Ignore, true);
                copiedPage.Properties.PAGE_NOMINATIOMN = pageDescription.GetMultiLangString();

                MessageBox.Show(string.Format("페이지 [{0}]이(가) 성공적으로 복사되었습니다.", copiedPage.Name), "페이지 복사 성공", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return copiedPage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "페이지 복사 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void SetProperties(Page targetPage)
        {
            try
            {
                targetPage.Properties[Properties.Page.PAGE_CUSTOM_SUPPLEMENTARYFIELD05] = "API Ext Sample 장치명".GetMultiLangString();
                targetPage.Properties[Properties.Page.PAGE_CUSTOM_SUPPLEMENTARYFIELD02] = "API Ext Sample 메인 공급".GetMultiLangString();
                targetPage.Properties[Properties.Page.PAGE_CUSTOM_SUPPLEMENTARYFIELD07] = "도면번호-001".GetMultiLangString();

                targetPage.Properties[Properties.Page.PAGE_CUSTOM_SUPPLEMENTARYFIELD03] = "API Ext Sample 보조공급장치 #1".GetMultiLangString();
                targetPage.Properties[Properties.Page.PAGE_CUSTOM_SUPPLEMENTARYFIELD04] = "API Ext Sample 보조공급장치 #2".GetMultiLangString();

                MessageBox.Show("페이지 속성이 성공적으로 설정되었습니다.", "페이지 속성 설정", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "페이지 속성 설정 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }
    }
}
