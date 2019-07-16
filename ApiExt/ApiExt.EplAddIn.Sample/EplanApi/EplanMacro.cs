using ApiExt.EplAddIn.Sample.Extensions;
using Eplan.EplApi.Base;
using Eplan.EplApi.DataModel;
using Eplan.EplApi.HEServices;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;
using Eplan.EplApi.DataModel.Graphics;

namespace ApiExt.EplAddIn.Sample.EplanApi
{
    public class EplanMacro
    {
        public StorableObject[] InsertPageMacro(string macroFullPath, Page insertAfterPage, Project targetProject)
        {
            try
            {
                Insert macroInsert = new Insert();
                var storableObjects = macroInsert.PageMacro(macroFullPath, insertAfterPage, targetProject, true);

                MessageBox.Show("페이지 매크로가 성공적으로 배치되었습니다.", "페이지 매크로", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return storableObjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "페이지 매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public StorableObject[] InsertWindowMacro(Page targetPage, string macroFullPath, int positionX, int positionY, int nVariant)
        {
            try
            {
                Insert macroInsert = new Insert();
                PointD macroPoint = new PointD(positionX, positionY);

                var storableObjects = macroInsert.WindowMacro(macroFullPath, nVariant, targetPage, macroPoint, Insert.MoveKind.Absolute);

                MessageBox.Show("윈도우 매크로가 성공적으로 배치되었습니다.", "윈도우 매크로", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return storableObjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "윈도우 매크로 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public void SetMacroProperties(Page targetPage, string macroFullPath, int positionX, int positionY, int nVariant)
        {
            try
            {
                StorableObject[] storableObjects = InsertWindowMacro(targetPage, macroFullPath, positionX, positionY, nVariant);
                SetDisplayTextForStorableObjects(storableObjects);

                MessageBox.Show("윈도우 매크로 속성이 성공적으로 설정 되었습니다.", "윈도우 매크로 속성", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "윈도우 매크로 속성 에러", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        public StorableObject[] InsertPlaceHolderMacro(Page targetPage, string macroFullPath, int positionX, int positionY, int nVariant, string placeHolderRecordName, Dictionary<string, string> placeHolderParameters)
        {
            try
            {
                StorableObject[] storableObjects = InsertWindowMacro(targetPage, macroFullPath, positionX, positionY, nVariant);
                SetPlaceHolderValues(storableObjects, placeHolderRecordName, placeHolderParameters);

                MessageBox.Show("위치 지정 객체가 성공적으로 배치되었습니다.", "위치 지정 객체", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return storableObjects;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "위치 지정 객체", MessageBoxButtons.OK, MessageBoxIcon.Error);

                throw ex;
            }
        }

        #region Private Methos

        private void SetPlaceHolderValues(IEnumerable<StorableObject> storableObjects, string recordName, Dictionary<string, string> placeHolderParameters)
        {
            if (placeHolderParameters == null || placeHolderParameters.Count == 0) // 설정할 값이 없으면 바로 종료
                return;

            foreach (PlaceHolder placeHolder in storableObjects.Where(s => s is PlaceHolder))
            {
                placeHolder.AddRecord(recordName);

                foreach (string key in placeHolderParameters.Keys)
                {
                    string placeHolderValue = placeHolderParameters[key] ?? string.Empty;
                    placeHolder.SetValue(recordName, key, placeHolderValue.GetMultiLangString());
                }

                placeHolder.ApplyRecord(recordName);
            }
        }

        private void SetDisplayTextForStorableObjects(StorableObject[] storableObjects)
        {
            int dtTextIndex = 0;
            string dtTextPrefix = DateTime.Now.ToString("MMddHHmmss"); 

            foreach (var storableObject in storableObjects)
            {
                if (storableObject is BoxedDevice) {
                    SetDisplayTextForBoxedDevice((BoxedDevice)storableObject, dtTextPrefix, dtTextIndex++);

                    return;
                }

                if (storableObject is InterruptionPoint) {
                    SetDisplayTextForInterruptionPoint((InterruptionPoint)storableObject, dtTextPrefix, dtTextIndex++);

                    return;
                }

                if (storableObject is Function) {
                    SetDisplayTextForFunctionItem((Function)storableObject, dtTextPrefix, dtTextIndex++);

                    return;
                }
            }
        }

        private void SetDisplayTextForBoxedDevice(BoxedDevice boxedDevice, string dtTextPrefix, int dtTextIndex)
        {
            string dtText = string.Format("{0}_{1,3:D3}", dtTextPrefix, dtTextIndex);

            boxedDevice.Name = dtText;
            boxedDevice.VisibleName = dtText;
            boxedDevice.Properties[Properties.Function.FUNC_TECHNICAL_CHARACTERISTIC] = string.Format("BlackBox_기술특성_{0}", dtTextPrefix);
        }

        private void SetDisplayTextForInterruptionPoint(InterruptionPoint interruptionPoint, string dtTextPrefix, int dtTextIndex)
        {
            string dtText = string.Format("{0}_{1,3:D3}", dtTextPrefix, dtTextIndex);

            interruptionPoint.Name = dtText;
            interruptionPoint.VisibleName = dtText;
            interruptionPoint.Properties[Properties.InterruptionPoint.INTERRUPTIONPOINT_DESCRIPTION] = string.Format("INTERRUPTIONPOINT_DESCRIPTION_{0}", dtTextPrefix).GetMultiLangString();
        }

        private void SetDisplayTextForFunctionItem(Function functionItem, string dtTextPrefix, int dtTextIndex)
        {
            string dtText = string.Format("{0}_{1,3:D3}", dtTextPrefix, dtTextIndex);

            functionItem.Name = dtText;
            functionItem.VisibleName = dtText;
            functionItem.Properties[Properties.Function.FUNC_TEXT] = string.Format("FUNC_TEXT_{0}", dtTextPrefix).GetMultiLangString();
        }

        #endregion
    }
}
