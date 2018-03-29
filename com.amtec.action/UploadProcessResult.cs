using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;

namespace com.amtec.action
{
    public class UploadProcessResult
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private int error;
        private MainView view;

        public UploadProcessResult(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int UploadProcessResultCall(String[] serialNumberArray, int processLayer)
        {
            String[] serialNumberUploadKey = new String[] { "ERROR_CODE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberUploadValues = new String[] { };
            String[] serialNumberResultValues = new String[] { };
            serialNumberUploadValues = serialNumberArray;
            error = imsapi.trUploadState(sessionContext, init.configHandler.StationNumber, 2, "-1", "-1", 0, 1, -1, 0, serialNumberUploadKey, serialNumberUploadValues, out serialNumberResultValues);

            if ((error != 0) && (error != 210))
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error + "(" + errorString + ")", "");
                return error;
            }
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error, "");
            return error;
        }

        public int UploadProcessResultCall(string serialNumber, int processLayer)
        {
            String[] serialNumberUploadKey = new String[] { "ERROR_CODE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberUploadValues = new String[] { };
            String[] serialNumberResultValues = new String[] { };
            error = imsapi.trUploadState(sessionContext, init.configHandler.StationNumber, processLayer, serialNumber, "-1", 0, 1, -1, 0, serialNumberUploadKey, serialNumberUploadValues, out serialNumberResultValues);
            LogHelper.Info("Api trUploadState: serial number =" + serialNumber + ", result code =" + error);
            if ((error != 0) && (error != 210))
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error + "(" + errorString + ")", "");
                return error;
            }
            error = 0;
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error, "");
            return error;
        }

        public int UploadFailureAndResultData(long bookDate, string serialNumber, string serialNumberPos, int processLayer, int serialNumberState, int duplicateSerialNumber
       , string[] measureValues, string[] failureValues)
        {
            string[] measureKeys = new string[] { "ERROR_CODE", "MEASURE_FAIL_CODE", "MEASURE_NAME", "MEASURE_VALUE" };
            string[] failureKeys = new string[] { "ERROR_CODE", "FAILURE_TYPE_CODE", "COMP_NAME" };
            string[] failureSlipKeys = new string[] { "ERROR_CODE", "TEST_STEP_NAME" };
            string[] failureSlipValues = new string[] { };
            string[] measureResultValues = new string[] { };
            string[] failureResultValues = new string[] { };
            string[] failureSlipResultValues = new string[] { };
            error = imsapi.trUploadFailureAndResultData(sessionContext, init.configHandler.StationNumber, processLayer, serialNumber, serialNumberPos,
                serialNumberState, duplicateSerialNumber, 0, bookDate, measureKeys, measureValues, out measureResultValues, failureKeys, failureValues, out failureResultValues,
                failureSlipKeys, failureSlipValues, out failureSlipResultValues);
            if (failureValues != null && failureValues.Length > 0)
            {
                foreach (var item in failureValues)
                {
                    LogHelper.Info(item);
                }
            }
            LogHelper.Info("Api trUploadFailureAndResultData (serial number:" + serialNumber + ",pos:" + serialNumberPos + ",process layer:" + processLayer + ",state:" + serialNumberState + ",result code:" + error);
            if (error == 0 || error == 210)
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadFailureAndResultData " + error + "(" + errorString + ")", "");
                error = 0;
            }
            else
            {
                string errorString = ""; //UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadFailureAndResultData " + error + "," + errorString, "");
            }
            return error;
        }

        public int UploadResultDataAndRecipe(long bookDate, string serialNum, string serialNumPos, string serialnumState, string[] listItem, int multiboard, int processLayer)
        {
            var uploadValues = new string[] { };
            string[] UploadKeyN = { "ErrorCode", "LowerLimit", "MeasureFailCode", "MeasureName", "MeasureValue", "Unit", "UpperLimit" };//"Nominal", "Remark", "Tolerance",
            uploadValues = listItem;
            string[] SNStateResultValues = new string[] { };
            //recipeVersionMode  要从-1 改为0 之前是有问题的  第二次更改从0改为1 感觉没什么区别.
            //recipeVersionMode =-1 或者为空 recipeVersionId=-1 measurename 名称长度不能超过80个字符 超过 报-10  duplicateserialnumber =0 表示只上传大板中某一个位置的SN =1/-1 上传时都会复制给大板中其他位置的小板
            int error = imsapi.trUploadResultDataAndRecipe(sessionContext, init.configHandler.StationNumber, processLayer, -1, serialNum,// 不知道为什么processLayer 只能=1 以前的都是2的
                serialNumPos, Convert.ToInt32(serialnumState), multiboard, bookDate, 0, -1, UploadKeyN, uploadValues, out SNStateResultValues);//snState 默认设置为1   
            LogHelper.Info("API trUploadResultDataAndRecipe: serial number:" + serialNum + ", serialnumState:" + serialnumState + ", result code =" + error);
            if (error == 0||error==210)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadResultDataAndRecipe " + error, "");
                error = 0;
            }
            else
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadResultDataAndRecipe " + error + "(" + errorString + ")", "");
            }

            return error;
        }
    }
}
