using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;

namespace com.amtec.action
{
    public class CheckSerialNumberState
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public CheckSerialNumberState(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public List<SerialNumberStateEntity> GetSerialNumberData(string serialNumber)
        {
            List<SerialNumberStateEntity> snList = new List<SerialNumberStateEntity>();
            String[] serialNumberStateResultKeys = new String[] { "LOCK_STATE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            LogHelper.Info("end api trCheckSerialNumberState (result code = " + error + ")");
            if ((error != 0) && (error != 5) && (error != 6) && (error != 204) && (error != 207) && (error != 212))
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "," + errorString, "");
            }
            else
            {
                int counter = serialNumberStateResultValues.Length;
                int loop = serialNumberStateResultKeys.Length;
                for (int i = 0; i < counter; i += loop)
                {
                    snList.Add(new SerialNumberStateEntity(serialNumberStateResultValues[i], serialNumberStateResultValues[i + 1], serialNumberStateResultValues[i + 2], serialNumberStateResultValues[i + 3]));
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
            }
            return snList;
        }
        public bool CheckSNStateNew(string serialNumber, int processLayer, out string errorMsg)
        {
           
            //view.ScrapSNlist.Clear();
            //SNPanelCount = 0;
            int errorcode = 0;
            //errorMsg = "";
            errorMsg = "";
             //return true;
          // snlist = new List<string>();
            String[] serialNumberStateResultKeys = new String[] { "ERROR_CODE", "SERIAL_NUMBER_POS", "SERIAL_NUMBER" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            LogHelper.Info("end api trCheckSerialNumberState (errorcode = " + error + ")");
            if (error == 0)
            {
                //int counter = serialNumberStateResultValues.Length;
                //int loop = serialNumberStateResultKeys.Length;
                //for (int i = 0; i < counter; i += loop)
                //{
                //    SNPanelCount++;
                //    snlist.Add(serialNumberStateResultValues[i + 2]);
                return true;
            }else if (error== 5 || error == 6)
              {
                errorMsg = "";
                int looplength = serialNumberStateResultKeys.Length;
                int alllength = serialNumberStateResultValues.Length;
                for (int i = 0; i < alllength; i += looplength)
                {
                     
                    errorcode = Convert.ToInt32(serialNumberStateResultValues[i]);
                    if (errorcode != 0)
                    {
                        if (errorcode == 202 || errorcode == 203)
                        {

                            string workstepdesc = GetNextProductionStep(serialNumber);
                            string errorString = UtilityFunction.GetZHSErrorString(errorcode, init, sessionContext);
                            errorMsg = errorcode + ";" + errorString + "(" + workstepdesc + ")";
                        }
                        else if (errorcode == -201 || errorcode == 204 || errorcode == 207 || errorcode == 212)//scrap
                        {
                            //if (!view.ScrapSNlist.Contains(serialNumberStateResultValues[i + 2]))
                            //{
                            //    view.ScrapSNlist.Add(serialNumberStateResultValues[i + 2]);
                            //}
                        }
                        else
                        {
                            string errorString = UtilityFunction.GetZHSErrorString(errorcode, init, sessionContext);
                            errorMsg = errorcode + ";" + errorString;
                        }
                    }
                    //else
                    //{
                    //    snlist.Add(serialNumberStateResultValues[i + 2]);
                    //}
                }
                if (errorMsg != "")
                {
                    view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "(" + errorMsg + ")", "");
                    return false;
                }
                else
                {
                    error = 0;
                    view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
                    return true;
                }

            }
            else
            {
                errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "(" + errorMsg + ")", "");
                return false;
            }

        }
        public bool CheckSNState(string serialNumber, int processLayer,out string errorString)
        {
             errorString = "";
             int checkscrap = 0;
            String[] serialNumberStateResultKeys = new String[] { "ERROR_CODE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ", Process layer:" + processLayer + ")");
            int errorCode = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            LogHelper.Info("end api trCheckSerialNumberState (Result code = " + errorCode + ")");
            if ((errorCode != 0) && (errorCode != 5) && (errorCode != 6) && (errorCode != 204) && (errorCode != 207) && (errorCode != 212))
            {
                errorString = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + errorCode + "," + errorString, "");
                return false;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + errorCode, "");
                if (errorCode == 5|| errorCode == 6)//202 Serial no. is invalid for this station; it was not seen by the previous station
                {
                    foreach (var item in serialNumberStateResultValues)
                    {
                        if (item != "0")
                        {
                            LogHelper.Info("Sub error code = " + item);
                            errorCode = Convert.ToInt32(item);
                            break;
                        }
                    }
                   
                    if (errorCode == -201 || errorCode == 204 || errorCode == 207 || errorCode == 212)//scrap
                    {
                        return true;
                    }
                    errorString = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
                    if (errorCode == 202 || errorCode == 203)
                    {
                        checkscrap++;
                        string strNextStep = GetNextProductionStep(serialNumber);
                        errorString = errorString + " -> " + strNextStep;
                    }
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }
        public string GetNextProductionStep(string serialNumber)
        {
            string strNextStep = "";
            string[] productionStepResultKeys = new string[] { "WORKSTEP_DESC" };
            string[] productionStepResultValues = new string[] { };
            int errorCode = imsapi.trGetNextProductionStep(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", 0, 0, 0, productionStepResultKeys, out productionStepResultValues);
            LogHelper.Info("Api trGetNextProductionStep serial number = " + serialNumber + " result code = " + errorCode + ", result string = " + UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext));
            if (errorCode == 0)
            {
                strNextStep = productionStepResultValues[0];
            }
            return strNextStep;
        }
        public int GetProcessLayerBySN(string serialNumber, string stationNumber)
        {
            int iProcessLayer = -1;
            //string currentWS = "";
            //serialNumber = serialNumber + "001";
            //string[] uploadInfoResultKeys = new string[] { "WORKSTEP_NUMBER" };
            //string[] uploadInfoResultValues = new string[] { };
            //int errorCode = imsapi.trGetSerialNumberUploadInfo(sessionContext, stationNumber, -1, serialNumber, "-1", 0, uploadInfoResultKeys, out uploadInfoResultValues);
            //LogHelper.Info("Api trGetSerialNumberUploadInfo: station number =" + stationNumber + ",serial number =" + serialNumber + ", result code =" + errorCode);
            //if (errorCode == 0)
            //{
            //    currentWS = uploadInfoResultValues[0];
            //    LogHelper.Debug("work step number :" + currentWS);
            //}

            //KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("FUNC_MODE", "0"), new KeyValue("WORKSTEP_FLAG", "1"), new KeyValue("SERIAL_NUMBER", serialNumber) };
            //string[] workplanDataResultKeys = new string[] { "WORKSTEP_NUMBER", "PROCESS_LAYER" };
            //string[] workplanDataResultValues = new string[] { };
            //int errorWP = imsapi.mdataGetWorkplanData(sessionContext, stationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            //LogHelper.Info("Api mdataGetWorkplanData: serial number =" + serialNumber + ", station number =" + stationNumber + ", result code =" + errorWP);
            //if (errorWP == 0)
            //{
            //    foreach (var item in workplanDataResultValues)
            //    {
            //        LogHelper.Debug(item);
            //    }
            //    for (int i = 0; i < workplanDataResultValues.Length; i += 2)
            //    {
            //        string strWSNO = workplanDataResultValues[i];
            //        string strPL = workplanDataResultValues[i + 1];
            //        if (Convert.ToInt32(strWSNO) > Convert.ToInt32(currentWS))
            //        {
            //            iProcessLayer = Convert.ToInt32(strPL);
            //            break;
            //        }
            //    }
            //}
            GetStationSettingModel stationSetting = new GetStationSettingModel();
            String[] stationSettingResultKey = new String[] { "BOM_VERSION", "WORKORDER_NUMBER", "PART_NUMBER", "WORKORDER_STATE", "PROCESS_VERSION", "PROCESS_LAYER", "ATTRIBUTE_2", "QUANTITY" };
            String[] stationSettingResultValues;
            LogHelper.Info("begin api trGetStationSetting (Station number:" + init.configHandler.StationNumber + ")");
            int errorCode = imsapi.trGetStationSetting(sessionContext, init.configHandler.StationNumber, stationSettingResultKey, out stationSettingResultValues);
            LogHelper.Info("end api trGetStationSetting (result code = " + errorCode + ")");
            if (errorCode == 0)
            {
                iProcessLayer = int.Parse(stationSettingResultValues[5]);
            }
            LogHelper.Info("Get Process Layer =" + iProcessLayer);
            return iProcessLayer;
        }
    }
}
