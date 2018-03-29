using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml;
using System.Linq;

namespace com.amtec.configurations
{
    public class ApplicationConfiguration
    {
        public String StationNumber { get; set; }

        public String Client { get; set; }

        public String RegistrationType { get; set; }

        public String SerialPort { get; set; }

        public String BaudRate { get; set; }

        public String Parity { get; set; }

        public String StopBits { get; set; }

        public String DataBits { get; set; }

        public String NewLineSymbol { get; set; }

        public String High { get; set; }

        public String Low { get; set; }

        public String EndCommand { get; set; }

        public String DLExtractPattern { get; set; }

        public String MBNExtractPattern { get; set; }

        public String MDAPath { get; set; }

        public String EquipmentExtractPattern { get; set; }

        public String OpacityValue { get; set; }

        public String LocationXY { get; set; }

        public String ThawingDuration { get; set; }

        public String LockTime { get; set; }

        public String UsageTime { get; set; }

        public String ThawingCheck { get; set; }

        public String GateKeeperTimer { get; set; }

        public String SolderPasteValidity { get; set; }

        public String IPAddress { get; set; }

        public String Port { get; set; }

        public String OpenControlBox { get; set; }

        public String StencilPrefix { get; set; }

        public String TimerSpan { get; set; }

        public String StartTrigerStr { get; set; }

        public String EndTrigerStr { get; set; }

        public String NoRead { get; set; }

        public String LogFileFolder { get; set; }

        public String LogTransOK { get; set; }

        public String LogTransError { get; set; }

        public String ChangeFileName { get; set; }

        public String CheckListFolder { get; set; }

        public String LogInType { get; set; }

        public String LoadExtractPattern { get; set; }

        public String Language { get; set; }

        public string AUTH_TEAM { get; set; }

        public string FilterByFileName { get; set; }

        public string FileNamePattern { get; set; }

        public string RefreshWO { get; set; }

        public string IPI_STATUS_CHECK { get; set; }

        public string Production_Inspection_CHECK { get; set; }
        public String FILE_CLEANUP { get; set; }

        public String FILE_CLEANUP_TREAD_TIMER { get; set; }
        public String FILE_CLEANUP_FOLDER_TYPE { get; set; }
        public String LIGHT_CHANNEL_ON { get; set; }
        public String LIGHT_CHANNEL_OFF { get; set; }
        public String IO_BOX_CONNECT { get; set; }
        public String IOSerialPort { get; set; }
        public String IOBaudRate { get; set; }
        public String IOParity { get; set; }
        public String IOStopBits { get; set; }
        public String IODataBits { get; set; }
        public String OutputEnter { get; set; }
        public String DataOutputInterface { get; set; }
        public String LAYER_DISPLAY { get; set; }
        public String CHECKCONECTTIME { get; set; }
        public String CEHCKTXTFILE { get; set; }


        //为了给回流焊添加设备,必须在配置中加上如下几个配置对象       郑培聪     20180228
        #region 新版上设备使用

        public String ReduceEquType { get; set; }

        #endregion


        #region checklist
        public String CHECKLIST_IPAddress { get; set; }
        public String CHECKLIST_Port { get; set; }
        public String CHECKLIST_SOURCE { get; set; }
        public String AUTH_CHECKLIST_APP_TEAM { get; set; }
        public String CHECKLIST_FREQ { get; set; }
        public String SHIFT_CHANGE_TIME { get; set; }
        public String RESTORE_TIME { get; set; }
        public String RESTORE_TREAD_TIMER { get; set; }
        #endregion

        Dictionary<string, string> dicConfig = null;

        public ApplicationConfiguration()
        {
            try
            {
                CommonModel commonModel = ReadIhasFileData.getInstance();
                XDocument config = XDocument.Load("ApplicationConfig.xml");
                StationNumber = commonModel.Station;
                Client = commonModel.Client;
                RegistrationType = commonModel.RegisterType;
                SerialPort = GetParameterValues(config, "SerialPort"); 
                BaudRate = GetParameterValues(config, "BaudRate"); 
                Parity = GetParameterValues(config, "Parity");
                StopBits = GetParameterValues(config, "StopBits"); 
                DataBits = GetParameterValues(config, "DataBits"); 
                NewLineSymbol = GetParameterValues(config, "NewLineSymbol"); 
                High = GetParameterValues(config, "High"); 
                Low = GetParameterValues(config, "Low"); 
                EndCommand = GetParameterValues(config, "EndCommand"); 
                DLExtractPattern = GetParameterValues(config, "DLExtractPattern"); 
                MBNExtractPattern = GetParameterValues(config, "MBNExtractPattern"); 
                EquipmentExtractPattern = GetParameterValues(config, "EquipmentExtractPattern"); 
                OpacityValue = GetParameterValues(config, "OpacityValue");
                LocationXY = GetParameterValues(config, "LocationXY");
                NoRead = GetParameterValues(config, "NoRead"); 
                LogFileFolder = GetParameterValues(config, "LogFileFolder"); 
                LogTransOK = GetParameterValues(config, "LogTransOK");
                LogTransError = GetParameterValues(config, "LogTransError"); 
                ChangeFileName = GetParameterValues(config, "ChangeFileName");
                CheckListFolder = GetParameterValues(config, "CheckListFolder"); 
                LoadExtractPattern = GetParameterValues(config, "LoadExtractPattern");
                LogInType = GetParameterValues(config, "LogInType"); 
                Language = GetParameterValues(config, "Language");
                MDAPath = GetParameterValues(config, "MDAPath"); 
                IPAddress = GetParameterValues(config, "IPAddress");
                Port = GetParameterValues(config, "Port");
                FilterByFileName = GetParameterValues(config, "FilterByFileName");
                FileNamePattern = GetParameterValues(config, "FileNamePattern");
                RefreshWO = GetParameterValues(config, "RefreshWO");
                FILE_CLEANUP = GetParameterValues(config,"FILE_CLEANUP");
                IPI_STATUS_CHECK = GetParameterValues(config, "IPI_STATUS_CHECK");
                Production_Inspection_CHECK = GetParameterValues(config, "Production_Inspection_CHECK");
                //FILE_CLEANUP = GetParameterValues(config,"FILE_CLEANUP");
                FILE_CLEANUP_TREAD_TIMER = GetParameterValues(config,"FILE_CLEANUP_TREAD_TIMER");
                FILE_CLEANUP_FOLDER_TYPE = GetParameterValues(config, "FILE_CLEANUP_FOLDER_TYPE");
                LIGHT_CHANNEL_ON = GetParameterValues(config, "LIGHT_CHANNEL_ON");
                LIGHT_CHANNEL_OFF = GetParameterValues(config, "LIGHT_CHANNEL_OFF");
                IO_BOX_CONNECT = GetParameterValues(config,"IO_BOX_CONNECT");
                ReduceEquType = GetParameterValues(config, "ReduceEquType");//config.Descendants("ReduceEquType").First().Value;
                StencilPrefix = GetParameterValues(config, "StencilPrefix");//config.Descendants("StencilPrefix").First().Value;
                if (IO_BOX_CONNECT != null && IO_BOX_CONNECT.Split(';').Length >= 6)
                {
                    string[] infos = IO_BOX_CONNECT.Split(';');
                    IOSerialPort = "COM" + infos[0];
                    IOBaudRate = infos[1];
                    IOStopBits = infos[4];
                    IODataBits = infos[2];
                    IOParity = infos[3];
                }
                OutputEnter = GetParameterValues(config, "OutputEnter");
                DataOutputInterface = GetParameterValues(config, "DataOutputInterface");
                LAYER_DISPLAY = GetParameterValues(config,"LAYER_DISPLAY");
                CHECKCONECTTIME = GetParameterValues(config, "CHECKCONECTTIME");
                CEHCKTXTFILE = GetParameterValues(config, "CEHCKTXTFILE");

                #region checklist
                CHECKLIST_IPAddress = GetParameterValues(config, "CHECKLIST_IPAddress");
                CHECKLIST_Port = GetParameterValues(config, "CHECKLIST_Port");
                CHECKLIST_SOURCE = GetParameterValues(config, "CHECKLIST_SOURCE");
                AUTH_CHECKLIST_APP_TEAM = GetParameterValues(config, "AUTH_CHECKLIST_APP_TEAM");
                CHECKLIST_FREQ = GetParameterValues(config, "CHECKLIST_FREQ");
                SHIFT_CHANGE_TIME = GetParameterValues(config, "SHIFT_CHANGE_TIME");
                RESTORE_TIME = GetParameterValues(config, "RESTORE_TIME");
                RESTORE_TREAD_TIMER = GetParameterValues(config, "RESTORE_TREAD_TIMER");
                AUTH_TEAM = GetParameterValues(config, "AUTH_TEAM");
                #endregion
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }

        public ApplicationConfiguration(IMSApiSessionContextStruct sessionContext, MainView view)
        {
            try
            {
                dicConfig = new Dictionary<string, string>();
                ConfigManage configHandler = new ConfigManage(sessionContext, view);
                CommonModel commonModel = ReadIhasFileData.getInstance();
                if (commonModel.UpdateConfig == "L")
                {
                    XDocument config = XDocument.Load("ApplicationConfig.xml");
                    StationNumber = commonModel.Station;
                    Client = commonModel.Client;
                    RegistrationType = commonModel.RegisterType;
                    SerialPort = GetParameterValues(config, "SerialPort");
                    BaudRate = GetParameterValues(config, "BaudRate");
                    Parity = GetParameterValues(config, "Parity");
                    StopBits = GetParameterValues(config, "StopBits");
                    DataBits = GetParameterValues(config, "DataBits");
                    NewLineSymbol = GetParameterValues(config, "NewLineSymbol");
                    High = GetParameterValues(config, "High");
                    Low = GetParameterValues(config, "Low");
                    EndCommand = GetParameterValues(config, "EndCommand");
                    DLExtractPattern = GetParameterValues(config, "DLExtractPattern");
                    MBNExtractPattern = GetParameterValues(config, "MBNExtractPattern");
                    EquipmentExtractPattern = GetParameterValues(config, "EquipmentExtractPattern");
                    OpacityValue = GetParameterValues(config, "OpacityValue");
                    LocationXY = GetParameterValues(config, "LocationXY");
                    NoRead = GetParameterValues(config, "NoRead");
                    LogFileFolder = GetParameterValues(config, "LogFileFolder");
                    LogTransOK = GetParameterValues(config, "LogTransOK");
                    LogTransError = GetParameterValues(config, "LogTransError");
                    ChangeFileName = GetParameterValues(config, "ChangeFileName");
                    CheckListFolder = GetParameterValues(config, "CheckListFolder");
                    LoadExtractPattern = GetParameterValues(config, "LoadExtractPattern");
                    LogInType = GetParameterValues(config, "LogInType");
                    Language = GetParameterValues(config, "Language");
                    MDAPath = GetParameterValues(config, "MDAPath");
                    IPAddress = GetParameterValues(config, "IPAddress");
                    Port = GetParameterValues(config, "Port");
                    FileNamePattern = GetParameterValues(config, "FileNamePattern");
                    FilterByFileName = GetParameterValues(config, "FilterByFileName");
                    RefreshWO = GetParameterValues(config, "RefreshWO");
                    IPI_STATUS_CHECK = GetParameterValues(config, "IPI_STATUS_CHECK");
                    Production_Inspection_CHECK = GetParameterValues(config, "Production_Inspection_CHECK");
                    FILE_CLEANUP = GetParameterValues(config, "FILE_CLEANUP");
                    FILE_CLEANUP_TREAD_TIMER = GetParameterValues(config,"FILE_CLEANUP_TREAD_TIMER");
                    FILE_CLEANUP_FOLDER_TYPE = GetParameterValues(config, "FILE_CLEANUP_FOLDER_TYPE");
                    LIGHT_CHANNEL_ON = GetParameterValues(config, "LIGHT_CHANNEL_ON");
                    LIGHT_CHANNEL_OFF = GetParameterValues(config, "LIGHT_CHANNEL_OFF");
                    IO_BOX_CONNECT = GetParameterValues(config, "IO_BOX_CONNECT");
                    ReduceEquType = GetParameterValues(config ,"ReduceEquType" );// config.Descendants("ReduceEquType").First().Value;
                    StencilPrefix = GetParameterValues(config, "StencilPrefix");//config.Descendants("StencilPrefix").First().Value;
                    if (IO_BOX_CONNECT != null && IO_BOX_CONNECT.Split(';').Length >= 6)
                    {
                        string[] infos = IO_BOX_CONNECT.Split(';');
                        IOSerialPort = "COM" + infos[0];
                        IOBaudRate = infos[1];
                        IOStopBits = infos[4];
                        IODataBits = infos[2];
                        IOParity = infos[3];
                    }
                    OutputEnter = GetParameterValues(config, "OutputEnter");
                    DataOutputInterface = GetParameterValues(config, "DataOutputInterface");
                    LAYER_DISPLAY = GetParameterValues(config, "LAYER_DISPLAY");
                    CHECKCONECTTIME = GetParameterValues(config, "CHECKCONECTTIME");
                    CEHCKTXTFILE = GetParameterValues(config, "CEHCKTXTFILE");

                    #region checklist
                    CHECKLIST_IPAddress = GetParameterValues(config, "CHECKLIST_IPAddress");
                    CHECKLIST_Port = GetParameterValues(config, "CHECKLIST_Port");
                    CHECKLIST_SOURCE = GetParameterValues(config, "CHECKLIST_SOURCE");
                    AUTH_CHECKLIST_APP_TEAM = GetParameterValues(config, "AUTH_CHECKLIST_APP_TEAM");
                    CHECKLIST_FREQ = GetParameterValues(config, "CHECKLIST_FREQ");
                    SHIFT_CHANGE_TIME = GetParameterValues(config, "SHIFT_CHANGE_TIME");
                    RESTORE_TIME = GetParameterValues(config, "RESTORE_TIME");
                    RESTORE_TREAD_TIMER = GetParameterValues(config, "RESTORE_TREAD_TIMER");
                    AUTH_TEAM = GetParameterValues(config, "AUTH_TEAM");
                    #endregion
                }
                else
                {
                    if (commonModel.UpdateConfig == "Y")
                    {
                        //int error = configHandler.DeleteConfigParameters(commonModel.APPTYPE);
                        //if (error == 0 || error == -3303 || error == -3302)
                        //{
                        //    WriteParameterToiTac(configHandler);
                        //}
                        string[] parametersValue = configHandler.GetParametersForScope(commonModel.APPTYPE);
                        if (parametersValue != null && parametersValue.Length > 0)
                        {
                            foreach (var parameterID in parametersValue)
                            {
                                configHandler.DeleteConfigParametersExt(parameterID);
                            }
                        }
                        WriteParameterToiTac(configHandler);
                    }
                    List<ConfigEntity> getvalues = configHandler.GetConfigData(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station);
                    if (getvalues != null)
                    {
                        foreach (var item in getvalues)
                        {
                            if (item != null)
                            {
                                string[] strs = item.PARAMETER_NAME.Split(new char[] { '.' });
                                dicConfig.Add(strs[strs.Length - 1], item.CONFIG_VALUE);
                            }
                        }
                    }

                    StationNumber = commonModel.Station;
                    Client = commonModel.Client;
                    RegistrationType = commonModel.RegisterType;
                    SerialPort = GetParameterValue("SerialPort");
                    BaudRate = GetParameterValue("BaudRate");
                    Parity = GetParameterValue("Parity");
                    StopBits = GetParameterValue("StopBits");
                    DataBits = GetParameterValue("DataBits");
                    NewLineSymbol = GetParameterValue("NewLineSymbol");
                    High = GetParameterValue("High");
                    Low = GetParameterValue("Low");
                    EndCommand = GetParameterValue("EndCommand");
                    DLExtractPattern = GetParameterValue("DLExtractPattern");
                    MBNExtractPattern = GetParameterValue("MBNExtractPattern");
                    EquipmentExtractPattern = GetParameterValue("EquipmentExtractPattern");
                    OpacityValue = GetParameterValue("OpacityValue");
                    LocationXY = GetParameterValue("LocationXY");
                    ThawingDuration = GetParameterValue("ThawingDuration");
                    ThawingCheck = GetParameterValue("ThawingCheck");
                    LockTime = GetParameterValue("LockOutTime");
                    UsageTime = GetParameterValue("UsageDurationSetting");
                    GateKeeperTimer = GetParameterValue("GateKeeperTimer");
                    SolderPasteValidity = GetParameterValue("SolderPasteValidity");
                    OpenControlBox = GetParameterValue("OpenControlBox");
                    StencilPrefix = GetParameterValue("StencilPrefix");
                    TimerSpan = GetParameterValue("TimerSpan");
                    StartTrigerStr = GetParameterValue("StartTrigerStr");
                    EndTrigerStr = GetParameterValue("EndTrigerStr");
                    NoRead = GetParameterValue("NoRead");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private string GetParameterValue(string parameterName)
        {
            if (dicConfig.ContainsKey(parameterName))
            {
                return dicConfig[parameterName];
            }
            else
            {
                return "";
            }
        }

        private string GetParameterValues(XDocument config, string parameterName)
        {
            string value = null;
            if (config.Descendants(parameterName).FirstOrDefault() == null)
            {
                value = "";
                LogHelper.Info("Parameter is not exist." + parameterName);
            }
            else
            {
                value = config.Descendants(parameterName).FirstOrDefault().Value;
            }
            return value;
        }

        private int GetIntValue(string text)
        {
            int value = 0;
            if (string.IsNullOrEmpty(text))
            {

            }
            else
            {
                value = Convert.ToInt32(text);
            }
            return value;
        }

        private void WriteParameterToiTac(ConfigManage configHandler)
        {
            GetApplicationDatas getData = new GetApplicationDatas();
            List<ParameterEntity> entityList = getData.GetApplicationEntity();
            string[] strs = GetParameterString(entityList);
            string[] strvalues = GetValueString(entityList);
            if (strs != null && strs.Length > 0)
            {
                int errorCode = configHandler.CreateConfigParameter(strs);
                if (errorCode == 0 || errorCode == 5)
                {
                    CommonModel commonModel = ReadIhasFileData.getInstance();
                    int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
                }
            }

            //if (entityList.Count > 0)
            //{
            //    List<ParameterEntity> entitySubList = null;
            //    CommonModel commonModel = ReadIhasFileData.getInstance();
            //    foreach (var entity in entityList)
            //    {
            //        entitySubList = new List<ParameterEntity>();
            //        entitySubList.Add(entity);
            //        string[] strs = GetParameterString(entitySubList);
            //        string[] strvalues = GetValueString(entitySubList);
            //        if (strs != null && strs.Length > 0)
            //        {
            //            int errorCode = configHandler.CreateConfigParameter(strs);
            //            if (errorCode == 0 || errorCode == 5)
            //            {                           
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //            else if (errorCode == -3301)//Parameter already exists
            //            {
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //        }
            //    }
            //}
        }

        private string[] GetParameterString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                strList.Add(entity.PARAMETER_DESCRIPTION);
                strList.Add(entity.PARAMETER_DIMPATH);
                strList.Add(entity.PARAMETER_DISPLAYNAME);
                strList.Add(entity.PARAMETER_NAME);
                strList.Add(entity.PARAMETER_PARENT_NAME);
                strList.Add(entity.PARAMETER_SCOPE);
                strList.Add(entity.PARAMETER_TYPE_NAME);
            }
            return strList.ToArray();
        }

        private string[] GetValueString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                if (entity.PARAMETER_VALUE == "")
                    continue;
                strList.Add(entity.PARAMETER_VALUE);
                strList.Add(entity.PARAMETER_NAME);

            }
            return strList.ToArray();
        }
    }
}
