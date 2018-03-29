using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Globalization;
using System.Text;

namespace com.amtec.action
{
    public class UtilityFunction
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private static IMSApiSessionContextStruct sessionContext;
        private string StationNumber = "";

        public UtilityFunction(IMSApiSessionContextStruct _sessionContext, string stationNo)
        {
            sessionContext = _sessionContext;
            this.StationNumber = stationNo;
        }

        public DateTime GetServerDateTime()
        {
            var calendarDataResultKeys = new string[] { "CURRENT_TIME_MILLIS" };
            var calendarDataResultValues = new string[] { };
            int error = imsapi.mdataGetCalendarData(sessionContext, StationNumber, calendarDataResultKeys, out calendarDataResultValues);
            if (error != 0)
            {
                LogHelper.Info("API mdataGetCalendarData error code = " + error);
                return DateTime.Now;
            }
            long numer = long.Parse(calendarDataResultValues[0]);
            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(numer).ToLocalTime();
            return date;
        }

        /// <summary>
        /// 字符串编码转换
        /// </summary>
        /// <param name="srcEncoding">原编码</param>
        /// <param name="dstEncoding">目标编码</param>
        /// <param name="srcBytes">原字符串</param>
        /// <returns>字符串</returns>
        public static string TransferEncoding(Encoding srcEncoding, Encoding dstEncoding, string srcStr)
        {
            byte[] srcBytes = srcEncoding.GetBytes(srcStr);
            byte[] bytes = Encoding.Convert(srcEncoding, dstEncoding, srcBytes);
            return dstEncoding.GetString(bytes);
        }

        /// <summary>
        /// 字节数组转为字符串
        /// 将指定的字节数组的每个元素的数值转换为它的等效十六进制字符串表示形式。
        /// </summary>
        /// <param name="bytes"></param>
        /// <returns></returns>
        public static string BitToString(byte[] bytes)
        {
            if (bytes == null)
            {
                return null;
            }
            //将指定的字节数组的每个元素的数值转换为它的等效十六进制字符串表示形式。
            return BitConverter.ToString(bytes);
        }

        /// <summary>
        /// 将十六进制字符串转为字节数组
        /// </summary>
        /// <param name="bitStr"></param>
        /// <returns></returns>
        public static byte[] FromBitString(string bitStr)
        {
            if (bitStr == null)
            {
                return null;
            }

            string[] sInput = bitStr.Split("-".ToCharArray());
            byte[] data = new byte[sInput.Length];
            for (int i = 0; i < sInput.Length; i++)
            {
                data[i] = byte.Parse(sInput[i], NumberStyles.HexNumber);
            }

            return data;
        }
        //Encoding.UTF8.GetString(FromBitString(result)); 

        public string GetTeamNumberByUser(string userName)
        {
            string teamNo = "";
            int resultCode = -1;
            bool hasMore = false;
            string[] mdataGetUserDataKeys = new string[] { "TEAM_NUMBER" };
            string[] mdataGetUserDataValues = new string[] { };
            KeyValue[] mdataGetUserDataFilter = new KeyValue[] { new KeyValue("USER_NAME", userName) };
            resultCode = imsapi.mdataGetUserData(sessionContext, StationNumber, mdataGetUserDataFilter, mdataGetUserDataKeys, out mdataGetUserDataValues, out hasMore);
            LogHelper.Info("Api mdataGetUserData user name =" + userName + " ,result code =" + resultCode);
            if (resultCode == 0)
            {
                teamNo = mdataGetUserDataValues[0];
            }
            return teamNo; ;
        }

        public int GetProcessLayerBySN(string serialNumber, string stationNumber)
        {
            int iProcessLayer = -1;
            string currentWS = "";
            string[] uploadInfoResultKeys = new string[] { "WORKSTEP_NUMBER" };
            string[] uploadInfoResultValues = new string[] { };
            int errorCode = imsapi.trGetSerialNumberUploadInfo(sessionContext, stationNumber, -1, serialNumber, "-1", 0, uploadInfoResultKeys, out uploadInfoResultValues);
            LogHelper.Info("Api trGetSerialNumberUploadInfo: station number =" + stationNumber + ",serial number =" + serialNumber + ", result code =" + errorCode);
            if (errorCode == 0)
            {
                currentWS = uploadInfoResultValues[0];
                LogHelper.Info("work step number :" + currentWS);
            }

            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("FUNC_MODE", "0"), new KeyValue("WORKSTEP_FLAG", "1"), new KeyValue("SERIAL_NUMBER", serialNumber) };
            string[] workplanDataResultKeys = new string[] { "WORKSTEP_NUMBER", "PROCESS_LAYER" };
            string[] workplanDataResultValues = new string[] { };
            int errorWP = imsapi.mdataGetWorkplanData(sessionContext, stationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            LogHelper.Info("Api mdataGetWorkplanData: serial number =" + serialNumber + ", station number =" + stationNumber + ", result code =" + errorWP);
            if (errorWP == 0)
            {
                foreach (var item in workplanDataResultValues)
                {
                    LogHelper.Info(item);
                }
                for (int i = 0; i < workplanDataResultValues.Length; i += 2)
                {
                    string strWSNO = workplanDataResultValues[i];
                    string strPL = workplanDataResultValues[i + 1];
                    if (Convert.ToInt32(strWSNO) > Convert.ToInt32(currentWS))
                    {
                        iProcessLayer = Convert.ToInt32(strPL);
                        break;
                    }
                }
            }
            LogHelper.Info("Get Process Layer =" + iProcessLayer);
            return iProcessLayer;
        }

        public static string GetZHSErrorString(int iErrorCode, InitModel _init, IMSApiSessionContextStruct _sessionContext)
        {
            string errorString = "";
            if (_init.configHandler.Language == "US")
            {
                imsapi.imsapiGetErrorText(sessionContext, iErrorCode, out errorString);
            }
            else
            {
                if (_init.ErrorCodeZHS.ContainsKey(iErrorCode))
                    errorString = _init.ErrorCodeZHS[iErrorCode];
                else
                {
                    imsapi.imsapiGetErrorText(sessionContext, iErrorCode, out errorString);
                }
            }

            return errorString;
        }

    }
}
