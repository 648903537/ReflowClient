using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class GetSerialNumberInfo
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetSerialNumberInfo(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public string[] GetSNInfo(string serialNumber)
        {
            int error = 0;
            string[] serialNumberResultKeys = new string[] { "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" };
            string[] serialNumberResultValues = new string[] { };
            LogHelper.Info("begin api trGetSerialNumberInfo (serial number =" + serialNumber + ")");
            error = imsapi.trGetSerialNumberInfo(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", serialNumberResultKeys, out serialNumberResultValues);
            LogHelper.Info("end api trGetSerialNumberInfo (result code = " + error + ")");
            string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error + "," + errorString, "");
            }
            return serialNumberResultValues;
        }

        public SerialNumberData[] GetPCBSNData(string serialNumber)
        {
            SerialNumberData[] serialNumberArray = new SerialNumberData[] { };
            int errorCode = imsapi.trGetSerialNumberBySerialNumberRef(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", out serialNumberArray);
            LogHelper.Info("Api trGetSerialNumberBySerialNumberRef serial number =" + serialNumber + ", result code =" + errorCode);
            return serialNumberArray;
        }
        public int GetSerialNumberByref(string serialNumber)
        {
            SerialNumberData[] serialNumArray = new SerialNumberData[] { };
            int errorSubSN = imsapi.trGetSerialNumberBySerialNumberRef(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", out serialNumArray);
            LogHelper.Info("trGetSerialNumberBySerialNumberRef(): SN:" + serialNumber + ",ResultCode :" + errorSubSN + "");
            return errorSubSN;
        }
    }
}
