using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace com.amtec.action
{
    public class AttributeManager
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public AttributeManager(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int AppendAttributeForAll(string stationNumber, int objectType, string objectNumber, string objectDetail, string attributeCode, string attributeValue)
        {
            int error = 0;
            string[] attributeUploadKeys = new string[] { "ATTRIBUTE_CODE", "ATTRIBUTE_VALUE", "ERROR_CODE" };
            string[] attributeUploadValues = new string[] { attributeCode, attributeValue, "0" };
            string[] attributeResultValues = new string[] { };
            error = imsapi.attribAppendAttributeValues(sessionContext, stationNumber, objectType, objectNumber, objectDetail, -1, 1, attributeUploadKeys, attributeUploadValues, out attributeResultValues);
            LogHelper.Info("Api attribAppendAttributeValues error=" + error + ",object type=" + objectType + ",object number=" + objectNumber + ",object detail=" + objectDetail + ",attribute code=" + attributeCode + ",attribute value=" + attributeValue + ", station number =" + stationNumber);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " attribAppendAttributeValues " + error, "");
            }
            else
            {
                if (attributeResultValues.Length > 0)
                {
                    if (attributeResultValues[2] == "-901")//attribute code not exist, create
                    {
                        imsapi.attribCreateAttribute(sessionContext, init.configHandler.StationNumber, objectType, attributeCode, attributeCode, "N");
                        error = imsapi.attribAppendAttributeValues(sessionContext, init.configHandler.StationNumber, objectType, objectNumber, objectDetail, -1, 1, attributeUploadKeys, attributeUploadValues, out attributeResultValues);
                        string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                        if (error == 0)
                        {
                            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " attribAppendAttributeValues " + error, "");
                        }
                        else
                        {
                            view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " attribAppendAttributeValues " + error + "(" + errorString + ")", "");
                        }
                    }
                }
                else
                {
                    string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                    view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " attribAppendAttributeValues " + error + "(" + errorString + ")", "");
                }
            }
            return error;
        }

        public string[] GetAttributeValueForAll(string stationNumber, int objectType, string objectNumber, string objectDetail, string attributeCode)
        {
            string[] attributeCodeArray = new string[] { attributeCode };
            string[] attributeResultKeys = new string[] { "ATTRIBUTE_CODE", "ATTRIBUTE_VALUE", "ERROR_CODE" };
            string[] attributeResultValues = new string[] { };
            int error = imsapi.attribGetAttributeValues(sessionContext, stationNumber, objectType, objectNumber, objectDetail, attributeCodeArray, 0, attributeResultKeys, out attributeResultValues);
            LogHelper.Info("api attribGetAttributeValues (object type =" + objectType + ",object number =" + objectNumber + ",attribute code =" + attributeCode + ", result code =" + error + ")");
            return attributeResultValues;
        }

        public Dictionary<string, string> GetAttributeValueForAllSec(string stationNumber, int objectType, string objectNumber, string objectDetail, string attributeCode)
        {
            string[] attributeCodeArray = new string[] { attributeCode };
            string[] attributeResultKeys = new string[] { "ATTRIBUTE_CODE", "ATTRIBUTE_VALUE", "ERROR_CODE" };
            string[] attributeResultValues = new string[] { };
            Dictionary<string, string> dicValues = new Dictionary<string, string>();
            int error = imsapi.attribGetAttributeValues(sessionContext, stationNumber, objectType, objectNumber, objectDetail, attributeCodeArray, 0, attributeResultKeys, out attributeResultValues);
            LogHelper.Info("api attribGetAttributeValues (object type =" + objectType + ",object number =" + objectNumber + ",attribute code =" + attributeCode + ", result code =" + error + ")");
            if (error == 0)
            {
                int loop = attributeResultKeys.Length;
                int count = attributeResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    dicValues[attributeResultValues[i]] = attributeResultValues[i + 1];
                }
            }
            return dicValues;
        }

        public int RemoveAttributeForAll(string stationNumber, int objectType, string objectNumber, string objectDetail, string attributeCode)
        {
            int error = imsapi.attribRemoveAttributeValue(sessionContext, stationNumber, objectType, objectNumber, objectDetail, attributeCode, "-1");
            LogHelper.Info("api attribRemoveAttributeValue (object type =" + objectType + ",object number =" + objectNumber + ",attribute code =" + attributeCode + ", error code =" + error + ")");
            return error;
        }

        public string[] GetAttributeValueForWorkStep(string attributeCode, string workplanid, string workstep)
        {
            string[] attributeCodeArray = new string[] { attributeCode };
            string[] attributeResultKeys = new string[] { "ATTRIBUTE_CODE", "ATTRIBUTE_VALUE", "ERROR_CODE" };
            string[] attributeResultValues = new string[] { };
            int error = imsapi.attribGetAttributeValues(sessionContext, init.configHandler.StationNumber, 12, workplanid, workstep, attributeCodeArray, 0, attributeResultKeys, out attributeResultValues);
            LogHelper.Info("api attribGetAttributeValues (workplanid =" + workplanid + ", workstep =" + workstep + ", error code =" + error + ")");
            return attributeResultValues;
        }
    }
}
