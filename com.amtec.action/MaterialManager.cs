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
    public class MaterialManager
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public MaterialManager(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int CreateNewMaterialBin(string materialBinNo, string partNo, string qty, string lotNo)
        {
            string[] materialBinUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "SUPPLIER_CHARGE_NUMBER" };
            string[] materialBinUploadValues = new string[] { "0", materialBinNo, partNo, qty, lotNo };
            string[] materialBinResultValues = new string[] { };
            int error = imsapi.mlCreateNewMaterialBin(sessionContext, init.configHandler.StationNumber, materialBinUploadKeys, materialBinUploadValues, out materialBinResultValues);
            string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("Api mlCreateNewMaterialBin (errorcode = " + error + ",error message = " + errorString + "),material bin number = " + materialBinNo + ", part number =" + partNo + ", quantity =" + qty);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlCreateNewMaterialBin " + error, "");
            }
            else
            {
                foreach (var item in materialBinResultValues)
                {
                    LogHelper.Info(item);
                }
              
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlCreateNewMaterialBin " + error + "(" + errorString + ")", "");
            }
            return error;
        }

        public int SplitMaterialBin(string originalMBN, string newMNB, string qty)
        {
            string[] splitMaterialBinKeys = new string[] { "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_QTY_ACTUAL" };
            string[] splitMaterialBinUploadValues = new string[] { newMNB, qty };
            string[] splitMaterialBinResultKeys = new string[] { "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_QTY_ACTUAL" };
            string[] splitMaterialBinResultValues = new string[] { };
            int error = imsapi.mlSplitMaterialBin(sessionContext, init.configHandler.StationNumber, originalMBN, splitMaterialBinKeys, splitMaterialBinUploadValues, splitMaterialBinResultKeys, out splitMaterialBinResultValues);
            string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("Api mlSplitMaterialBin (errorcode = " + error + ",error message = " + errorString + "),original material bin number = " + originalMBN + ", new material bin number =" + newMNB);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlSplitMaterialBin " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlSplitMaterialBin " + error + "(" + errorString + ")", "");
            }
            return error;
        }

        public string[] GetMaterialBinDataDetails(string materialBinNo, out int errorCode)
        {
            KeyValue[] materialBinFilters = new KeyValue[] { new KeyValue("MATERIAL_BIN_NUMBER", materialBinNo) };
            AttributeInfo[] attributes = new AttributeInfo[] { };
            string[] materialBinResultKeys = new string[] { "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "SUPPLIER_CHARGE_NUMBER", "PART_DESC", "MSL_FLOOR_LIFETIME_REMAIN", "EXPIRATION_DATE" };
            string[] materialBinResultValues = new string[] { };
            LogHelper.Info("begin api mlGetMaterialBinData (Material bin number:" + materialBinNo + ")");
            int error = imsapi.mlGetMaterialBinData(sessionContext, init.configHandler.StationNumber, materialBinFilters, attributes, materialBinResultKeys, out materialBinResultValues);
            LogHelper.Info("end api mlGetMaterialBinData (result code = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            else
            {
                string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error + "(" + errorString + ")", "");
            }
            errorCode = error;
            return materialBinResultValues;
        }

        public string GetPartNumberFromMBN(string materialBinNumber)
        {
            string strPartNumber = "";
            KeyValue[] materialBinFilters = new KeyValue[] { new KeyValue("MATERIAL_BIN_NUMBER", materialBinNumber) };
            AttributeInfo[] attributes = new AttributeInfo[] { };
            string[] materialBinResultKeys = new string[] { "MATERIAL_BIN_PART_NUMBER" };
            string[] materialBinResultValues = new string[] { };
            LogHelper.Info("begin api mlGetMaterialBinData (Material bin number:" + materialBinNumber + ")");
            int error = imsapi.mlGetMaterialBinData(sessionContext, init.configHandler.StationNumber, materialBinFilters, attributes, materialBinResultKeys, out materialBinResultValues);
            LogHelper.Info("end api mlGetMaterialBinData (result code = " + error + ")");
            if (error == 0)
            {
                strPartNumber = materialBinResultValues[0];
                //view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            else
            {
                //view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            return strPartNumber;
        }

        public decimal GetMaterialQty(string materialBinNumber)
        {
            decimal qty = 0;
            KeyValue[] materialBinFilters = new KeyValue[] { new KeyValue("MATERIAL_BIN_NUMBER", materialBinNumber) };
            AttributeInfo[] attributes = new AttributeInfo[] { };
            string[] materialBinResultKeys = new string[] { "MATERIAL_BIN_QTY_ACTUAL" };
            string[] materialBinResultValues = new string[] { };
            LogHelper.Info("begin api mlGetMaterialBinData (Material bin number:" + materialBinNumber + ")");
            int error = imsapi.mlGetMaterialBinData(sessionContext, init.configHandler.StationNumber, materialBinFilters, attributes, materialBinResultKeys, out materialBinResultValues);
            LogHelper.Info("end api mlGetMaterialBinData (result code = " + error + ")");
            if (error == 0)
            {
                qty = Convert.ToDecimal(materialBinResultValues[0]);
                //view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            else
            {
                //view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            return qty;
        }
    }
}
