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
    public class LockManager
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public LockManager(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int LockObjects(int objectType, string lockGroupName, string lockInformation, string materialBinNo)
        {
            string[] objectUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER" };
            string[] objectUploadValues = new string[] { "0", materialBinNo };
            string[] objectResultValues = new string[] { };
            int errorCode = imsapi.lockObjects(sessionContext, init.configHandler.StationNumber, objectType, lockGroupName, lockInformation, -1, 0, objectUploadKeys, objectUploadValues, out objectResultValues);
            LogHelper.Info("Api lockObjects object type =" + objectType + ", lock group name =" + lockGroupName + ", material bin number =" + materialBinNo + ", result code =" + errorCode);
            return errorCode;
        }

        public int UnLockObjects(int objectType, string lockGroupName, string unLockInformation, string materialBinNo)
        {
            string[] objectUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER" };
            string[] objectUploadValues = new string[] { "0", materialBinNo };
            string[] objectResultValues = new string[] { };
            int errorCode = imsapi.lockUnlockObjects(sessionContext, init.configHandler.StationNumber, objectType, lockGroupName, unLockInformation, 0, -1, 0, objectUploadKeys, objectUploadValues, out objectResultValues);
            LogHelper.Info("Api lockUnlockObjects object type =" + objectType + ", lock group name =" + lockGroupName + ", material bin number =" + materialBinNo + ", result code =" + errorCode);
            return errorCode;
        }
    }
}
