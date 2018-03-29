using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class ActivateWorkorder
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private int error;
        private MainView view;
        private InitModel init;

        public ActivateWorkorder(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int ActivateWorkorderResultcall(string workorder, int processLayer)
        {
            int activationResult = imsapi.trActivateWorkOrder(sessionContext, init.configHandler.StationNumber, workorder, "-1", "-1", processLayer, 2);//1 = Activate work order for the station only;2 = Activate work order for entire line
            if (activationResult == 0)
            {
                error = activationResult;
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trActivateWorkOrder " + activationResult, "");
            }
            else
            {
                error = activationResult;
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trActivateWorkOrder " + activationResult, "");
            }
            return activationResult;
        }
    }
}
