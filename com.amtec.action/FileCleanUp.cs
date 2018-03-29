using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
namespace com.amtec.action
{
    public class FileCleanUp
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public FileCleanUp(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }
        static string filecleanTime = "";
        public void InitFileCleanUpAttribute()
        {
            if (init.configHandler.FILE_CLEANUP == null || init.configHandler.FILE_CLEANUP == "")
                return;
            GetAttributeValue getattribute = new GetAttributeValue(sessionContext, init, view);
            string[] attributeResultValues = getattribute.GetAttributeValueForAll(7, init.configHandler.StationNumber, "-1", "NEXT_FILE_CLEANUP");
            if (attributeResultValues.Length > 0)
            {
                filecleanTime = attributeResultValues[1];
            }
            else
            {
                DateTime dt = DateTime.Now.AddDays(Convert.ToDouble(init.configHandler.FILE_CLEANUP));
                AppendAttribute appendAttri = new AppendAttribute(sessionContext, init, view);
                appendAttri.AppendAttributeForAll(7, init.configHandler.StationNumber, "-1", "NEXT_FILE_CLEANUP", dt.ToString("yyyy/MM/dd"));
                filecleanTime = dt.ToString("yyyy/MM/dd");
            }
        }

        public void DeleteFolderFile(string deletefiletype)
        {
          if (Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd")) < Convert.ToDateTime(filecleanTime))
                return;
            if (Directory.Exists(init.configHandler.LogTransOK))
            {
                if (deletefiletype == "1"|| deletefiletype == "2")
                {
                    string[] fileNames = Directory.GetFiles(init.configHandler.LogTransOK);
                    if (fileNames != null)
                    {
                        List<FileInfo> fileList = new List<FileInfo>();
                        foreach (var fileName in fileNames)
                        {
                            FileInfo fi = new FileInfo(fileName);
                            fileList.Add(fi);
                        }
                        List<FileInfo> filesOrderedASC = fileList.OrderBy(p => p.LastWriteTime).ToList();

                        foreach (var file in filesOrderedASC)
                        {
                            if (file.LastWriteTime < Convert.ToDateTime(filecleanTime))
                            {

                                file.Delete();
                                LogHelper.Info("delete file:" + file.Name);
                            }
                            else
                                break;
                        }
                    }
                }
                
            }
            if (Directory.Exists(init.configHandler.LogTransError))
            {
                if (deletefiletype == "0"|| deletefiletype == "2")
                {
                    string[] fileNames = Directory.GetFiles(init.configHandler.LogTransError);
                    if (fileNames != null)
                    {
                        List<FileInfo> fileList = new List<FileInfo>();
                        foreach (var fileName in fileNames)
                        {
                            FileInfo fi = new FileInfo(fileName);
                            fileList.Add(fi);
                        }
                        List<FileInfo> filesOrderedASC = fileList.OrderBy(p => p.LastWriteTime).ToList();

                        foreach (var file in filesOrderedASC)
                        {
                            if (file.LastWriteTime < Convert.ToDateTime(filecleanTime))
                            {
                                file.Delete();
                                LogHelper.Info("delete file:" + file.Name);
                            }
                            else
                                break;
                        }
                    }
                }
                
            }
            DateTime dt = DateTime.Now.AddDays(Convert.ToDouble(init.configHandler.FILE_CLEANUP));
            AppendAttribute appendAttri = new AppendAttribute(sessionContext, init, view);
            appendAttri.AppendAttributeForAll(7, init.configHandler.StationNumber, "-1", "NEXT_FILE_CLEANUP", dt.ToString("yyyy/MM/dd"));
            filecleanTime = dt.ToString("yyyy/MM/dd");
        }
    }
}
