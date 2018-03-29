using com.amtec.action;
using com.amtec.configurations;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using com.itac.oem.common.container.imsapi.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Timers;
using Reflow.com.amtec.forms;
using com.amtec.device;

namespace com.amtec.forms 
{
    public partial class MainView : Form
    {
        public ApplicationConfiguration config;
        IMSApiSessionContextStruct sessionContext;
        public bool isScanProcessEnabled = false;
        private InitModel initModel;
        private LanguageResources res;
        public string UserName = "";
        private string indata = "";
        private DateTime PFCStartTime = DateTime.Now;
        List<SerialNumberData> serialNumberArray = new List<SerialNumberData>();
        public delegate void HandleInterfaceUpdateTopMostDelegate(string sn, string message);
        public HandleInterfaceUpdateTopMostDelegate topmostHandle;
        public TopMostForm topmostform = null;
        CommonModel commonModel = null;
        public string CaptionName = "";
        private System.Timers.Timer CheckConnectTimer = new System.Timers.Timer();
        private int iProcessLayer = 2;
        private System.Timers.Timer FileCleanUpTimerdelete = null;

        private System.Timers.Timer RestoreMaterialTimer = null;
        string Supervisor_OPTION = "1";
        string IPQC_OPTION = "1";
        private SocketClientHandler2 checklist_cSocket = null;
        bool isStartLineCheck = true;//开线点检已经获取=true. 过程点检=false

        //记录当前设备的工单号,当工单改变时,刷新"刷新工单"按钮时会更新设备,否则则不然      郑培聪     20180228
        string workOrderForEquipment = string.Empty;


        #region Init
        public MainView(string userName, DateTime dTime, IMSApiSessionContextStruct _sessionContext)
        {
            InitializeComponent();
            sessionContext = _sessionContext;
            UserName = userName;
            commonModel = ReadIhasFileData.getInstance();
            this.lblLoginTime.Text = dTime.ToString("yyyy/MM/dd HH:mm:ss");
            this.lblUser.Text = userName == "" ? commonModel.Station : userName;
            this.lblStationNO.Text = commonModel.Station;
        }

        private void MainView_Shown(object sender, EventArgs e)
        {
            BackgroundWorker bgWorker = new BackgroundWorker();
            bgWorker.DoWork += new DoWorkEventHandler(bgWorkerInit);
            bgWorker.RunWorkerAsync();
        }
        bool isOK = true;
        private void bgWorkerInit(object sender, DoWorkEventArgs args)
        {
            errorHandler(0, "Application start...", "");
            errorHandler(0, "Version :" + Assembly.GetExecutingAssembly().GetName().Version.ToString(), "");
            res = new LanguageResources();
            config = new ApplicationConfiguration(sessionContext, this);
            InitializeMainGUI init = new InitializeMainGUI(sessionContext, config, this, res);
          
            this.InvokeEx(x =>
            {
                //this.tabDocument.Parent = null;
                this.Text = res.MAIN_TITLE + " (" + Assembly.GetExecutingAssembly().GetName().Version.ToString() + ")";
                CaptionName = res.MAIN_TITLE + System.Environment.NewLine + config.StationNumber;
                this.tabLCR.Parent = null;
                //this.tabCheckList.Parent = null;
                this.tabSetup.Parent = null;

                //添加设备页签      郑培聪     20180227
                //this.tabEquipment.Parent = null;
                

                //this.tabDocument.Parent = null;
                if (config.RefreshWO != "Y")
                {
                    this.btnRefreshWo.Visible = false;
                }
                #region add by qy
                SystemVariable.CurrentLangaugeCode = config.Language;
                InitCintrolLanguage();
                #endregion
                Application.DoEvents();
                cSocket = new SocketClientHandler(this);
                bool isOK = cSocket.connect(config.IPAddress, config.Port);
                initModel = init.Initialize();
                if (isOK)
                {
                    GetTimerStart();
                    FileCleanUp fileup = new FileCleanUp(sessionContext, initModel, this);
                    fileup.InitFileCleanUpAttribute();
                    GetFileCleanUpTimerStart();
                    LoadYield();
                    //InitSetupGrid();
                    //InitTaskData();
                    if (config.AUTH_CHECKLIST_APP_TEAM != "" && config.AUTH_CHECKLIST_APP_TEAM != null)
                    {
                        string[] teams = config.AUTH_CHECKLIST_APP_TEAM.Split(';');
                        string[] items = teams[0].Split(',');
                        string Super = items[0];
                        Supervisor_OPTION = items[1];
                        string[] IPQCitems = teams[1].Split(',');
                        string IP = IPQCitems[0];
                        IPQC_OPTION = IPQCitems[1];
                    }


                    //添加读取上设备配置文件       郑培聪     20180228
                    if (config.RESTORE_TIME != "" && config.RESTORE_TREAD_TIMER != "")
                    {
                        //GetRestoreTimerStart();
                        ReadRestoreFile();
                    }


                    if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")//20161208 edit by qy
                    {
                        InitShift2(txbCDAMONumber.Text);
                        InitWorkOrderType();
                        this.tabCheckList.Parent = null;
                        checklist_cSocket = new SocketClientHandler2(this);
                        isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                        if (isOK)
                        {
                            //if (!CheckShiftChange2())
                            //{
                            //    InitTaskData_SOCKET("开线点检;设备点检");
                            //    isStartLineCheck = true;
                            //}
                            //else
                            //{
                            //    if (!ReadCheckListFile())//20161214 edit by qy
                            //    {
                            //        InitTaskData_SOCKET("开线点检");
                            //        isStartLineCheck = true;
                            //    }
                            //}
                        }
                    }
                    else
                    {
                        InitTaskData();
                        this.tabCheckListTable.Parent = null;
                    }
                    InitCheckResultMapping();

                    //获取设备信息        郑培聪     20180227
                    InitEquipmentGridEXT();
                    workOrderForEquipment = txbCDAMONumber.Text;

                    InitDocumentGrid();
                    //StrippedEquipmentFromStation();
                    ShowTopWindow();
                    SetTipMessage(MessageType.OK, Message("msg_Initialize Success"));
                    this.txbCDADataInput.Focus();
                   
                }
                else
                {
                    errorHandler(3, config.IPAddress + "connect faill", "");
                }
            });
          
        }

        #region add by qy
        private void InitCintrolLanguage()
        {
            MutiLanguages lang = new MutiLanguages();
            foreach (Control ctl in this.Controls)
            {
                lang.InitLangauge(ctl);
                if (ctl is TabControl)
                {
                    lang.InitLangaugeForTabControl((TabControl)ctl);
                }
            }

            //Controls不包含ContextMenuStrip，可用以下方法获得
            System.Reflection.FieldInfo[] fieldInfo = this.GetType().GetFields(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            for (int i = 0; i < fieldInfo.Length; i++)
            {
                switch (fieldInfo[i].FieldType.Name)
                {
                    case "ContextMenuStrip":
                        ContextMenuStrip contextMenuStrip = (ContextMenuStrip)fieldInfo[i].GetValue(this);
                        lang.InitLangauge(contextMenuStrip);
                        break;
                }
            }
        }

        public string Message(string messageId)
        {
            return MutiLanguages.ParserString("$" + messageId);
        }
        #endregion
        #endregion

        #region delegate
        public delegate void errorHandlerDel(int typeOfError, String logMessage, String labelMessage);
        public void errorHandler(int typeOfError, String logMessage, String labelMessage)
        {
            if (txtConsole.InvokeRequired)
            {
                errorHandlerDel errorDel = new errorHandlerDel(errorHandler);
                Invoke(errorDel, new object[] { typeOfError, logMessage, labelMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        LogHelper.Info(logMessage);
                        break;
                    case 1:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Blue;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        LogHelper.Info(logMessage);
                        break;
                    case 2:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        SetTopWindowMessage("Error", logMessage);
                        LogHelper.Error(logMessage);
                        break;
                    case 3:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        SetTopWindowMessage("Error", logMessage);
                        LogHelper.Error(logMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        LogHelper.Error(logMessage);
                        break;
                }
                SetStatusLabelText(logMessage);
                txtConsole.AppendText(errorBuilder);
                txtConsole.ScrollToCaret();
            }
        }

        public delegate void SetTipMessageDel(MessageType strType, string strMessage);
        private void SetTipMessage(MessageType strType, string strMessage)
        {
            if (this.messageControl1.InvokeRequired)
            {
                SetTipMessageDel messageDel = new SetTipMessageDel(SetTipMessage);
                Invoke(messageDel, new object[] { strType, strMessage });
            }
            else
            {
                switch (strType)
                {
                    case MessageType.OK:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Error:
                        this.messageControl1.BackColor = Color.Red;
                        this.messageControl1.PicType = @"pic\Close.png";
                        this.messageControl1.Title = "Error Message";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Instruction:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\Instruction.png";
                        this.messageControl1.Title = "Instruction";
                        this.messageControl1.Content = strMessage;
                        break;
                    default:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                }
            }
        }

        public delegate void SetConnectionTextDel(int typeOfError, string strMessage);
        public void SetConnectionText(int typeOfError, string strMessage)
        {
            if (txtConnection.InvokeRequired)
            {
                SetConnectionTextDel connectDel = new SetConnectionTextDel(SetConnectionText);
                Invoke(connectDel, new object[] { typeOfError, strMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConnection.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        if (strMessage.Contains("PING") || strMessage.Contains("PONG"))
                        {
                            break;
                        }
                        LogHelper.Info(strMessage);
                        break;
                    case 1:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        LogHelper.Error(strMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        break;
                }

                txtConnection.AppendText(errorBuilder);
                txtConnection.ScrollToCaret();
            }
        }

        public void SetStatusLabelText(string strText)
        {
            this.InvokeEx(x => this.lblStatus.Text = strText);
        }

        public string GetWorkOrderValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAMONumber.Text);
            return str;
        }

        public string GetPartNumberValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAPartNumber.Text);
            return str;
        }

        public void SetWorkorderValue(string strText)
        {
            this.InvokeEx(x => this.txbCDAMONumber.Text = strText);
        }

        public void SetPartNumberValue(string strText)
        {
            this.InvokeEx(x => this.txbCDAPartNumber.Text = strText);
        }

        public void SetLCRInfoValue(string strText)
        {
            this.InvokeEx(x => this.txtLCRInfo.Text = strText);
        }

        public void SetLCRTypeValue(string strText)
        {
            this.InvokeEx(x => this.txtLCRType.Text = strText);
        }

        public void SetLCRValueValue(string strText)
        {
            this.InvokeEx(x => this.txtLCRValue.Text = strText);
        }

        public void SetContainerNoValue(string strText)
        {
            this.InvokeEx(x => this.txtContainerNo.Text = strText);
        }

        public void SetDataInputValue(string strText)
        {
            this.InvokeEx(x => this.txbCDADataInput.Text = strText);
        }

        public TextBox getFieldPartNumber()
        {
            return this.txbCDAPartNumber;
        }

        public TextBox getFieldWorkorder()
        {
            return this.txbCDAMONumber;
        }

        public Label getFieldLabelUser()
        {
            return lblUser;
        }

        public Label getFieldLabelTime()
        {
            return lblLoginTime;
        }
        #endregion

        public void DataRecivedHeandlerlight(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            try
            {
                if (!CheckCheckList())
                {
                    return;
                }
                Thread.Sleep(200);
                Byte[] bt1 = new Byte[sp.BytesToRead];
                sp.Read(bt1, 0, sp.BytesToRead);
                indata = System.Text.Encoding.ASCII.GetString(bt1).Trim();
                //indata = sp.ReadLine();
                LogHelper.Info("Scan number(original): " + indata);

                 
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + ";" + ex.StackTrace);
            }
            finally
            {
                initModel.scannerHandler.handler2().DiscardInBuffer();
            }
        }


        #region Data process function
        public void DataRecivedHeandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            try
            {
                if (isFormOutPoump)
                {
                    return;
                }
                if (!VerifyCheckList())
                {
                    return;
                }
                Thread.Sleep(200);
                Byte[] bt = new Byte[sp.BytesToRead];
                sp.Read(bt, 0, sp.BytesToRead);
                indata = System.Text.Encoding.ASCII.GetString(bt).Trim();
                //indata = sp.ReadLine();
                LogHelper.Info("Scan number(original): " + indata);
              
                /*
                 * 
                indata = indata.Replace("?", "").Replace("K", "").Replace("X", "").Replace("H", "").Replace("Q", "").Replace("T1", "").Replace("D1", "").Trim();
                if (indata.Length <= 2)
                {
                    initModel.scannerHandler.handler().DiscardInBuffer();
                    return;
                }

                SetDataInputValue(indata);
                if (indata.TrimEnd() == config.NoRead)
                {
                    initModel.scannerHandler.sendLowExt();
                    errorHandler(2, Message("msg_NO READ"), "");
                    initModel.scannerHandler.handler().DiscardInBuffer();
                    return;
                }
                //match material bin number
                Match match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, Message("msg_Scan material bin number"));
                    }));
                    ProcessMaterialBinNo(match.ToString());
                    return;
                }
                //match equipment
                match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, Message("msg_Scan equipment number"));
                    }));
                    ProcessEquipmentData(match.ToString());
                    return;
                }
                //match serial number
                match = Regex.Match(indata, config.DLExtractPattern);
                if (match.Success)
                {
                    if (CheckEquipmentSetup() && CheckMaterialSetUp())
                    {
                        this.Invoke(new MethodInvoker(delegate
                        {
                            SetTipMessage(MessageType.OK, Message("msg_Scan serial number"));
                        }));
                        ProcessSerialNumber(match.ToString());
                        return;
                    }
                }
                errorHandler(3, Message("msg_wrong barcode"), "");           
                 */
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + ";" + ex.StackTrace);
            }
            finally
            {
                initModel.scannerHandler.handler().DiscardInBuffer();
            }
        }
        #endregion

        #region Event
        private void MainView_Load(object sender, EventArgs e)
        {
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath) + @"\pic\Chart_Column_Silver.png";
            NetworkChange.NetworkAvailabilityChanged += AvailabilityChanged;
        }

        private void MainView_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show(Message("msg_Do you want to close the application"), Message("msg_Quit Application"), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.OK)
            {
                if (this.txbCDAMONumber.Text != "")
                {

                    //启用设备注销        郑培聪     20180228
                    EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                    foreach (DataGridViewRow row in dgvEquipment.Rows)
                    {
                        string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                        string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                        if (string.IsNullOrEmpty(equipmentNo))
                            continue;
                        int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                        RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                    }

                    //SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                    //setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 2);
                }
                LogHelper.Info("Application end...");
                System.Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void txbCDADataInput_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {

                if (!VerifyCheckList())
                {
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    //errorHandler(3, Message("$msg_checklist_first"), "");
                    return;
                }
                indata = this.txbCDADataInput.Text.Trim();
                /*
                //match material bin number
                Match match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, Message("msg_Scan material bin number"));
                    }));
                    ProcessMaterialBinNo(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }*/


                //match equipment       回流焊启用机台样件验证页签功能     郑培聪     2018/02/27
                #region 上设备
                Match match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, Message("msg_Scan equipment number"));
                    }));
                    ProcessEquipmentDataEXT(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                #endregion

                /*
                //match serial number
                match = Regex.Match(indata, config.DLExtractPattern);
                if (match.Success)
                {
                    if (CheckEquipmentSetup() && CheckMaterialSetUp())
                    {
                        this.Invoke(new MethodInvoker(delegate
                        {
                            //this.tabControl1.SelectedTab = this.tabShipping;
                            SetTipMessage(MessageType.OK, Message("msg_Scan serial number"));
                        }));
                        //ProcessSerialNumber(match.ToString());
                        this.txbCDADataInput.SelectAll();
                        this.txbCDADataInput.Focus();
                        return;
                    }
                }
                this.txbCDADataInput.SelectAll();
                this.txbCDADataInput.Focus();
                errorHandler(3, Message("msg_wrong barcode"), "");
                 */
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedTab.Name == "tabSetup")
            {
                this.gridSetup.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabEquipment")
            {
                this.dgvEquipment.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabDocument")
            {
                this.gridDocument.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabSNLog")
            {
                this.gridSNLog.ClearSelection();
            }
        }

        private void gridDocument_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            long documentID = Convert.ToInt64(gridDocument.Rows[e.RowIndex].Cells[0].Value.ToString());
            string fileName = gridDocument.Rows[e.RowIndex].Cells[1].Value.ToString();
            SetDocumentControlForDoc(documentID, fileName);

        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (this.dgvEquipment.SelectedRows.Count > 0)
            {
                DataGridViewRow row = this.dgvEquipment.SelectedRows[0];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["EquipmentIndex"].Value = "";
                row.Cells["Status"].Value = ReflowClient.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                //remove attribute "attribEquipmentHasRigged"
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                this.dgvEquipment.ClearSelection();
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, this);
            initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
            if (initModel.currentSettings != null && initModel.currentSettings.workorderNumber != this.txbCDAMONumber.Text)
            {
                this.gridSetup.Rows.Clear();
                //this.dgvEquipment.Rows.Clear();
                this.txbCDAMONumber.Text = initModel.currentSettings.workorderNumber;
                this.txbCDAPartNumber.Text = initModel.currentSettings.partNumber;
                LoadYield();
                InitSetupGrid();
                //InitEquipmentGrid();
                ShowTopWindow();
                SetTipMessage(MessageType.OK, Message("msg_Refresh Success"));
                this.txbCDADataInput.Focus();
            }
            if (initModel.currentSettings == null)
            {
                this.gridSetup.Rows.Clear();
                //this.dgvEquipment.Rows.Clear();
                this.txbCDAMONumber.Text = "";
                this.txbCDAPartNumber.Text = "";
            }
        }

        int iIndexItem = -1;
        private void dgvEquipment_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvEquipment.Rows.Count == 0)
                    return;
                this.dgvEquipment.ContextMenuStrip = contextMenuStrip1;
                iIndexItem = ((DataGridView)sender).CurrentRow.Index;
            }
        }

        private void removeEquipment_Click(object sender, EventArgs e)
        {
            if (iIndexItem > -1)
            {
                DataGridViewRow row = this.dgvEquipment.Rows[iIndexItem];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["ScanTime"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["Status"].Value = ReflowClient.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                this.dgvEquipment.ClearSelection();
            }
        }

        private void btnPassBoard_Click(object sender, EventArgs e)
        {
            PassBoard();
        }

        private void btnFetch_Click(object sender, EventArgs e)
        {
            //LCRSendData("function:impa?");
            //Thread.Sleep(1000);
            //LCRSendData("fetch?");
        }
        #endregion

        #region Listen file
        public void ListenFile(string filePath)
        {
            //if (!VerifyCheckList())
            //{
            //    return;
            //}
            if (isFormOutPoump)
            {
                return;
            }
            LogHelper.Info("start analysis file " + filePath);
            if (!File.Exists(filePath))
            {
                errorHandler(0, filePath + "not exist", "");
                return;
            }
            string[] lines = File.ReadAllLines(filePath, Encoding.Default);
            List<string> lineList = new List<string>();
            foreach (var line in lines)
            {
                lineList.Add(line);
            }
            try
            {


                int firstArea = lineList.IndexOf("[Process Info]");
                int secondArea = lineList.IndexOf("[Statistics Limits]");
                int thirdArea = lineList.IndexOf("[Baseline Profile Info]");
                int fourthArea = lineList.IndexOf("[VP_DATA]");
                string[] partValues = lineList[1].Split(new char[] { '=' });
                string partNumber = partValues[1];
                string[] values = partNumber.Split(new char[] { '-' });
                string strPL = values[2].Substring(0, 1);
                int iProcessLayer = strPL == "T" ? 0 : 1;
                string serialNumber = lineList[fourthArea + 4].Split(new char[] { '=' })[1];
                if (string.IsNullOrEmpty(serialNumber))
                {
                    MoveFileToErrorFolder(filePath, "serial number cann't be empty");
                    errorHandler(2, "Analysis file fail", "");
                    return;
                }
                //get measurement standard value
                Dictionary<string, MeasurementEntity> dicMeasureEntity = new Dictionary<string, MeasurementEntity>();
                for (int i = secondArea; i < thirdArea; i++)
                {
                    string line = lineList[i].Trim();
                    if (line == "[Statistics Limits]" || line == "[Baseline Profile Info]" || line == "")
                        continue;
                    string[] values1 = line.Split(new char[] { '=' });
                    string strValue = values1[1];
                    string strName = values1[0];
                    string strMeasureName = strName.Substring(0, strName.LastIndexOf('_'));
                    string strTemp = strName.Substring(strName.LastIndexOf('_') + 1);
                    MeasurementEntity entity = null;
                    if (dicMeasureEntity.ContainsKey(strMeasureName))
                    {
                        entity = dicMeasureEntity[strMeasureName];
                        switch (strTemp)
                        {
                            case "LOW":
                                entity.MinValue = strValue;
                                break;
                            case "HIGH":
                                entity.MaxValue = strValue;
                                break;
                            case "TARGET":
                                entity.TargetValue = strValue;
                                break;
                            default:
                                break;
                        }
                    }
                    else
                    {
                        entity = new MeasurementEntity();
                        entity.MeasurementName = strMeasureName;
                        switch (strTemp)
                        {
                            case "LOW":
                                entity.MinValue = strValue;
                                break;
                            case "HIGH":
                                entity.MaxValue = strValue;
                                break;
                            case "TARGET":
                                entity.TargetValue = strValue;
                                break;
                            default:
                                break;
                        }
                        dicMeasureEntity[strMeasureName] = entity;
                    }
                }
                //measurement data without max & min
                List<string> measurementData1 = new List<string>();
                for (int j = thirdArea + 4; j < fourthArea; j++)
                {
                    string line = lineList[j];
                    if (line.Contains("="))
                    {
                        string[] meaValues = line.Split(new char[] { '=' });
                        string meaName = meaValues[0];
                        string meaValue = meaValues[1];
                        measurementData1.Add("0");
                        measurementData1.Add("0");
                        measurementData1.Add(meaName);
                        measurementData1.Add(meaValue);
                    }
                }
                string date = lineList[fourthArea + 1].Split(new char[] { '=' })[1];
                string time = lineList[fourthArea + 2].Split(new char[] { '=' })[1];
                DateTime bookDate = Convert.ToDateTime(date + " " + time);
                long lBookData = ConvertDateToStamp(bookDate);
              
                //measurement data with max & min
                List<string> keys = dicMeasureEntity.Keys.ToList();
                List<string> measurementData2 = new List<string>();
                for (int t = fourthArea + 5; t < lineList.Count; t++)
                {
                    string[] values4 = lineList[t].Split(new char[] { '=' });
                    string measName = values4[0];
                    string measValue = values4[1];
                    string key = "";
                    bool hasStandard = FindStandardInMeasurement(measName, keys, out key);
                    if (hasStandard)
                    {
                        MeasurementEntity entity = dicMeasureEntity[key];
                        if (entity != null)
                        {
                            double maxValue = Convert.ToDouble(entity.MaxValue);
                            double minValue = Convert.ToDouble(entity.MinValue);
                            double meaValue = Convert.ToDouble(measValue);
                            string measFailCode = "";
                            if (meaValue >= minValue && meaValue <= maxValue)
                            {
                                measFailCode = "0";
                            }
                            else
                            {
                                measFailCode = "1";
                            }
                            //"ErrorCode", "LowerLimit", "MeasureFailCode", "MeasureName", "MeasureValue", "Unit", "UpperLimit"
                            measurementData2.Add("0");
                            measurementData2.Add(entity.MinValue);
                            measurementData2.Add(measFailCode);
                            measurementData2.Add(measName);
                            measurementData2.Add(measValue);
                            measurementData2.Add("-1");
                            measurementData2.Add(entity.MaxValue);
                        }
                    }
                    else
                    {
                        measurementData1.Add("0");
                        measurementData1.Add("0");
                        measurementData1.Add(measName);
                        measurementData1.Add(measValue);
                    }
                }
                LogHelper.Info("end analysis file");
                //upload state
                UploadProcessResult uploadHandler = new UploadProcessResult(sessionContext, initModel, this);
                uploadHandler.UploadProcessResultCall(serialNumber, iProcessLayer);

                //扣除相应的设备数量     郑培聪     20180228
                UpdateGridDataAfterUploadState();

                this.InvokeEx(x =>
                {
                    LoadYield();
                });

                GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                // Initial Product Inspection 
                string[] snInfo = getSNHandler.GetSNInfo(serialNumber);//{ "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" };
                if (snInfo != null)
                {
                    SetWorkorderValue(snInfo[2]);
                    SetPartNumberValue(snInfo[1]);
                    CheckIPIStatus();
                }

                //

                //SerialNumberData[] serialNumberArray = getSNHandler.GetPCBSNData(serialNumber);
                //string serialNumberPCB = "";
                //if (serialNumberArray != null && serialNumberArray.Length > 0)
                //{
                //    serialNumberPCB = serialNumberArray[0].serialNumber;
                //}
                int errorCode1 = uploadHandler.UploadFailureAndResultData(lBookData, serialNumber, "-1", iProcessLayer, -1, 1, measurementData1.ToArray(), new string[] { });
                int errorCode2 = uploadHandler.UploadResultDataAndRecipe(lBookData, serialNumber, "-1", "-1", measurementData2.ToArray(), 1, iProcessLayer);
                if (errorCode1 == 0 && errorCode2 == 0)
                {
                    UpdateIPIStatus("0");//首件检查 成功
                    UpdateDataToLogGrid(serialNumber, "解析文件成功", Convert.ToString(iProcessLayer));
                    MoveFileToOKFolder(filePath);
                    errorHandler(0, "Analysis file success", "");
                    SetTopWindowMessage("Analysis file success" + filePath, "");
                }
                else
                {
                    UpdateIPIStatus("1");//首件检查 失败
                    UpdateIPIStatusForProductionInspection("1", serialNumber);
                    UpdateDataToLogGrid(serialNumber, "解析文件失败", Convert.ToString(iProcessLayer));
                    MoveFileToErrorFolder(filePath, "Upload state error");
                    errorHandler(2, "Upload state error", "");
                }

            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                //UpdateDataToLogGrid(serialNumber, "解析文件失败");
                MoveFileToErrorFolder(filePath, ex.Message);
                errorHandler(2, "Analysis file fail", "");
            }
        }

        private bool FindStandardInMeasurement(string measureName, List<string> keys, out string key)
        {
            key = "";
            foreach (var item in keys)
            {
                if (measureName.Contains(item) && !measureName.Contains("_PWI") && !measureName.Contains("_CPK"))
                {
                    key = item;
                    return true;
                }
            }
            return false;
        }

        private void MoveFileToOKFolder(string filepath)
        {
            string OkFolder = config.LogTransOK;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to ok folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }

                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to ok folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = config.LogTransOK + strDirName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, Message("msg_Move file:") + filepath + Message(" msg_to OK folder success."), "");
            }
            catch (Exception e)
            {
                errorHandler(2, Message("msg_move file error ") + e.Message, "");
            }
        }

        private void MoveFileToErrorFolder(string filepath, string errorMsg)
        {
            string errorFolder = config.LogTransError;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to error folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }
                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to error folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = errorFolder + strDirName + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, Message("msg_Move file:") + filepath + Message("msg_to OK folder success."), "");
            }
            catch (Exception e)
            {
                errorHandler(2, Message("msg_move file error ") + e.Message, "");
            }
        }
        #endregion

        #region Other functions
        private void LoadYield()
        {
            GetProductQuantity getProductHandler = new GetProductQuantity(sessionContext, initModel, this);
            if (!string.IsNullOrEmpty(this.txbCDAMONumber.Text))
            {
                ProductEntity entity = getProductHandler.GetProductQty(Convert.ToInt32(initModel.currentSettings.processLayer), this.txbCDAMONumber.Text);
                if (entity != null)
                {
                    int totalQty = Convert.ToInt32(entity.QUANTITY_PASS) + Convert.ToInt32(entity.QUANTITY_FAIL) + Convert.ToInt32(entity.QUANTITY_SCRAP);
                    this.lblPass.Text = entity.QUANTITY_PASS;
                    this.lblFail.Text = entity.QUANTITY_FAIL;
                    this.lblScrap.Text = entity.QUANTITY_SCRAP;
                    this.lblYield.Text = "0%";
                    //this.lblAllCount.Text = totalQty + "";
                    if (totalQty > 0)
                    {
                        this.lblYield.Text = Math.Round(Convert.ToDecimal(lblPass.Text) / Convert.ToDecimal(totalQty) * 100, 2) + "%";
                    }
                }
            }
        }

        private void ShowTopWindow()
        {
            if (topmostform == null)
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
            }
        }

        private void SetTopWindowMessage(string text, string errorMsg)
        {
            if (topmostform != null)
            {
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
            else
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
        }

        List<PartMappingEntity> mappingDataList = null;
        private List<PartMappingEntity> GetPartMappingData(DataTable dt)
        {
            mappingDataList = new List<PartMappingEntity>();
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    PartMappingEntity entity = new PartMappingEntity();
                    entity.PartNo = row[0].ToString();
                    entity.PartDesc = row[1].ToString();
                    entity.MeasurementType = row[2].ToString();
                    entity.MeasurementUnit = row[3].ToString();
                    entity.MinValue = row[4].ToString();
                    entity.MaxValue = row[5].ToString();
                    entity.TestCount = row[6].ToString();
                    mappingDataList.Add(entity);
                }
            }
            return mappingDataList;
        }

        private void StrippedEquipmentFromStation()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            EquipmentManager equipmentHandler = new EquipmentManager(sessionContext, initModel, this);
            List<EquipmentEntityExt> entityList = equipmentHandler.GetSetupEquipmentDataByStation(config.StationNumber);
            if (entityList != null && entityList.Count > 0)
            {
                foreach (var item in entityList)
                {
                    equipmentHandler.UpdateEquipmentData(item.EQUIPMENT_INDEX, item.EQUIPMENT_NUMBER, 1);
                    RemoveAttributeForEquipment(item.EQUIPMENT_NUMBER, item.EQUIPMENT_INDEX, "attribEquipmentHasRigged");
                }
            }
        }

        Dictionary<string, string> dicAttris = new Dictionary<string, string>();
        private void GetWorkOrderAttris()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            dicAttris = getAttriHandler.GetAllAttributeValuesForWO(this.txbCDAMONumber.Text);
        }

        private void AddDataToSNGrid(object[] values)
        {


            this.Invoke(new MethodInvoker(delegate
            {

               
                bool isExist = false;
                bool isStatus = false;
                for (int i = 0; i < this.gridSNLog.Rows.Count; i++)
                {
                    if (gridSNLog.Rows[i].Cells["LSerialNumber"].Value.ToString() == values[0].ToString() && gridSNLog.Rows[i].Cells["LProcessLayer"].Value.ToString() == values[1].ToString())
                    {
                        isExist = true;
                        if (gridSNLog.Rows[i].Cells["LStatus"].Value.ToString() != values[5].ToString())
                        {
                            gridSNLog.Rows[i].Cells["LStatus"].Value = values[5].ToString();
                        }
                        break;
                    }

                }
                if (!isExist)
                    this.gridSNLog.Rows.Insert(0, values);

                if (this.gridSNLog.Rows.Count > 100)
                {
                    this.gridSNLog.Rows.RemoveAt(100);
                }
                this.gridSNLog.ClearSelection();
            }));
        }

        private delegate void UpdateDataToLogGridHandle(string serialNumber, string message, string iProcessLayer);
        private void UpdateDataToLogGrid(string serialNumber, string message, string iProcessLayer)
        {

            String showprocessLayer = "";
            if (iProcessLayer == "0")
            {
                showprocessLayer = "T";
            }
            if (iProcessLayer == "1")
            {
                showprocessLayer = "B";
            }
            if (iProcessLayer == "2")
            {
                showprocessLayer = "TB";
            }
            try
            {
                if (this.gridSNLog.InvokeRequired)
                {
                    UpdateDataToLogGridHandle updateDataDel = new UpdateDataToLogGridHandle(UpdateDataToLogGrid);
                    Invoke(updateDataDel, new object[] { serialNumber, message, iProcessLayer });
                }
                else
                {
                    bool isExist = false;
                    for (int i = 0; i < this.gridSNLog.Rows.Count; i++)
                    {
                        if (gridSNLog.Rows[i].Cells["LSerialNumber"].Value.ToString() == serialNumber && gridSNLog.Rows[i].Cells["LProcessLayer"].Value.ToString() == showprocessLayer)
                        {
                            gridSNLog.Rows[i].Cells["LStatus"].Value = message;
                            isExist = true;
                            break;
                        }
                    }

                    if (!isExist)
                    {
                        this.gridSNLog.Rows.Insert(0, new object[] { serialNumber, showprocessLayer, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), message });
                    }
                    this.gridSNLog.ClearSelection();
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void InitSetupGrid()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            GetMaterialBinData getMaterial = new GetMaterialBinData(sessionContext, initModel, this);
            DataTable dt = getMaterial.GetBomMaterialData(this.txbCDAMONumber.Text);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    this.gridSetup.Rows.Add(new object[10] { ReflowClient.Properties.Resources.Close, "", row["PartNumber"], row["PartDesc"], "", "", row["CompName"], row["Quantity"], "", "" });
                }
                this.gridSetup.ClearSelection();
            }
        }
        //verify the serial number's part number is equals the current part number
        private bool VerifySerialNumber(string serialNumber)
        {
            bool isValid = true;
            if (this.txbCDAPartNumber.Text.Trim() == "")
                return isValid;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                string snPartNumber = snValues[1];
                LogHelper.Info("Current part number:" + this.txbCDAPartNumber.Text.Trim());
                LogHelper.Info("Serial number part number:" + snPartNumber);
                if (snPartNumber == this.txbCDAPartNumber.Text.Trim())
                { }
                else
                {
                    errorHandler(3, Message("msg_The serial number's part number is not equals the current part number"), "");
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, Message("msg_The serial number: ") + serialNumber + Message("msg_ is invalid"), "");
                SetTopWindowMessage(serialNumber, "The serial number is invalid.");
                isValid = false;
            }
            return isValid;
        }

        private void PassBoard()
        {
            initModel.scannerHandler.sendHighExt();
            Thread.Sleep(Convert.ToInt32(config.GateKeeperTimer));
            initModel.scannerHandler.sendLowExt();
            //initModel.scannerHandler.sendHigh();
        }

        private void ProcessMaterialBinNo(string materialBinNo)
        {
            MaterialManager matHnadler = new MaterialManager(sessionContext, initModel, this);
            string strPartNumber = matHnadler.GetPartNumberFromMBN(materialBinNo);
            SetContainerNoValue(materialBinNo);
            List<PartMappingEntity> valueList = mappingDataList.Where(t => t.PartNo == strPartNumber).ToList();
            if (valueList != null && valueList.Count > 0)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    this.txtPartNo.Text = valueList[0].PartNo;
                    this.txtPartDesc.Text = valueList[0].PartDesc;
                    this.txtTestCount.Text = valueList[0].TestCount;
                    this.txtMeasurementType.Text = valueList[0].MeasurementType;
                    this.txtMeasurementUnit.Text = valueList[0].MeasurementUnit;
                    this.txtMinValue.Text = valueList[0].MinValue;
                    this.txtMaxValue.Text = valueList[0].MaxValue;
                }));
            }
        }

        private bool ProcessSerialNumber(string serialNumber, int processLayer)
        {
            //verify material&equipment is ok
            if (!VerifyEquipment())
            {
                SendSN(config.LIGHT_CHANNEL_ON);
                return false;
            }
            if (!VerifyCheckList())
            {
                SendSN(config.LIGHT_CHANNEL_ON);
                return false;
            }
            //if (!VerifySerialNumber(serialNumber))
            //    return false;
            //set work order value & part numebr value
            //GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            //string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            //if (snValues != null && snValues.Length > 0)
            //{
            //    SetWorkorderValue(snValues[2]);
            //    SetPartNumberValue(snValues[1]);
            //}
            //verify wo 
            if (!VerifySerialNumberByWo(serialNumber))
            {
                SendSN(config.LIGHT_CHANNEL_ON);
                return false;
            }
            //gate keeper,check serial state
            CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
            processLayer = checkHandler.GetProcessLayerBySN(serialNumber, initModel.configHandler.StationNumber);
            if (processLayer == -1)
            {
                errorHandler(2, "没有找到激活工单", "");
                //SetTopWindowMessage(serialNumber, "没有找到激活工单.");
                SendSN(config.LIGHT_CHANNEL_ON);
                return false;
            }
            string errorstring = "";
            bool isOK = checkHandler.CheckSNStateNew(serialNumber, processLayer, out errorstring);
            String showprocessLayer = "";
            if (processLayer==0)
            {
                showprocessLayer = "T";
            }
            if (processLayer == 1)
            {
                showprocessLayer = "B";
            }
            if (processLayer == 2)
            {
                showprocessLayer = "TB";
            }

            if (isOK)
            {
                if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                {
                    SendSN(config.LIGHT_CHANNEL_OFF);
                }
                SetTopWindowMessage(serialNumber, "");
                AddDataToSNGrid(new object[7] { serialNumber, showprocessLayer, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", "Check Serial Number state success" });
                if (config.CEHCKTXTFILE != "YES")
                {
                    UploadProcessResult uploadHandler = new UploadProcessResult(sessionContext, initModel, this);
                    uploadHandler.UploadProcessResultCall(serialNumber, processLayer);
                    GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                    // Initial Product Inspection 
                    LoadYield();
                    UpdateIPIStatus("0");//首件检查 成功
                   // UpdateDataToLogGrid(serialNumber, "解析文件成功", Convert.ToString(processLayer));
                    errorHandler(0, "update state success", "");
                    SetTopWindowMessage("upload state success", "");
                    string[] snInfo = getSNHandler.GetSNInfo(serialNumber);//{ "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" };
                    if (snInfo != null)
                    {
                        SetWorkorderValue(snInfo[2]);
                        SetPartNumberValue(snInfo[1]);
                        CheckIPIStatus();
                    }
                    return true;
                }
                return true;
            }
            else
            {
                //SetTopWindowMessage(serialNumber, "Check Serial Number State Error.");
                errorHandler(2, "序列号状态错误" + errorstring, "");
                AddDataToSNGrid(new object[7] { serialNumber, showprocessLayer, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NG", "Check Serial Number State Error" });
                if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                    SendSN(config.LIGHT_CHANNEL_ON);
                return false;
            }
            return true;
        }

        private void UpdateGridDataAfterUploadState()
        {

            #region 参照锡膏印刷的设备数量扣取       郑培聪     20180228
            if (dgvEquipment.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvEquipment.Rows)
                {
                    string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                    Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
                        {
                            int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
                            if (config.ReduceEquType == "1")
                            {
                                row.Cells["UsCount"].Value = iQty - 1;
                                int usagecount = 0;
                                GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
                                string[] valuesusage = getAttribHandler.GetAttributeValueForEquipment("USAGE_COUNT", equipmentNo, "0");
                                if (valuesusage != null && valuesusage.Length != 0)
                                {
                                    usagecount = Convert.ToInt32(valuesusage[1]);
                                }

                                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                                appendAttri.AppendAttributeValuesForEquipment("USAGE_COUNT", Convert.ToString(usagecount + 1), equipmentNo);
                            }
                            else
                            {
                                row.Cells["UsCount"].Value = iQty - initModel.numberOfSingleBoards;//
                            }

                        }
                    }
                    else
                    {
                        if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
                        {
                            int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
                            row.Cells["UsCount"].Value = iQty - initModel.numberOfSingleBoards;//
                        }
                    }
                }
            }
            #endregion

            //if (dgvEquipment.Rows.Count > 0)
            //{
            //    foreach (DataGridViewRow row in dgvEquipment.Rows)
            //    {
            //        if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
            //        {
            //            int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
            //            row.Cells["UsCount"].Value = iQty - initModel.numberOfSingleBoards;//
            //        }
            //    }
            //}

            //string materialBinNo = FindMaterialBinNumber();
            //if (materialBinNo != null && materialBinNo != "")
            //{
            //    Double iPPNQty = Convert.ToDouble(Convert.ToDecimal(gridSetup.Rows[0].Cells["PPNQty"].Value));//* initModel.numberOfSingleBoards;//todo
            //    LogHelper.Info("Consumption material bin number:" + materialBinNo);
            //    LogHelper.Info("Consumption quantity:" + materialBinNo);
            //    UpdateMaterialGridData(materialBinNo, iPPNQty);
            //}
        }

        private void UpdateMaterialGridData(string materialBinNumber, double qty)
        {
            ProcessMaterialBinData materialHandler = new ProcessMaterialBinData(sessionContext, initModel, this);
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString() == materialBinNumber)
                {
                    double iQty = Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value);
                    if (iQty >= qty)
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = iQty - qty;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, GetWorkOrderValue(), -qty);
                        if (iQty == qty)//update 2015/6/24
                        {
                            if (i + 1 < this.gridSetup.Rows.Count)
                            {
                                string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                                string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                                string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                                //setup material
                                SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                                setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                                setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);
                            }
                            else
                            {
                                errorHandler(3, Message("msg_The solder paste not enough."), "");
                            }
                        }
                        break;
                    }
                    else
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = 0;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, GetWorkOrderValue(), -iQty);
                        if (i + 1 < this.gridSetup.Rows.Count)
                        {
                            string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                            string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                            string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                            //setup material
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                            setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);
                            UpdateMaterialGridData(nextMaterialBinNo, qty - iQty);
                        }
                        else
                        {
                            //warming no material
                            errorHandler(3, Message("msg_The solder paste not enough."), "");
                        }
                    }
                }
            }
        }

        private long ConvertDateToStamp(DateTime dt)
        {
            DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            TimeSpan toNow = dt.Subtract(dtStart);
            return Convert.ToInt64(toNow.TotalMilliseconds);
        }

        private string FindMaterialBinNumber()
        {
            string materialBinNo = "";
            if (gridSetup.Rows.Count > 0)
            {
                for (int i = 0; i < this.gridSetup.Rows.Count; i++)
                {
                    if (Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value) > 0)
                    {
                        materialBinNo = gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString();
                        break;
                    }

                }
            }
            return materialBinNo;
        }

        private bool CheckMaterialBinHasSetup(string materailBinNo)
        {
            bool isExist = false;
            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value.ToString() == materailBinNo)
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }

        private bool VerifyEquipment()
        {
            bool isValid = true;
            if (this.dgvEquipment.Rows.Count == 0)
                return isValid;
            EquipmentManager equipmentHandler = new EquipmentManager(sessionContext, initModel, this);
            int errorCode = equipmentHandler.CheckEquipmentData(this.txbCDAMONumber.Text);
            if (errorCode != 0)
            {
                errorHandler(3, Message("msg_Check equipment data error :") + errorCode, "");
                return false;
            }
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                if (Convert.ToInt32(item.Cells["UsCount"].Value) <= 0)
                {
                    isValid = false;
                    item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                    errorHandler(3, Message("Equipment usage count cann''t less then 0"), "");
                    break;
                }
                else if (Convert.ToDateTime(item.Cells["NextMaintenance"].Value) <= DateTime.Now)
                {
                    string equipmentNo = item.Cells["EquipNo"].Value.ToString();
                    Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        isValid = false;
                        item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                        errorHandler(3, Message("msg_Stencil need to clean"), "");
                        //add attribute lock time
                        //add attribute to equipment
                        //AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                        //appendAttri.AppendAttributeValuesForEquipment("StencilLockTime", DateTime.Now.AddHours(Convert.ToDouble(config.LockTime)).ToString("yyyy/MM/dd HH:mm:ss"), equipmentNo);
                        break;
                    }
                }
            }
            if (isValid)//continue check material expiry date
            {
                foreach (DataGridViewRow itemM in this.gridSetup.Rows)
                {
                    DateTime dtExpiry = Convert.ToDateTime(itemM.Cells["ExpiryTime"].Value);
                    if (DateTime.Now > dtExpiry)
                    {
                        isValid = false;
                        errorHandler(3, Message("msg_The solder paste has expiry."), "");
                        break;
                    }
                }
            }
            return isValid;
        }

        string snWorkOrder = "";
        private bool VerifySerialNumberByWo(string serialNumber)
        {
            bool isValid = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                snWorkOrder = snValues[2];

                if (snWorkOrder == this.txbCDAMONumber.Text.Trim())
                { }
                else
                {
                    errorHandler(3, Message("msg_WO_NotMatch"), "");
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, Message("msg_The serial number: ") + serialNumber + Message("msg_ is invalid"), "");
                SetTopWindowMessage(serialNumber, Message("msg_The serial number is invalid."));
                isValid = false;
            }
            return isValid;
        }

        private bool VerifyActivatedWO()
        {
            bool isValid = true;
            GetCurrentWorkorder getActivatedWOHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
            GetStationSettingModel stationSetting = getActivatedWOHandler.GetCurrentWorkorderResultCall();
            if (stationSetting != null && stationSetting.workorderNumber != null)
            {
                if (stationSetting.workorderNumber == this.txbCDAMONumber.Text)
                {
                    isValid = true;
                }
                else
                {
                    isValid = false;
                    errorHandler(2, Message("msg_The current activated work order has changed, please refresh."), "");
                }
            }
            return isValid;
        }

        private bool VerifyMaterialBinData(string materialBinNo, string partNumber)
        {
            bool isValid = true;
            if (config.ThawingCheck != "Enable")
            {
                return true;
            }
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] values = getAttriHandler.GetAttributeValueForContainer("SP_THAW_COMPLETE", materialBinNo);
            if (values != null && values.Length > 0)
            {
                DateTime dtCompleteThawing = Convert.ToDateTime(values[1]);//.AddMinutes(GetThawingTime(partNumber))
                if (DateTime.Now < dtCompleteThawing)
                {
                    isValid = false;
                    errorHandler(3, Message("msg_The solder paste thawing not complete."), "");
                }
            }
            else
            {
                isValid = false;
                errorHandler(3, Message("msg_The solder paste thawing not complete."), "");
            }
            return isValid;
        }

        private double GetThawingTime(string partNumber)
        {
            double iValue = 0;
            //get solder paste part number attribute
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] attriCodes = new string[] { "SPThawingTime" };
            Dictionary<string, string> attriValues = getAttriHandler.GetAttributeValueForPart(attriCodes, partNumber);
            if (attriValues != null)
            {
                foreach (var key in attriValues.Keys)
                {
                    if (!string.IsNullOrEmpty(attriValues[key]))
                    {
                        if (key == "SPThawingTime")
                        {
                            iValue = Convert.ToDouble(attriValues[key]);
                        }
                    }
                }
            }
            if (iValue == 0)
            {
                iValue = Convert.ToDouble(config.ThawingDuration);
            }
            return iValue;
        }

        private string ConvertDateFromStamp(string timeStamp)
        {
            double d = Convert.ToDouble(timeStamp);
            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(d).ToLocalTime();
            return date.ToString();
        }

        private string ConverToHourAndMin(int number)
        {
            int iHour = number / 60;
            int iMin = number % 60;
            return iHour + "hr " + iMin + "min";
        }

        private bool CheckMaterialSetUp()
        {
            bool isValid = true;
            double iQty = 0;
            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString().Length == 0)
                {
                    errorHandler(3, Message("msg_Material setup required."), "");
                    isValid = false;
                    break;
                }
                iQty += Convert.ToDouble(row.Cells["Qty"].Value);
            }
            if (iQty <= 1 && gridSetup.Rows.Count > 0)
            {
                errorHandler(3, Message("msg_Material quantity is not enough."), "");
                isValid = false;
            }
            return isValid;
        }

        #region Compare
        private class RowComparer : System.Collections.IComparer
        {
            private static int sortOrderModifier = 1;

            public RowComparer(SortOrder sortOrder)
            {
                if (sortOrder == SortOrder.Descending)
                {
                    sortOrderModifier = -1;
                }
                else if (sortOrder == SortOrder.Ascending) { sortOrderModifier = 1; }
            }
            public int Compare(object x, object y)
            {
                DataGridViewRow DataGridViewRow1 = (DataGridViewRow)x;
                DataGridViewRow DataGridViewRow2 = (DataGridViewRow)y;
                // Try to sort based on the Scan time column.
                string value1 = DataGridViewRow1.Cells["colItemCode"].Value.ToString();
                string value2 = DataGridViewRow2.Cells["colItemCode"].Value.ToString();
                string type1 = DataGridViewRow1.Cells["colType"].Value.ToString();
                string type2 = DataGridViewRow2.Cells["colType"].Value.ToString();
                int CompareResult = 0;
                if (type1 == type2)
                {
                    CompareResult = value1.CompareTo(value2);
                }
                else
                {
                    CompareResult = type1.CompareTo(type2);
                }
                return CompareResult * sortOrderModifier;
            }
        }
        #endregion

        #region Document
        static string cachePN = "";
        private void InitDocumentGrid()
        {
            if (config.FilterByFileName == "disable") //by station
            {
                if (gridDocument.Rows.Count <= 0)
                {
                    GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                    List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                    if (listDoc != null && listDoc.Count > 0)
                    {
                        foreach (DocumentEntity item in listDoc)
                        {
                            gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                        }
                    }
                }
            }
            else //by station & filename(partno)
            {
                if (this.txbCDAPartNumber.Text == "" || cachePN == this.txbCDAPartNumber.Text)
                    return;
                cachePN = this.txbCDAPartNumber.Text;
                gridDocument.Rows.Clear();
                this.Invoke(new MethodInvoker(delegate
                {
                    webBrowser1.Navigate("about:blank");
                }));
                GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                if (listDoc != null && listDoc.Count > 0)
                {
                    foreach (DocumentEntity item in listDoc)
                    {
                        string filename = item.MDA_FILE_NAME;
                        Match name = Regex.Match(filename, config.FileNamePattern);
                        if (name.Success)
                        {
                            if (name.Groups.Count > 1)
                            {
                                string partno = name.Groups[1].ToString();
                                if (partno == this.txbCDAPartNumber.Text)
                                {
                                    gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                                }
                            }
                        }
                    }
                }
            }
        }

        private void GetDocumentCollections()
        {
            GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
            //get advice id
            Advice[] adviceArray = getDocument.GetAdviceByStationAndPN(this.txbCDAPartNumber.Text);
            if (adviceArray != null && adviceArray.Length > 0)
            {
                int iAdviceID = adviceArray[0].id;
                List<DocumentEntity> list = getDocument.GetDocumentDataByAdvice(iAdviceID);
                if (list != null && list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        string docID = item.MDA_DOCUMENT_ID;
                        string fileName = item.MDA_FILE_NAME;
                        SetDocumentControl(docID, fileName);
                        break;
                    }
                }
            }
        }

        private void SetDocumentControl(string docID, string fileName)
        {
            GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
            byte[] content = documentHandler.GetDocumnetContentByID(Convert.ToInt64(docID));
            if (content != null)
            {
                string path = config.MDAPath;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                string filePath = path + @"/" + fileName;
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                Encoding.GetEncoding("gb2312");
                fs.Write(content, 0, content.Length);
                fs.Flush();
                fs.Close();
            }
        }

        private void SetDocumentControlForDoc(long documentID, string fileName)
        {
            string path = config.MDAPath;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filePath = path + @"/" + fileName;
            if (!File.Exists(filePath))
            {
                GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
                byte[] content = documentHandler.GetDocumnetContentByID(documentID);
                if (content != null)
                {
                    FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                    fs.Write(content, 0, content.Length);
                    fs.Flush();
                    fs.Close();
                }
            }
            this.webBrowser1.Navigate(filePath);
        }
        #endregion

        #region IO TOWER LIGHT
        private bool SendSN(string serialNumber)
        {
            try
            {
                //Thread.Sleep(20);
                if (config.DataOutputInterface == "COM")
                {
                    try
                    {
                        initModel.scannerHandler.handler2().Write(strToToHexByte(serialNumber), 0, strToToHexByte(serialNumber).Length);
                        LogHelper.Info("Send command:" + serialNumber);
                        return true;
                    }
                    catch (Exception e)
                    {
                        LogHelper.Error(e);
                        return false;
                    }
                }
                else
                {
                    if (config.OutputEnter == "1")
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber + "\r"); //大写键总是被按起。。。。
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber + "\r");
                        }
                    }
                    else
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber);
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber);
                        }
                    }
                    SendKeys.Flush();
                    Thread.Sleep(300);
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                return false;
            }
        }
        private static byte[] strToToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }
        #endregion

        #region Equipment
        private void InitEquipmentGrid()
        {
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            List<EquipmentEntity> listEntity = eqManager.GetRequiredEquipmentData(this.txbCDAMONumber.Text);
            if (listEntity != null)
            {
                foreach (var item in listEntity)
                {
                    this.dgvEquipment.Rows.Add(new object[8] { ReflowClient.Properties.Resources.Close, item.PART_NUMBER, item.EQUIPMENT_DESCRIPTION, "", "", "", "", "0" });
                }
            }
            this.dgvEquipment.ClearSelection();
        }

        //添加锡膏印刷部分上设备代码     郑培聪     20180228
        //如果assign 的是设备，就不能上其他设备
        List<string> equipmentList = new List<string>();
        private void InitEquipmentGridEXT()
        {
            //if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
            //    return;

            //EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            //foreach (DataGridViewRow row in dgvEquipment.Rows)
            //{
            //    string equipmentNo = row.Cells["EquipNo"].Value.ToString();
            //    string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
            //    if (string.IsNullOrEmpty(equipmentNo))
            //        continue;
            //    int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
            //    RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
            //}

            //this.dgvEquipment.Rows.Clear();
            //equipmentList.Clear();
            //List<string> equipList = new List<string>();
            //Dictionary<string, List<EquipmentEntity>> DiclistEntity = eqManager.GetRequiredEquipmentDataDic(this.txbCDAMONumber.Text);
            //if (DiclistEntity != null)
            //{
            //    foreach (var key in DiclistEntity.Keys)
            //    {
            //        List<EquipmentEntity> listEntity = DiclistEntity[key.ToString()];
            //        foreach (var item in listEntity)
            //        {
            //            equipmentList.Add(item.EQUIPMENT_NUMBER);
            //            //if (equipList.Contains(item.PART_NUMBER))
            //            //    continue;
            //            equipList.Add(item.PART_NUMBER);

            //            string partdesc = eqManager.GePartDescData(item.PART_NUMBER);
            //            this.dgvEquipment.Rows.Add(new object[8] { ReflowClient.Properties.Resources.Close, item.PART_NUMBER, partdesc, "", "", "", "", "0" });
            //        }
            //    }
            //}
            //this.dgvEquipment.ClearSelection();

            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            foreach (DataGridViewRow row in dgvEquipment.Rows)
            {
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string UsCount = row.Cells["UsCount"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                if (string.IsNullOrEmpty(UsCount) || string.IsNullOrEmpty(equipmentNo))
                    continue;
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
            }

            this.dgvEquipment.Rows.Clear();
            equipmentList.Clear();
            List<string> equipList = new List<string>();
            Dictionary<string, List<EquipmentEntity>>  DiclistEntity = eqManager.GetRequiredEquipmentDataDicEXT(this.txbCDAMONumber.Text);
            if (DiclistEntity != null)
            {
                List<string> groupList = new List<string>();
                foreach (var key in DiclistEntity.Keys)
                {
                    string[] keys = key.Split(';');
                    string group = keys[1];
                    if (groupList.Contains(group))
                        continue;
                    groupList.Add(group);

                    #region 设备原本的代码,为配合分组只上一个的原则,这边会做代码修改       郑培聪     2017/11/15


                    //List<EquipmentEntity> listEntity = DiclistEntity[key.ToString()];
                    //foreach (var item in listEntity)
                    //{
                    //    equipmentList.Add(item.EQUIPMENT_NUMBER);
                    //    //if (equipList.Contains(item.PART_NUMBER))
                    //    //    continue;
                    //    equipList.Add(item.PART_NUMBER);

                    //    string partdesc = eqManager.GePartDescData(item.PART_NUMBER);
                    //    this.dgvEquipment.Rows.Add(new object[8] { ScreenPrinter.Properties.Resources.Close, item.PART_NUMBER, partdesc, item.EQUIPMENT_NUMBER, "", "", "", "0" });
                    //}


                    List<EquipmentEntity> listEntity = DiclistEntity[key.ToString()];
                    if (listEntity == null || listEntity.Count < 1)
                        continue;
                    string partdesc = string.Empty;
                    string partNumber = string.Empty;
                    string equipmentNumber = string.Empty;
                    foreach (var item in listEntity)
                    {
                        if (!partNumber.Contains(item.PART_NUMBER))
                        {
                            partNumber += "," + item.PART_NUMBER;
                            partdesc += "," + eqManager.GePartDescData(item.PART_NUMBER);
                        }

                        if (!equipmentNumber.Contains(item.EQUIPMENT_NUMBER))
                            equipmentNumber += "," + item.EQUIPMENT_NUMBER;

                        equipmentList.Add(item.EQUIPMENT_NUMBER);
                    }

                    this.dgvEquipment.Rows.Add(new object[9] { ReflowClient.Properties.Resources.Close, partNumber.TrimStart(','), partdesc.TrimStart(','), equipmentNumber.TrimStart(','), "", "", "", "0", key });


                    #endregion
                }
            }
            this.dgvEquipment.ClearSelection();
        }

        public bool CheckEquipmentSetup()
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length == 0)
                {
                    errorHandler(3, Message("msg_Equipment setup required."), "");
                    return false;
                }
            }
            return true;
        }

        public void ProcessEquipmentData(string equipmentNo)
        {
            if (!VerifyActivatedWO())
                return;
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            string ePartNumber = "";
            string eIndex = "0";
            if (!CheckEquipmentDuplication(values, ref ePartNumber, ref eIndex))
            {
                errorHandler(3, Message("msg_The equipment") + equipmentNo + Message("msg_ has more Available states."), "");
                return;
            }
            if (!CheckEquipmentValid(ePartNumber))
            {
                errorHandler(3, Message("msg_The equipment is invalid"), "");
                return;
            }
            //check equipment number  whether need to setup?
            if (!CheckEquipmentIsExist(ePartNumber))
                return;
            //check equipment number has rigged on others station
            if (CheckEquipmentHasSetup(equipmentNo, eIndex, "attribEquipmentHasRigged"))
            {
                errorHandler(3, Message("msg_The equipment has rigged on others station."), "");
                return;
            }
            string strEquipmentIndex = eIndex;
            //check the equipment whether is cleaning?
            GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesEquip = getAttribHandler.GetAttributeValueForEquipment("StencilLockTime", equipmentNo, strEquipmentIndex);
            if (valuesEquip != null && valuesEquip.Length > 0)
            {
                try
                {
                    //if (Convert.ToDateTime(valuesEquip[1]) > DateTime.Now)
                    //{
                    errorHandler(3, Message("msg_The equipment is cleaning"), "");
                    return;
                    //}
                }
                catch (Exception exx)
                {

                    LogHelper.Error(exx);
                }
            }
            int errorCode = eqManager.UpdateEquipmentData(strEquipmentIndex, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                //add attribue command the equipment is uesd
                AppendAttributeForEquipment(equipmentNo, strEquipmentIndex, "attribEquipmentHasRigged");
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    entityExt.PART_NUMBER = ePartNumber;
                    entityExt.EQUIPMENT_INDEX = strEquipmentIndex;
                    SetEquipmentGridData(entityExt);
                    SetTipMessage(MessageType.OK, Message("msg_Process equipment number ") + equipmentNo + Message("msg_SUCCESS"));
                }
            }
        }


        /// <summary>
        /// 引用锡膏印刷部分的上设备代码        郑培聪     20180228
        /// </summary>
        /// <param name="equipmentNo"></param>
        public void ProcessEquipmentDataEXT(string equipmentNo)
        {
             if (!VerifyActivatedWO())
                return;
            if (!equipmentList.Contains(equipmentNo))
            {
                string rightequipment = "";
                foreach (var equ in equipmentList)
                {
                    rightequipment += equ + ";";
                }
                errorHandler(2, Message("msg_equ is not belong to the list") + rightequipment.TrimEnd(';'), "");
                return;
            }
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            string ePartNumber = "";
            string eIndex = "0";
            int uasedcount = 0;
            if (values != null && values.Length > 0)
            {
                uasedcount = Convert.ToInt32(values[5]);
            }

            if (!CheckEquipmentDuplication(values, ref ePartNumber, ref eIndex))
            {
                errorHandler(3, Message("msg_The equipment") + equipmentNo + Message("msg_ has more Available states."), "");
                return;
            }
            if (!CheckEquipmentValid(ePartNumber))
            {
                errorHandler(3, Message("msg_The equipment is invalid"), "");
                return;
            }
            //check equipment number  whether need to setup?
            if (!CheckEquipmentIsExist(ePartNumber))
                return;
            //check equipment number has rigged on others station
            if (CheckEquipmentHasSetup(equipmentNo, eIndex, "attribEquipmentHasRigged"))
            {
                errorHandler(3, Message("msg_The equipment has rigged on others station."), "");
                return;
            }
            string strEquipmentIndex = eIndex;
            //check the equipment whether is cleaning?
            int equicount = 0;
            Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
            if (matchStencil.Success)
            {
                GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesTest = getAttribHandler.GetAttributeValueForEquipment("TEST_ITEM", equipmentNo, strEquipmentIndex);
                if (valuesTest == null || valuesTest.Length == 0)
                {
                    errorHandler(3, Message("msg_The equipment is not test"), "");
                    return;
                }

                string[] valuesEquip = getAttribHandler.GetAttributeValueForEquipment("STENCIL_LOCK_TIME", equipmentNo, strEquipmentIndex);
                if (valuesEquip != null && valuesEquip.Length > 0)
                {
                    try
                    {
                        //if (Convert.ToDateTime(valuesEquip[1]) > DateTime.Now)
                        //{
                        errorHandler(3, Message("msg_The equipment is cleaning"), "");
                        return;
                        //}
                    }
                    catch (Exception exx)
                    {

                        LogHelper.Error(exx);
                    }
                }
                string[] valuesTest2 = getAttribHandler.GetAttributeValueForEquipment("NEXT_TEST_DATE", equipmentNo, "0");
                DateTime NextTestDate = DateTime.Now;
                if (valuesTest2 != null && valuesTest2.Length != 0)
                    NextTestDate = Convert.ToDateTime(valuesTest2[1]);
                if (NextTestDate < DateTime.Now)
                {
                    errorHandler(3, Message("msg_Stencil need to clean"), "");
                    AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    appendAttri.AppendAttributeValuesForEquipment("STENCIL_LOCK_TIME", DateTime.Now.AddHours(Convert.ToDouble(config.LockTime)).ToString("yyyy/MM/dd HH:mm:ss"), equipmentNo);
                    return;
                }

                int usagecount = 0;
                int maxcount = 0;
                if (config.ReduceEquType == "1")
                {
                    string[] valuesusage = getAttribHandler.GetAttributeValueForEquipment("USAGE_COUNT", equipmentNo, strEquipmentIndex);
                    if (valuesusage != null && valuesusage.Length != 0)
                    {
                        usagecount = Convert.ToInt32(valuesusage[1]);
                    }
                    string[] valuesmax = getAttribHandler.GetAttributeValueForEquipment("MAX_USAGE", equipmentNo, strEquipmentIndex);
                    if (valuesmax != null && valuesmax.Length != 0)
                    {
                        maxcount = Convert.ToInt32(valuesmax[1]);
                    }
                    equicount = maxcount - usagecount;
                    if (equicount <= 0)
                    {
                        errorHandler(3, Message("Equipment usage count cann''t less then 0"), "");
                        return;
                    }
                }
                else
                {
                    if (uasedcount <= 0)
                    {
                        errorHandler(3, Message("Equipment usage count cann''t less then 0"), "");
                        return;
                    }
                }
            }
            else
            {
                if (uasedcount <= 0)
                {
                    errorHandler(3, Message("Equipment usage count cann''t less then 0"), "");
                    return;
                }
            }
            int errorCode = eqManager.UpdateEquipmentData(strEquipmentIndex, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                //add attribue command the equipment is uesd
                AppendAttributeForEquipment(equipmentNo, strEquipmentIndex, "attribEquipmentHasRigged");
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    entityExt.PART_NUMBER = ePartNumber;
                    entityExt.EQUIPMENT_INDEX = strEquipmentIndex;
                    SetEquipmentGridData(entityExt, equicount, values[6]);
                    SetTipMessage(MessageType.OK, Message("msg_Process equipment number ") + equipmentNo + Message("msg_SUCCESS"));
                    SaveEquAndMaterial();
                    if (CheckMaterialSetUp() && CheckEquipmentSetup())
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_OFF);
                    }
                }
            }
        }

        /// <summary>
        /// 引用锡膏印刷部分的上设备代码      郑培聪     20180228
        /// restore.txt文本保存设备数据信息
        /// </summary>
        private void SaveEquAndMaterial()
        {
            try
            {
                string path = @"restore.txt";
                string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                string equipment = "";
                string material = "";
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(datetime);
                sb.AppendLine(txbCDAMONumber.Text + ";" + initModel.currentSettings.processLayer);
                foreach (DataGridViewRow row in dgvEquipment.Rows)
                {
                    string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                    if (string.IsNullOrEmpty(equipmentNo))
                        continue;
                    equipment += equipmentNo + ";";
                }
                sb.AppendLine(equipment.TrimEnd(';'));
                foreach (DataGridViewRow row in gridSetup.Rows)
                {
                    string materialbin = row.Cells["MaterialBinNo"].Value.ToString();
                    if (string.IsNullOrEmpty(materialbin))
                        continue;
                    material += materialbin + ";";
                }
                sb.AppendLine(material.TrimEnd(';'));
                FileStream fs = new FileStream(path, FileMode.Create);
                byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
                fs.Seek(0, SeekOrigin.Begin);
                fs.Write(bt, 0, bt.Length);
                fs.Flush();
                fs.Close();
                LogHelper.Info("Save restore file success.");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        /// <summary>
        /// 引用锡膏印刷部分的上设备代码      郑培聪     20180228
        /// 读取restore.txt文本保存的设备数据信息,并将之上设备
        /// </summary>
        private void ReadRestoreFile()
        {
            if (this.txbCDAMONumber.Text != "")
            {
                if (config.RESTORE_TIME == "")
                    return;
                string path = @"restore.txt";
                if (File.Exists(path))
                {
                    string[] linelist = File.ReadAllLines(path);
                    string datetimespan = linelist[0];
                    string workorder = linelist[1];
                    string equipment = linelist[2];
                    string material = linelist[3];

                    TimeSpan span = DateTime.Now - Convert.ToDateTime(datetimespan);

                    if (span.TotalMinutes > Convert.ToInt32(config.RESTORE_TIME))//判断是否大于10分钟，大于10分钟则不自动上料和设备
                    {

                    }
                    else
                    {
                        string[] workorders = workorder.Split(';');
                        if (workorders.Length > 1)
                        {
                            if (workorders[0] == this.txbCDAMONumber.Text)//判断工单是否有变化，无变化则自动上料和设备
                            {
                                if (workorders[1] == initModel.currentSettings.processLayer.ToString())//判断面次是否有变化，无变化则自动上料和设备
                                {
                                    bool isOK = false;

                                    #region setup equ
                                    EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                                    string[] equs = equipment.Split(';');
                                    if (equipment.Replace(";", "").Trim() != "")
                                    {
                                        foreach (var equipmentNo in equs)
                                        {
                                            string equipmentnumber = equipmentNo.ToString();
                                            if (string.IsNullOrEmpty(equipmentnumber))
                                                continue;
                                            int errorCode = eqManager.UpdateEquipmentData("0", equipmentnumber, 1);
                                            RemoveAttributeForEquipment(equipmentnumber, "0", "attribEquipmentHasRigged");
                                            ProcessEquipmentDataEXT(equipmentnumber);
                                        }
                                        isOK = true;
                                    }
                                    #endregion
                                    
                                    if (isOK)
                                        errorHandler(0, Message("msg_Material and Equipment setup has been restored"), "");
                                }

                            }
                        }

                    }
                }
            }
        }

        private void SetEquipmentGridData(EquipmentEntityExt entityExt)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == entityExt.PART_NUMBER
                    && (row.Cells["EquipNo"].Value == null || row.Cells["EquipNo"].Value.ToString() == ""))
                {
                    Match matchStencil = Regex.Match(entityExt.EQUIPMENT_NUMBER, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        row.Cells["NextMaintenance"].Value = DateTime.Now.AddHours(Convert.ToDouble(config.UsageTime)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        row.Cells["NextMaintenance"].Value = DateTime.Now.AddSeconds(Convert.ToDouble(entityExt.SECONDS_BEFORE_EXPIRATION)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    row.Cells["ScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                    row.Cells["EquipNo"].Value = entityExt.EQUIPMENT_NUMBER;
                    row.Cells["EquipmentIndex"].Value = entityExt.EQUIPMENT_INDEX;
                    row.Cells["Status"].Value = ReflowClient.Properties.Resources.ok;
                    row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(0, 192, 0);
                }
            }
        }

        /// <summary>
        /// 添加提供给新版本使用的SetEquipmentGridData多态       郑培聪     20180228
        /// </summary>
        /// <param name="entityExt"></param>
        /// <param name="usagecount"></param>
        /// <param name="expireddate"></param>
        private void SetEquipmentGridData(EquipmentEntityExt entityExt, int usagecount, string expireddate)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                //if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == entityExt.PART_NUMBER
                //    && (row.Cells["EquipNo"].Value == null || row.Cells["EquipNo"].Value.ToString() == ""))
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString().Contains(entityExt.PART_NUMBER)
                    && (row.Cells["EquipNo"].Value != null && row.Cells["EquipNo"].Value.ToString().Contains(entityExt.EQUIPMENT_NUMBER)))
                {
                    Match matchStencil = Regex.Match(entityExt.EQUIPMENT_NUMBER, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        if (config.ReduceEquType == "1")
                            row.Cells["UsCount"].Value = usagecount;
                        else
                        {
                            row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                        }

                        row.Cells["NextMaintenance"].Value = DateTime.Now.AddHours(Convert.ToDouble(config.UsageTime)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                        //row.Cells["NextMaintenance"].Value = DateTime.Now.AddSeconds(Convert.ToDouble(entityExt.SECONDS_BEFORE_EXPIRATION)).ToString("yyyy/MM/dd HH:mm:ss");
                        row.Cells["NextMaintenance"].Value = Convert.ToDateTime("1970-01-01 08:00:00").AddMilliseconds(Convert.ToDouble(expireddate)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    row.Cells["ScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["EquipNo"].Value = entityExt.EQUIPMENT_NUMBER;
                    row.Cells["EquipmentIndex"].Value = entityExt.EQUIPMENT_INDEX;
                    row.Cells["Status"].Value = ReflowClient.Properties.Resources.ok;
                    row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(0, 192, 0);

                    //因为同一组可能需要上两个,所以做上设备的时候只会显示一行成OK       郑培聪     2017/11/18
                    break;
                }
            }
        }

        private bool CheckEquipmentIsExist(string partNumber)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == partNumber && row.Cells["EquipNo"].Value != null && row.Cells["EquipNo"].Value.ToString() == partNumber
                    && row.Cells["UsCount"].Value.ToString() != "")
                {
                    errorHandler(3, Message("msg_The equipment already exist."), "");
                    return false;
                }
            }
            return true;
        }

        private bool CheckEquipmentValid(string ePartNumber)
        {
            bool isValid = false;
            if (string.IsNullOrEmpty(ePartNumber))
                return false;
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                if (item.Cells["eqPartNumber"].Value.ToString() == ePartNumber)// "EQUIPMENT_STATE", "ERROR_CODE", "PART_NUMBER"
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }

        private bool CheckEquipmentDuplication(string[] values, ref string ePartNumber, ref string eIndex)
        {
            int iCount = 0;
            ePartNumber = "";
            eIndex = "0";
            for (int i = 0; i < values.Length; i += 6)
            {
                if (values[i] == "0")
                {
                    ePartNumber = values[i + 2];
                    eIndex = values[i + 3];
                    iCount++;
                }

            }
            if (iCount > 1)
                return false;
            else
                return true;
        }

        private bool CheckEquipmentHasSetup(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            bool hasSetup = false;
            GetAttributeValue getAttributeHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] values = getAttributeHandler.GetAttributeValueForAll(15, equipmentNumber, equipmentIndex, attributeCode);
            if (values != null && values.Length > 0)
            {
                hasSetup = true;
            }
            return hasSetup;
        }

        private void AppendAttributeForEquipment(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            //取消属性不做值分配     郑培聪     20180228
            AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
            appendAttriHandler.AppendAttributeForAll(15, equipmentNumber, equipmentIndex, attributeCode, "Y");
        }

        private void RemoveAttributeForEquipment(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            RemoveAttributeValue removeAttriHandler = new RemoveAttributeValue(sessionContext, initModel, this);
            removeAttriHandler.RemoveAttributeForAll(15, equipmentNumber, equipmentIndex, attributeCode);
        }
        #endregion

        #endregion

        #region Network status
        private string strNetMsg = "Network Connected";
        private void picNet_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.Show(strNetMsg, this.picNet);
        }

        private void AvailabilityChanged(object sender, NetworkAvailabilityEventArgs e)
        {
            if (e.IsAvailable)
            {
                this.picNet.Image = ReflowClient.Properties.Resources.NetWorkConnectedGreen24x24;
                this.toolTip1.Show("Network Connected", this.picNet);
                strNetMsg = "Network Connected";
            }
            else
            {
                this.picNet.Image = ReflowClient.Properties.Resources.NetWorkDisconnectedRed24x24;
                this.toolTip1.Show("Network Disconnected", this.picNet);
                strNetMsg = "Network Disconnected";
            }
        }
        #endregion

        #region CheckList
        private void btnAddTask_Click(object sender, EventArgs e)
        {
            int iHour = DateTime.Now.Hour;
            if (8 <= iHour && iHour <= 18)
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
            }
            else
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
            }
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clResult1"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clSeq"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clDate"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clShift"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clStatus"].ReadOnly = true;
            gridCheckList.ClearSelection();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    //errorHandler(2, Message("$msg_checklist_first"), "");
            //    return;
            //}
            CheckListsCreate();
            #region
            //if (gridCheckList.Rows.Count > 0)
            //{
            //    string targetFileName = "";
            //    string shortFileName = config.StationNumber + "_" + this.gridCheckList.Rows[0].Cells["clShift"].Value.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            //    bool isOK = CreateTemplate(shortFileName, ref targetFileName);
            //    if (isOK)
            //    {
            //        Excel.Application xlsApp = null;
            //        Excel._Workbook xlsBook = null;
            //        Excel._Worksheet xlsSheet = null;
            //        try
            //        {
            //            GC.Collect();
            //            xlsApp = new Excel.Application();
            //            xlsApp.DisplayAlerts = false;
            //            xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //            xlsBook = xlsApp.ActiveWorkbook;
            //            xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;

            //            int iBeginIndex = 7;
            //            Excel.Range range = null;
            //            foreach (DataGridViewRow row in gridCheckList.Rows)
            //            {
            //                range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
            //                range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //                string strSeq = row.Cells["clSeq"].Value.ToString();
            //                string strItemName = row.Cells["clItemName"].Value.ToString();
            //                string strItemPoint = row.Cells["clPoint"].Value.ToString();
            //                string strItemStandard = row.Cells["clStandard"].Value.ToString();
            //                string strItemMethod = row.Cells["clMethod"].Value.ToString();
            //                string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
            //                string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            //                string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
            //                string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
            //                string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
            //                string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
            //                string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
            //                string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();
            //                xlsSheet.Cells[iBeginIndex, 1] = strSeq;
            //                xlsSheet.Cells[iBeginIndex, 2] = strItemName;
            //                xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
            //                xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
            //                xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
            //                xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
            //                xlsSheet.Cells[iBeginIndex, 7] = strCheckDate;
            //                xlsSheet.Cells[iBeginIndex, 8] = strException;
            //                xlsSheet.Cells[iBeginIndex, 9] = strHappendTime;
            //                xlsSheet.Cells[iBeginIndex, 10] = strProcessContent;
            //                xlsSheet.Cells[iBeginIndex, 11] = strProcessPersion;
            //                xlsSheet.Cells[iBeginIndex, 12] = strOperator;
            //                xlsSheet.Cells[iBeginIndex, 13] = strLeader;
            //                iBeginIndex++;
            //            }
            //            xlsBook.Save();
            //            errorHandler(0, "Save Production Check List success.(" + targetFileName + ")", "");
            //        }
            //        catch (Exception ex)
            //        {
            //            LogHelper.Error(ex);
            //        }
            //        finally
            //        {
            //            xlsBook.Close(false, Type.Missing, Type.Missing);
            //            xlsApp.Quit();
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

            //            xlsSheet = null;
            //            xlsBook = null;
            //            xlsApp = null;

            //            GC.Collect();
            //            GC.WaitForPendingFinalizers();
            //        }
            //    }
            //}
            #endregion
        }

        #region add by qy
        private void CheckListsCreate()
        {
            if (gridCheckList.Rows.Count > 0)
            {
                string targetFileName = "";
                string shortFileName = config.StationNumber + "_Reflow_" + DateTime.Now.ToString("yyyyMM");
                bool isOK = CreateTemplate(shortFileName, ref targetFileName);
                if (isOK)
                {
                    Excel.Application xlsApp = null;
                    Excel._Workbook xlsBook = null;
                    Excel._Worksheet xlsSheet = null;
                    try
                    {
                        GC.Collect();
                        xlsApp = new Excel.Application();
                        xlsApp.DisplayAlerts = false;
                        xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        xlsBook = xlsApp.ActiveWorkbook;
                        xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;
                        int count = xlsSheet.UsedRange.Cells.Rows.Count;

                        int iBeginIndex = count;
                        Excel.Range range = null;
                        foreach (DataGridViewRow row in gridCheckList.Rows)
                        {
                            range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
                            range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            string strSeq = row.Cells["clSeq"].Value.ToString();
                            string strItemName = row.Cells["clItemName"].Value.ToString();
                            string strItemPoint = row.Cells["clPoint"].Value.ToString();
                            string strItemStandard = row.Cells["clStandard"].Value.ToString();
                            string strItemMethod = row.Cells["clMethod"].Value.ToString();
                            string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
                            string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            string strShift = row.Cells["clShift"].Value.ToString();
                            string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
                            string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
                            string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
                            string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
                            string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
                            string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();

                            xlsSheet.Cells[iBeginIndex, 1] = iBeginIndex - 7;
                            xlsSheet.Cells[iBeginIndex, 2] = strItemName;
                            xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
                            xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
                            xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
                            xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
                            xlsSheet.Cells[iBeginIndex, 7] = strShift;
                            xlsSheet.Cells[iBeginIndex, 8] = strCheckDate;
                            xlsSheet.Cells[iBeginIndex, 9] = strException;
                            xlsSheet.Cells[iBeginIndex, 10] = strHappendTime;
                            xlsSheet.Cells[iBeginIndex, 11] = strProcessContent;
                            xlsSheet.Cells[iBeginIndex, 12] = strProcessPersion;
                            xlsSheet.Cells[iBeginIndex, 13] = strOperator;
                            xlsSheet.Cells[iBeginIndex, 14] = strLeader;

                            iBeginIndex++;
                        }
                        xlsBook.Save();
                        errorHandler(0, Message("msg_Save_CheckList_Success") + ".(" + targetFileName + ")", "");
                        SetTipMessage(MessageType.OK, Message("msg_Save_CheckList_Success") + ".(" + targetFileName + ")");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex);
                    }
                    finally
                    {
                        xlsBook.Close(false, Type.Missing, Type.Missing);
                        xlsApp.Quit();
                        KillSpecialExcel(xlsApp);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

                        xlsSheet = null;
                        xlsBook = null;
                        xlsApp = null;

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }

        private bool CheckCheckList()
        {
            bool result = true;
            foreach (DataGridViewRow row in gridCheckList.Rows)
            {
                string status = row.Cells["clStatus"].Value.ToString();
                if (status != "OK")
                {
                    result = false;
                    errorHandler(2, Message("msg_Verify_CheckList"), "");
                    break;
                }
            }
            return result;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        private static void KillSpecialExcel(Excel.Application m_objExcel)
        {
            try
            {
                if (m_objExcel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(m_objExcel.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private string GetCheckItemResult(string result1, string result2)
        {
            if (string.IsNullOrEmpty(result1))
                return result2;
            if (string.IsNullOrEmpty(result2))
                return result1;
            else
                return "NA";
        }

        private void InitTaskData()
        {
            try
            {
                Dictionary<string, List<CheckListItemEntity>> dicTask = new Dictionary<string, List<CheckListItemEntity>>();
                XDocument xdc = XDocument.Load("TaskFile.xml");
                var stationNodes = from item in xdc.Descendants("StationNumber")
                                   where item.Attribute("value").Value == config.StationNumber
                                   select item;
                XElement stationNode = stationNodes.FirstOrDefault();
                var tasks = from item in stationNode.Descendants("shift")
                            select item;
                foreach (XElement node in tasks.ToList())
                {
                    string shiftValue = GetNoteAttributeValues(node, "value");
                    List<CheckListItemEntity> itemList = new List<CheckListItemEntity>();
                    var items = from item in node.Descendants("Item")
                                select item;
                    foreach (XElement subItem in items.ToList())
                    {
                        CheckListItemEntity entity = new CheckListItemEntity();
                        entity.ItemName = GetNoteAttributeValues(subItem, "name");
                        entity.ItemPoint = GetNoteAttributeValues(subItem, "point");
                        entity.ItemStandard = GetNoteAttributeValues(subItem, "standard");
                        entity.ItemMethod = GetNoteAttributeValues(subItem, "method");
                        entity.ItemInputType = GetNoteAttributeValues(subItem, "inputType");
                        itemList.Add(entity);
                    }
                    if (!dicTask.ContainsKey(shiftValue))
                    {
                        dicTask[shiftValue] = itemList;
                    }
                }
                //init check list grid
                string strInputValue = GetNoteDescendantsValues(stationNode, "DataInputType");
                string[] strInputValues = strInputValue.Split(new char[] { ',' });
                DataTable dtInput = new DataTable();
                dtInput.Columns.Add("name");
                dtInput.Columns.Add("value");
                DataRow rowEmpty = dtInput.NewRow();
                rowEmpty["name"] = "";
                rowEmpty["value"] = "";
                dtInput.Rows.Add(rowEmpty);
                foreach (var strValues in strInputValues)
                {
                    DataRow row = dtInput.NewRow();
                    row["name"] = strValues;
                    row["value"] = strValues;
                    dtInput.Rows.Add(row);
                }
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DataSource = dtInput;
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DisplayMember = "Name";
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).ValueMember = "Value";

                int iHour = DateTime.Now.Hour;
                int seq = 1;
                if (8 <= iHour && iHour <= 18)
                {
                    if (dicTask.ContainsKey("白班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["白班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "白班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
                else
                {
                    if (dicTask.ContainsKey("晚班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["晚班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }

        private string GetNoteAttributeValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Attribute(attributename).Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private string GetNoteDescendantsValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Descendants(attributename).First().Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private string GetNoteDescendantsAttributeValues(XElement node, string nodeName, string attributeName)
        {
            string strValue = "";
            try
            {
                strValue = node.Descendants(nodeName).First().Attribute(attributeName).Value;
            }
            catch (Exception ex)
            {
                //MeasuredOctet 
                strValue = node.Descendants("RepairAction").First().Attribute("repairKey").Value;
                LogHelper.Info(node.ToString());
                LogHelper.Info("Node Name: " + nodeName);
                LogHelper.Info("Attribute Name: " + attributeName);
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private void SetCheckListInputStatus()
        {
            foreach (DataGridViewRow row in this.gridCheckList.Rows)
            {
                if (row.Cells["clInputType"].Value.ToString() == "1")
                {
                    row.Cells["clResult1"].ReadOnly = true;
                }
                else if (row.Cells["clInputType"].Value.ToString() == "2")
                {
                    row.Cells["clResult2"].ReadOnly = true;
                }
                row.Cells["clSeq"].ReadOnly = true;
                row.Cells["clDate"].ReadOnly = true;
                row.Cells["clShift"].ReadOnly = true;
                row.Cells["clStatus"].ReadOnly = true;
            }
        }

        private void gridCheckList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (this.gridCheckList.Columns[e.ColumnIndex].Name == "clResult1" && this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString() != "")
            {
                //verify the input value
                string strRegex = @"^(\d{0,9}.\d{0,9})-(\d{0,9}.\d{0,9}).*$";
                string strResult1 = this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString();
                string strStandard = this.gridCheckList.Rows[e.RowIndex].Cells["clStandard"].Value.ToString();
                Match match = Regex.Match(strStandard, strRegex);
                if (match.Success)
                {
                    if (match.Groups.Count > 2)
                    {
                        double iMin = Convert.ToDouble(match.Groups[1].Value);
                        double iMax = Convert.ToDouble(match.Groups[2].Value);
                        double iResult = Convert.ToDouble(strResult1);
                        if (iResult >= iMin && iResult <= iMax)
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "OK";
                        }
                        else
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                        }
                    }
                }
                else
                {
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                }
            }
            //else if (this.gridCheckList.Columns[e.ColumnIndex].Name == "clResult2" && this.gridCheckList.Rows[e.RowIndex].Cells["clResult2"].Value != null
            //    && this.gridCheckList.Rows[e.RowIndex].Cells["clResult2"].Value.ToString() != "")
            //{
            //    //verify the select value
            //    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
            //    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "OK";
            //}
        }

        #region Grid ComboBox
        int iRowIndex = -1;
        private void gridCheckList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            }
        }

        public void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(combox_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "OK";
                    }
                    else
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.White;
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void combox_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
        }
        #endregion

        int iIndexCheckList = -1;
        private void gridCheckList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.gridCheckList.Rows.Count == 0)
                    return;
                this.gridCheckList.ContextMenuStrip = contextMenuStrip2;
                iIndexCheckList = ((DataGridView)sender).CurrentRow.Index;
                ((DataGridView)sender).CurrentRow.Selected = true;
            }
        }

        private bool CreateTemplate(string strFileName, ref string targetFileName)
        {
            bool bFlag = true;
            targetFileName = "";
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath);
            string strExportPath = _appDir + @"\CheckListFiles\";
            //临时文件目录
            if (Directory.Exists(strExportPath) == false)
            {
                Directory.CreateDirectory(strExportPath);
            }
            string strSourceFileName = strExportPath + @"CheckListTemplate.xls";
            string strTargetFileName = config.CheckListFolder + strFileName + ".xls";
            targetFileName = strTargetFileName;
            if (!Directory.Exists(config.CheckListFolder))
                Directory.CreateDirectory(config.CheckListFolder);
            if (File.Exists(targetFileName))
            {
                return true;
            }
            if (System.IO.File.Exists(strSourceFileName))
            {
                try
                {
                    System.IO.File.Copy(strSourceFileName, strTargetFileName, true);
                    //去掉文件Readonly,避免不可写
                    FileInfo file = new FileInfo(strTargetFileName);
                    if ((file.Attributes & FileAttributes.ReadOnly) > 0)
                    {
                        file.Attributes ^= FileAttributes.ReadOnly;
                    }
                }
                catch (Exception ex)
                {
                    bFlag = false;
                    LogHelper.Error(ex);
                    throw ex;
                }
            }
            else
            {
                bFlag = false;
            }

            return bFlag;
        }

        private void checkListAdd_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                int iHour = DateTime.Now.Hour;
                if (8 <= iHour && iHour <= 18)
                {
                    gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
                }
                else
                {
                    gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
                }
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clResult1"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clSeq"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clDate"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clShift"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clStatus"].ReadOnly = true;
                gridCheckList.ClearSelection();
            }
        }

        private void checkListDelete_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                this.gridCheckList.Rows.RemoveAt(iIndexCheckList);
                int seq = 1;
                foreach (DataGridViewRow row in this.gridCheckList.Rows)
                {
                    row.Cells["clSeq"].Value = seq;
                    seq++;
                }
                this.gridCheckList.ClearSelection();
            }
        }
        #endregion

        #region PFC
        private SocketClientHandler cSocket = null;
        private object _lock1 = new Object();
        public void ProcessPFCMessage(string pfcMsg)
        {
            //#!PONGCRLF
            //#!BCTOPBARCODECRLF
            //#!BCBOTTOMBARCODECRLF
            //#!BOARDAVCRLF
            //#!TRANSFERBARCODECRLF
            lock (_lock1)
            {
                if (isFormOutPoump)
                {
                    return;
                }
                if (!VerifyCheckList())
                {
                    return;
                }

                #region 验证设备        郑培聪     20180227
                if (!CheckEquipmentSetup())
                {
                    return;
                }
                #endregion

                //errorHandler(0, "Receive message from PFC " + pfcMsg.TrimEnd(), "");
                SetConnectionText(0, "Receive message from PFC " + pfcMsg.TrimEnd());
                LogHelper.Info("Receive message from PFC " + pfcMsg.TrimEnd());
                if (pfcMsg.Length >= 10)
                {
                    bool isOK = true;
                    string messageType = pfcMsg.Substring(2, 8).TrimEnd();
                    switch (messageType)
                    {
                        case "PONG":
                            PFCStartTime = DateTime.Now;
                            break;
                        case "BCTOP":
                            string serialNumber = pfcMsg.Substring(10).TrimEnd();
                            isOK = ProcessSerialNumber(serialNumber, 0);
                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber);
                            }
                            break;
                        case "BCBOTTOM":
                            string serialNumber1 = pfcMsg.Substring(10).TrimEnd();
                            isOK = ProcessSerialNumber(serialNumber1, 1);
                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber1);
                            }
                            break;
                        case "BOARDAV"://todo
                            //isOK = ProcessSerialNumberData();
                            //if (isOK)
                            //{
                            //    SendMsessageToPFC(PFCMessage.GO, "");
                            //    BoardCome = false;
                            //}
                            break;
                        case "TRANSFER":

                            pfcMsg = pfcMsg.Substring(10).TrimEnd().Replace("#!PONG", "").Trim();
                            string serialNumber2 = pfcMsg.Replace("/r", "");
                            SendMsessageToPFC(PFCMessage.COMPLETE, serialNumber2);
                            SendMessageToCOM(serialNumber2);
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Receive message length less then 10.");
                }
            }
        }

        private object _lockSend = new Object();

        PosGraghicsForm frm = null;
      
        public void SendMsessageToPFC(PFCMessage msgType, string serialNumber)
        {
            lock (_lockSend)
            {
                //#!PINGCRLF
                //#!GOBARCODECRLF
                //#!COMPLETEBARCODECRLF
                string prefix = "#!";
                string suffix = HexToStr1("0D") + HexToStr1("0A");
                string sendMessage = "";
                switch (msgType)
                {
                    case PFCMessage.PING:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                    case PFCMessage.GO:
                        sendMessage = prefix + PFCMessage.GO.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.COMPLETE:
                        sendMessage = prefix + PFCMessage.COMPLETE.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.CONFIRM:
                        sendMessage = prefix + PFCMessage.CONFIRM.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    default:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                }
                //send message through socket
                try
                {
                    if (DateTime.Now.Subtract(PFCStartTime).Seconds >= 20)
                    {
                        cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                        PFCStartTime = DateTime.Now;
                        Thread.Sleep(1000);
                    }
                    bool isOK = cSocket.send(sendMessage);
                    if (isOK)
                    {
                        //errorHandler(1, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    else
                    {
                        //errorHandler(2, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                        bool isConnectOK = cSocket.connect(config.IPAddress, config.Port);
                        if (isConnectOK)
                        {
                            isOK = cSocket.send(sendMessage);
                            if (isOK)
                            {
                                SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                            else
                            {
                                SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                        }
                        else
                        {
                            SetConnectionText(1, "Conncet to PFC error");
                        }
                    }
                }
                catch (Exception ex)
                {
                    cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                    bool isOK = cSocket.send(sendMessage);
                    if (isOK)
                    {
                        //errorHandler(0, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    else
                    {
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    LogHelper.Error(ex.Message, ex);
                }
            }
        }

        public static string HexToStr1(string mHex) // 返回十六进制代表的字符串
        {
            mHex = mHex.Replace(" ", "");
            if (mHex.Length <= 0) return "";
            byte[] vBytes = new byte[mHex.Length / 2];
            for (int i = 0; i < mHex.Length; i += 2)
                if (!byte.TryParse(mHex.Substring(i, 2), NumberStyles.HexNumber, null, out vBytes[i / 2]))
                    vBytes[i / 2] = 0;
            return ASCIIEncoding.Default.GetString(vBytes);
        }

        public void GetTimerStart()
        {
            // 循环间隔时间(1分钟)
            CheckConnectTimer.Interval = Convert.ToInt16(config.CHECKCONECTTIME) * 1000;
            // 允许Timer执行
            CheckConnectTimer.Enabled = true;
            // 定义回调
            CheckConnectTimer.Elapsed += new ElapsedEventHandler(CheckConnectTimer_Elapsed);
            // 定义多次循环
            CheckConnectTimer.AutoReset = true;
        }

        private void CheckConnectTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (isFormOutPoump)
            {
                return;
            }
            SendMsessageToPFC(PFCMessage.PING, "");
        }

        private void SendMessageToCOM(string msg)
        {
            try
            {
                initModel.scannerHandler.handler().Write(msg + "\r\n");
                LogHelper.Info("Send message to COM :" + msg);
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }
        #endregion

        #region 首件检查
        string IPIstatus = "";
        private void CheckIPIStatus()
        {
            if (config.IPI_STATUS_CHECK == "ENABLE" || config.Production_Inspection_CHECK == "ENABLE")
            {
                GetWorkPlanData workplanHandle = new GetWorkPlanData(sessionContext, initModel, this);
                string[] workplanDataResultValues = workplanHandle.GetWorkplanDataForStation(GetWorkOrderValue());
                string workplanid = workplanDataResultValues[2];
                string workstep = workplanDataResultValues[3];

                AttributeManager attribHandler = new AttributeManager(sessionContext, initModel, this);
                string[] attributeResultValues = attribHandler.GetAttributeValueForWorkStep("IPI_STATE_UPDATE", workplanid, workstep);
                if (attributeResultValues.Length > 0)
                {
                    if (attributeResultValues[1] == "Y")
                    {
                        string[] valuesAttri = attribHandler.GetAttributeValueForAll(initModel.configHandler.StationNumber, 1, GetWorkOrderValue(), "-1", "IPI_STATUS");
                        if (valuesAttri != null && valuesAttri.Length > 0)
                        {
                            IPIstatus = valuesAttri[1];
                        }
                    }
                }
            }
        }
        private void UpdateIPIStatus(string result)
        {
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                if (IPIstatus == "0")
                {
                    AttributeManager attribHandler = new AttributeManager(sessionContext, initModel, this);
                    if (result == "0")//成功
                        attribHandler.AppendAttributeForAll(initModel.configHandler.StationNumber, 1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "1");
                    else//失败
                        attribHandler.AppendAttributeForAll(initModel.configHandler.StationNumber, 1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "-1");

                    IPIstatus = "";
                }
            }
        }
        #endregion

        private void btnRefreshWo_Click(object sender, EventArgs e)
        {
            //if (!CheckCheckList())
            //{
            //    return;
            //}



            string processlayerPre = "";
            string WorkorderPre = this.txbCDAMONumber.Text;
            if (initModel.currentSettings != null)
                processlayerPre = initModel.currentSettings.processLayer.ToString();
            GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, this);
            initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
            if (initModel.currentSettings == null)
            {
                return;
            }
            else
            {
                #region 添加P板数量      郑培聪     20180301
                GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(initModel.currentSettings.partNumber);
                if (listData != null && listData.Count > 0)
                {
                    MdataGetPartData mData = listData[0];
                    initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                }
                #endregion
            }
            this.txbCDAMONumber.Text = initModel.currentSettings.workorderNumber;
            this.txbCDAPartNumber.Text = initModel.currentSettings.partNumber;
            iProcessLayer = initModel.currentSettings.processLayer;
            this.Invoke(new MethodInvoker(delegate
            {
                LoadYield();
            }));
            InitDocumentGrid();
            if (config.RESTORE_TIME != "" && config.RESTORE_TREAD_TIMER != "")
            {
                GetRestoreTimerStart();
            }
            if (initModel.currentSettings.workorderNumber != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer.ToString())
            {
                strShiftChecklist = "";
                InitWorkOrderType();
                InitShift2(WorkorderPre);//20161215 add by qy
                if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                {
                    if (!CheckShiftChange2())
                    {
                        InitTaskData_SOCKET("开线点检;设备点检");
                        isStartLineCheck = true;
                    }
                    else
                    {
                        if (!ReadCheckListFile())//20161214 edit by qy
                        {
                        InitTaskData_SOCKET("开线点检");
                        isStartLineCheck = true;
                    }
                    }
                }
            }

            //获取设备信息        郑培聪     20180227
            if (workOrderForEquipment != txbCDAMONumber.Text)
            {
                InitEquipmentGridEXT();
                workOrderForEquipment = txbCDAMONumber.Text;
            }

            CheckIPIStatus();
        }

        #region production inspection
        private void UpdateIPIStatusForProductionInspection(string result, string serialnumber)
        {
            if (config.Production_Inspection_CHECK == "ENABLE")
            {
                if (IPIstatus == "1")
                {
                    GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                    int error = getHandle.GetSerialNumberByref(serialnumber);
                    if (error == -203)
                        serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);

                    GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                    string[] valuesAttri = getattribute.GetAttributeValueForAll(0, serialnumber, "-1", "IPI");
                    if (valuesAttri != null && valuesAttri.Length > 0)
                    {
                        AttributeManager appendAttri = new AttributeManager(sessionContext, initModel, this);
                        if (result != "0")
                            appendAttri.AppendAttributeForAll(initModel.configHandler.StationNumber, 1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "-2");

                        IPIstatus = "";
                    }

                }
            }
        }

        #endregion

        #region
        string OKlist = "";
        string NGlist = "";
        private System.Timers.Timer FileCleanUpTimer;

        private void InitCheckResultMapping()
        {
            string[] LineList = File.ReadAllLines("CheckResultMappingFile.txt", Encoding.Default);
            foreach (var line in LineList)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    string[] strs = line.Split(new char[] { ';' });
                    if (strs[0] == "OK")
                    {
                        OKlist = OKlist + "," + strs[1];
                    }
                    else
                    {
                        NGlist = NGlist + "," + strs[1];
                    }
                }
            }
        }
        public void GetFileCleanUpTimerStart()
        {
            FileCleanUpTimerdelete= new System.Timers.Timer();
            if (config.FILE_CLEANUP == null || config.FILE_CLEANUP == "")
                return;
            // 循环间隔时间(1分钟)
            FileCleanUpTimerdelete.Interval = Convert.ToInt32(config.FILE_CLEANUP_TREAD_TIMER) * 1000;
            // 允许Timer执行
            FileCleanUpTimerdelete.Enabled = true;
            // 定义回调
            FileCleanUpTimerdelete.Elapsed += new ElapsedEventHandler(FileCleanUpTimer_Elapsed);
            // 定义多次循环
            FileCleanUpTimerdelete.AutoReset = true;
        }

        private void FileCleanUpTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            string deletefiletype = config.FILE_CLEANUP_FOLDER_TYPE;
            FileCleanUp fileup = new FileCleanUp(sessionContext, initModel, this);
            fileup.DeleteFolderFile(deletefiletype);
        }

        #endregion

        private void btnShowPCB_Click(object sender, EventArgs e)
        {
            if (frm != null && frm.pictureBox1.Image != null)
            {
                frm.Hide();
            }
            if (txbCDAPartNumber.Text == "")
            {
                errorHandler(2, Message("msg_No activated work order"), "");
                return;
            }
            frm = new PosGraghicsForm(this, sessionContext, initModel, txbCDAPartNumber.Text, initModel.currentSettings == null ? "-1" : initModel.currentSettings.processLayer.ToString());
            frm.Show();
            if (frm.pictureBox1.Image == null)
                frm.Hide();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SendSN(config.LIGHT_CHANNEL_OFF);
        }

        #region checklist from OA
        bool Supervisor = false;
        bool IPQC = true;
        private void InitTaskData_SOCKET(string djclass)
        {
            try
            {
                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, Message("msg_no active wo"), "");
                    return;
                }
                this.Invoke(new MethodInvoker(delegate
                {
                    try
                    {
                        this.dgvCheckListTable.Rows.Clear();

                        Supervisor = false;
                        IPQC = true;
                        GetWorkPlanData handle = new GetWorkPlanData(sessionContext, initModel, this);
                        int firstSN = int.Parse(this.lblPass.Text) + int.Parse(this.lblFail.Text) + int.Parse(this.lblScrap.Text);
                        if (firstSN == 0)
                        {
                            djclass = djclass + ";首末件点检";
                        }

                        string workstep_text = handle.GetWorkStepInfobyWorkPlan(this.txbCDAMONumber.Text, initModel.currentSettings.processLayer);
                        if (workstep_text != "")
                        {
                            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                            string[] processCode = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "TTE_PROCESS_CODE");
                            if (processCode != null && processCode.Length > 0)
                            {
                                string process = processCode[1];
                                string sedmessage = "{getCheckListItem;" + PartNumber + ";" + process + ";[" + workstep_text + "];[" + djclass + "];" + "}";
                                string returnMsg = checklist_cSocket.SendData(sedmessage);

                                if (returnMsg != "" && returnMsg != null)
                                {
                                    string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                    string status = values[1];
                                    if (status == "0")//“0” , or “-1” (error)  
                                    {
                                        int seq = 1;
                                        string itemregular = @"\{[^\{\}]+\}"; //@"\[[^\[\]]+\]";
                                        MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                        if (match.Count <= 0)
                                        {
                                            errorHandler(2, Message("msg_No checklist data"), "");
                                            return;
                                        }
                                        else
                                            SetTipMessage(MessageType.OK, "");
                                        for (int i = 0; i < match.Count; i++)
                                        {
                                            string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                            //string[] datas = data.Split(';');
                                            string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                            string sourceclass = datas[4];//数据来源
                                            string formno = datas[0];//对应单号
                                            string itemno = datas[1];//机种品号
                                            string itemnname = datas[2];//机种品名
                                            string sbno = datas[5];//设备编号
                                            string sbname = datas[6];//设备名称
                                            string gcno = datas[7];//过程编号
                                            string gcname = datas[8];//过程名称
                                            string lbclass = datas[9];//类别
                                            string djxmname = datas[10];//点检项目
                                            string specvalue = datas[11];//规格值
                                            string djkind = datas[12];//点检类型
                                            string maxvalue = datas[14];//上限值
                                            string minvalue = datas[13];//下限值
                                            string djclase = datas[15];//点检类别
                                            string djversion = datas[3];//版本
                                            string dataclass = datas[16];//状态

                                            object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                            this.dgvCheckListTable.Rows.Add(objValues);
                                            seq++;
                                            SetCheckListInputStatusTable();

                                            if (djkind == "判断值")
                                            {
                                                string[] strInputValues = new string[] { "Y", "N" };
                                                DataTable dtInput = new DataTable();
                                                dtInput.Columns.Add("name");
                                                dtInput.Columns.Add("value");
                                                DataRow rowEmpty = dtInput.NewRow();
                                                rowEmpty["name"] = "";
                                                rowEmpty["value"] = "";
                                                dtInput.Rows.Add(rowEmpty);
                                                foreach (var strValues in strInputValues)
                                                {
                                                    DataRow row = dtInput.NewRow();
                                                    row["name"] = strValues;
                                                    row["value"] = strValues;
                                                    dtInput.Rows.Add(row);
                                                }

                                                DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                ComboBoxCell.DataSource = dtInput;
                                                ComboBoxCell.DisplayMember = "Name";
                                                ComboBoxCell.ValueMember = "Value";
                                                dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                            }

                                            this.dgvCheckListTable.ClearSelection();
                                        }
                                    }
                                    else
                                    {
                                        string errormsg = values[1];
                                        errorHandler(2, errormsg, "");
                                    }
                                }
                                else
                                {
                                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                                    returnMsg = checklist_cSocket.SendData(sedmessage);

                                    if (returnMsg != "" && returnMsg != null)
                                    {
                                        string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                        string status = values[1];
                                        if (status == "0")//“0” , or “-1” (error)  
                                        {
                                            int seq = 1;
                                            string itemregular = @"\{[^\{\}]+\}";
                                            MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                            if (match.Count <= 0)
                                            {
                                                errorHandler(2, Message("msg_No checklist data"), "");
                                                return;
                                            }
                                            else
                                                SetTipMessage(MessageType.OK, "");
                                            for (int i = 0; i < match.Count; i++)
                                            {
                                                string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                                //string[] datas = data.Split(';');
                                                string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                                string sourceclass = datas[4];//数据来源
                                                string formno = datas[0];//对应单号
                                                string itemno = datas[1];//机种品号
                                                string itemnname = datas[2];//机种品名
                                                string sbno = datas[5];//设备编号
                                                string sbname = datas[6];//设备名称
                                                string gcno = datas[7];//过程编号
                                                string gcname = datas[8];//过程名称
                                                string lbclass = datas[9];//类别
                                                string djxmname = datas[10];//点检项目
                                                string specvalue = datas[11];//规格值
                                                string djkind = datas[12];//点检类型
                                                string maxvalue = datas[14];//上限值
                                                string minvalue = datas[13];//下限值
                                                string djclase = datas[15];//点检类别
                                                string djversion = datas[3];//版本
                                                string dataclass = datas[16];//状态

                                                object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                                this.dgvCheckListTable.Rows.Add(objValues);
                                                seq++;
                                                SetCheckListInputStatusTable();

                                                if (djkind == "判断值")
                                                {
                                                    string[] strInputValues = new string[] { "Y", "N" };
                                                    DataTable dtInput = new DataTable();
                                                    dtInput.Columns.Add("name");
                                                    dtInput.Columns.Add("value");
                                                    DataRow rowEmpty = dtInput.NewRow();
                                                    rowEmpty["name"] = "";
                                                    rowEmpty["value"] = "";
                                                    dtInput.Rows.Add(rowEmpty);
                                                    foreach (var strValues in strInputValues)
                                                    {
                                                        DataRow row = dtInput.NewRow();
                                                        row["name"] = strValues;
                                                        row["value"] = strValues;
                                                        dtInput.Rows.Add(row);
                                                    }

                                                    DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                    ComboBoxCell.DataSource = dtInput;
                                                    ComboBoxCell.DisplayMember = "Name";
                                                    ComboBoxCell.ValueMember = "Value";
                                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                                }

                                                this.dgvCheckListTable.ClearSelection();
                                            }
                                        }
                                        else
                                        {
                                            string errormsg = values[1];
                                            errorHandler(2, errormsg, "");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                errorHandler(2, Message("msg_no TTE_PROCESS_CODE"), "");//20161213 edit by qy
                                return;
                            }
                        }
                        else
                        {
                            errorHandler(2, Message("msg_no workstep text"), "");
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex.Message, ex);
                    }
                }));

            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void SetCheckListInputStatusTable()
        {
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjkind"].Value.ToString() == "判断值")
                {
                    row.Cells["tabResult1"].ReadOnly = true;
                }
                else if (row.Cells["tabdjkind"].Value.ToString() == "输入值" || row.Cells["tabdjkind"].Value.ToString() == "范围值")
                {
                    row.Cells["tabResult2"].ReadOnly = true;
                }
                row.Cells["tabNo"].ReadOnly = true;
                row.Cells["tabStatus"].ReadOnly = true;
            }
        }

        private void btnSupervisor_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQC_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }

        public void SupervisorConfirm(string user)//班长确认
        {
            DialogResult dr = MessageBox.Show(Message("msg_produtc or not"), Message("msg_Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                Supervisor = true;
                errorHandler(0, Message("msg_supervisor confirm OK"), "");
            }
            else
            {
                Supervisor = false;
                errorHandler(2, Message("msg_supervisor confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;1;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }

        }

        public void IPQCConfirm(string user)//IPQC巡检
        {
            DialogResult dr = MessageBox.Show(Message("msg_IPQC produtc or not"), Message("msg_Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                IPQC = true;
                errorHandler(0, Message("msg_IPQC confirm OK"), "");
            }
            else
            {
                IPQC = false;
                errorHandler(2, Message("msg_IPQC confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;2;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }
        }

        private void btnAddCheckListTable_Click(object sender, EventArgs e)
        {
            dgvCheckListTable.Rows.Add(new object[] { this.dgvCheckListTable.Rows.Count + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" });

            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult1"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabNo"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabStatus"].ReadOnly = true;
            dgvCheckListTable.ClearSelection();
        }
        string sequece = "";
        private void dgvCheckListTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //20161208 edit by qy
            try
            {
                if (e.RowIndex == -1)
                    return;
                if (this.dgvCheckListTable.Columns[e.ColumnIndex].Name == "tabResult1" && this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() != "")
                {
                    if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "范围值")
                    {
                        //verify the input value
                        string strRegex = @"^\d{0,9}\.\d{0,9}|-\d{0,9}\.\d{0,9}";//@"^(\d{0,9}.\d{0,9})～(\d{0,9}.\d{0,9}).*$";"^(\-|\+?\d{0,9}.\d{0,9})～(\-|\+?\d{0,9}.\d{0,9})$"
                        string strResult1 = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString();
                        string strStandard = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString().Replace("（", "").Replace("）", "");
                        string strMax = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabmaxvalue"].Value.ToString();
                        string strMin = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabminvalue"].Value.ToString();
                        Match match1 = Regex.Match(strMax, strRegex);
                        Match match2 = Regex.Match(strMin, strRegex);
                        if (match1.Success && match2.Success)
                        {
                            //if (match.Groups.Count > 2)
                            //{
                            //double iMin = Convert.ToDouble(match.Groups[1].Value);
                            //double iMax = Convert.ToDouble(match.Groups[2].Value);
                            double iMin = Convert.ToDouble(match2.ToString());
                            double iMax = Convert.ToDouble(match1.ToString());
                            double iResult = Convert.ToDouble(strResult1);
                            if (iResult >= iMin && iResult <= iMax)
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                            }
                            else
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                            }
                            //}
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                    else if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "输入值")
                    {
                        if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() == this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString())
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }

        }

        private void dgvCheckListTable_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBoxTable_SelectedIndexChanged);
            }
        }

        public void ComboBoxTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(comboxtable_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "OK";
                    }
                    else
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.White;
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void comboxtable_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBoxTable_SelectedIndexChanged);
        }

        int iIndexCheckListTable = -1;
        private void dgvCheckListTable_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvCheckListTable.Rows.Count == 0)
                    return;
                ((DataGridView)sender).CurrentRow.Selected = true;
                iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
                this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;

                if (iIndexCheckListTable == -1)
                    this.dgvCheckListTable.ContextMenuStrip = null;

            }
        }
        private void dgvCheckListTable_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    if (this.dgvCheckListTable.Rows.Count == 0)
            //        return;

            //    if (e.RowIndex == -1)
            //    {
            //        this.dgvCheckListTable.ContextMenuStrip = null;
            //        return;
            //    }

            //    iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
            //    this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;
            //    ((DataGridView)sender).CurrentRow.Selected = true;
            //}
        }

        private void SaveCheckList()
        {
            try
            {
                string path = @"CheckList.txt";
                string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(datetime);
                sb.AppendLine(txbCDAMONumber.Text + ";" + initModel.currentSettings.processLayer);
                sb.AppendLine(Supervisor.ToString());
                sb.AppendLine(IPQC.ToString());
                sb.AppendLine(sequece);
                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                {
                    string sourceclass = row.Cells["tabsourceclass"].Value.ToString();//数据来源
                    string formno = row.Cells["tabformno"].Value.ToString();//对应单号
                    string itemno = row.Cells["tabitemno"].Value.ToString();//机种品号
                    string itemnname = row.Cells["tabitemname"].Value.ToString();//机种品名
                    string sbno = row.Cells["tabsbno"].Value.ToString();//设备编号
                    string sbname = row.Cells["tabsnname"].Value.ToString();//设备名称
                    string gcno = row.Cells["tabgcno"].Value.ToString();//过程编号
                    string gcname = row.Cells["tabgcname"].Value.ToString();//过程名称
                    string lbclass = row.Cells["tablbclass"].Value.ToString();//类别
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();//点检项目
                    string specvalue = row.Cells["tabspecname"].Value.ToString();//规格值
                    string result1 = row.Cells["tabResult1"].Value.ToString();
                    string result2 = row.Cells["tabResult2"].Value == null ? "" : row.Cells["tabResult2"].Value.ToString();// row.Cells["tabResult2"].Value.ToString();
                    string status = row.Cells["tabstatus"].Value.ToString();//结果
                    string djkind = row.Cells["tabdjkind"].Value.ToString();//点检类型
                    string maxvalue = row.Cells["tabmaxvalue"].Value.ToString();//上限值
                    string minvalue = row.Cells["tabminvalue"].Value.ToString();//下限值
                    string djclase = row.Cells["tabdjclass"].Value.ToString();//点检类别
                    string djversion = row.Cells["tabdjversion"].Value.ToString();//版本
                    string dataclass = row.Cells["tabdataclass"].Value.ToString();//状态

                    //string cell13 = row.Cells[13].Value == null ? "" : row.Cells[13].Value.ToString();
                    string linedata = sourceclass + "￥" + formno + "￥" + itemno + "￥" + itemnname + "￥" + sbno + "￥" + sbname + "￥" + gcno + "￥" + gcname + "￥" + lbclass + "￥" + djxmname + "￥" + specvalue + "￥" + result1 + "￥" + result2 + "￥" + status + "￥" + djkind + "￥" + maxvalue + "￥" + minvalue + "￥" + djclase + "￥" + djversion + "￥" + dataclass;
                    //string linedata = row.Cells[1].Value.ToString() + ";" + row.Cells[2].Value.ToString() + ";" + row.Cells[3].Value.ToString() + ";" + row.Cells[4].Value.ToString() + ";" + row.Cells[5].Value.ToString() + ";" + row.Cells[6].Value.ToString() + ";" + row.Cells[7].Value.ToString() + ";" + row.Cells[8].Value.ToString() + ";" + row.Cells[9].Value.ToString() + ";" + row.Cells[10].Value.ToString() + ";" + row.Cells[11].Value.ToString() + ";" + row.Cells[12].Value.ToString() + ";" + cell13 + ";" + row.Cells[14].Value.ToString() + ";" + row.Cells[15].Value.ToString() + ";" + row.Cells[16].Value.ToString() + ";" + row.Cells[17].Value.ToString() + ";" + row.Cells[18].Value.ToString() + ";" + row.Cells[19].Value.ToString() + ";" + row.Cells[20].Value.ToString() + ";" + djkind;
                    sb.AppendLine(linedata);
                }

                FileStream fs = new FileStream(path, FileMode.Create);
                byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
                fs.Seek(0, SeekOrigin.Begin);
                fs.Write(bt, 0, bt.Length);
                fs.Flush();
                fs.Close();
                LogHelper.Info("Save checklist file success.");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private bool ReadCheckListFile()
        {
            try
            {
                string path = @"CheckList.txt";
                if (File.Exists(path))
                {
                    string[] linelist = File.ReadAllLines(path);
                    string datetimespan = linelist[0];
                    string workorder = linelist[1];
                    Supervisor = Convert.ToBoolean(linelist[2]);
                    IPQC = Convert.ToBoolean(linelist[3]);
                    sequece = linelist[4];
                    TimeSpan span = DateTime.Now - Convert.ToDateTime(datetimespan);

                    if (span.TotalMinutes > Convert.ToInt32(config.RESTORE_TIME))//判断是否大于10分钟，大于10分钟则不自动点检
                    {
                        return false;
                    }
                    else
                    {
                        string[] workorders = workorder.Split(';');
                        if (workorders.Length > 1)
                        {
                            if (workorders[0] == this.txbCDAMONumber.Text)//判断工单是否有变化，无变化则自动点检
                            {
                                //if (workorders[1] == initModel.currentSettings.processLayer.ToString())//判断面次是否有变化
                                //{
                                #region setup checklist
                                int seq = 1;
                                if (linelist.Count() <= 6)
                                    return false;
                                this.dgvCheckListTable.Rows.Clear();
                                for (int i = 5; i < linelist.Count(); i++)
                                {
                                    string line = linelist[i];
                                    if (string.IsNullOrEmpty(line.Trim()))
                                        continue;

                                    string[] datas = line.Split('￥');
                                    object[] objValues = new object[] { seq, datas[17], datas[9], datas[7], datas[10], datas[11], "", datas[13], datas[14], datas[6], datas[15], datas[16], datas[8], datas[0], datas[1], datas[2], datas[3], datas[4], datas[5], datas[18], datas[19], "" };
                                    this.dgvCheckListTable.Rows.Add(objValues);
                                    seq++;
                                    SetCheckListInputStatusTable();
                                    if (datas[14] == "判断值")
                                    {
                                        string[] strInputValues = new string[] { "Y", "N" };
                                        DataTable dtInput = new DataTable();
                                        dtInput.Columns.Add("name");
                                        dtInput.Columns.Add("value");
                                        DataRow rowEmpty = dtInput.NewRow();
                                        rowEmpty["name"] = "";
                                        rowEmpty["value"] = "";
                                        dtInput.Rows.Add(rowEmpty);
                                        foreach (var strValues in strInputValues)
                                        {
                                            DataRow row = dtInput.NewRow();
                                            row["name"] = strValues;
                                            row["value"] = strValues;
                                            dtInput.Rows.Add(row);
                                        }

                                        DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                        ComboBoxCell.DataSource = dtInput;
                                        ComboBoxCell.DisplayMember = "Name";
                                        ComboBoxCell.ValueMember = "Value";
                                        dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                    }
                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"].Value = datas[12];
                                    this.dgvCheckListTable.ClearSelection();

                                }
                                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                                {
                                    if (row.Cells["tabStatus"].Value.ToString() == "OK")
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                    }
                                    else if ((row.Cells["tabStatus"].Value.ToString() == "NG"))
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.Red;
                                    }

                                }
                                return true;
                                #endregion
                                //}
                                //else
                                //{
                                //    return false;
                                //}
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }
        }
        string strShift = "";
        private void WriteIntoShift2()
        {
            string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            strShift = datetime;
            string path = @"CheckListShiftTemp.txt";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(datetime + ";" + this.txbCDAMONumber.Text);
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
            fs.Seek(0, SeekOrigin.Begin);
            fs.Write(bt, 0, bt.Length);
            fs.Flush();
            fs.Close();
        }

        //检查有没有到换班时间，如果到换班时间
        string strShiftChecklist = "";
        private bool CheckShiftChange2()
        {
            try
            {
                bool isValid = false;
                if (strShiftChecklist == "")
                    return false;

                string[] shifchangetimes = config.SHIFT_CHANGE_TIME.Split(';');
                List<string> shiftList = new List<string>();
                string nowDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                for (int i = 0; i < shifchangetimes.Length; i++)
                {

                    shiftList.Add(DateTime.Now.ToString("yyyy/MM/dd ") + shifchangetimes[i].Substring(0, 2) + ":" + shifchangetimes[i].Substring(2, 2));

                }

                shiftList.Sort();

                for (int j = shiftList.Count - 1; j < shiftList.Count; j--)
                {
                    if (j == -1)
                        break;
                    LogHelper.Info("shift time: " + shiftList[j]);
                    string shitftime = shiftList[j];

                    if (Convert.ToDateTime(nowDate) > Convert.ToDateTime(shiftList[j])) //当前时间与设定的时间做比较，如果到换班时间则比较上次点检的时间
                    {
                        if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        {
                            isValid = true;
                        }
                        break;
                    }
                    else
                    {
                        if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        {
                            string covert_datetime = nowDate;
                            if (j == shiftList.Count - 1)
                            {
                                covert_datetime = shiftList[j - 1];
                            }
                            else if (j == 0)
                            {
                                covert_datetime = shiftList[j];
                            }
                            if (Convert.ToDateTime(strShiftChecklist) < Convert.ToDateTime(nowDate) && Convert.ToDateTime(nowDate) < Convert.ToDateTime(covert_datetime))
                            {
                                isValid = true;
                            }
                            break;
                        }

                        //if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        //{
                        //    shitftime = Convert.ToDateTime(shitftime).AddDays(-1).ToString("yyyy/MM/dd HH:mm:ss");

                        //    if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        //    {
                        //        isValid = true;
                        //    }
                        //    break;
                        //}
                    }
                }

                return isValid;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }

        }

        private void InitShift2(string wo)
        {
            string path = @"CheckListShiftTemp.txt";
            if (File.Exists(path))
            {
                string[] content = File.ReadAllLines(path);

                foreach (var item in content)
                {
                    if (item != "")
                    {
                        string[] items = item.Split(';');
                        //if (items[1] == wo)
                        //{
                        strShiftChecklist = items[0];
                        break;
                        //}
                    }
                }
            }
        }
        DateTime next_checklist_time = DateTime.Now;
        string checklist_freq_time = "";
        private void InitWorkOrderType()
        {

            Dictionary<string, string> dicfreq = new Dictionary<string, string>();
            string CHECKLIST_FREQ = config.CHECKLIST_FREQ;
            string[] freqs = CHECKLIST_FREQ.Split(';');
            foreach (var item in freqs)
            {
                string[] items = item.Split(',');
                string key = items[0];
                if (key == "")
                    key = "OTHERS";
                dicfreq[key] = items[1];
            }

            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "WORKORDER_TYPE");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string value = valuesAttri[1];

                if (CHECKLIST_FREQ.Contains(value))
                {
                    checklist_freq_time = dicfreq[value];
                }
                else
                {
                    checklist_freq_time = dicfreq["OTHERS"];
                }
            }
            else
            {
                checklist_freq_time = dicfreq["OTHERS"];
            }
            if (strShiftChecklist != "")
            {
                next_checklist_time = Convert.ToDateTime(strShiftChecklist).AddMinutes(double.Parse(checklist_freq_time) * 60);
            }
            else
            {
                next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
            }

        }

        private void InitProductionChecklist()
        {
            if (DateTime.Now > next_checklist_time)
            {
                if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                {
                    InitTaskData_SOCKET("过程点检");
                    isStartLineCheck = false;
                    next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
                }
            }
        }

        public void GetRestoreTimerStart()
        {

            if (RestoreMaterialTimer != null && RestoreMaterialTimer.Enabled)
                return;
            RestoreMaterialTimer = new System.Timers.Timer();
            // 循环间隔时间(1分钟)
            RestoreMaterialTimer.Interval = Convert.ToInt32(config.RESTORE_TREAD_TIMER) * 1000;
            // 允许Timer执行
            RestoreMaterialTimer.Enabled = true;
            // 定义回调
            RestoreMaterialTimer.Elapsed += new ElapsedEventHandler(RestoreMaterialTimer_Elapsed);
            // 定义多次循环
            RestoreMaterialTimer.AutoReset = true;
        }

        private void RestoreMaterialTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SaveCheckList();
            InitProductionChecklist();
            InitShiftCheckList();
        }

        bool IsGetShiftCheckList = false;
        private void InitShiftCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //InitShift2(txbCDAMONumber.Text);
                if (!CheckShiftChange2())
                {
                    if (this.dgvCheckListTable.Rows.Count <= 0 || (this.dgvCheckListTable.Rows.Count > 0 && !isStartLineCheck))//!IsShiftCheck()
                    {
                        InitTaskData_SOCKET("开线点检;设备点检");
                        isStartLineCheck = true;
                    }
                }
            }
        }
        private bool IsShiftCheck()//true 表示已经带出开线点检的内容了
        {
            bool isValid = false;
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjclass"].Value.ToString() == "开线点检")
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }

        public void OpenScanPort()
        {
            initModel.scannerHandler = new ScannerHeandler(initModel, this);
            initModel.scannerHandler.handler().DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
            initModel.scannerHandler.handler().Open();
        }

        private void btnSupervisorTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQCTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }

        private void btnConfirmTable_Click(object sender, EventArgs e)
        {
            try
            {

                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, Message("msg_no active wo"), "");
                    return;
                }
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value == null || row.Cells["tabStatus"].Value.ToString() == "")
                    {
                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return;
                    }
                }

                string headmessage = "{appendCheckListResult;" + PartNumber;
                string sedmessage = "";
                string date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    string gdcode = this.txbCDAMONumber.Text;
                    string itemno = PartNumber;
                    string itemname = initModel.currentSettings.partdesc;
                    string gczcode = initModel.configHandler.StationNumber;
                    string gczname = "";
                    string lineclass = "";
                    string lbclass = row.Cells["tablbclass"].Value.ToString();
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();
                    string specvalue = "";
                    if (row.Cells["tabResult1"].Value.ToString() != "")
                        specvalue = row.Cells["tabResult1"].Value.ToString();
                    else
                        specvalue = row.Cells["tabResult2"].Value.ToString();
                    string djkind = row.Cells["tabdjkind"].Value.ToString();
                    string maxvalues = row.Cells["tabmaxvalue"].Value.ToString();
                    string minvalues = row.Cells["tabminvalue"].Value.ToString();
                    string djclass = row.Cells["tabdjclass"].Value.ToString();
                    string djversion = row.Cells["tabdjversion"].Value.ToString();
                    string djuser = lblUser.Text;
                    string djremark = "";
                    string djdate = date;
                    string jcuser = lblUser.Text;
                    string qruser = "";
                    string pguser = "";

                    string msgrow = "{" + gdcode + "#!#" + itemno + "#!#" + itemname + "#!#" + gczcode + "#!#" + gczname + "#!#" + lineclass + "#!#" + lbclass + "#!#" + djxmname + "#!#" + specvalue + "#!#" + djkind + "#!#" + maxvalues + "#!#" + minvalues + "#!#" + djclass + "#!#" + djversion + "#!#" + djuser + "#!#" + djremark + "#!#" + djdate + "#!#" + jcuser + "#!#" + qruser + "#!#" + pguser + "}";
                    if (sedmessage == "")
                        sedmessage = msgrow;
                    else
                        sedmessage = sedmessage + ";" + msgrow;
                }
                string endsendmessage = headmessage + ";" + sedmessage + "}";
                string returnMsg = checklist_cSocket.SendData(endsendmessage);
                if (returnMsg != null && returnMsg != "")
                {
                    returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                    string[] Msgs = returnMsg.Split(';');
                    if (Msgs[1] == "0")
                    {
                        if (Supervisor_OPTION == "1")
                        {
                            Supervisor = true;
                            errorHandler(0, Message("msg_Send_CheckList_Success"), "");
                        }
                        else
                        {
                            errorHandler(0, Message("msg_Send_CheckList_Success,please supervisor confirm"), "");
                        }

                        sequece = Msgs[3];
                        SaveCheckList();
                        WriteIntoShift2();
                        InitShift2(txbCDAMONumber.Text);
                    }
                    else
                    {
                        errorHandler(2, Message("msg_Send_CheckList_fail"), "");
                    }
                }
                else
                {
                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                    returnMsg = checklist_cSocket.SendData(endsendmessage);
                    if (returnMsg != null && returnMsg != "")
                    {
                        returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                        string[] Msgs = returnMsg.Split(';');
                        if (Msgs[1] == "0")
                        {
                            if (Supervisor_OPTION == "1")
                            {
                                Supervisor = true;
                                errorHandler(0, Message("msg_Send_CheckList_Success"), "");
                            }
                            else
                            {
                                errorHandler(0, Message("msg_Send_CheckList_Success,please supervisor confirm"), "");
                            }

                            sequece = Msgs[3];
                            SaveCheckList();
                            WriteIntoShift2();
                            InitShift2(txbCDAMONumber.Text);
                        }
                        else
                        {
                            errorHandler(2, Message("msg_Send_CheckList_fail"), "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private bool VerifyCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //if (!CheckShiftChange2())
                //{
                //    if (this.dgvCheckListTable.Rows.Count <= 0 && dgvCheckListTable.Rows[0].Cells["tabdjclass"].Value.ToString() != "开线点检")
                //    {
                //        InitTaskData_SOCKET("开线点检");
                //    }

                //}
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value.ToString() != "OK")
                    {
                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return false;
                    }
                }
                if (this.dgvCheckListTable.Rows.Count > 0)
                {
                    if (!Supervisor)
                    {
                        errorHandler(2, Message("msg_Superivisor_check_fail"), "");
                        return false;
                    }
                    if (!IPQC)
                    {

                        errorHandler(2, Message("msg_IPQC_check_fail"), "");
                        return false;
                    }
                }

                return true;
            }
            else
            {
                foreach (DataGridViewRow row in gridCheckList.Rows)
                {
                    if (row.Cells["clStatus"].Value.ToString() != "OK")
                    {

                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return false;
                    }
                }
                if (this.gridCheckList.Rows.Count > 0)
                {
                    if (!Supervisor)
                    {

                        errorHandler(2, Message("msg_Superivisor_check_fail"), "");
                        return false;
                    }
                    if (!IPQC)
                    {

                        errorHandler(2, Message("msg_IPQC_check_fail"), "");
                        return false;
                    }
                }
                return true;
            }
        }

        #endregion

        public bool isFormOutPoump = false;
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.txbCDADataInput.Text = "";
            LogoutForm frmOut = new LogoutForm(UserName, this, initModel, sessionContext);
            DialogResult dr = frmOut.ShowDialog();

            if (dr == DialogResult.OK)
            {
                UserName = frmOut.UserName;
                lblUser.Text = UserName;
                lblLoginTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                sessionContext = frmOut.sessionContext;
                if (config.LogInType == "COM")
                {
                    SerialPort serialPort = new SerialPort();
                    serialPort.PortName = config.SerialPort;
                    serialPort.BaudRate = int.Parse(config.BaudRate);
                    serialPort.Parity = (Parity)int.Parse(config.Parity);
                    serialPort.StopBits = (StopBits)1;
                    serialPort.Handshake = Handshake.None;
                    serialPort.DataBits = int.Parse(config.DataBits);
                    serialPort.NewLine = "\r";
                    serialPort.DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
                    serialPort.Open();
                    initModel.scannerHandler.SetSerialPortData(serialPort);
                }
            }
        }
    }
}
