using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Unitech.R1000.Reader;
using Unitech.R1000.Reader.Constants;
using Unitech.R1000.Reader.Structures;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Threading;
using System.Collections.Specialized;
using System.Collections;




namespace RFIDDemoCS
{
    public partial class frmMain : Form
    {
        private Boolean m_bOpen = false;
        private Boolean m_bStart = false;
        private Boolean m_bContinue = false;
        private Boolean m_Encontrado = false;

        private delegate void StopInvDelegate();
        private delegate void InsertItemDelegate(ACCESS_DATA lpAccessData, int nAntenna, Boolean bAdd);
        private delegate void ErrorMessageDelegate(Result nStatus, Int32 nErrorCode, String strMsg);
        private ScanTrigger scanTrigger = new ScanTrigger();
        private CallbackDelegate m_fnStopProc = null;

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            #region codes ========================================
            Result nRet = Result.FAILURE;
            String strVersion = String.Empty;
            nRet = R1000Reader.RFIDCreate(ref strVersion);
            if (nRet != Result.OK)
            {
                ErrorMessage(nRet, 0, "Initialized Reader Failed.");
                return;
            }
            //lblVersion.Text = "rfid.dll Ver:" + strVersion;

            btnOpen_Click(sender, e);
            scanTrigger.TriggerDown += new TriggerEventHandle(TriggerProc);
           // m_fnStopProc = new CallbackDelegate(InvStopProc);
            m_fnStopProc = new CallbackDelegate(InvStopProc1);
            #endregion //end codes
        }

        private void frmMain_Closed(object sender, EventArgs e)
        {
            #region codes ========================================
            R1000Reader.RFIDClose(0);
            R1000Reader.RFIDDestroy();
            scanTrigger.Dispose();
            #endregion //end codes
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            #region codes ========================================
            Result nRet = Result.FAILURE;

            if (!m_bOpen)
            {
                nRet = R1000Reader.RFIDOpen(0);
                if (nRet != Result.OK)
                {
                    ErrorMessage(nRet, 0, String.Empty);
                    return;
                }

                //set anntena parameter for best Inventory
                AntennaPortConfig pConfig = new AntennaPortConfig();
                R1000Reader.RFIDGetAntennaPortConfiguration(0, ref pConfig);
                //pConfig.powerLevel = 280;
                //pConfig.dwellTime = 300;
                //pConfig.numberInventoryCycles = 8192;
                //R1000Reader.RFIDSetAntennaPortConfiguration(0, ref pConfig);

                //For speed up block write below settings are necessary.
                //FIXEDQ_PARMS FixedParms = new FIXEDQ_PARMS();
                //FixedParms.qValue = 0;
                //FixedParms.retryCount = 0;
                //FixedParms.toggleTarget = 0;
                //FixedParms.repeatUntilNoTags = 0;

                //SINGULATION_ALGORITHM_PARMS AlgParms = (SINGULATION_ALGORITHM_PARMS)FixedParms;

                //R1000Reader.RFIDSingulationAlgorithmParameters(SingulationAlgorithm.FIXEDQ, ref AlgParms, true);

                //SingulationAlgorithm nAlgorithm = SingulationAlgorithm.FIXEDQ;
                //R1000Reader.RFIDSingulationAlgorithm(ref nAlgorithm, true);

                btnOpen.Text = "Close";
                m_bOpen = true;

                ErrorMessage(0, 0, "Ready");
            }//end if (!m_bOpen)
            else
            {
                R1000Reader.RFIDClose(0);
                btnOpen.Text = "Open";
                m_bOpen = false;
                ErrorMessage(0, 0, "Close");
            }
            #endregion //end codes
        }

        public static byte[] StrToByteArray(string str)
        {
            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            return encoding.GetBytes(str);
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            #region codes ========================================

            byte[] bytes = new byte[1024];
            if (m_bOpen)
            {
                IPAddress address = IPAddress.Parse(textBoxIP.Text);
                IPEndPoint lep = new IPEndPoint(address, 35500);
                Socket socket = new Socket(AddressFamily.InterNetwork,
                                   SocketType.Stream,
                                         ProtocolType.Tcp);

                socket.Connect(lep);

                for (int i = 0; i < lstView.Items.Count; i++)
                {
                    string ID = lstView.Items[i].SubItems[2].Text;
                    string times = lstView.Items[i].SubItems[1].Text;
                    string space = "&";
                    string end = "#";

                    if (times.Length != 4)
                    {
                        if (times.Length == 1)
                            times = String.Concat("000", times);
                        if (times.Length == 2)
                            times = String.Concat("00", times);
                        if (times.Length == 3)
                            times = String.Concat("0", times);
                        if (times.Length > 4)
                            times = "MAX";
                    }


                    string concat = String.Concat(ID, space, times, end);

                    if (i == (lstView.Items.Count - 1))
                        concat = String.Concat(concat, "%");
                    byte[] temp = StrToByteArray(concat);

                    // Encode the data string into a byte array.
                    byte[] msg = Encoding.ASCII.GetBytes(concat);

                    // Send the data through the socket.
                    int bytesSent = socket.Send(msg);
                }
                // Receive the response from the remote device.
                //int bytesRec = socket.Receive(bytes);
                //Console.WriteLine("Echoed test = {0}",
                //                  Encoding.ASCII.GetString(bytes, 0, bytesRec));

                // Release the socket.
                socket.Shutdown(SocketShutdown.Both);
                socket.Close();
                ErrorMessage(0, 0, "Completed!");

            }
            #endregion //end codes
        }

        private static ManualResetEvent sendDone =
        new ManualResetEvent(false);

        private static void SendCallback(IAsyncResult ar)
        {
            try
            {
                // Retrieve the socket from the state object.
                Socket client = (Socket)ar.AsyncState;

                // Complete sending the data to the remote device.
                int bytesSent = client.EndSend(ar);
                Console.WriteLine("Sent {0} bytes to server.", bytesSent);

                // Signal that all bytes have been sent.
                sendDone.Set();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            #region codes ========================================
            if (m_bOpen)
            {
                R1000Reader.RFIDCancelOperation();
            }
            #endregion //end codes
        }

        private void TriggerProc(object sender, TriggerEventArgs args)
        {
            #region codes ========================================
            if (m_bOpen && !m_bStart)
            {
                m_bStart = true;
                RFID_INVENTORY stInventory = new RFID_INVENTORY();

                ACCESS_STATUS stAccessStatus = new ACCESS_STATUS();

                //operation in Non-blocking mode
                //stInventory.lpfnStartProc = new CallbackDelegate(InvStartProc);
                //stInventory.lpfnStopProc = new CallbackDelegate(InvStopProc);
                stInventory.lpfnStopProc = m_fnStopProc;
                R1000Reader.RFIDInventory(stInventory, ref stAccessStatus, false, 0);
            }
            scanTrigger.DoneTrigger();
            #endregion //end codes
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
           
        }

        private void StopInventory()
        {
            #region codes ========================================
            m_bStart = false;
            btnStart.Text = "Inico";
            #endregion //end codes
        }

        public Int32 InvStopProc1([In] IntPtr hWnd, [In] UInt32 nMessage, [In, Out] UInt32 wParam, ref ACCESS_STATUS lpAccessStatus)
        {
            #region codes ========================================
            ErrorMessageDelegate errorMessage = new ErrorMessageDelegate(ErrorMessage);
            this.Invoke(errorMessage, new Object[] { (Result)lpAccessStatus.dwStatus, (Int32)lpAccessStatus.dwErrorCode, String.Empty });
            m_Encontrado = false;
            InsertItemDelegate insertItem = new InsertItemDelegate(BuscarCodigo);
            RetrieveData(ref lpAccessStatus, BuscarCodigo, true);

            if (m_bContinue)
            {
                if ((Result)lpAccessStatus.dwStatus == Result.OPERATION_CANCELLED)
                {
                    this.Invoke(new StopInvDelegate(StopInventory));
                }
            }
            else
            {
                this.Invoke(new StopInvDelegate(StopInventory));
            }

            return 1;
            #endregion //end codes
        }

        public Int32 InvStopProc([In] IntPtr hWnd, [In] UInt32 nMessage, [In, Out] UInt32 wParam, ref ACCESS_STATUS lpAccessStatus)
        {
            #region codes ========================================
            ErrorMessageDelegate errorMessage = new ErrorMessageDelegate(ErrorMessage);
            this.Invoke(errorMessage, new Object[] { (Result)lpAccessStatus.dwStatus, (Int32)lpAccessStatus.dwErrorCode, String.Empty });

            InsertItemDelegate insertItem = new InsertItemDelegate(InsertItem);
            RetrieveData(ref lpAccessStatus, insertItem, true);

            if (m_bContinue)
            {
                if ((Result)lpAccessStatus.dwStatus == Result.OPERATION_CANCELLED)
                {
                    this.Invoke(new StopInvDelegate(StopInventory));
                }
            }
            else
            {
                this.Invoke(new StopInvDelegate(StopInventory));
            }

            return 1;
            #endregion //end codes
        }


        private void InsertItem(ACCESS_DATA accessData, int nAntenna, Boolean bAdd)
        {
            #region codes ========================================

            int i, nCount = 0;
            int nSize = lstView.Items.Count;
            Boolean bFind = false;

            Byte[] bEPC = new Byte[accessData.unEPCLength - 4];
            Utils.Copy(accessData.pnEPC, 2, bEPC, 0, accessData.unEPCLength - 4);

            String strEPC = String.Empty;
            Utils.Byte2Hex(bEPC, ref strEPC);

            for (i = 0; i < nSize; i++)
            {
                ListViewItem listItem = lstView.Items[i];
                if (listItem.SubItems[2].Text == strEPC)
                {
                    if (bAdd == true)
                    {
                        nCount = Convert.ToInt32(listItem.SubItems[1].Text);
                        nCount++;
                        listItem.SubItems[1].Text = nCount.ToString();
                        lstView.Items[i] = listItem;
                    }
                    bFind = true;
                    break;
                }
            }

            if (bFind == false)
            {
                nCount = 1;
                ListViewItem listItem = new ListViewItem(nSize.ToString());
                listItem.SubItems.Add(nCount.ToString());
                listItem.SubItems.Add(strEPC);
                lstView.Items.Add(listItem);


                const string PAPY_IP = "192.168.92.136";
                const UInt16 PAPY_PORT = 1999;

                //try
                //{
                //    string data = "e" + strEPC + ";" + DateTime.Now;
                //    Socket sender = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);

                //    EndPoint ep = new IPEndPoint(PAPY_IP, PAPY_PORT);
                //    sender.Connect(ep);
                //    sender.Send(Encoding.ASCII.GetBytes(data), SocketFlags.None);
                //    sender.Close();
                //}
                //catch (Exception ex)
                //{
                //    lblStatus.Text = ex.Message;
                //}

            }

            #endregion //end codes
        }

        private void RetrieveData(ref ACCESS_STATUS lpAccessStatus, InsertItemDelegate insertItem, Boolean bAdd)
        {
            #region codes ========================================
            if (lpAccessStatus.dwStatus == 0 && lpAccessStatus.dwErrorCode == 0)
            {
                for (int i = 0; i < lpAccessStatus.unAntennas; i++)
                {
                    ANTENNA_STATUS stAntennaStatus = new ANTENNA_STATUS();
                    R1000Reader.RFIDGetAntennaStatus(i, ref stAntennaStatus);
                    for (int j = 0; j < stAntennaStatus.unCount; j++)
                    {
                        ACCESS_DATA accessData = new ACCESS_DATA();
                        UInt32 nRet = R1000Reader.RFIDGetAccessData(i, j, ref accessData);
                        if (nRet == 1 && accessData.unEPCLength > 0)
                        {
                            this.Invoke(insertItem, new object[] { accessData, stAntennaStatus.unAntenna, bAdd });
                        }
                    }
                }
            }
            #endregion //end codes
        }

        private void tabMain_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region codes ========================================
            if (tabMain.SelectedIndex == 0)
            {
                lblStatus.Text = String.Empty;
            }
            else if (tabMain.SelectedIndex == 1)
            {
                btnRefresh_Click(sender, e);
                lblStatus.Text = String.Empty;
            }
            #endregion
        }

        private void cmbResponseMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region codes ========================================
            if (m_bOpen)
            {
                ResponseMode responseMode = ResponseMode.UNKNOWN;
                if (cmbResponseMode.SelectedIndex == 0)
                {
                    responseMode = ResponseMode.COMPACT;
                }
                else
                {
                    responseMode = ResponseMode.NORMAL;
                }
                Result nRet = R1000Reader.RFIDSetResponseMode(responseMode);
                if (nRet == Result.OK)
                {
                    ErrorMessage(0, 0, "Set Response Mode OK");
                }
                else
                {
                    ErrorMessage(0, 0, "Set Response Mode Failure");
                }
            }
            #endregion //end codes
        }

        private void cmbOperationMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region codes ========================================
            Result nRet = Result.FAILURE;
            if (m_bOpen)
            {
                if (cmbOperationMode.SelectedIndex == 0)
                {
                    nRet = R1000Reader.RFIDSetOperationMode(RadioOperationMode.CONTINUOUS);
                    m_bContinue = true;
                }
                else
                {
                    nRet = R1000Reader.RFIDSetOperationMode(RadioOperationMode.NONCONTINUOUS);
                    m_bContinue = false;
                }

                if (nRet == Result.OK)
                {
                    ErrorMessage(0, 0, "Set Operation Mode OK");
                }
                else
                {
                    ErrorMessage(0, 0, "Set Operation Mode Failure");
                }
            }
            #endregion //end codes
        }

        private void numTagStopCount_ValueChanged(object sender, EventArgs e)
        {
            #region codes ========================================
            if (m_bOpen)
            {
                R1000Reader.RFIDSetStopCount(Convert.ToInt32(numTagStopCount.Value));
            }
            #endregion //end codes
        }

        private void traPowerLevel_ValueChanged(object sender, EventArgs e)
        {
            #region codes ========================================
            lbldBm.Text = traPowerLevel.Value.ToString() + "dBm";
            #endregion //end codes
        }

        private void btnSet_Click(object sender, EventArgs e)
        {
            #region codes ========================================
            Result nRet = Result.FAILURE;

            AntennaPortConfig antPortConfig = new AntennaPortConfig();
            nRet = R1000Reader.RFIDGetAntennaPortConfiguration(0, ref antPortConfig);
            if (nRet != Result.OK)
            {
                ErrorMessage(0, 0, "Get Antenna Config Failure");
                return;
            }

            antPortConfig.dwellTime = Convert.ToUInt32(txtdwelltime.Text);
            antPortConfig.numberInventoryCycles = Convert.ToUInt32(txtInvRounds.Text);
            antPortConfig.physicalRxPort = 3;
            antPortConfig.physicalTxPort = 3;
            antPortConfig.powerLevel = (UInt32)traPowerLevel.Value * 10;

            nRet = R1000Reader.RFIDSetAntennaPortState(0, AntennaPortState.ENABLED);
            if (nRet != Result.OK)
            {
                ErrorMessage(0, 0, "Set Antenna Port State Failure");
                return;
            }

            nRet = R1000Reader.RFIDSetAntennaPortConfiguration(0, ref antPortConfig);
            if (nRet != Result.OK)
            {
                ErrorMessage(0, 0, "Set Antenna Config Failure");
            }
            else
            {
                ErrorMessage(0, 0, "Set Config OK");
            }
            #endregion //end codes
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            #region codes ========================================
            if (m_bOpen)
            {
                //Response Mode
                ResponseMode responseMode = ResponseMode.UNKNOWN;
                R1000Reader.RFIDGetResponseMode(ref responseMode);
                if (responseMode == ResponseMode.COMPACT)
                {
                    cmbResponseMode.SelectedIndex = 0;
                }
                else
                {
                    cmbResponseMode.SelectedIndex = 1;
                }

                //Operation Mode
                RadioOperationMode operationMode = RadioOperationMode.UNKNOWN;
                operationMode = R1000Reader.RFIDGetOperationMode();
                if (operationMode == RadioOperationMode.CONTINUOUS)
                {
                    cmbOperationMode.SelectedIndex = 0;
                    m_bContinue = true;
                }
                else
                {
                    cmbOperationMode.SelectedIndex = 1;
                    m_bContinue = false;
                }

                //tag stop count
                numTagStopCount.Value = R1000Reader.RFIDGetStopCount();

                AntennaPortConfig antPortConfig = new AntennaPortConfig();
                R1000Reader.RFIDGetAntennaPortConfiguration(0, ref antPortConfig);
                txtdwelltime.Text = antPortConfig.dwellTime.ToString();
                txtInvRounds.Text = antPortConfig.numberInventoryCycles.ToString();
                traPowerLevel.Value = (int)antPortConfig.powerLevel / 10;
                lbldBm.Text = traPowerLevel.Value.ToString() + "dBm";

                ErrorMessage(0, 0, "Config Refreshed");
            }
            #endregion //end codes
        }

        private void ErrorMessage(Result nStatus, Int32 nErrorCode, String strMsg)
        {
            #region codes ========================================
            String strError = String.Empty;

            if (strMsg != String.Empty)
            {
                lblStatus.Text = strMsg;
                return;
            }

            if (nErrorCode != 0x00)
            {
                switch (nErrorCode)
                {
                    case 0x00:
                        strError = "RFID_STATUS_OK";
                        break;
                    case 0x01:	//Read after write verify failed.
                        strError = "Read after write verify failed.";
                        break;
                    case 0x02:	//problem transmitting tag command
                        strError = "problem transmitting tag command";
                        break;
                    case 0x03:	//CRC error on tag response to a write
                        strError = "CRC error on tag response to a write";
                        break;
                    case 0x04:	//CRC error on the read packet when verifying the write
                        strError = "CRC error on the read packet when verifying the write";
                        break;
                    case 0x05:	//Maximum retry's on the write exceeded
                        strError = "Maximum retry's on the write exceeded";
                        break;
                    case 0x06:	//Failed waiting for read data from tag, possible timeout.
                        strError = "Failed waiting for read data from tag, possible timeout.";
                        break;
                    case 0x07:	//Failure requesting a new tag handle.
                        strError = "Failure requesting a new tag handle.";
                        break;
                    case 0x0A:	//error waiting for tag response, possible timeout
                        strError = "Error waiting for tag response, possible timeout";
                        break;
                    case 0x0B:	//CRC error on tag response to a kill
                        strError = "CRC error on tag response to a kill";
                        break;
                    case 0x0C:	//problem transmitting 2nd half of tag kill.
                        strError = "problem transmitting 2nd half of tag kill.";
                        break;
                    case 0x0D:	//tag responded with an invalid handle on first kill command
                        strError = "tag responded with an invalid handle on first kill command";
                        break;
                    case 0xFA:
                        strError = "tag has insufficient power to perform the memory write";
                        break;
                    case 0xFB:
                        strError = "specified memory location is locked and/or permalocked";
                        break;
                    case 0xFC: //specified memory location does not exist of the PC value is not supported by the tag.
                        strError = "specified memory location does not exist";
                        break;
                    case 0xFD:
                        strError = "Tag failed to response within timeout";
                        break;
                    case 0xFE:
                        strError = "CRC was invalid";
                        break;
                    case 0xFF: //general error
                        //strError = "general error"));
                        break;
                    default:
                        break;
                }
            }
            else if (nStatus != 0x00)
            {
                switch (nStatus)
                {
                    case Result.OK:
                        strError = "RFID_STATUS_OK";
                        break;
                    case Result.NOT_INITIALIZED:
                        strError = "RFID_ERROR_NOT_INITIALIZED";
                        break;
                    case Result.INVALID_PARAMETER:
                        strError = "RFID_ERROR_INVALID_PARAMETER";
                        break;
                    case Result.INVALID_HANDLE:
                        strError = "RFID_ERROR_INVALID_HANDLE";
                        break;
                    case Result.NO_SUCH_RADIO:
                        strError = "RFID_ERROR_NO_SUCH_RADIO";
                        break;
                    case Result.ALREADY_OPEN:
                        strError = "RFID_ERROR_ALREADY_OPEN";
                        break;
                    case Result.DRIVER_MISMATCH:
                        strError = "RFID_ERROR_DRIVER_MISMATCH";
                        break;
                    case Result.OUT_OF_MEMORY:
                        strError = "RFID_ERROR_OUT_OF_MEMORY";
                        break;
                    case Result.CURRENTLY_NOT_ALLOWED:
                        strError = "RFID_ERROR_CURRENTLY_NOT_ALLOWED";
                        break;
                    case Result.RADIO_NOT_PRESENT:
                        strError = "RFID_ERROR_RADIO_NOT_PRESENT";
                        break;
                    case Result.RADIO_FAILURE:
                        strError = "RFID_ERROR_RADIO_FAILURE";
                        break;
                    case Result.RADIO_BUSY:
                        strError = "RFID_ERROR_RADIO_BUSY";
                        break;
                    case Result.RADIO_NOT_RESPONDING:
                        strError = "RFID_ERROR_RADIO_NOT_RESPONDING";
                        break;
                    case Result.EMULATION_MODE:
                        strError = "RFID_ERROR_EMULATION_MODE";
                        break;
                    case Result.BUFFER_TOO_SMALL:
                        strError = "RFID_ERROR_BUFFER_TOO_SMALL";
                        break;
                    case Result.FAILURE:
                        strError = "RFID_ERROR_FAILURE";
                        break;
                    case Result.DRIVER_LOAD:
                        strError = "RFID_ERROR_DRIVER_LOAD";
                        break;
                    case Result.NOT_SUPPORTED:
                        strError = "RFID_ERROR_NOT_SUPPORTED";
                        break;
                    case Result.OPERATION_CANCELLED:
                        strError = "RFID_ERROR_OPERATION_CANCELLED";
                        break;
                    default:
                        strError = "Unknow Error";
                        break;
                }
            }
            else if (nStatus == 0 && nErrorCode == 0)
            {
                strError = "Operation OK";
            }

            lblStatus.Text = strError;

            #endregion //end codes
        }

       
        
        private void textBoxIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(lstView.Items.Count.ToString());
            for (int i = 0; i < lstView.Items.Count; i++)
            {
                MessageBox.Show(lstView.Items[i].SubItems[2].Text);



            }
        }


        [DllImport("CoreDll.dll")]
        public static extern void MessageBeep(int code);

        public static void MessageBeep()
        {
         //    MessageBeep(-1);  // Default beep code is -1
            for (int i = 0; i < 10; i++)
            {
                MessageBeep(1);  // Default beep code is -1
            }
        }

        private void BuscarCodigo (ACCESS_DATA accessData, int nAntenna, Boolean bAdd)
        {
            #region codes ========================================


 
            
            int nSize = lstView.Items.Count;
            int largo;
            String Dato;
            Byte[] bEPC = new Byte[accessData.unEPCLength - 4];
            Utils.Copy(accessData.pnEPC, 2, bEPC, 0, accessData.unEPCLength - 4);

            String strEPC = String.Empty;
            Utils.Byte2Hex(bEPC, ref strEPC);

           // MessageBox.Show(strEPC.ToString ());
            txtBuscar.Text =  txtBuscar.Text.Replace("l2","13")  ;
            txtBuscar.Text = txtBuscar.Text.Replace("L2", "13");
            Dato = txtBuscar.Text;
            largo = Dato.Length;


            
            //MessageBox.Show (strEPC.Substring(1, largo ) ) ;

            if (txtBuscar.Text ==  (strEPC.Substring(0, largo )   )  & m_Encontrado ==false )
            {

                MessageBeep();
                MessageBox.Show("Encontrado " + strEPC);
               
                m_Encontrado = true;

                
                
            }

            
            

            #endregion //end codes
        }




        private void button2_Click(object sender, EventArgs e)
        {
            
            
        }

        private void lstView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnStart_Click_1(object sender, EventArgs e)
        {
            #region codes ========================================
            if (m_bOpen)
            {
                if (!m_bStart)
                {
                    m_bStart = true;
                    m_Encontrado = false;


                    ErrorMessage(0, 0, "Inventorying...");
                    btnStart.Text = "Parar";

                    lstView.Items.Clear();

                    RFID_INVENTORY stInventory = new RFID_INVENTORY();
                    ACCESS_STATUS stAccessStatus = new ACCESS_STATUS();

                    //operation in Non-blocking mode
                    stInventory.hWnd = this.Handle;
                    //stInventory.lpfnStartProc = new CallbackDelegate(InvStartProc);
                    //stInventory.lpfnStopProc = new CallbackDelegate(InvStopProc);                  
                    stInventory.lpfnStopProc = m_fnStopProc;

                    R1000Reader.RFIDInventory(stInventory, ref stAccessStatus, false, 0);
                }
                else
                {
                    R1000Reader.RFIDAbortOperation();
                    //StopInventory();
                }
            }//end if (m_bOpen)
            #endregion //end codes
        }

        private void txtBuscar_TextChanged(object sender, EventArgs e)
        {

        }
    }
}