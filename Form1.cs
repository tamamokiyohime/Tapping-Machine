using System;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net;               //載入網路
using System.Net.Sockets;
using Excel = Microsoft.Office.Interop.Excel;

using PCI_DMC;
using PCI_DMC_ERR;
using System.Collections.Generic;

namespace DMC_NET
{
    //2018/09/24 update
     
    public partial class Form1 : Form
    {
        Thread ThWorking, ThWorking_PLC, ThHome;
        String X0Message = "", X1Message = "";
        bool KtrBoolClear = false;
        bool newturn = false;
        int OneCirclePluse = 128000;
        int TransmissionRate = 1020;
        int cmd1 = 0, pos1 = 0, cmd2 = 0, pos2 = 0;
        short spd1 = 0, spd2 = 0, toe1 = 0, toe2 = 0;
        uint err1 = 0, err2 = 0;
        List<double> motorTorque1 = new List<double>();
        List<double> motorTorque2 = new List<double>();
        List<double> motorRpm1 = new List<double>();
        List<double> motorRpm2 = new List<double>();

        short existcard = 0, rc;
        ushort gCardNo = 0, DeviceInfo = 0, gnodeid;
        ushort[] gCardNoList = new ushort[16];
        uint[] SlaveTable = new uint[4];
        ushort[] NodeID = new ushort[32];
        byte[] value = new byte[10];
        ushort gNodeNum;
        bool  gIsServoOn;
        TextBox[] txtIoSts = new TextBox[16];

        Thread th;
        Socket T;
        int delayMotorDeg = 0;
        int rpmRate1 = 200; //ktr比例
        int rpmRate2 = 2;
        int torqueRate = 5;
        int count = 0;      //一個行程資料數目
        int excelTime = 0;  //excel陣列數目
        ushort node1 = 2, node2 = 1;    //節點    虎尾3.4 中山1.2
        //ushort node1 = 3, node2 = 4;
        bool b ;
        int com1, com2;
        List<double> ktrTorque1 = new List<double>();
        List<double> ktrTorque2 = new List<double>();
        List<double> ktrRpm1 = new List<double>();
        List<double> ktrRpm2 = new List<double>();
        double[,] rpm_1 = new double[90000, 10];
        double[,] rpm_2 = new double[90000, 10];
        double[] rpm_motor1 = new double[90000];
        double[] rpm_motor2 = new double[90000];
        double[,] torque_1 = new double[90000, 10];
        double[,] torque_2 = new double[90000, 10];
        double[] torque_motor1 = new double[90000];
        double[] torque_motor2 = new double[90000];

        public Form1()
        {
            InitializeComponent();
        }

        private void btninitial_Click(object sender, EventArgs e)
        {
            ushort i, card_no = 0;

            btnralm.Enabled = false;
            btnstop.Enabled = false;
            btnreset1.Enabled = false;
            btnNmove.Enabled = false;
            btnPmove.Enabled = false;
            chksvon.Enabled = false;

            for (i = 0; i < 4; i++)
            {
                SlaveTable[i] = 0;
            }
            btnFindSlave.Enabled = false;
            txtSlaveNum.Text = "0";
            CmbCardNo.Items.Clear();
            cmbNodeID.Items.Clear();

            rc = CPCI_DMC.CS_DMC_01_open(ref existcard);

            if (existcard <= 0)
                MessageBox.Show("No DMC-NET card can be found!");
            else
            {

                for (i = 0; i < existcard; i++)
                {
                    rc = CPCI_DMC.CS_DMC_01_get_CardNo_seq(i, ref card_no);
                    gCardNoList[i] = card_no;

                    CmbCardNo.Items.Insert(i, card_no);

                }

                btnFindSlave.Enabled = true;        //2011.08.05
                CmbCardNo.SelectedIndex = 0;
                gCardNo = gCardNoList[0];

                for (i = 0; i < existcard; i++)
                {
                    rc = CPCI_DMC.CS_DMC_01_pci_initial(gCardNoList[i]);
                    if (rc != 0)
                        MessageBox.Show("Can't boot PCI_DMC Master Card!");

                    rc = CPCI_DMC.CS_DMC_01_initial_bus(gCardNoList[i]);
                    if (rc != 0)
                    {
                        MessageBox.Show("Initial Failed!");
                    }
                    else
                    {
                        rc = CPCI_DMC.CS_DMC_01_start_ring(gCardNo, 0);                      //Start communication                      
                        rc = CPCI_DMC.CS_DMC_01_get_device_table(gCardNo, ref DeviceInfo);   //Get Slave Node ID 
                        rc = CPCI_DMC.CS_DMC_01_get_node_table(gCardNo, ref SlaveTable[0]);
                    }
                }
            }

        }
        private void chksvon_CheckedChanged(object sender, EventArgs e)
        {
            gIsServoOn = chksvon.Checked;
            gnodeid = ushort.Parse(cmbNodeID.Text);
            //btnWork.Enabled = true;
            rc =CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node1, 0, 1);
            rc =CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node1, 0, 1);     
            rc =CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node1, 0, (ushort)(gIsServoOn ? 1 : 0));

            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_ipulser_mode(gCardNo, node2, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_set_rm_04pi_opulser_mode(gCardNo, node2, 0, 1);
            rc = CPCI_DMC.CS_DMC_01_ipo_set_svon(gCardNo, node2, 0, (ushort)(gIsServoOn ? 1 : 0));
        }

        private void btnralm_Click(object sender, EventArgs e)
        {
            gnodeid = ushort.Parse(cmbNodeID.Text);
            rc =CPCI_DMC.CS_DMC_01_set_ralm(gCardNo, gnodeid, 0);
        }

        private void btnstop_Click(object sender, EventArgs e)
        {
            rc = CPCI_DMC.CS_DMC_01_emg_stop(gCardNo, node1, 0);
            rc = CPCI_DMC.CS_DMC_01_emg_stop(gCardNo, node2, 0);
            if (th!=null)
                th.Abort();
            ThWorking_PLC.Abort();
            ThWorking.Abort();
        }

        private void btnreset_Click(object sender, EventArgs e)
        {   
            gnodeid = ushort.Parse(cmbNodeID.Text);
            CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
            CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
            CPCI_DMC.CS_DMC_01_set_position(gCardNo, node2, 0, 0);
            CPCI_DMC.CS_DMC_01_set_command(gCardNo, node2, 0, 0);
            btnralm.Enabled = true;
            btnstop.Enabled = true;
            btnreset1.Enabled = true;
            btnNmove.Enabled = true;
            btnPmove.Enabled = true;
            chksvon.Checked = false;
            chksvon.Enabled = true;

            count = 0;
            excelTime = 0;
            ExcelPath.Text = "路徑:";
        }

        //將陣列歸0
        public void ArrayReset(double[] a)
        {
            for(int i=0;i<a.Length;i++)
            {
                a[i] = 0;
            }
        }
        public void ArrayReset(double[,] a)
        {
            for(int i=0;i<excelTime;i++)
                for(int j=0;j<10;j++)
                {
                    a[i, j] = 0;
                }
        }

        private void btnexit_Click(object sender, EventArgs e)
        {
            ushort i;
            for (i = 0; i < existcard; i++) rc = CPCI_DMC.CS_DMC_01_reset_card(gCardNoList[i]);          
            CPCI_DMC.CS_DMC_01_close();
            Application.Exit();
        }
        private void btnPmove_Click(object sender, EventArgs e)
        {
            double m_Tacc = Double.Parse(txtTacc.Text),m_Tdec = Double.Parse(txtTdec.Text);
            int m_Rpm = Int16.Parse(txtRpm1.Text);
            gnodeid = ushort.Parse(cmbNodeID.Text);
            /* Set up Velocity mode parameter */
            rc =CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node2, 0, m_Tacc, m_Tdec);
            //* Start Velocity move: rpm > 0 move forward , rpm < 0 move negative */
            rc =CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node2, 0, m_Rpm);
        }

        private void btnConnectPLC_Click(object sender, EventArgs e)
        {
            string IP = txtIPToPLC.Text;                //設定變數IP，其字串
            int Port = int.Parse(txtPortToPLC.Text);    //設定變數Port，為整數
            try
            {
                //IPAddress是IP，如" 127.0.0.1"  ;IPEndPoint是ip和端口對的組合，如"127.0.0.1: 1000 "  
                IPEndPoint EP = new IPEndPoint(IPAddress.Parse(IP), Port);
                //new Socket( 通訊協定家族IP4 , 通訊端類型 , 通訊協定TCP)
                T = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                T.Connect(EP); //建立連線
                lblConnectStatus.Text = "已連線至PLC";
                btnWork.Enabled = true;
            }
            catch (Exception)
            {
                lblConnectStatus.Text = "無法連線至PLC,請檢查線路或IP";
                return;
            }
        }

        private void CmbCardNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            gCardNo = Convert.ToUInt16(CmbCardNo.SelectedItem);
        }

        private void btnAutoHome_Click(object sender, EventArgs e)
        {
            ThHome = new Thread(AutoHome);
            ThHome.Start();
        }

        private void AutoHome()
        {
            while(true)
            {
                homeSend("000000000006" + "010204000001");
                homeListen();
                showMotorState();
                if (label31.Text== "01-02-01-01")
                {
                    rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node2, 0, 0.1, 0.1);
                    rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node2, 0, 1300);
                }
                else
                {
                    rc = CPCI_DMC.CS_DMC_01_sd_stop(gCardNo, node2, 0, 0.01);
                    CPCI_DMC.CS_DMC_01_set_position(gCardNo, node2, 0, 0);
                    CPCI_DMC.CS_DMC_01_set_command(gCardNo, node2, 0, 0);
                    break;
                }
            }
        }
        private void homeSend(string Str)
        {
            byte[] A = new byte[1]; //初始需告陣列(因不知道資料大小，下面會做陣列調整)
            for (int i = 0; i < Str.Length / 2; i++)
            {
                Array.Resize(ref A, Str.Length / 2);  //Array.Resize(ref 陣列名稱, 新的陣列大小)  
                string str2 = Str.Substring(i * 2, 2);
                A[i] = Convert.ToByte(str2, 16); //字串依照"frombase"轉換數字(Byte)
            }
            T.Send(A, 0, Str.Length / 2, SocketFlags.None);
        }
        //================接收訊息========================================
        private void homeListen()
        {
            EndPoint ServerEP = (EndPoint)T.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;

            try
            {
                inLen = T.ReceiveFrom(B, ref ServerEP);
            }
            catch (Exception)
            {
                T.Close();
                MessageBox.Show("伺服器中斷連線!");
                //btn_Plc_Connect.Enabled = true;
            }
            label31.Text = BitConverter.ToString(B, 6, inLen - 6);
        }
        private void btnSaveExcel_Click(object sender, EventArgs e)
        {
            String FileStr = "D:\\";
            FileStr += DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
            Excel.Application Excel_app1 = new Excel.Application();
            Excel.Workbook Excel_WB1 = Excel_app1.Workbooks.Add();
            Excel.Worksheet Excel_WS1 = new Excel.Worksheet();
            Excel_app1.Cells[1, 1] = "馬達牙攻轉速";
            Excel_app1.Cells[1, 2] = "馬達凸輪轉速";
            Excel_app1.Cells[1, 3] = "馬達牙攻扭矩";
            Excel_app1.Cells[1, 4] = "馬達凸輪扭矩";
            Excel_app1.Cells[1, 5] = "KTR牙攻轉速";
            Excel_app1.Cells[1, 6] = "KTR凸輪轉速";
            Excel_app1.Cells[1, 7] = "KTR牙攻扭矩";
            Excel_app1.Cells[1, 8] = "KTR凸輪扭矩";
            for(int i=0;i<motorRpm1.Count;i++)
            {
                Excel_app1.Cells[i + 2, 1] = motorRpm1[i];
                Excel_app1.Cells[i + 2, 2] = motorRpm2[i];
                Excel_app1.Cells[i + 2, 3] = motorTorque1[i];
                Excel_app1.Cells[i + 2, 4] = motorTorque2[i];
            }
            for (int i = 0; i < ktrRpm1.Count; i++)
            {
                Excel_app1.Cells[i + 2, 5] = ktrRpm1[i];
                Excel_app1.Cells[i + 2, 6] = ktrRpm2[i];
                Excel_app1.Cells[i + 2, 7] = ktrTorque1[i];
                Excel_app1.Cells[i + 2, 8] = ktrTorque2[i];
            }
            Excel_WB1.SaveAs(FileStr);
            Excel_WB1.Close();
            Excel_WB1 = null;
            Excel_app1.Quit();
            Excel_app1 = null;
            ExcelPath.Text += FileStr;
        }

        private void chart3_Click(object sender, EventArgs e)
        {

        }

        private void btnWork_Click(object sender, EventArgs e)
        {
            chart1.Series[0].Points.Clear();
            chart2.Series[0].Points.Clear();
            chart3.Series[0].Points.Clear();
            chart4.Series[0].Points.Clear();
            chart5.Series[0].Points.Clear();
            chart6.Series[0].Points.Clear();
            chart7.Series[0].Points.Clear();
            chart8.Series[0].Points.Clear();
            ktrRpm1.Clear();
            ktrRpm2.Clear();
            ktrTorque1.Clear();
            ktrTorque2.Clear();
            motorRpm1.Clear();
            motorRpm2.Clear();
            motorTorque1.Clear();
            motorTorque2.Clear();

            ThWorking = new Thread(working);
            ThWorking.Start();

            ThWorking_PLC = new Thread(working_PLC);
            ThWorking_PLC.Start();
            
        }

        private void working_PLC()
        {
            while(true)
            {
                if (!ThWorking.IsAlive)
                {
                    lblcount.Text = "exit";
                    break;
                }
                Send("000000000006" + "010313000004");
                Listen();
                if (KtrBoolClear)
                {
                    if (newturn)
                    {
                        //ktrRpm1.Clear();
                        //ktrRpm2.Clear();
                        //ktrTorque1.Clear();
                        //ktrTorque2.Clear();
                        //chart5.Series[0].Points.Clear();
                        //chart6.Series[0].Points.Clear();
                        //chart7.Series[0].Points.Clear();
                        //chart8.Series[0].Points.Clear();
                        
                        newturn = false;
                    }
                    KtrBoolClear = !KtrBoolClear;
                }
                else
                {
                    //chart5.Series[0].Points.AddXY(ktrRpm1.Count, ktrRpm1[ktrRpm1.Count - 1]);
                    //chart6.Series[0].Points.AddXY(ktrRpm2.Count, ktrRpm2[ktrRpm2.Count - 1]);
                    //chart7.Series[0].Points.AddXY(ktrTorque1.Count, ktrTorque1[ktrTorque1.Count - 1]);
                    //chart8.Series[0].Points.AddXY(ktrTorque2.Count, ktrTorque2[ktrTorque2.Count - 1]);
                }
            }
        }

        public void working()
        {
            int Amount = Convert.ToInt32(txtAmount.Text);

            for (int i = 0; i < Amount; i++)
            {
                //ktrRpm1.Clear();
                //ktrRpm2.Clear();
                //ktrTorque1.Clear();
                //ktrTorque2.Clear();
                //motorRpm1.Clear();
                //motorRpm2.Clear();
                //motorTorque1.Clear();
                //motorTorque2.Clear();
                chart1.Series[0].Points.Clear();
                chart2.Series[0].Points.Clear();
                chart3.Series[0].Points.Clear();
                chart4.Series[0].Points.Clear();
                chart5.Series[0].Points.Clear();
                chart6.Series[0].Points.Clear();
                chart7.Series[0].Points.Clear();
                chart8.Series[0].Points.Clear();
                bool boolGo = true;
                newturn = true;
                rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node2, 0, ref cmd1);
                rc = CPCI_DMC.CS_DMC_01_start_tr_move(gCardNo, node2, 0, (int)(OneCirclePluse * TransmissionRate), 0, Convert.ToInt32(2133.3333 * -Int16.Parse(txtRpm1.Text)), 0.1, 0.1);
                while (true)
                {
                    showMotorState();
                    saveMotorData();
                    showChart();
                    limitSend("000000000006" + "010204000001");
                    limitX0Listen();
                    limitSend("000000000006" + "010204010001");
                    limitX1Listen();
                    label30.Text = X0Message;
                    if (X1Message == "01-02-01-01" & boolGo)
                    {
                        rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, Int32.Parse(txtRpm2.Text));
                    }
                    else if (X1Message == "01-02-01-00")
                    {
                        rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0);
                        boolGo = false;
                    }
                    else if (X1Message == "01-02-01-01" & !boolGo & X0Message != "01-02-01-00") 
                    {
                        rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, -Int32.Parse(txtRpm2.Text));
                    }
                    else if (X0Message == "01-02-01-00" & !boolGo)
                    {
                        KtrBoolClear = true;
                        label29.Text = "in";
                        showMotorState();
                        rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0);
                        Thread.Sleep(500);
                        CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
                        CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
                        //CPCI_DMC.CS_DMC_01_set_position(gCardNo, node2, 0, 0);
                        //CPCI_DMC.CS_DMC_01_set_command(gCardNo, node2, 0, 0);
                        break;
                    }
                }
            }
        }
        private void limitSend(string Str)
        {
            byte[] A = new byte[1]; //初始需告陣列(因不知道資料大小，下面會做陣列調整)
            for (int i = 0; i < Str.Length / 2; i++)
            {
                Array.Resize(ref A, Str.Length / 2);  //Array.Resize(ref 陣列名稱, 新的陣列大小)  
                string str2 = Str.Substring(i * 2, 2);
                A[i] = Convert.ToByte(str2, 16); //字串依照"frombase"轉換數字(Byte)
            }
            T.Send(A, 0, Str.Length / 2, SocketFlags.None);
        }
        private void limitX0Listen()
        {
            EndPoint ServerEP = (EndPoint)T.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;
            while (true)
            {
                try
                {
                    inLen = T.ReceiveFrom(B, ref ServerEP);
                    break;
                }
                catch (Exception)//當try發生問題時重新向PLC發送請求(18.10.25)
                {
                    //T.Close();
                    //MessageBox.Show("伺服器中斷連線!");
                    limitSend("000000000006" + "010204000001");
                    //btn_Plc_Connect.Enabled = true;
                    //break;
                }
            }
            X0Message = BitConverter.ToString(B, 6, inLen - 6);
        }
        private void limitX1Listen()
        {
            EndPoint ServerEP = (EndPoint)T.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;
            while (true)
            {
                try
                {
                    inLen = T.ReceiveFrom(B, ref ServerEP);
                    break;
                }
                catch (Exception)//當try發生問題時重新向PLC發送請求(18.10.25)
                {
                    //T.Close();
                    //MessageBox.Show("伺服器中斷連線!");
                    limitSend("000000000006" + "010204010001");
                    //break;
                    //btn_Plc_Connect.Enabled = true;
                }
            }
            X1Message = BitConverter.ToString(B, 6, inLen - 6);
        }
        private void showMotorState()
        {
            rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node1, 0, ref cmd1);
            rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node2, 0, ref cmd2);
            //Command
            if (rc == 0)
            {
                txtcommand1.Text = cmd1.ToString();
                txtcommand2.Text = cmd2.ToString();
            }
            //Feedback
            rc = CPCI_DMC.CS_DMC_01_get_position(gCardNo, node1, 0, ref pos1);
            rc = CPCI_DMC.CS_DMC_01_get_position(gCardNo, node2, 0, ref pos2);
            if (rc == 0)
            {
                txtfeedback1.Text = pos1.ToString();
                txtfeedback2.Text = pos2.ToString();
            }
            //Speed
            rc = CPCI_DMC.CS_DMC_01_get_rpm(gCardNo, node1, 0, ref spd1);
            rc = CPCI_DMC.CS_DMC_01_get_rpm(gCardNo, node2, 0, ref spd2);
            if (rc == 0)
            {
                txtspeed1.Text = spd1.ToString();
                txtspeed2.Text = spd2.ToString();
            }
            //Torque
            rc = CPCI_DMC.CS_DMC_01_get_torque(gCardNo, node1, 0, ref toe1);
            rc = CPCI_DMC.CS_DMC_01_get_torque(gCardNo, node2, 0, ref toe2);
            if (rc == 0)
            {
                //扭矩是千分比
                txtTorque1.Text = ((double)toe1 / 1000 * 7.16).ToString();
                txtTorque2.Text = ((double)toe2 / 1000 * 7.16).ToString();
            }
            //err
            rc = CPCI_DMC.CS_DMC_01_get_alm_code(gCardNo, node1, 0, ref err1);
            rc = CPCI_DMC.CS_DMC_01_get_alm_code(gCardNo, node1, 0, ref err2);
            if (rc == 0)
            {
                txtERR1.Text = err1.ToString();
                txtERR2.Text = err2.ToString();
            }
        }
        private void saveMotorData()
        {
            motorTorque1.Add((double)toe1 / 1000 * 7.16);
            motorTorque2.Add((double)toe2 / 1000 * 7.16);
            motorRpm1.Add(spd1 / 10);
            motorRpm2.Add(spd2 / 10);
        }
        private void showChart()
        {
            chart1.Series[0].Points.AddXY(motorRpm1.Count, motorRpm1[motorRpm1.Count - 1]);
            chart2.Series[0].Points.AddXY(motorRpm2.Count, motorRpm2[motorRpm2.Count - 1]);
            chart3.Series[0].Points.AddXY(motorTorque1.Count, motorTorque1[motorTorque1.Count - 1]);
            chart4.Series[0].Points.AddXY(motorTorque2.Count, motorTorque2[motorTorque2.Count - 1]);
        }

        private void btnNmove_Click_1(object sender, EventArgs e)
        {
            double m_Tacc = Double.Parse(txtTacc.Text), m_Tdec = Double.Parse(txtTdec.Text);
            int m_Rpm = Int16.Parse(txtRpm1.Text);
            gnodeid = ushort.Parse(cmbNodeID.Text);
            /* Set up Velocity mode parameter */
            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node2, 0, m_Tacc, m_Tdec);
            //* Start Velocity move: rpm > 0 move forward , rpm < 0 move negative */
            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node2, 0, -1 * m_Rpm);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
        }

        private void btnFindSlave_Click(object sender, EventArgs e)
        {
            ushort i, lMask = 0x1,p=0;
            uint DeviceType = 0, IdentityObject = 0;
            btnreset1.Enabled = false;
            btnralm.Enabled = false;
            btnstop.Enabled = false;
            btnreset1.Enabled = false;
            btnNmove.Enabled = false;
            btnPmove.Enabled = false;
            chksvon.Enabled = false;
            gNodeNum = 0;               
            txtSlaveNum.Text = "0";
            cmbNodeID.Items.Clear();

            for (i = 0; i < 1; i++) NodeID[i] = 0;

            if (SlaveTable[0] == 0)
                MessageBox.Show("CardNo: " + gCardNo.ToString() + " No slave found!");      
            else
            {
                for (i = 0; i < 32; i++)
                {
                    if ((SlaveTable[0] & lMask) != 0)
                    {
                        NodeID[gNodeNum] = (ushort)(i + 1);
                        gNodeNum++;
                        rc = CPCI_DMC.CS_DMC_01_get_devicetype((short)gCardNo, (ushort)(i + 1), (ushort)0, ref DeviceType, ref IdentityObject);
                        if (rc != 0)
                        {
                            MessageBox.Show("get_devicetype failed - code=" + rc);
                        }
                        else
                        {
                            switch (DeviceType)
                            {
                                case 0x4020192:				//Servo A2 series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x6020192:				//Servo M series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x8020192:				//Servo A2R series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                                case 0x9020192:				//Servo S series
                                    cmbNodeID.Items.Add(i + 1);
                                    p++;
                                    break;
                            }
                        }
                    }
                    lMask <<= 1;
                }
                if (p == 0)
                {
                    MessageBox.Show("No A2 Servo Device Found!");
                }
                else
                {
                    txtSlaveNum.Text = gNodeNum.ToString();
                    cmbNodeID.SelectedIndex = 0;
                    btnreset1.Enabled = true;
                }
            }
        }
        //public void plcState()
        //{
        //    int m_Rpm = Int32.Parse(txtRpm2.Text);
        //    while (true)
        //    {
        //        CheckForIllegalCrossThreadCalls = false;
        //        count++;
        //        lblcount.Text = count.ToString();
        //        int cmd1 = 0, pos1 = 0, cmd2 = 0, pos2 = 0;
        //        short spd1 = 0, spd2 = 0, toe1 = 0, toe2 = 0;
        //        uint err1 = 0, err2 = 0;
        //        *******
        //        rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node1, 0, ref cmd1);
        //        rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node2, 0, ref cmd2);
        //        rc = CPCI_DMC.CS_DMC_01_get_command(gCardNo, node2, 0, ref com2);
        //        Command
        //        if (rc == 0)
        //        {
        //            txtcommand1.Text = cmd1.ToString();
        //            txtcommand2.Text = cmd2.ToString();
        //        }
        //        Feedback
        //        rc = CPCI_DMC.CS_DMC_01_get_position(gCardNo, node1, 0, ref pos1);
        //        rc = CPCI_DMC.CS_DMC_01_get_position(gCardNo, node2, 0, ref pos2);
        //        if (rc == 0)
        //        {
        //            txtfeedback1.Text = pos1.ToString();
        //            txtfeedback2.Text = pos2.ToString();
        //        }
        //        Speed
        //        rc = CPCI_DMC.CS_DMC_01_get_rpm(gCardNo, node1, 0, ref spd1);
        //        rc = CPCI_DMC.CS_DMC_01_get_rpm(gCardNo, node2, 0, ref spd2);
        //        if (rc == 0)
        //        {
        //            txtspeed1.Text = spd1.ToString();
        //            txtspeed2.Text = spd2.ToString();
        //        }
        //        Torque
        //        rc = CPCI_DMC.CS_DMC_01_get_torque(gCardNo, node1, 0, ref toe1);
        //        rc = CPCI_DMC.CS_DMC_01_get_torque(gCardNo, node2, 0, ref toe2);
        //        if (rc == 0)
        //        {
        //            扭矩是千分比
        //            txtTorque1.Text = toe1.ToString();
        //            txtTorque2.Text = toe2.ToString();
        //        }
        //        err
        //        rc = CPCI_DMC.CS_DMC_01_get_alm_code(gCardNo, node1, 0, ref err1);
        //        rc = CPCI_DMC.CS_DMC_01_get_alm_code(gCardNo, node1, 0, ref err2);
        //        if (rc == 0)
        //        {
        //            txtERR1.Text = err1.ToString();
        //            txtERR2.Text = err2.ToString();
        //        }
        //        rpm_motor1[excelTime] = spd1;
        //        rpm_motor2[excelTime] = spd2;
        //        torque_motor1[excelTime] = ((double)toe1 / 1000) * 7.16;
        //        torque_motor2[excelTime] = ((double)toe2 / 1000) * 7.16;
        //        Send("000000000006" + "010310280028");  //前段為TCP固定，後段為讀取暫存器D
        //        Send("000000000006" + "010310000028");
        //        Listen();                                 //解析收到的MODBUS
        //        *************
        //        if ((com2 <= 128000 * TransmissionRate / 2 + 16 * 3556 * TransmissionRate / 2 - delayMotorDeg * 3556 * TransmissionRate / 2) && !b)    //1280000*50為180deg +- deg*3556*50 +-delay deg
        //        {
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, m_Rpm);
        //        }
        //        else if (com2 >= (1280000 * TransmissionRate / 2 + 16 * 3556 * TransmissionRate / 2 - delayMotorDeg * 3556 * TransmissionRate / 2) && com2 <= 1280000 * TransmissionRate / 2 + 16 * 3556 * TransmissionRate / 2 + delayMotorDeg * 3556 * TransmissionRate / 2)
        //        {
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0);
        //        }
        //        else
        //        {
        //            b = true;
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
        //            rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0 - m_Rpm);
        //            if (spd2 < 100)
        //            {
        //                rc = CPCI_DMC.CS_DMC_01_sd_stop(gCardNo, node1, 0, 0.1);
        //                rc = CPCI_DMC.CS_DMC_01_sd_stop(gCardNo, node2, 0, 0.1);
        //                btnralm.Enabled = false;
        //                btnstop.Enabled = false;
        //                btnNmove.Enabled = false;
        //                btnPmove.Enabled = false;
        //                chksvon.Enabled = false;
        //                Thread.Sleep(300);
        //                chksvon.Checked = false;
        //                th.Abort();
        //            }
        //        }
        //    }
        //}
        private void Send(string Str)
        {
            byte[] A = new byte[1]; //初始需告陣列(因不知道資料大小，下面會做陣列調整)
            for (int i = 0; i < Str.Length / 2; i++)
            {
                Array.Resize(ref A, Str.Length / 2);  //Array.Resize(ref 陣列名稱, 新的陣列大小)  
                string str2 = Str.Substring(i * 2, 2);
                A[i] = Convert.ToByte(str2, 16); //字串依照"frombase"轉換數字(Byte)
            }
            T.Send(A, 0, Str.Length / 2, SocketFlags.None);
        }
        private void Listen()
        {

            EndPoint ServerEP = (EndPoint)T.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;
            while (true)
            {
                try
                {
                    inLen = T.ReceiveFrom(B, ref ServerEP);
                    break;
                }
                catch (Exception) //當try發生問題時重新向PLC發送請求(18.10.25)
                {
                    //T.Close();
                    Send("000000000006" + "010313000004");
                    //MessageBox.Show("伺服器中斷連線!");
                    //btnConnectPLC.Enabled = true;
                    //break;
                }
            }
            txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
            string[] ary = txtReceive.Text.Split('-');
            //將讀取到的16進制碼換成10進制碼，且切割後的陣列兩個為1組
            //double[] rpm1 = new double[5];
            //double[] rpm2 = new double[5];
            //double[] torque1 = new double[5];
            //double[] torque2 = new double[5];
            double rpm1, rpm2, torque1, torque2;
            try //嘗試轉換電壓資料，發生Exception時ary為null(18.10.25)
            {
                rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
                rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
                torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
                torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));
            }
            //因此重新發送請求給PLC(18.10.25)
            catch (Exception)
            {
                Send("000000000006" + "010313000004");
                inLen = T.ReceiveFrom(B, ref ServerEP);
                txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
                ary = txtReceive.Text.Split('-');
                rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
                rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
                torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
                torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));
            }
            rpm1 = (rpm1 * 10 / 8000) * rpmRate1;
            rpm2= (rpm2 * 10 / 8000) * rpmRate2;
            torque1 = (torque1 * 10 / 8000) * torqueRate/5;
            torque2 = (torque2 * 10 / 8000) * torqueRate;
            ktrRpm1.Add(rpm1);
            ktrRpm2.Add(rpm2);
            ktrTorque1.Add(torque1);
            ktrTorque2.Add(torque2);

        }
        public double changeVoltage0x16(double v)
        {
            if (v > 32767)
                return ((65535 - v + 1) * (-1));
            else
                return v;
        }
    }
}
