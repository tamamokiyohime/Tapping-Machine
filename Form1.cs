using System;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net;               //載入網路
using System.Net.Sockets;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Data;

using PCI_DMC;
using PCI_DMC_ERR;
using System.Collections.Generic;
using System.IO;

namespace DMC_NET
{
    //2018/09/24 update
     
    public partial class Form1 : Form
    {
        Thread ThWorking, ThWorking_PLC, ThHome, ThWorking_PLC_2;
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
        Socket T, T2;
        int delayMotorDeg = 0;
        int rpmRate1 = 200; //ktr比例
        int rpmRate2 = 2;
        int torqueRate1 = 1;
        int torqueRate2 = 5;
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

        List<double> ktrTorque1_off = new List<double>();
        List<double> ktrTorque2_off = new List<double>();
        List<double> ktrRpm1_off = new List<double>();
        List<double> ktrRpm2_off = new List<double>();

        List<char> source = new List<char>();
        double[,] rpm_1 = new double[90000, 10];
        double[,] rpm_2 = new double[90000, 10];
        double[] rpm_motor1 = new double[90000];
        double[] rpm_motor2 = new double[90000];
        double[,] torque_1 = new double[90000, 10];
        double[,] torque_2 = new double[90000, 10];
        double[] torque_motor1 = new double[90000];
        double[] torque_motor2 = new double[90000];

        delegate void UpdateUIDelegate();
        double[] data;  //緩衝器裡面的陣列
        DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();
            //data = new double[bufferedAiCtrl1.BufferCapacity];//設定DAQ 緩衝器
            //bufferedAiCtrl1.Streaming = true;
            //bufferedAiCtrl1.Prepare();
            dt.Columns.Add("RPM1");
            dt.Columns.Add("Torq1");
            dt.Columns.Add("RPM2");
            dt.Columns.Add("Torq2");
            dataGridView1.DataSource = dt;
            timer1.Interval = 100;
            timer1.Enabled = false;
            bufferedAiCtrl1.Prepare();

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
                lblConnectStatus.Text = "已連線至PLC1";
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
            string FileStr = "D:\\實驗數據\\";
            string FileStr2 = "";
            FileStr += DateTime.Now.ToString("yyyyMMdd");
            FileStr += " Experiment\\";
            if (!Directory.Exists(FileStr))
            {
                Directory.CreateDirectory(FileStr);
            }
            FileStr += DateTime.Now.ToString("yyMMdd-HHmm");
            FileStr2 = FileStr;
            FileStr2 += "_DAQ-C" + txtRpm1.Text + "-T" + txtRpm2.Text + notice.Text + ".csv";
            FileStr += "-C" + txtRpm1.Text + "-T" + txtRpm2.Text +notice.Text+ ".csv";
            txtReceive.Text = FileStr;

            StreamWriter file = new StreamWriter(FileStr, false, Encoding.Default);
            file.Write("Tapper RPM(M),Cam RPM(M),Tapper Torq(M),Cam Torq(M),Tapper RPM(KTR),Cam RPM(KTR),Tapper Torq(KTR),Cam Torq(KTR),");
            file.WriteLine(motorRpm1.Count.ToString() + "," + motorRpm2.Count.ToString() + "," + motorTorque1.Count.ToString() + "," + motorTorque2.Count.ToString() + "," + ktrRpm1.Count.ToString() + "," + ktrRpm2.Count.ToString() + "," + ktrTorque1.Count.ToString() + "," + ktrTorque2.Count.ToString());
            for (int i = 0; i < ktrRpm1.Count; i++)
            {
                if (i>motorRpm1.Count-1)
                {
                    file.WriteLine(",,,," + ktrRpm1[i] + "," + ktrRpm2[i] + "," + ktrTorque1[i] + "," + ktrTorque2[i] );
                }
                else
                {
                    file.WriteLine(motorRpm1[i] + "," + motorRpm2[i] + "," + motorTorque1[i] + "," + motorTorque2[i] + "," + ktrRpm1[i] + "," + ktrRpm2[i] + "," + ktrTorque1[i] + "," + ktrTorque2[i] );
                }
            }
            file.Close();

            StreamWriter file2 = new StreamWriter(FileStr2, false, Encoding.Default);
            string DAQ_out = "";
            foreach (DataColumn column in dt.Columns)
            {
                DAQ_out += column.ColumnName + ",";
            }
            DAQ_out += "\n";
            file2.Write(DAQ_out);
            DAQ_out = "";
            foreach(DataRow row in dt.Rows)
            {
                foreach(DataColumn col in dt.Columns)
                {
                    DAQ_out += row[col].ToString().Trim() + ",";
                }
                DAQ_out += "\n";
                file2.Write(DAQ_out);
                DAQ_out = "";
            }
            file2.Dispose();
            file2.Close();

        }

        private void btnConnectPLC2_Click(object sender, EventArgs e)
        {
            string IP = txtIPToPLC.Text;                //設定變數IP，其字串
            int Port = int.Parse(txtPortToPLC.Text);    //設定變數Port，為整數
            try
            {
                //IPAddress是IP，如" 127.0.0.1"  ;IPEndPoint是ip和端口對的組合，如"127.0.0.1: 1000 "  
                IPEndPoint EP2 = new IPEndPoint(IPAddress.Parse(IP), Port);
                //new Socket( 通訊協定家族IP4 , 通訊端類型 , 通訊協定TCP)
                T2 = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                T2.Connect(EP2); //建立連線
                lblConnectStatus2.Text = "已連線至PLC2";
                btnWork.Enabled = true;
            }
            catch (Exception)
            {
                lblConnectStatus.Text = "無法連線至PLC,請檢查線路或IP";
                return;
            }
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
            source.Clear();

            //chart5.DataSource = ktrRpm1;
            //chart6.DataSource = ktrRpm2;
            //chart7.DataSource = ktrTorque1;
            //chart8.DataSource = ktrTorque2;
            //chart5.Series[0].YValueMembers = "ktrRpm1";
            //chart6.Series[0].YValueMembers = "ktrRpm2";
            //chart7.Series[0].YValueMembers = "ktrTorque1";
            //chart8.Series[0].YValueMembers = "ktrTorque2";

            dt.Clear();

            data = new double[bufferedAiCtrl1.BufferCapacity];//設定DAQ 緩衝器
                                                              //bufferedAiCtrl1.Streaming = true;

            //bufferedAiCtrl1.Prepare();
            bufferedAiCtrl1.Start();
            //timer1.Enabled = true;

            rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, double.Parse(txtTacc.Text), double.Parse(txtTdec.Text));
            ThWorking = new Thread(working);
            ThWorking.Start();

            ThWorking_PLC = new Thread(working_PLC);
            ThWorking_PLC.Start();
            ThWorking_PLC_2 = new Thread(working_PLC2);
            ThWorking_PLC_2.Start();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            chart5.DataSource = dt;
            chart6.DataSource = dt;
            chart7.DataSource = dt;
            chart8.DataSource = dt;
            chart5.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart6.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart7.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart8.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart5.Series[0].YValueMembers = "RPM1";
            chart6.Series[0].YValueMembers = "RPM2";
            chart7.Series[0].YValueMembers = "Torq1";
            chart8.Series[0].YValueMembers = "Torq2";
            chart5.DataBind();
            chart6.DataBind();
            chart7.DataBind();
            chart8.DataBind();

        }

        private void chart5_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            ktrRpm1_off.Clear();
            ktrRpm2_off.Clear();
            ktrTorque1_off.Clear();
            ktrTorque2_off.Clear();
            int times = 0;
            while(times < 100)
            {
                Send("000000000006" + "010313000004");
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
                        Send("000000000006" + "010313000004");
                    }
                }
                string[] ary = BitConverter.ToString(B, 6, inLen - 6).Split('-');
                //double rpm1, rpm2, torque1, torque2;
                try //嘗試轉換電壓資料，發生Exception時ary為null(18.10.25)
                {
                    //rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
                    //rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
                    //torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
                    //torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));

                    ktrRpm1_off.Add(changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber)));
                    ktrRpm2_off.Add(changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber)));
                    ktrTorque1_off.Add(changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber)));
                    ktrTorque2_off.Add(changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber)));
                }
                //因此重新發送請求給PLC(18.10.25)
                catch (Exception)
                {
                    Send("000000000006" + "010313000004");
                    inLen = T.ReceiveFrom(B, ref ServerEP);
                    ary = BitConverter.ToString(B, 6, inLen - 6).Split('-');
                    //rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
                    //rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
                    //torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
                    //torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));
                    ktrRpm1_off.Add(changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber)));
                    ktrRpm2_off.Add(changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber)));
                    ktrTorque1_off.Add(changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber)));
                    ktrTorque2_off.Add(changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber)));

                }
                times++;

                rpm1_off.Text = (ktrRpm1_off.Average()*10/8192).ToString();
                rpm2_off.Text = (ktrRpm2_off.Average() * 10 / 8192).ToString();
                torq1_off.Text = (ktrTorque1_off.Average() * 10 / 8192).ToString();
                torq2_off.Text = (ktrTorque2_off.Average() * 10 / 8192).ToString();

            }

            
        }

        private void rpm2_off_TextChanged(object sender, EventArgs e)
        {

        }

        private void bufferedAiCtrl1_DataReady(object sender, Automation.BDaq.BfdAiEventArgs e)
        {
            bufferedAiCtrl1.GetData(e.Count, data);
            this.Invoke((UpdateUIDelegate)delegate ()
            {
                DataRow row = dt.NewRow();
                bool ready = false;
                for (int i = 0; i < e.Count; i++)
                {
                    switch (i % 4)
                    {
                        case 0:
                            row["RPM1"] = data[i] * rpmRate1;
                            //row["RPM1"] = data[i];
                            break;
                        case 1:
                            row["Torq1"] = data[i] * torqueRate1;
                            //row["Torq1"] = data[i];
                            break;
                        case 2:
                            row["RPM2"] = data[i] * rpmRate2;
                            //row["RPM2"] = data[i];
                            break;
                        case 3:
                            row["Torq2"] = data[i] * torqueRate2;
                            //row["Torq2"] = data[i];
                            ready = true;
                            break;
                    }
                    if (ready)
                    {
                        dt.Rows.Add(row);
                        dt.ImportRow(row);
                        ready = false;
                        dataGridView1.Update();
                        row = dt.NewRow();
                    }
                }
            });
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            bufferedAiCtrl1.Start();
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
        private void working_PLC2()
        {
            while (true)
            {
                if (!ThWorking.IsAlive)
                {
                    lblcount.Text = "exit";
                    break;
                }
                Send2("000000000006" + "010313000004");
                Listen2();
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
                rc = CPCI_DMC.CS_DMC_01_start_tr_move(gCardNo, node2, 0, OneCirclePluse * TransmissionRate, 0, Convert.ToInt32(2133.3333 * -Int16.Parse(txtRpm1.Text)), 0.1, 0.1);
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
                        //rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, Int32.Parse(txtRpm2.Text));
                    }
                    else if (X1Message == "01-02-01-00")
                    {
                        //rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0);
                        boolGo = false;
                    }
                    else if (X1Message == "01-02-01-01" & !boolGo & X0Message != "01-02-01-00")
                    {
                        //rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, -Int32.Parse(txtRpm2.Text));
                    }
                    else if (X0Message == "01-02-01-00" & !boolGo)
                    {
                        KtrBoolClear = true;
                        label29.Text = "in";
                        showMotorState();
                        //rc = CPCI_DMC.CS_DMC_01_set_velocity_mode(gCardNo, node1, 0, 0.1, 0.1);
                        rc = CPCI_DMC.CS_DMC_01_set_velocity(gCardNo, node1, 0, 0);
                        Thread.Sleep(500);
                        CPCI_DMC.CS_DMC_01_set_position(gCardNo, node1, 0, 0);
                        CPCI_DMC.CS_DMC_01_set_command(gCardNo, node1, 0, 0);
                        //CPCI_DMC.CS_DMC_01_set_position(gCardNo, node2, 0, 0);
                        //CPCI_DMC.CS_DMC_01_set_command(gCardNo, node2, 0, 0);
                        break;
                    }
                    
                }
                bufferedAiCtrl1.Stop();
                //bufferedAiCtrl1.Cleanup();
                //bufferedAiCtrl1.Release();
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

        private void Send2(string Str)
        {
            byte[] A = new byte[1]; //初始需告陣列(因不知道資料大小，下面會做陣列調整)
            for (int i = 0; i < Str.Length / 2; i++)
            {
                Array.Resize(ref A, Str.Length / 2);  //Array.Resize(ref 陣列名稱, 新的陣列大小)  
                string str2 = Str.Substring(i * 2, 2);
                A[i] = Convert.ToByte(str2, 16); //字串依照"frombase"轉換數字(Byte)
            }
            T2.Send(A, 0, Str.Length / 2, SocketFlags.None);
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
            //txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
            //string[] ary = txtReceive.Text.Split('-');
            string[] ary = BitConverter.ToString(B, 6, inLen - 6).Split('-');
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
                //txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
                //ary = txtReceive.Text.Split('-');
                ary = BitConverter.ToString(B, 6, inLen - 6).Split('-');
                rpm1 = changeVoltage0x16(Int32.Parse(ary[3] + ary[4], System.Globalization.NumberStyles.HexNumber));
                rpm2 = changeVoltage0x16(Int32.Parse(ary[5] + ary[6], System.Globalization.NumberStyles.HexNumber));
                torque1 = changeVoltage0x16(Int32.Parse(ary[7] + ary[8], System.Globalization.NumberStyles.HexNumber));
                torque2 = changeVoltage0x16(Int32.Parse(ary[9] + ary[10], System.Globalization.NumberStyles.HexNumber));
                
            }
            rpm1 = (rpm1 * 10 / 8192-double.Parse(rpm1_off.Text)) * rpmRate1;
            rpm2= (rpm2 * 10 / 8192-double.Parse(rpm2_off.Text)) * rpmRate2;
            torque1 = (torque1 * 10 / 8192-double.Parse(torq1_off.Text)) * torqueRate1;
            torque2 = (torque2 * 10 / 8192 - double.Parse(torq1_off.Text)) * torqueRate2;
            
            ktrRpm1.Add(rpm1);
            ktrRpm2.Add(rpm2);
            ktrTorque1.Add(torque1);
            ktrTorque2.Add(torque2);
            source.Add('A');

        }
        private void Listen2()
        {

            EndPoint ServerEP = (EndPoint)T2.RemoteEndPoint;
            byte[] B = new byte[1023];
            int inLen = 0;
            while (true)
            {
                try
                {
                    inLen = T2.ReceiveFrom(B, ref ServerEP);
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
            //txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
            //string[] ary = txtReceive.Text.Split('-');
            string[] ary2 = BitConverter.ToString(B, 6, inLen - 6).Split('-');
            //將讀取到的16進制碼換成10進制碼，且切割後的陣列兩個為1組
            //double[] rpm1 = new double[5];
            //double[] rpm2 = new double[5];
            //double[] torque1 = new double[5];
            //double[] torque2 = new double[5];
            double rpm1, rpm2, torque1, torque2;
            try //嘗試轉換電壓資料，發生Exception時ary為null(18.10.25)
            {
                rpm1 = changeVoltage0x16(Int32.Parse(ary2[3] + ary2[4], System.Globalization.NumberStyles.HexNumber));
                rpm2 = changeVoltage0x16(Int32.Parse(ary2[5] + ary2[6], System.Globalization.NumberStyles.HexNumber));
                torque1 = changeVoltage0x16(Int32.Parse(ary2[7] + ary2[8], System.Globalization.NumberStyles.HexNumber));
                torque2 = changeVoltage0x16(Int32.Parse(ary2[9] + ary2[10], System.Globalization.NumberStyles.HexNumber));
            }
            //因此重新發送請求給PLC(18.10.25)
            catch (Exception)
            {
                Send("000000000006" + "010313000004");
                inLen = T.ReceiveFrom(B, ref ServerEP);
                //txtReceive.Text = BitConverter.ToString(B, 6, inLen - 6);
                //ary = txtReceive.Text.Split('-');
                ary2 = BitConverter.ToString(B, 6, inLen - 6).Split('-');
                rpm1 = changeVoltage0x16(Int32.Parse(ary2[3] + ary2[4], System.Globalization.NumberStyles.HexNumber));
                rpm2 = changeVoltage0x16(Int32.Parse(ary2[5] + ary2[6], System.Globalization.NumberStyles.HexNumber));
                torque1 = changeVoltage0x16(Int32.Parse(ary2[7] + ary2[8], System.Globalization.NumberStyles.HexNumber));
                torque2 = changeVoltage0x16(Int32.Parse(ary2[9] + ary2[10], System.Globalization.NumberStyles.HexNumber));
            }
            rpm1 = (rpm1 * 10 / 8192 - double.Parse(rpm1_off.Text)) * rpmRate1;
            rpm2 = (rpm2 * 10 / 8192 - double.Parse(rpm2_off.Text)) * rpmRate2;
            torque1 = (torque1 * 10 / 8192 - double.Parse(torq1_off.Text)) * torqueRate1;
            torque2 = (torque2 * 10 / 8192 - double.Parse(torq1_off.Text)) * torqueRate2;
            ktrRpm1.Add(rpm1);
            ktrRpm2.Add(rpm2);
            ktrTorque1.Add(torque1);
            ktrTorque2.Add(torque2);
            source.Add('B');

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
