using System;
using System.Text;
using System.Runtime.InteropServices;

namespace PCI_DMC_ERR
{
    public class CPCI_DMC_ERR
    {
	//General
        public const short ERR_SameCardNumber          =   901; //Card IDNo Repeat(All the Card IDNo. must be the only).
        public const short ERR_CardType                =   902; //The Card not support the API.
        public const short ERR_CantFindDLL             =   903; //API connect DLL Failed.
        public const short ERR_SerchErrorCode          =   904; //The ErrorCode must be get by use _DMC_01_get_cardtype_errorcode API.

	//PCI_DMC_A01/PCI_DMC_B01
        public const short ERR_NoError                 =   0;
        public const short ERR_CardNoError             =   3; //Card No. is unable to recognize, please check DIP Switch on the card.
        public const short ERR_bootmodeErr             =   5; //Can't start DSP procedure.
        public const short ERR_downloadcode            =   6; //DSP Program Memory R/W Error.
        public const short ERR_downloadinit            =   7; //DSP Data Memory R/W Error.
        public const short ERR_PCI_boot_first          =   8; //It must run _DMC_01_pci_initial(U16 CardNo) first before run the API.
        public const short ERR_FeedRate_updata         =   9; //The value of FeedRate overwrite is constant.
        public const short ERR_AxisNoError             =  11; //Node number overflow.
        public const short ERR_IPO_First               =  12; //It must transfer to IPO mode.
        public const short ERR_Target_reach            =  13; //Run in Mode1, it need target arrived.
        public const short ERR_Servo_on_first          =  14; //Servo On first.
        public const short ERR_MPG_Mode                =  15; //It can't change position to zero becase in MPG mode.
        public const short ERR_PDO_TG		       =  16; //It can't receive message from module,while gave orders to module in PDO mode.
	public const short ERR_ConfigFileOpenError     =  17; //It is wrong,while open the debug information file.
        public const short ERR_Ctrl_value	       =  18; //It is wrong to use control command.
	public const short ERR_Security_Fifo           =  19; //Security Fpga write Error.
        public const short ERR_Security_Fifo_busy      =  20; //Security Fpga busy.
        public const short ERR_SpeedLimitError         =  21; //Speed over the maximal speed.
        public const short ERR_Security_Page	       =  22; //Security page must be smaller than 16.
        public const short ERR_Slave_Security_op       =  23; //Security slave_operate can't be execute.
        public const short ERR_channel_no	       =  24; //Channel no. is wrong.
        public const short ERR_start_ring_first        =  25; //It must run _DMC_01_start_ring(U16 CardNo, U8 RingNo) first before run the API.
        public const short ERR_NodeIDError             =  26; //It can't find the NodeID.
        public const short ERR_MailBoxErr              =  27; //DSP busy, Command can't be execute.
        public const short ERR_SdoData                 =  28; //SDO Data Send Request, But Without ACK.
        public const short ERR_IOCTL                   =  29; //Operating system can't process this IRP.
        public const short ERR_SdoSvonFirst            =  30; //Servo on first.
        public const short ERR_SlotIDError             =  31; //Can't find this slot number.
        public const short ERR_PDO_First               =  32; //Please transfer to PDO mode.
        public const short ERR_Protocal_build          =  33; //Communication Protocal must be established.
        public const short ERR_Maching_TimeOut         =  34; //Module match time out.
        public const short ERR_Maching_NG              =  35; //Module match failed.
        public const short ERR_Group_Num               =  36; //Maximal modules of the group is six.
        public const short ERR_Master_Alarm            =  37; //Error(Communication Error).
        public const short ERR_Alarm_reset = 38;

        public const short ERR_Master_Security_Wr      =  40; //Security Master Write useless.
        public const short ERR_Master_Security_Rd      =  41; //Security Master Read useless.
        public const short ERR_Master_Security_Pw      =  42; //Please input correct password first.
        public const short ERR_NonSupport_CardVer      =  50; //Card version is not correct, please contact the factor.
        public const short ERR_Compare_Source	       =  51; //Compare Source select error.
        public const short ERR_Compare_Direction       =  52; //Compare direction select error it must be 0 or 1.

        public const short ERR_GetDLLPath	       =  60; //Can't find DLL path.
        public const short ERR_GetDLLVersion	       =  61; //Can't find DLL version data.

        public const short ERR_04PISTOP_Timeout	       =  70; //04PI Stop Fifo time out 
        public const short ERR_ServoSTOP_Timeout       =  71; //Servo Stop Fifo time out
        public const short ERR_04PISTOP_status	       =  72; //04PI Stop MC_done not to Error
        public const short ERR_04PIHoming_err	       =  73; //04PI Home status error
        public const short ERR_04PISdo_trans	       =  74; //04PI sdo send but get data error

        public const short ERR_RangeError              = 112; //NodeID Error.
        public const short ERR_MotionBusy              = 114; //Motion is busy.
        public const short ERR_SpeedError              = 116; //Maximal speed can't be zero.
        public const short ERR_AccTimeError            = 117; //Acceleration time or Deceleration time over 1000 second.
        public const short ERR_PitchZero               = 124; //Helix mode parameter pitch = 0.
        public const short ERR_BufferFull              = 127; //Buffer full.
        public const short ERR_PathError               = 128; //Move command is wrong.
        public const short ERR_NoSupportMode           = 130; //Not support velocity change.
        public const short ERR_FeedHold_support        = 132; //In feedhold stop mode, can't accept new command.
        public const short ERR_SDStop_On               = 133; //Execute sd_stop command, can't accept new command.
        public const short ERR_VelChange_supper        = 134; //It can't execute velocity change under Feedhold and Sync_move and sd_stop mode.
        public const short ERR_Command_set             = 135; //Can't execute Feedhold repeat.
        public const short ERR_sdo_message_choke       = 136; //SDO command returned error, please check communction.
        public const short ERR_VelChange_buff_feedhold = 137; //Feedhold must be enabled first and can't execute velocity change.
        public const short ERR_VelChange_sync_move     = 138; //Card wait for sync_move command, so can't execute velocity change.
        public const short ERR_VelChange_SD_On         = 139; //Sd_stop in executing, so can't execute velocity change.
        public const short ERR_P_Change_Mode           = 140; //P_Change only execute at the max speed of the Position_TO_Position mode.
        public const short ERR_BufferLength            = 141; //while ececute Path_p_change and _Path_velocity_change_onfly and _Path_Start_Move_2seg, buffer length must be zero.
        public const short ERR_2segMove_Dist           = 142; //Execute the API, parameter distance must be the same direction.
        public const short ERR_CenterMatch             = 143; //The center of a circel must be match after calculated.
        public const short ERR_EndMatch                = 144; //The center of a circel must be match after calculated.
        public const short ERR_AngleCalcu              = 145; //Angle is wrong.
        public const short ERR_RedCalcu                = 146; //Radius is wrong.
        public const short ERR_GearSetting             = 147; //In gear setting element = 0 or denominator = 0.
        public const short ERR_CamTable                = 148; //Table Setting First Arrary Point Error.
        public const short ERR_AxesNum                 = 149; //Axis numbers must be greater than 2 .
        public const short ERR_SpiralPos               = 150; //The final position will be the circel of spiral.   
        public const short ERR_SpeedMode_Slave         = 151; //The slave can'nt execute motion command on speedcontinue mode.
        public const short ERR_SpeedMode_SlaveSet      = 152; //Only No.1~No.6 slave of every PCI_DMC_01 card can run on speedcontinue mode.   
        
        public const short ERR_Backlash_step 	       = 154; //accstep+contstep+decstep must small than 100
        public const short ERR_Backlash_status	       = 155;//motion done status must be 0 and buffer length must be 0 too.	
        
	public const short Compare_Cards_Not_Equal			= 201; //Card Imformation Not Equal.(Card No. && Card Numbers)
        public const short Compare_Nodes_Not_Equal			= 202; //Node Numbers Not Equal.
        public const short Compare_Node_ID_Not_Equal			= 203; //Node Imformation Not Equal.(NodeID)
        public const short Compare_Node_Device_Type_Not_Equal		= 204; //Slave Imformation Not Equal.(Type 2)
        public const short Compare_Node_Identity_Object_Not_Equal	= 205; //Slave Imformation Not Equal.(Type 1)
        public const short Compare_File_Path_NULL			= 206; //File Path Error.
        public const short Compare_File_Open_Fail			= 207; //Open File Error.(Please Input the correct File Path)
        public const short Compare_File_Not_Exist			= 208; //File Not Exist.


	//PCI_DMC_F01
	public const short ERR_NoCardFound				= 301; //The Card No. Not Found or Initial Not Finished.
	public const short ERR_CardInitial				= 302; //Initial Failed.
	public const short ERR_MemoryAccess_Failed			= 303; //Memory Read/Write Failed.
	public const short ERR_MemoryOutOfRange				= 304; //The Operation of Memory is over range.
	public const short ERR_UartTxIsBusy				= 305; //Uart Tx is busy
	public const short ERR_UartRxError				= 306; //Uart Rx Read Error.
	public const short ERR_UartRxIsNotReady				= 307; //Uart Rx Not Ready.
	public const short ERR_NotSupportFunc				= 308; //Can't use the function.
	public const short ERR_NoNodeFound				= 309; //Node No. Error
	public const short ERR_APIInputError				= 310; //The Parameter of the API input Error or Over range.
	public const short ERR_SDOFailed				= 311; //SDO send error.
	public const short ERR_SDOBusy					= 312; //SDO Busy or SDO Writed two commands at the same time.
	public const short ERR_APITypeErr				= 313; //The API can't run above the module.
	public const short ERR_ScaleFailed				= 314; //AD module adjust failed.
	public const short ERR_F_BufferFull				= 315; //MailBox Buffer Fulled.
	public const short ERR_ConnectErr				= 316; //Commuication exceptional.
	public const short ERR_MBWordChFailed				= 317; //MailBox Writed two commands in the same time.
	public const short ERR_MailBoxFailed				= 318; //MailBox send Command failed.
	public const short ERR_CantResetCard				= 319; //ResetCard Can't execute.
	public const short ERR_PDOFailed				= 320; //No Response from PDO.
	public const short ERR_MBCmding					= 321; //Command is processing by MailBox.
	public const short ERR_SVOFF					= 322; //Must be execute Servo On before action.
	public const short ERR_DriverError				= 323; //Must be execute reset Ralm before action because DeviceError to occur.
	public const short ERR_ConnReset_Failed				= 324; //Rest Card Failed.
	public const short ERR_F01SlotIDError				= 325; //Slot Error or SlotID Over Range.
	public const short ERR_UartData_NoMatch				= 326; //Download CodeFailed.
	public const short ERR_SVON					= 327; //You must turn off the Servo before execute the function.
	public const short ERR_Mpg_Already_On				= 328; //The NodeID is Enabled at MPG¡BJog or DDA Function¡AOther Function is Stoped Now.
	public const short ERR_MpgNumber_Over				= 329; //MPG or Jog Numbers is Over Range(MAX:3).
	public const short ERR_Mpg_Data_Failed				= 330; //MPG or Jog Failed in Data Transferred.
	public const short ERR_DDA_Buffer_Full				= 331; //DDA Buffer Full.
	public const short ERR_F_Slave_Security_op			= 332; //Slave Secutiry Write Error.		
	public const short ERR_F_Security_Page				= 333; //Page over the range.
	public const short ERR_F_GetDLLPath				= 334; //Can't find DLL path.
	public const short ERR_F_GetDLLVersion				= 335; //Can't find DLL version data.
	public const short F_Compare_File_Open_Fail			= 336; //Open File Error.(Please Input the correct File Path)	
	public const short F_Compare_File_Not_Exist			= 337; //File Not Exist.
	public const short F_Compare_Cards_Not_Equal			= 338; //Card Imformation Not Equal.(Card No. && Card Numbers)
	public const short F_Compare_File_Path_NULL			= 339; //File Path Error.
	public const short F_Compare_Node_ID_Not_Equal			= 340; //Node Imformation Not Equal.(NodeID)
	public const short F_Compare_Node_Device_Type_Not_Equal		= 341; //Slave Imformation Not Equal.(Type 2)
	public const short F_Compare_Node_Identity_Object_Not_Equal	= 342; //Slave Imformation Not Equal.(Type 1)
	public const short F_Compare_Nodes_Not_Equal			= 343; //Node Numbers Not Equal.
	public const short ERR_SecurityNoRet				= 344; //Security No Response.
	public const short ERR_SDORetTimeOut				= 345; //SDO Timeout.
	public const short ERR_Uart_Connect_Fail			= 346; //Uart failed of communication.


	public const short ERR_UseGetError				= 399; //System Error Occur,Please call DMC_F01_get_card_error_code
	//sub if main = 399
	public const short ERR_sub_NoError				= 0; //OK
	public const short ERR_sub_CantConnect				= 1; //DMC Commucation Connect Failed.
	public const short ERR_sub_SDOFailed				= 2; //SDO Failed.
	public const short ERR_sub_CantReset				= 3; //Reset DMC Commucation Failed.

    }
}