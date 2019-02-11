using System;
using System.Text;
using System.Runtime.InteropServices;

namespace PCI_DMC
{
    public class CPCI_DMC
    {


        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_cardtype_errorcode")]
        public static extern short CS_DMC_01_get_cardtype_errorcode(ref ushort returncode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_clear_cardtype_errorcode")]
        public static extern short CS_DMC_01_clear_cardtype_errorcode();

        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_open")]
        public static extern short CS_DMC_01_open(ref short existcard);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_CardNo_seq")]
        public static extern short CS_DMC_01_get_CardNo_seq(ushort CardNo_seq, ref ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_close")]
        public static extern   void CS_DMC_01_close();
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_pci_initial")]
        public static extern short CS_DMC_01_pci_initial(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_initial_bus")]
        public static extern short CS_DMC_01_initial_bus(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_initial_bus2")]
        public static extern short CS_DMC_01_initial_bus2();
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ring")]
        public static extern short CS_DMC_01_start_ring(ushort CardNo, byte RingNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_card_version")]
        public static extern short CS_DMC_01_get_card_version(ushort CardNo, ref ushort ver);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_devicetype")]
        public static extern short CS_DMC_01_get_devicetype(short CardNo, ushort NodeID, ushort SlotID, ref uint DeviceType, ref uint IdentityObject);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_device_table")]
        public static extern short CS_DMC_01_get_device_table(ushort CardNo, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_node_table")]
        public static extern short CS_DMC_01_get_node_table(ushort CardNo, ref uint NodeIDTable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_cycle_time")]
        public static extern short CS_DMC_01_get_cycle_time(ushort CardNo, ref int time);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_check_card_running")]
        public static extern short CS_DMC_01_check_card_running(ushort CardNo, ref ushort running);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_reset_card")]
        public static extern short CS_DMC_01_reset_card(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_check_nodeno")]
        public static extern short CS_DMC_01_check_nodeno(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort exist);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_master_connect_status")]
        public static extern short CS_DMC_01_get_master_connect_status(ushort CardNo, ref ushort Protocal);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_mailbox_Error")]
        public static extern short CS_DMC_01_get_mailbox_Error(ushort CardNo, ref uint error_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_mailbox_cnt")]
        public static extern short CS_DMC_01_get_mailbox_cnt(ushort CardNo, ref uint PC_cnt, ref uint DSP_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_dsp_cnt")]
        public static extern short CS_DMC_01_get_dsp_cnt(ushort CardNo, ref uint int_cnt, ref uint main_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_dio_output")]
        public static extern short CS_DMC_01_set_dio_output(ushort CardNo, ushort On_Off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_dio_output")]
        public static extern short CS_DMC_01_get_dio_output(ushort CardNo, ref ushort On_Off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_dio_input")]
        public static extern short CS_DMC_01_get_dio_input(ushort CardNo, ref ushort On_Off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_check_canopen_lock")]
        public static extern short CS_DMC_01_check_canopen_lock(ushort CardNo, ref ushort C_lock);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_canopen_ret")]
        public static extern short CS_DMC_01_get_canopen_ret(ushort CardNo, ref ushort COBID, ref byte value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_send_message")]
        public static extern short CS_DMC_01_send_message(ushort CardNo, ushort NodeID, ushort SlotID, ushort Index, ushort SubIdx, ushort DataType, ushort Value0, ushort Value1, ushort Value2, ushort Value3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_send_message2")]
        public static extern short CS_DMC_01_send_message2(ushort CardNo, ushort NodeID, ushort SlotID, ushort Index, ushort SubIdx, ushort DataType, ushort Value0, ushort Value1, ushort Value2, ushort Value3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_send_message3")]
        public static extern short CS_DMC_01_send_message3(short CardNo, ref ushort Index, ref ushort SubIdx, ref ushort DataType, ref ushort Value0, ref ushort Value1, ref ushort Value2, ref ushort Value3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_reset_sdo_choke")]
        public static extern short CS_DMC_01_reset_sdo_choke(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_sdo_retry_history")]
        public static extern short CS_DMC_01_get_sdo_retry_history(ushort CardNo, ref uint cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_message")]
        public static extern short CS_DMC_01_read_message(short CardNo, ref ushort Cmd, ref ushort COBID, ref ushort SubIdx, ref ushort Value0, ref ushort Value1, ref ushort Value2, ref ushort Value3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_message2")]
        public static extern short CS_DMC_01_read_message2(short CardNo, ushort NodeID, ref ushort Cmd, ref ushort COBID, ref ushort SubIdx, ref ushort Value0, ref ushort Value1, ref ushort Value2, ref ushort Value3, ref ushort cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_message")]
        public static extern short CS_DMC_01_get_message(short CardNo, ushort NodeID, ushort SlotID, ushort Index, ushort SubIdx, ref ushort Cmd, ref ushort COBID, ref ushort SubIndex, ref ushort Value0, ref ushort Value1, ref ushort Value2, ref ushort Value3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_sdo_driver_speed_profile")]
        public static extern short CS_DMC_01_set_sdo_driver_speed_profile(ushort CardNo, ushort NodeID, ushort SlotID, uint MaxVel, double acc, double dec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sdo_driver_r_move")]
        public static extern short CS_DMC_01_start_sdo_driver_r_move(ushort CardNo, ushort NodeID, ushort SlotID, int Distance);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sdo_driver_a_move")]
        public static extern short CS_DMC_01_start_sdo_driver_a_move(short CardNo, ushort NodeID, ushort SlotID, int Position);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sdo_driver_new_position_move")]
        public static extern short CS_DMC_01_start_sdo_driver_new_position_move(short CardNo, ushort NodeID, ushort SlotID, int Position, ushort abs_rel);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_home_config")]
        public static extern short CS_DMC_01_set_home_config(ushort CardNo, ushort NodeID, ushort SlotID, ushort Mode, int offset, ushort lowSpeed, ushort highSpeed, double acc);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_home_move")]
        public static extern short CS_DMC_01_set_home_move(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_home_config2")]
        public static extern short CS_DMC_01_set_home_config2(ushort CardNo, ushort NodeID, ushort SlotID, ushort Mode, int offset, ushort lowSpeed, ushort highSpeed, double acc);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_home_config_offset2")]
        public static extern short CS_DMC_01_set_home_config_offset2(ushort CardNo, ushort NodeID, ushort SlotID, int offset);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_home_move2")]
        public static extern short CS_DMC_01_set_home_move2(ushort CardNo, ushort NodeID, ushort SlotID, ushort Mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_escape_home_move")]
        public static extern short CS_DMC_01_escape_home_move(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_ipo_set_svon")]
        public static extern short CS_DMC_01_ipo_set_svon(ushort CardNo, ushort NodeID, ushort SlotID, ushort ON_OFF);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_motion_cnt")]
        public static extern short CS_DMC_01_motion_cnt(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort pc_mc_cnt, ref ushort dsp_mc_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_move")]
        public static extern short CS_DMC_01_start_tr_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_move")]
        public static extern short CS_DMC_01_start_sr_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_move")]
        public static extern short CS_DMC_01_start_ta_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_move")]
        public static extern short CS_DMC_01_start_sa_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_p_change")]
        public static extern short CS_DMC_01_p_change(ushort CardNo, ushort NodeID, ushort SlotID, int NewPos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_v_change")]
        public static extern short CS_DMC_01_v_change(ushort CardNo, ushort NodeID, ushort SlotID, int NewSpeed, double sec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_move_2seg")]
        public static extern short CS_DMC_01_start_tr_move_2seg(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_move_2seg")]
        public static extern short CS_DMC_01_start_sr_move_2seg(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_move_2seg")]
        public static extern short CS_DMC_01_start_ta_move_2seg(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_move_2seg")]
        public static extern short CS_DMC_01_start_sa_move_2seg(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_move_2seg2")]
        public static extern short CS_DMC_01_start_tr_move_2seg2(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_move_2seg2")]
        public static extern short CS_DMC_01_start_sr_move_2seg2(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_move_2seg2")]
        public static extern short CS_DMC_01_start_ta_move_2seg2(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_move_2seg2")]
        public static extern short CS_DMC_01_start_sa_move_2seg2(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int Dist2, int StrVel, int MaxVel, int MaxVel2, double Tacc, double Tsec, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_move_xy")]
        public static extern short CS_DMC_01_start_tr_move_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_move_xy")]
        public static extern short CS_DMC_01_start_sr_move_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_move_xy")]
        public static extern short CS_DMC_01_start_ta_move_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_move_xy")]
        public static extern short CS_DMC_01_start_sa_move_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_arc_xy")]
        public static extern short CS_DMC_01_start_tr_arc_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_arc_xy")]
        public static extern short CS_DMC_01_start_sr_arc_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_arc_xy")]
        public static extern short CS_DMC_01_start_ta_arc_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_arc_xy")]
        public static extern short CS_DMC_01_start_sa_arc_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_arc2_xy")]
        public static extern short CS_DMC_01_start_tr_arc2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int End_X, int End_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_arc2_xy")]
        public static extern short CS_DMC_01_start_sr_arc2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int End_X, int End_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_arc2_xy")]
        public static extern short CS_DMC_01_start_ta_arc2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int End_X, int End_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_arc2_xy")]
        public static extern short CS_DMC_01_start_sa_arc2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int End_X, int End_Y, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_arc3_xy")]
        public static extern short CS_DMC_01_start_tr_arc3_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int End_X, int End_Y, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_arc3_xy")]
        public static extern short CS_DMC_01_start_sr_arc3_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int End_X, int End_Y, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_arc3_xy")]
        public static extern short CS_DMC_01_start_ta_arc3_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int End_X, int End_Y, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_arc3_xy")]
        public static extern short CS_DMC_01_start_sa_arc3_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int End_X, int End_Y, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_move_xyz")]
        public static extern short CS_DMC_01_start_tr_move_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int DisZ, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_move_xyz")]
        public static extern short CS_DMC_01_start_sr_move_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int DisZ, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_move_xyz")]
        public static extern short CS_DMC_01_start_ta_move_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int DisZ, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_move_xyz")]
        public static extern short CS_DMC_01_start_sa_move_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int DisZ, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_tr_heli_xy")]
        public static extern short CS_DMC_01_start_tr_heli_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int Depth, int Pitch, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sr_heli_xy")]
        public static extern short CS_DMC_01_start_sr_heli_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int Depth, int Pitch, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_ta_heli_xy")]
        public static extern short CS_DMC_01_start_ta_heli_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int Depth, int Pitch, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_sa_heli_xy")]
        public static extern short CS_DMC_01_start_sa_heli_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int Depth, int Pitch, short Dir, int StrVel, int MaxVel, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_tv_move")]
        public static extern short CS_DMC_01_tv_move(ushort CardNo, ushort NodeID, ushort SlotID, int StrVel, int MaxVel, double Tacc, short Dir);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_sv_move")]
        public static extern short CS_DMC_01_sv_move(ushort CardNo, ushort NodeID, ushort SlotID, int StrVel, int MaxVel, double Tacc, short Dir);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_motion_error_code")]
        public static extern short CS_DMC_01_get_motion_error_code(ushort CardNo, ushort NodeID, ushort SlotID, ref short error_code);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_motion_error_cnt")]
        public static extern short CS_DMC_01_get_motion_error_cnt(ushort CardNo, ushort NodeID, ushort SlotID, ref short error_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_emg_stop")]
        public static extern short CS_DMC_01_emg_stop(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_sd_stop")]
        public static extern short CS_DMC_01_sd_stop(ushort CardNo, ushort NodeID, ushort SlotID, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_sd_abort")]
        public static extern short CS_DMC_01_sd_abort(ushort CardNo, ushort NodeID, ushort SlotID, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_motion_done")]
        public static extern short CS_DMC_01_motion_done(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort MC_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_motion_status")]
        public static extern short CS_DMC_01_motion_status(ushort CardNo, ushort NodeID, ushort SlotID, ref uint MC_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_interrutp_buffer")]
        public static extern short CS_DMC_01_set_interrutp_buffer(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_command")]
        public static extern short CS_DMC_01_get_command(ushort CardNo, ushort NodeID, ushort SlotID, ref int cmd);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_target_pos")]
        public static extern short CS_DMC_01_get_target_pos(ushort CardNo, ushort NodeID, ushort SlotID, ref int pos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_current_speed")]
        public static extern short CS_DMC_01_get_current_speed(ushort CardNo, ushort NodeID, ushort SlotID, ref int speed);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_position")]
        public static extern ushort CS_DMC_01_set_position(ushort CardNo, ushort NodeID, ushort SlotID, int pos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_command")]
        public static extern short CS_DMC_01_set_command(ushort CardNo, ushort NodeID, ushort SlotID, int cmd);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_position")]
        public static extern short CS_DMC_01_get_position(ushort CardNo, ushort NodeID, ushort SlotID, ref int pos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_current_speed_rpm")]
        public static extern short CS_DMC_01_get_current_speed_rpm(ushort CardNo, ushort NodeID, ushort SlotID, ref int rpm);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_buffer_length")]
        public static extern short CS_DMC_01_get_buffer_length(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort bufferLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_gear")]
        public static extern short CS_DMC_01_set_gear(ushort CardNo, ushort NodeID, ushort SlotID, short numerator, short denominator, ushort Enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_torque_mode")]
        public static extern short CS_DMC_01_set_torque_mode(ushort CardNo, ushort NodeID, ushort SlotID, uint slope);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_torque")]
        public static extern short CS_DMC_01_set_torque(ushort CardNo, ushort NodeID, ushort SlotID, int ratio);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_torque_stop")]
        public static extern short CS_DMC_01_set_torque_stop(ushort CardNo, ushort NodeID, ushort SlotID, ushort stop);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_torque_velocity_feed_forward")]
        public static extern short CS_DMC_01_set_torque_velocity_feed_forward(ushort CardNo, ushort NodeID, ushort SlotID, uint velocity_feed_forward);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_torque_velocity_limit")]
        public static extern short CS_DMC_01_set_torque_velocity_limit(ushort CardNo, ushort NodeID, ushort SlotID, uint velocity_limit);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_velocity_mode")]
        public static extern short CS_DMC_01_set_velocity_mode(ushort CardNo, ushort NodeID, ushort SlotID, double Tacc, double Tdec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_velocity")]
        public static extern short CS_DMC_01_set_velocity(ushort CardNo, ushort NodeID, ushort SlotID, int rpm);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_velocity_stop")]
        public static extern short CS_DMC_01_set_velocity_stop(ushort CardNo, ushort NodeID, ushort SlotID, ushort stop);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_velocity_torque_feed_forward")]
        public static extern short CS_DMC_01_set_velocity_torque_feed_forward(ushort CardNo, ushort NodeID, ushort SlotID, uint torque_feed_forward);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_velocity_torque_limit")]
        public static extern short CS_DMC_01_set_velocity_torque_limit(ushort CardNo, ushort NodeID, ushort SlotID, uint torque_limit);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_torque")]
        public static extern short CS_DMC_01_get_torque(ushort CardNo, ushort NodeID, ushort SlotID, ref short torque);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_rpm")]
        public static extern short CS_DMC_01_get_rpm(ushort CardNo, ushort NodeID, ushort SlotID, ref short rpm);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_servo_parameter")]
        public static extern short CS_DMC_01_read_servo_parameter(ushort CardNo, ushort NodeID, ushort SlotID, ushort group, ushort idx, ref uint data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_write_servo_parameter")]
        public static extern short CS_DMC_01_write_servo_parameter(ushort CardNo, ushort NodeID, ushort SlotID, ushort group, ushort idx, uint data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_ralm")]
        public static extern short CS_DMC_01_set_ralm(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_ralm_servo_all")]
        public static extern short CS_DMC_01_set_ralm_servo_all(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_alm_code")]
        public static extern short CS_DMC_01_get_alm_code(ushort CardNo, ushort NodeID, ushort SlotID, ref uint alm_code);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_master_alm_code")]
        public static extern short CS_DMC_01_master_alm_code(ushort CardNo, ref ushort alm_code);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_slave_error")]
        public static extern short CS_DMC_01_slave_error(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort alm_cnt);

        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_app_get_circle_endpoint")]
        public static extern short CS_misc_app_get_circle_endpoint(int Start_X, int Start_Y, int Center_X, int Center_Y, double Angle, ref int end_x, ref int end_y);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_app_get_circle_center_point")]
        public static extern short CS_misc_app_get_circle_center_point(int Start_X, int Start_Y, int End_x, int End_y, double Angle, ref int Center_X, ref int Center_Y);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_set_record_debuging")]
        public static extern short CS_misc_set_record_debuging(ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_open_record_debuging_file")]
        public static extern short CS_misc_open_record_debuging_file(ushort Card_Type, string file_name, ushort open);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_security")]
        public static extern short CS_misc_security(uint OtherWord0, uint OtherWord1, uint SyntekWord0, uint SyntekWord1, ref uint Password0, ref uint Password1, ref uint Password2, ref uint Password3);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_get_serialno")]
        public static extern short CS_misc_slave_get_serialno(ushort CardNo, ushort NodeID, ushort SlotID, ref uint SerialNO);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_check_verifykey")]
        public static extern short CS_misc_slave_check_verifykey(ushort CardNo, ushort NodeID, ushort SlotID, ref uint VerifyKey, ref ushort Lock_state);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_check_userpassword")]
        public static extern short CS_misc_slave_check_userpassword(ushort CardNo, ushort NodeID, ushort SlotID, ref uint Password_data, ref ushort Password_state);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_write_userpassword")]
        public static extern short CS_misc_slave_write_userpassword(ushort CardNo, ushort NodeID, ushort SlotID, ref uint Password_data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_write_verifykey")]
        public static extern short CS_misc_slave_write_verifykey(ushort CardNo, ushort NodeID, ushort SlotID, ref uint VerifyKey);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_user_data_buffer_write")]
        public static extern short CS_misc_slave_user_data_buffer_write(ushort CardNo, ushort NodeID, ushort SlotID, ushort Address, ref uint Data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_user_data_to_flash")]
        public static extern short CS_misc_slave_user_data_to_flash(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_misc_slave_user_data_buffer_read")]
        public static extern short CS_misc_slave_user_data_buffer_read(ushort CardNo, ushort NodeID, ushort SlotID, ushort Address, ref uint Data);
        //I16 _stdcall _misc_app_get_spiral_endpoint(I32 start_X, I32 start_Y, I32 center_X, I32 center_Y, I32 spiral_interval, I32 spiral_angle, I32* end_x, I32* end_y);

        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_rm_input_value")]
        public static extern short CS_DMC_01_get_rm_input_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_input_filter")]
        public static extern short CS_DMC_01_set_rm_input_filter(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_input_filter_enable")]
        public static extern short CS_DMC_01_set_rm_input_filter_enable(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_output_value")]
        public static extern short CS_DMC_01_set_rm_output_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_rm_output_value")]
        public static extern short CS_DMC_01_get_rm_output_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_output_value_error_handle")]
        public static extern short CS_DMC_01_set_rm_output_value_error_handle(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_rm_output_value_error_handle")]
        public static extern short CS_DMC_01_get_rm_output_value_error_handle(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_output_active")]
        public static extern short CS_DMC_01_set_rm_output_active(ushort CardNo, ushort NodeID, ushort SlotID, ushort Enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_mpg_axes_enable")]
        public static extern short CS_DMC_01_set_rm_mpg_axes_enable(ushort CardNo, ushort MasterNodeID, ushort MasterSlotID, ref ushort NodeID, ref ushort SlotID, ushort enable, ushort pulser_ratio, ref uint ratio, ref uint slop);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_mpg_axes_enable2")]
        public static extern short CS_DMC_01_set_rm_mpg_axes_enable2(ushort CardNo, ushort MasterNodeID, ushort MasterSlotID, ref ushort NodeID, ref ushort SlotID, ushort enable, ushort pulser_ratio, ref uint ratio, ref uint slop, ref ushort denominator);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_jog_axes_enable")]
        public static extern short CS_DMC_01_set_rm_jog_axes_enable(ushort CardNo, ushort MasterNodeID, ushort MasterSlotID, ref ushort NodeID, ref ushort SlotID, ushort enable, ushort jog_mode, ref int jog_speed, ref double sec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_command_buf_clear")]
        public static extern short CS_DMC_01_command_buf_clear(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_buf_dwell")]
        public static extern short CS_DMC_01_buf_dwell(short CardNo, ushort NodeID, ushort SlotID, int dwell_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_sync_move")]
        public static extern short CS_DMC_01_sync_move(short CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_sync_move_config")]
        public static extern short CS_DMC_01_sync_move_config(short CardNo, ushort NodeID, ushort SlotID, short enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_soft_limit")]
        public static extern short CS_DMC_01_set_soft_limit(ushort CardNo, ushort NodeID, ushort SlotID, int PLimit, int NLimit);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_soft_limit")]
        public static extern short CS_DMC_01_enable_soft_limit(ushort CardNo, ushort NodeID, ushort SlotID, short Action);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_disable_soft_limit")]
        public static extern short CS_DMC_01_disable_soft_limit(ushort CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_soft_limit_status")]
        public static extern short CS_DMC_01_get_soft_limit_status(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort PLimit_sts, ref ushort NLimit_sts);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_slave_version")]
        public static extern short CS_DMC_01_get_slave_version(short CardNo, ushort NodeID, ushort SlotID, ref ushort version);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_slave_subversion")]
        public static extern short CS_DMC_01_get_slave_subversion(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort SubVersion);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_monitor")]
        public static extern short CS_DMC_01_set_monitor(ushort CardNo, ushort NodeID, ushort SlotID, ushort monitorw);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_monitor")]
        public static extern short CS_DMC_01_get_monitor(ushort CardNo, ushort NodeID, ushort SlotID, ref uint value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_servo_command")]
        public static extern short CS_DMC_01_get_servo_command(ushort CardNo, ushort NodeID, ushort SlotID, ref int servo_cmd);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_servo_DI")]
        public static extern short CS_DMC_01_get_servo_DI(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort servo_DI);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_servo_DO")]
        public static extern short CS_DMC_01_get_servo_DO(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort servo_DO);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_trigger_buf_function")]
        public static extern short CS_DMC_01_set_trigger_buf_function(short CardNo, ushort NodeID, ushort SlotID, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_multi_axes_move")]
        public static extern short CS_DMC_01_multi_axes_move(ushort CardNo, ushort AxisNum, ref ushort NodeID, ref ushort SlotID, ref int DistArrary, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_spiral_xy")]
        public static extern short CS_DMC_01_start_spiral_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int spiral_interval, int spiral_angle, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_spiral2_xy")]
        public static extern short CS_DMC_01_start_spiral2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int center_x, int center_y, int end_x, int end_y, ushort dir, ushort circlenum, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_ipulser_mode")]
        public static extern short CS_DMC_01_set_rm_04pi_ipulser_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_opulser_mode")]
        public static extern short CS_DMC_01_set_rm_04pi_opulser_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_svon_polarity")]
        public static extern short CS_DMC_01_set_rm_04pi_svon_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ushort polarity);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_DO2")]
        public static extern short CS_DMC_01_set_rm_04pi_DO2(ushort CardNo, ushort NodeID, ushort SlotID, ushort ON_OFF);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_04pi_set_poweron")]
        public static extern short CS_DMC_01_04pi_set_poweron(ushort CardNo, ushort NodeID, ushort SlotID, ushort ON_OFF);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_homing_ratio")]
        public static extern short CS_DMC_01_set_rm_04pi_homing_ratio(ushort CardNo, ushort NodeID, ushort SlotID, ushort ratio);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_get_MEL_polarity")]
        public static extern short CS_DMC_01_rm_04pi_get_MEL_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_set_MEL_polarity")]
        public static extern short CS_DMC_01_rm_04pi_set_MEL_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ushort inverse);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_get_PEL_polarity")]
        public static extern short CS_DMC_01_rm_04pi_get_PEL_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_set_PEL_polarity")]
        public static extern short CS_DMC_01_rm_04pi_set_PEL_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ushort inverse);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_move")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_v_move")]
        public static extern short CS_DMC_01_rm_04pi_md1_v_move(ushort CardNo, ushort NodeID, ushort SlotID, int StrVel, int MaxVel, double Tacc, short dir, ushort m_curve);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_line2")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_line2(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Dist, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_arc")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_arc(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Center, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_line3")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_line3(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Dist, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_line4")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_line4(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Dist, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_arc2")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_arc2(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int End, double Angle, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_arc3")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_arc3(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Center, ref int End, short dir, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_start_heli")]
        public static extern short CS_DMC_01_rm_04pi_md1_start_heli(ushort CardNo, ushort NodeID, ref ushort SlotID, ref int Center, int Depth, int Pitch, short dir, int StrVel, int MaxVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_set_gear")]
        public static extern short CS_DMC_01_rm_04pi_md1_set_gear(ushort CardNo, ushort NodeID, ushort SlotID, short numerator, short denominator, ushort Enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_p_change")]
        public static extern short CS_DMC_01_rm_04pi_md1_p_change(ushort CardNo, ushort NodeID, ushort SlotID, int NewPos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_v_change")]
        public static extern short CS_DMC_01_rm_04pi_md1_v_change(ushort CardNo, ushort NodeID, ushort SlotID, int NewSpeed, double sec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_set_soft_limit")]
        public static extern short CS_DMC_01_rm_04pi_md1_set_soft_limit(ushort CardNo, ushort NodeID, ushort SlotID, int PLimit, int NLimit, ushort Enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_get_soft_limit_status")]
        public static extern short CS_DMC_01_rm_04pi_md1_get_soft_limit_status(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort NLimit_status, ref ushort PLimit_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_set_sld")]
        public static extern short CS_DMC_01_rm_04pi_md1_set_sld(ushort CardNo, ushort NodeID, ushort SlotID, short enable, short sd_logic, short mode);
        //I16 _stdcall _DMC_01_04pi_md1_get_mc_error_code(U16 CardNo, U16 NodeID, U16 SlotID, U16 *error_code);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_04pi_ref_counter")]
        public static extern short CS_DMC_01_set_rm_04pi_ref_counter(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04PI_get_buffer")]
        public static extern short CS_DMC_01_rm_04PI_get_buffer(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort bufferLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04pi_md1_get_mc_error_code")]
        public static extern short CS_DMC_01_rm_04pi_md1_get_mc_error_code(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort error_code);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_link_interrupt")]
        public static extern short CS_DMC_01_link_interrupt(ushort CardNo, ushort __stdcal);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_int_factor")]
        public static extern short CS_DMC_01_set_int_factor(ushort CardNo, ushort NodeID, ushort int_factor);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_int_enable")]
        public static extern short CS_DMC_01_int_enable(ushort CardNo, ushort NodeID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_int_count")]
        public static extern short CS_DMC_01_get_int_count(ushort CardNo, ushort NodeID, ref ushort count);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_int_disable")]
        public static extern short CS_DMC_01_int_disable(ushort CardNo, ushort NodeID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_int_status")]
        public static extern short CS_DMC_01_get_int_status(ushort CardNo, ushort NodeID, ref ushort event_int_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_security")]
        public static extern short CS_DMC_01_read_security(ushort CardNo, ushort page, ref ushort array);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_security_status")]
        public static extern short CS_DMC_01_read_security_status(ushort CardNo, ref ushort status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_write_security_status")]
        public static extern short CS_DMC_01_write_security_status(ushort CardNo, ushort status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_write_security")]
        public static extern short CS_DMC_01_write_security(ushort CardNo, ushort page, ref ushort array);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_check_userpassword")]
        public static extern short CS_DMC_01_check_userpassword(ushort CardNo, ref uint password_data, ref ushort password_state);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_write_userpassword")]
        public static extern short CS_DMC_01_write_userpassword(ushort CardNo, ref uint password_data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_check_verifykey")]
        public static extern short CS_DMC_01_check_verifykey(ushort CardNo, ref uint VerifyKey, ref ushort state);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_write_verifykey")]
        public static extern short CS_DMC_01_write_verifykey(ushort CardNo, ref uint verifykey);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_read_serialno")]
        public static extern short CS_DMC_01_read_serialno(ushort CardNo, ref uint SerialNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_value")]
        public static extern short CS_DMC_01_rm_04da_set_output_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_get_return_code")]
        public static extern short CS_DMC_01_rm_04da_get_return_code(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort returncode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_range")]
        public static extern short CS_DMC_01_rm_04da_set_output_range(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort range);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_enable")]
        public static extern short CS_DMC_01_rm_04da_set_output_enable(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_overrange")]
        public static extern short CS_DMC_01_rm_04da_set_output_overrange(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_error_clear")]
        public static extern short CS_DMC_01_rm_04da_set_output_error_clear(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_read_control_register")]
        public static extern short CS_DMC_01_rm_04da_read_control_register(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort control_register_data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_get_output_value")]
        public static extern short CS_DMC_01_rm_04da_get_output_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_error_handle")]
        public static extern short CS_DMC_01_rm_04da_set_output_error_handle(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_set_output_offset_value")]
        public static extern short CS_DMC_01_rm_04da_set_output_offset_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, short value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_get_output_offset_value")]
        public static extern short CS_DMC_01_rm_04da_get_output_offset_value(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref short value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_rm_04da_read_data")]
        public static extern short CS_DMC_01_rm_04da_read_data(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_04ad_input_range")]
        public static extern short CS_DMC_01_set_04ad_input_range(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort range);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_04ad_input_range")]
        public static extern short CS_DMC_01_get_04ad_input_range(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort range);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_04ad_conversion_time")]
        public static extern short CS_DMC_01_set_04ad_conversion_time(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_04ad_conversion_time")]
        public static extern short CS_DMC_01_get_04ad_conversion_time(ushort CardNo, ushort NodeID, ushort SlotID, ref ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_04ad_data")]
        public static extern short CS_DMC_01_get_04ad_data(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_04ad_average_mode")]
        public static extern short CS_DMC_01_set_04ad_average_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_04ad_average_mode")]
        public static extern short CS_DMC_01_get_04ad_average_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ref ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_04ad_input_enable")]
        public static extern short CS_DMC_01_set_04ad_input_enable(ushort CardNo, ushort NodeID, ushort SlotID, ushort channelno, ushort ON_OFF);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_group")]
        public static extern short CS_DMC_01_set_group(ushort CardNo, ref ushort NodeID, ref ushort SlotID, ushort NodeID_Num, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_speed_continue")]
        public static extern short CS_DMC_01_speed_continue(ushort CardNo, ushort NodeID, ushort SlotID, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_speed_continue_mode")]
        public static extern short CS_DMC_01_speed_continue_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_speed_continue_combine_ratio")]
        public static extern short CS_DMC_01_speed_continue_combine_ratio(ushort CardNo, ushort NodeID, ushort SlotID, ushort ratio);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_sd_mode")]
        public static extern short CS_DMC_01_set_sd_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_scurve_rate")]
        public static extern short CS_DMC_01_set_scurve_rate(ushort CardNo, ushort NodeID, ushort SlotID, ushort scurve_rate);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_dda_mode")]
        public static extern short CS_DMC_01_enable_dda_mode(ushort CardNo, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_dda_data")]
        public static extern short CS_DMC_01_set_dda_data(ushort CardNo, ref int abs_pos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_dda_cnt")]
        public static extern short CS_DMC_01_get_dda_cnt(ushort CardNo, ref ushort dda_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_feedrate_overwrite")]
        public static extern short CS_DMC_01_feedrate_overwrite(ushort CardNo, ushort NodeID, ushort SlotID, ushort Mode, int New_Speed, double sec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_move")]
        public static extern short CS_DMC_01_start_v3_move(ushort CardNo, ushort NodeID, ushort SlotID, int Dist, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_move_xy")]
        public static extern short CS_DMC_01_start_v3_move_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_arc_xy")]
        public static extern short CS_DMC_01_start_v3_arc_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, double Angle, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_arc2_xy")]
        public static extern short CS_DMC_01_start_v3_arc2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int End_X, int End_Y, double Angle, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_arc3_xy")]
        public static extern short CS_DMC_01_start_v3_arc3_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int End_x, int End_y, short Dir, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_move_xyz")]
        public static extern short CS_DMC_01_start_v3_move_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int DisX, int DisY, int DisZ, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_heli_xy")]
        public static extern short CS_DMC_01_start_v3_heli_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int Depth, int Pitch, short Dir, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_multi_axes")]
        public static extern short CS_DMC_01_start_v3_multi_axes(ushort CardNo, ushort AxisNum, ref ushort NodeID, ref ushort SlotID, ref int DistArrary, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_spiral_xy")]
        public static extern short CS_DMC_01_start_v3_spiral_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int Center_X, int Center_Y, int spiral_interval, uint spiral_angle, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_spiral2_xy")]
        public static extern short CS_DMC_01_start_v3_spiral2_xy(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int center_x, int center_y, int end_x, int end_y, ushort dir, ushort circlenum, int StrVel, int ConstVel, int EndVel, double Tphase1, double Tphase2, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_fifo_cnt")]
        public static extern short CS_DMC_01_get_fifo_cnt(ushort CardNo, ref ushort fifo_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_liner_speed_master")]
        public static extern short CS_DMC_01_liner_speed_master(ushort CardNo, ushort NodeID, ushort SlotID, ushort Mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_record_mode")]
        public static extern short CS_DMC_01_enable_record_mode(ushort CardNo, ushort enable, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_record_data")]
        public static extern short CS_DMC_01_get_record_data(ushort CardNo, ref uint record_data);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_dec_acceleration")]
        public static extern short CS_DMC_01_set_dec_acceleration(ushort CardNo, ushort NodeID, ushort SlotID, double pss);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_sd_time")]
        public static extern short CS_DMC_01_set_sd_time(ushort CardNo, ushort NodeID, ushort SlotID, double sd_dec);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_DLL_path")]
        public static extern short CS_DMC_01_get_DLL_path(string lpFilePath, uint nSize, ref uint nLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_DLL_version")]
        public static extern short CS_DMC_01_get_DLL_version(string lpBuf, uint nSize, ref uint nLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_DLL_path_Single")]
        public static extern short CS_DMC_01_get_DLL_path_Single(ushort Card_Type, string lpFilePath, uint nSize, ref uint nLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_DLL_version_Single")]
        public static extern short CS_DMC_01_get_DLL_version_Single(ushort Card_Type, string lpBuf, uint nSize, ref uint nLength);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_switch_buffer")]
        public static extern short CS_DMC_01_enable_switch_buffer(ushort CardNo, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_save_device_info")]
        public static extern short CS_DMC_01_save_device_info(ushort Card_Type, string path);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_compare_device_info")]
        public static extern short CS_DMC_01_compare_device_info(ushort Card_Type, string path, string pathLog);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_sld_stop")]
        public static extern short CS_DMC_01_set_sld_stop(ushort CardNo, ushort NodeID, ushort SlotID, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_sld_mode")]
        public static extern short CS_DMC_01_set_sld_mode(ushort CardNo, ushort NodeID, ushort SlotID, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_pdo_Rx_error")]
        public static extern short CS_DMC_01_pdo_Rx_error(ushort CardNo, ushort NodeID, ushort SlotID, ref uint Rx_error1, ref uint Rx_error2, ushort set_get);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_pdo_Tx_error")]
        public static extern short CS_DMC_01_pdo_Tx_error(ushort CardNo, ushort NodeID, ushort SlotID, ref uint Tx_error, ushort set_get);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_start_v3_sphere_xyz")]
        public static extern short CS_DMC_01_start_v3_sphere_xyz(ushort CardNo, ref ushort NodeID, ref ushort SlotID, int pos1_x, int pos1_y, int pos1_z, int pos2_x, int pos2_y, int pos2_z, int StrVel, int MaxVel, int EndVel, double Tacc, double Tdec, ushort m_curve, ushort m_r_a);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_function_time")]
        public static extern short CS_DMC_01_set_function_time(ushort CardNo, uint ms1000);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_mc_fifo_mode")]
        public static extern short CS_DMC_01_set_mc_fifo_mode(ushort CardNo, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_axis_mode")]
        public static extern short CS_DMC_01_enable_axis_mode(ushort CardNo, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_clear_dda_data")]
        public static extern short CS_DMC_01_clear_dda_data(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_pitch")]
        public static extern short CS_DMC_01_set_cam_pitch(ushort CardNo, ushort NodeID, ushort SlotID, int pitch);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_org")]
        public static extern short CS_DMC_01_set_cam_org(short CardNo, ushort NodeID, ushort SlotID, short Dir, int orgpos);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_enable")]
        public static extern short CS_DMC_01_set_cam_enable(short CardNo, ushort NodeID, ushort SlotID, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_mode")]
        public static extern short CS_DMC_01_set_cam_mode(short CardNo, ushort NodeID, ushort SlotID, ushort Mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_table")]
        public static extern short CS_DMC_01_set_cam_table(short CardNo, ushort NodeID, ushort SlotID, short Dir, ref int table);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cam_table2")]
        public static extern short CS_DMC_01_set_cam_table2(short CardNo, ushort NodeID, ushort SlotID, short Dir, ref int table, int Num);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_dmc_mode")]
        public static extern short CS_DMC_01_set_dmc_mode(short CardNo, ushort NodeID, ushort SlotID);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_backlash")]
        public static extern short CS_DMC_01_set_backlash(short CardNo, ushort NodeID, ushort SlotID, short enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_backlash_info")]
        public static extern short CS_DMC_01_set_backlash_info(short CardNo, ushort NodeID, ushort SlotID, short backlash, ushort accstep, ushort contstep, ushort decstep);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_electcam")]
        public static extern short CS_DMC_01_enable_electcam(short CardNo, ushort NodeID, ushort SlotID, ushort enable, ushort axisbit, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_cv_mode")]
        public static extern short CS_DMC_01_set_cv_mode(short CardNo, ushort NodeID, ushort SlotID, ushort flag_const_v, ushort flag_dec_c);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_ex_position")]
        public static extern short CS_DMC_01_enable_ex_position(short CardNo, ushort NodeID, ushort SlotID, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_6x_mode")]
        public static extern short CS_DMC_01_enable_6x_mode(ushort CardNo, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_enable_always_monitorid")]
        public static extern short CS_DMC_01_enable_always_monitorid(ushort CardNo, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_nrline_status")]
        public static extern short CS_DMC_01_get_nrline_status(short CardNo, ushort NodeID, ushort SlotID, ref ushort nrline_sts);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_gear")]
        public static extern short CS_DMC_01_get_gear(ushort CardNo, ushort NodeID, ushort SlotID, ref short numerator, ref short denominator, ref ushort Enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_01ph_alm_polarity")]
        public static extern short CS_DMC_01_set_01ph_alm_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ushort polarity);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_01ph_org_polarity")]
        public static extern short CS_DMC_01_set_01ph_org_polarity(ushort CardNo, ushort NodeID, ushort SlotID, ushort polarity);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_rm_output_value_HPF")]
        public static extern short CS_DMC_01_get_rm_output_value_HPF(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ref ushort value);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_rm_output_value_HPF")]
        public static extern short CS_DMC_01_set_rm_output_value_HPF(ushort CardNo, ushort NodeID, ushort SlotID, ushort port, ushort value);

        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_position")]
        public static extern short CS_DMC_01_set_compare_channel_position(ushort CardNo, ushort compare_channel, int position);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_compare_channel_position")]
        public static extern short CS_DMC_01_get_compare_channel_position(ushort CardNo, ushort compare_channel, ref int position);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_ipulser_mode")]
        public static extern short CS_DMC_01_set_compare_ipulser_mode(ushort CardNo, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_direction")]
        public static extern short CS_DMC_01_set_compare_channel_direction(ushort CardNo, ushort compare_channel, ushort dir);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_trigger_time")]
        public static extern short CS_DMC_01_set_compare_channel_trigger_time(ushort CardNo, ushort compare_channel, uint time_us);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_one_shot")]
        public static extern short CS_DMC_01_set_compare_channel_one_shot(ushort CardNo, ushort compare_channel);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_source")]
        public static extern short CS_DMC_01_set_compare_channel_source(ushort CardNo, ushort compare_channel, ushort source);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel0_position_cmp")]
        public static extern short CS_DMC_01_channel0_position_cmp(ushort CardNo, int start, ushort dir, ushort interval, uint trigger_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel0_position_cmp_by_gpio")]
        public static extern short CS_DMC_01_channel0_position_cmp_by_gpio(ushort CardNo, ushort dir, ushort interval, int trigger_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_output_enable")]
        public static extern short CS_DMC_01_channel1_output_enable(ushort CardNo, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_output_mode")]
        public static extern short CS_DMC_01_channel1_output_mode(ushort CardNo, ushort mode);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_get_io_status")]
        public static extern short CS_DMC_01_channel1_get_io_status(ushort CardNo, ref ushort io_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_set_gpio_out")]
        public static extern short CS_DMC_01_channel1_set_gpio_out(ushort CardNo, ushort on_off);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_get_fifo_status")]
        public static extern short CS_DMC_01_channel1_get_fifo_status(ushort CardNo, ref ushort fifo_status);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_get_fifo_counter")]
        public static extern short CS_DMC_01_channel1_get_fifo_counter(ushort CardNo, ref ushort fifo_cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_position_compare_table")]
        public static extern short CS_DMC_01_channel1_position_compare_table(ushort CardNo, ref int pos_table, uint table_size);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_position_compare_table_level")]
        public static extern short CS_DMC_01_channel1_position_compare_table_level(ushort CardNo, ref int pos_table, ref uint level_table, uint table_size);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_position_compare_table_cnt")]
        public static extern short CS_DMC_01_channel1_position_compare_table_cnt(ushort CardNo, ref uint cnt);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_polarity")]
        public static extern short CS_DMC_01_set_compare_channel_polarity(ushort CardNo, ushort inverse);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_position_re_compare_table")]
        public static extern short CS_DMC_01_channel1_position_re_compare_table(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_channel1_position_re_compare_table_level")]
        public static extern short CS_DMC_01_channel1_position_re_compare_table_level(ushort CardNo);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_set_compare_channel_enable")]
        public static extern short CS_DMC_01_set_compare_channel_enable(ushort CardNo, ushort compare_channel, ushort enable);
        [DllImport("PCI_DMC.dll", EntryPoint = "_DMC_01_get_compare_channel_direction")]
        public static extern short CS_DMC_01_get_compare_channel_direction(ushort CardNo, ushort compare_channel, ref ushort dir);


    }
}
