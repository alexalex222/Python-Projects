# -*- coding: utf-8 -*-
"""
Created on Tue May 22 11:04:35 2018

@author: kchen
"""

import sys
import pyodbc
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
import matplotlib.dates
from datetime import timedelta
import win32com.client as win32
import os
from scipy.interpolate import interp1d

class ColdBar:
    def __init__(self, page, item, start_date, end_date):
        self.__user = 'kchen'
        self.__password = 'kuilin1'
        self.page = page
        self.item = item
        self.start_date = start_date
        self.end_date = end_date
        self.pass_schedule_change = False
        self.head_body_temp_bias = False
        self.fmet_tight = False
        self.big_model_err =False
        self.cold_from_furnace = False
        self.email_text = "This is an automatically generated Email. <br />"

        self.piece_query_string = """
        SELECT 
        HMILL_PCE.GEN_EST, 
        HMILL_PCE.HMILL_PCE_NO,
        HMILL_PCE.PROD_BLK_NO, 
        MHS_PROD.MH_ORD_IDENT, 
        MHS_PROD.SDT_NO, 
        HMILL_PCE.SER_NO, 
        HMILL_PCE.PAGE_NO, 
        HMILL_PCE.ITEM_NO, 
        HMILL_PCE.HT_NO, 
        HMILL_PCE.CAST_PCE_NO, 
        HMILL_PCE.PART_NO, 
        MHS_PROD.FUR_NO, 
        MHS_PROD.NORTH_CHRG_PYRO_A_TEMP, 
        MHS_PROD.SOUTH_CHRG_PYRO_A_TEMP, 
        MHS_PROD.SOUTH_CHRG_PYRO_B_TEMP, 
        MHS_PROD.MEAS_SLAB_CHRG_TEMP, 
        MHS_PROD.SLAB_CHRG_TEMP, 
        MHS_PROD.SLAB_TRACK_DUR, 
        MHS_PROD.PLAN_PROCESS_CD, 
        HMILL_PROD.BOTLNCK_AFE_ZON_NO, 
        HMILL_AFE.FM_GAP_VERN_CD, 
        HMILL_PROD.AFE_ALLOW_FLG, 
        HMILL_PROD.AFE_RM_FLG, 
        HMILL_PROD.AFE_DE_SEL_DUR, 
        MHS_PROD.AUTO_EXTR_FLG, 
        HMILL_PROD.AFE_FM_GAP_OFFST_CD, 
        HMILL_AFE.LV2_REQ_ST1_IDLE_TIME, 
        HMILL_PROD.ST1_IDLE_DUR, 
        HMILL_AFE.COILR_1_AVAIL_AT_EXTR_REQ_FLG, 
        HMILL_AFE.COILR_2_AVAIL_AT_EXTR_REQ_FLG, 
        HMILL_AFE.COILR_3_AVAIL_AT_EXTR_REQ_FLG, 
        HMILL_QMS_PROD.COIL_NO, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_1_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_2_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_3_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_4_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_5_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_6_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_7_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_8_PACE_DUR, 
        HMILL_AFE.NXT_EXTR_PRED_ZON_9_PACE_DUR, 
        HMILL_PROD.ACT_EXTR_RME_PU_DUR, 
        HMILL_PROD.ACT_RME_PU_RM_SIM_DUR, 
        HMILL_PROD.ACT_RM_SIM_RMX_PYRO_PU_DUR, 
        HMILL_PROD.ACT_RMX_PYRO_PU_RM_SIM_DO_DUR, 
        HMILL_PROD.ACT_RMX_PYRO_PU_FME_PYRO_DUR, 
        HMILL_PROD.ACT_FME_PYRO_PU_ST1_SIM_DUR, 
        HMILL_PROD.ACT_ST1_SIM_FME_PYRO_DO_DUR, 
        MHS_PROD.SLAB_THK, MHS_PROD.SLAB_LGT, 
        MHS_PROD.SLAB_WDT, MHS_PROD.SLAB_WT, 
        MHS_PROD.HEAT_CD, 
        MHS_PROD.SLAB_GRD_FAM, 
        HMILL_FM_MODEL.FM_FAM_NO, 
        HMILL_FM_MODEL.GRT_INDEX_NO, 
        MHS_PROD.SLAB_GRD_CD, 
        MHS_PROD.SLAB_HT_PRAC_NO, 
        MHS_PROD.SLAB_DLY_DUR, 
        MHS_PROD.SLAB_CHRG_EST, 
        MHS_PROD.PREHT_ZON_EXIT_EST, 
        MHS_PROD.CHRG_ZON_EXIT_EST, 
        MHS_PROD.INTERM_ZON_EXIT_EST, 
        MHS_PROD.SLAB_PHOTOEYE_EST, 
        MHS_PROD.EXTR_READY_EST, 
        MHS_PROD.SLAB_PACE_READY_EST, 
        MHS_PROD.HM2_REQ_EST, 
        MHS_PROD.SLAB_EXTR_EST, 
        MHS_PROD.SLAB_FUR_DUR, 
        HMILL_PCE.ACPT_EST, 
        HMILL_FM_PROD.ST1_PU_EST, 
        MHS_PROD.SLAB_PHOTOEYE_EST - MHS_PROD.SLAB_EXTR_EST AS PHOTOEYE_DUR,
        MHS_PROD.PACE_READY_DUR/(24*3600) AS Count_Down,
        MHS_PROD.SLAB_PHOTOEYE_EST - MHS_PROD.SLAB_EXTR_EST - MHS_PROD.PACE_READY_DUR/(24*3600) AS PHOTO_COUNT,
        MHS_PROD.MAX_ALLOW_PACE_RATE, 
        MHS_PROD.BASE_PACE_RATE, 
        MHS_PROD.EXTR_AIM_PACE_RATE, 
        MHS_PROD.EXTR_RAW_AIM_PACE_RATE, 
        MHS_PROD.EXTR_LV1_PACE_RATE, 
        MHS_PROD.EXTR_PACE_RATE, 
        MHS_PROD.SCHED_HT_PRAC_CD, 
        MHS_PROD.SCHED_HT_INDEX_NO, 
        MHS_PROD.INTGRT_HT_NO, 
        MHS_PROD.INTGRT_EFCT_HT_NO, 
        MHS_PROD.SCHED_EXTR_TEMP, 
        MHS_PROD.INTGRT_AIM_FUR_EXTR_TEMP, 
        MHS_PROD.CALC_SLAB_EXTR_TEMP, 
        HMILL_RM_PHYS.RM_BCK_CALC_FUR_EXTR_TEMP -15 AS RM_BCK_CALC_FUR_EXTR_TEMP, 
        HMILL_RM_PROD2.RM_BCK_CALC_FUR_EXTR_TEMP_CD, 
        HMILL_RM_PHYS.RM_BCK_CALC_FUR_EXTR_TEMP - 15 - MHS_PROD.INTGRT_AIM_FUR_EXTR_TEMP AS TEMP_ERR,
        HMILL_RM_PHYS.RM_BCK_CALC_FUR_EXTR_TEMP - 15 - MHS_PROD.CALC_SLAB_EXTR_TEMP AS MODEL_ERR,
        HMILL_RM_PHYS.RM_BODY_2_SKID_MARK_TEMP, 
        MHS_PROD.RMX_TEMP, 
        HMILL_RM_FM.RMX_PRED_LONG_SCAN_TEMP, 
        HMILL_RM_FM.RM_PRED_RMX_VERN_TEMP, 
        HMILL_RM_FM.RM_PRED_RMX_TEMP, 
        HMILL_RM_PHYS.RM_PRED_RME_RMX_DELTA_TEMP, 
        HMILL_RM_FM.RM_PRED_FM_ARR_VERN_TEMP, 
        MHS_PROD.INTGRT_AIM_FM_ENT_TEMP, 
        HMILL_FSU_PRED.BASE_FME_TEMP, 
        HMILL_FM_ENT_PDI.AIM_ENT_TEMP, 
        HMILL_RM_PHYS.PRED_FME_TEMP, 
        HMILL_RM_FM.FUR_FBK_PROLL_LONG_SCAN_TEMP, 
        HMILL_FM_ENT_PDI.MEAS_FME_TEMP, 
        HMILL_FSU_PRED.PRED_FME_TEMP, 
        HMILL_FM_PROD.CALC_COURSE_2_NO, 
        HMILL_RM_FM.AUTO_BAR_COOL_HOLD_DUR, 
        HMILL_RM_FM.AUTO_BAR_COOL_SEL_FLG, 
        MHS_TEMP_FBK.APPL_FUR_POS_EXTR_TEMP_AIM_ADJ, 
        MHS_PROD.SLAB_DE_PILR_MEAS_LGT, 
        MHS_PROD.SLAB_POS_MEAS_LGT, 
        MHS_PROD.SLAB_POS_LEAD_CHRG_POS, 
        MHS_PROD.SLAB_POS_TRAIL_CHRG_POS, 
        MHS_PROD.SLAB_ALIGN_VALUE, 
        MHS_PROD.CHRG_ZON_ENT_ALIGN_COMPL_SECND, 
        MHS_PROD.SLAB_POS_LEAD_SAFE_ZON_NO, 
        MHS_PROD.SLAB_POS_TRAIL_SAFE_ZON_NO, 
        MHS_PROD.SLAB_POS_SOUTH_CHRG_REQ_FLG, 
        MHS_PROD.FUR_RAD_COEFF_OFFST, 
        MHS_PROD.SLAB_CHAR_RAD_COEFF_OFFST, 
        MHS_PROD.ZON_1_RAD_COEFF, 
        MHS_PROD.ZON_2_RAD_COEFF, 
        MHS_PROD.ZON_3_RAD_COEFF, 
        MHS_PROD.ZON_4_RAD_COEFF, 
        MHS_PROD.ZON_5_RAD_COEFF, 
        MHS_PROD.ZON_6_RAD_COEFF, 
        MHS_PROD.ZON_7_RAD_COEFF, 
        MHS_PROD.ZON_8_RAD_COEFF, 
        MHS_PROD.ZON_9_RAD_COEFF, 
        MHS_PROD.ZON_10_RAD_COEFF, 
        MHS_PROD.ZON_11_RAD_COEFF, 
        MHS_PROD.ZON_12_RAD_COEFF, 
        HMILL_RM_PHYS.SLAB_DSCALE_AVG_TOP_TEMP, 
        HMILL_RM_PHYS.SLAB_DSCALE_AVG_BOT_TEMP, 
        HMILL_RM_PHYS.SLAB_DSCALE_TOP_TEMP_RANGE, 
        HMILL_RM_PHYS.SLAB_DSCALE_BOT_TEMP_RANGE, 
        HMILL_RM_PHYS.ENT_TEMP, 
        HMILL_RM_PHYS.ENT_FBK_TEMP, 
        MHS_PROD.HMILL_AIM_OFFST_TEMP, 
        MHS_PROD.PREHT_ZON_CTRL_FLG, 
        MHS_PROD.CHRG_ZON_CTRL_FLG, 
        MHS_PROD.INTERM_ZON_CTRL_FLG, 
        MHS_PROD.SOAK_ZON_CTRL_FLG, 
        MHS_PROD.SLAB_CHRG_DIST, 
        MHS_PROD.AIM_CHRG_GAP_DIST, 
        MHS_PROD.CHRG_GAP_DIST, 
        MHS_PROD.AIM_PREHT_ZON_WORK_TEMP, 
        MHS_PROD.AIM_CHRG_ZON_WORK_TEMP, 
        MHS_PROD.AIM_INTERM_ZON_WORK_TEMP, 
        MHS_PROD.AIM_SOAK_ZON_WORK_TEMP, 
        MHS_PROD.PACE_READY_DUR, 
        MHS_PROD.PACE_READY_OFFST_DUR, 
        MHS_PROD.HT_READY_OFFST_DUR, 
        MHS_PROD2.ORIGNL_PACE_READY_DUR, 
        MHS_PROD2.TEMP_READY_OFFST_DUR, 
        MHS_PROD.PREHT_ZON_PACE_FLG, 
        MHS_PROD.CHRG_ZON_PACE_FLG, 
        MHS_PROD.INTERM_ZON_PACE_FLG, 
        MHS_PROD.SOAK_ZON_PACE_FLG, 
        MHS_PROD.PREHT_ZON_PACE_LIM_CNT, 
        MHS_PROD.CHRG_ZON_PACE_LIM_CNT, 
        MHS_PROD.INTERM_ZON_PACE_LIM_CNT, 
        MHS_PROD.SOAK_ZON_PACE_LIM_CNT, 
        MHS_PROD.CHRG_PERMISV_FAIL_FLG, 
        MHS_PROD.PUR_SLAB_HT_PCE_NO, 
        MHS_HEATING_PRACTICE.AIM_PREHT_ZON_SLAB_TEMP, 
        MHS_HEATING_PRACTICE.AIM_CHRG_ZON_SLAB_TEMP, 
        MHS_HEATING_PRACTICE.AIM_INTERM_ZON_SLAB_TEMP, 
        MHS_HEATING_PRACTICE.AIM_SOAK_ZON_SLAB_TEMP, 
        MHS_PROD.PRED_EWK_PASS_CNT, 
        MHS_PROD.PRED_RGH_PASS_CNT, 
        HMILL_AFE.COURSE_0_PRED_EWK_PASS_CNT, 
        HMILL_AFE.COURSE_0_PRED_RGH_PASS_CNT, 
        HMILL_RM_PROD.OP_RM_PASS_NO, 
        HMILL_QMS_PROD.EWK_PASS_CNT, 
        HMILL_QMS_PROD.RM_PASS_CNT, 
        HMILL_AFE.RM_LAST_PASS_DLY_REASON_CD, 
        MHS_PROD.RECHRG_FLG, 
        MHS_PROD.GRAVEL_ON_SLAB_FLG, 
        HMILL_PROD.SLAB_DISP_CD, 
        HMILL_PROD.SLAB_THEOR_DENS, 
        HMILL_PROD.SLAB_THEOR_DENS_WT, 
        MHS_PROD.SLAB_DENS, 
        HMILL_PROD.HMILL_SLAB_WT, 
        HMILL_PCE.HMILL_COIL_WT, 
        HMILL_PCE.COIL_WT, 
        HMILL_FM_PROD.THEOR_WT_BY_YLD, 
        HMILL_FM_PROD.THEOR_WT_BY_VOL, 
        HMILL_PCE.HMILL_ENT_WT, 
        HMILL_PCE.HMILL_EXIT_WT, 
        HMILL_FM_ACS.ACS_LEAD_CROP_WT, 
        HMILL_FM_ACS.ACS_TRAIL_CROP_WT, 
        MHS_PROD.CLIP_SLAB_FLG, 
        MHS_PROD.CONFIRM_CLIP_SLAB_FLG, 
        HMILL_FSU_PRED.AUTO_ALTRN_THK_DSCALE_DSEL_FLG, 
        HMILL_FSU_PRED.AUTO_ALTRN_THK_USE_FLG, 
        HMILL_FSU_PRED.OP_ALTRN_THK_USE_FLG, 
        HMILL_FSU_PRED.AUTO_ALTRN_THK_STATUS_CD, 
        HMILL_PCE.CUST_NAME, 
        HMILL_PROD.NUMERIC_ORD_HMILL_NO, 
        HMILL_PROD.NEXT_OP_CD, 
        MHS_SLAB_TEMP.TOP_PROFL_SLAB_TEMP, 
        MHS_SLAB_TEMP.CENTER_PROFL_SLAB_TEMP, 
        MHS_SLAB_TEMP.BOT_PROFL_SLAB_TEMP, 
        MHS_SLAB_TEMP.AVG_PROFL_SLAB_TEMP, 
        MHS_SLAB_TEMP.SLAB_PACE_RATE, 
        HMILL_RM_FM.HMSTC_ACPT_ZON_CALC_REQ_FLG, 
        HMILL_RM_FM.HMSTC_ACPT_ZON_PRED_SLAB_TEMP, 
        HMILL_RM_FM.HMSTC_PRED_FME_TEMP, 
        HMILL_RM_FM.HMSTC_PRED_FME_TEMP_CORR_VALUE, 
        HMILL_RM_FM.HMSTC_BEST_OPTION_COST, 
        HMILL_RM_FM.HMSTC_ACTION_CD, 
        HMILL_RM_FM.HMSTC_HT_READY_TIM_OFFST_DUR, 
        MHS_PROD2.CHRG_CALC_RM_PROCESS_DUR, 
        MHS_PROD2.CHRG_CALC_FM_PROCESS_DUR, 
        MHS_PROD2.LOW_SETPT_LIM_FLG, 
        MHS_PROD2.PREH_ZON_HI_SETPT_LIM, 
        MHS_PROD2.TARGT_RM_GAP_DUR, 
        MHS_PROD2.TARGT_FM_GAP_DUR, 
        MHS_PROD2.FUR_LGT_LOW_SETPT, 
        MHS_PROD2.MIN_LOW_SETPT_DUR, 
        MHS_PROD2.TOT_AVAIL_HT_DUR, 
        MHS_TEMP_FBK.PRED_AMT_FUR_RME_DROPT_TEMP, 
        MHS_TEMP_FBK.PRED_AMT_RME_RMX_DROPT_TEMP, 
        MHS_TEMP_FBK.PRED_AMT_RMX_FME_DROPT_TEMP, 
        MHS_TEMP_FBK.RM_PRED_FUR_RME_DELTA_TEMP, 
        MHS_TEMP_FBK.RM_PRED_RME_RMX_DELTA_TEMP, 
        MHS_TEMP_FBK.RM_PRED_RMX_FMA_DELTA_TEMP, 
        MHS_TEMP_FBK.ACT_FUR_RME_DELTA_TEMP, 
        MHS_TEMP_FBK.ACT_RME_RMX_DELTA_TEMP, 
        MHS_TEMP_FBK.ACT_RMX_FMA_DELTA_TEMP, 
        MHS_PROD.NORM_HT_PRIORITY_CD, 
        MHS_PROD.SPECIAL_HT_PRIORITY_CD, 
        MHS_PROD.SPECIAL_PRAC_FLG, 
        MHS_PROD.SPECIAL_PRAC_CD, 
        HMILL_RM_PHYS.RMX_LEAD_SCAN_TEMP, 
        HMILL_RM_PHYS.RMX_LEAD_FME_SCAN_TEMP, 
        HMILL_RM_PHYS.EXIT_BODY_TEMP, 
        MHS_TEMP_FBK.APPL_ACT_COURSE_2_OFFST_TEMP, 
        MHS_TEMP_FBK.APPL_FBK_OFFST_TEMP, 
        MHS_TEMP_FBK.APPL_FMA_ERR_OFFST_TEMP,
        MHS_TEMP_FBK.APPL_FUR_ERR_OFFST_TEMP,
        MHS_TEMP_FBK.APPL_FUR_POS_EXTR_TEMP_AIM_ADJ, 
        MHS_TEMP_FBK.APPL_LEAD_FMA_OFFST_TEMP, 
        MHS_TEMP_FBK.APPL_OP_OFFST_TEMP, 
        MHS_TEMP_FBK.APPL_RM_CTRL_EFCT_TEMP, 
        MHS_TEMP_FBK.APPL_RM_MODEL_OFFST_TEMP, 
        HMILL_QMS_TREND.FM_ARR_AVG_LEAD_TEMP, 
        HMILL_QMS_TREND.AIM_FM_ARR_AVG_LEAD_TEMP
        FROM WIPHM2_PRD.HMILL_AFE HMILL_AFE, 
        WIPHM2_PRD.HMILL_FM_ACS HMILL_FM_ACS, 
        WIPHM2_PRD.HMILL_FM_ENT_PDI HMILL_FM_ENT_PDI, 
        WIPHM2_PRD.HMILL_FM_MODEL HMILL_FM_MODEL, 
        WIPHM2_PRD.HMILL_FM_PROD HMILL_FM_PROD, 
        WIPHM2_PRD.HMILL_FSU_PRED HMILL_FSU_PRED, 
        WIPHM2_PRD.HMILL_PCE HMILL_PCE, 
        WIPHM2_PRD.HMILL_PROD HMILL_PROD, 
        WIPHM2_PRD.HMILL_QMS_PROD HMILL_QMS_PROD, 
        WIPHM2_PRD.HMILL_QMS_TREND HMILL_QMS_TREND, 
        WIPHM2_PRD.HMILL_RM_FM HMILL_RM_FM, 
        WIPHM2_PRD.HMILL_RM_PHYS HMILL_RM_PHYS, 
        WIPHM2_PRD.HMILL_RM_PROD HMILL_RM_PROD, 
        WIPHM2_PRD.HMILL_RM_PROD2 HMILL_RM_PROD2, 
        MHS_PRD.MHS_HEATING_PRACTICE MHS_HEATING_PRACTICE,
        MHS_PRD.MHS_PROD MHS_PROD, 
        MHS_PRD.MHS_PROD2 MHS_PROD2, 
        MHS_PRD.MHS_SLAB_TEMP MHS_SLAB_TEMP, 
        MHS_PRD.MHS_TEMP_FBK MHS_TEMP_FBK      
        WHERE 
        MHS_PROD.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_PROD.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_AFE.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_FSU_PRED.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_FM_ENT_PDI.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_RM_FM.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_RM_PHYS.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND MHS_TEMP_FBK.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND MHS_HEATING_PRACTICE.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_RM_PROD.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_FM_MODEL.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_QMS_PROD.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_FM_PROD.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_FM_ACS.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_RM_PROD2.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND MHS_SLAB_TEMP.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND MHS_PROD2.HMILL_PCE_NO = HMILL_PCE.HMILL_PCE_NO 
        AND HMILL_PCE.HMILL_PCE_NO = HMILL_QMS_TREND.HMILL_PCE_NO 
        AND ((HMILL_PCE.GEN_EST>to_date('{0}', 'yyyy-mm-dd')
        AND HMILL_PCE.GEN_EST<to_date('{1}', 'yyyy-mm-dd')) 
        AND HMILL_PCE.PAGE_NO = {2}
        AND HMILL_PCE.ITEM_NO = {3}
        AND (MHS_HEATING_PRACTICE.PROD_RATE_NO=1) AND (MHS_SLAB_TEMP.FUR_POS_CD=28))
        ORDER BY HMILL_PCE.GEN_EST asc;
        """
        
        self.comb_query_string = """
        SELECT 
        MHS_FUR_COMB_ZON.FILE_EST, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,1,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_1, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,2,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_2, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,3,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_3, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,4,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_4, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,5,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_5, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,6,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_6, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,7,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_7, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,8,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_8, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,9,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_9, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,10,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_10, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,11,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_11, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,12,MHS_FUR_COMB_ZON.AVG_ZON_SETPT_TEMP,0))/4) AS SETPT_TEMP_12, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,1,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_1, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,2,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_2, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,3,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_3, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,4,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_4, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,5,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_5, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,6,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_6, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,7,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_7, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,8,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_8, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,9,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_9, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,10,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_10, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,11,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_11, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,12,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCA_TEMP,0))/4) AS TCA_12, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,1,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_1, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,2,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_2, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,3,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_3, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,4,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_4, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,5,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_5, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,6,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_6, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,7,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_7, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,8,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_8, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,9,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_9, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,10,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_10, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,11,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_11, 
        (Sum(decode(MHS_FUR_COMB_ZON.FUR_COMB_ZON_NO,12,MHS_FUR_COMB_ZON.AVG_FUR_CTRL_TCB_TEMP,0))/4) AS TCB_12, 
        (Sum(decode(MHS_FUR_CTRL_ZON.FUR_CTRL_ZON_NO,1,decode(MHS_FUR_CTRL_ZON.AUTO_CTRL_SEL_FLG,'Y',1,0)))/12) AS AUTO_CTRL_PRE, 
        (Sum(decode(MHS_FUR_CTRL_ZON.FUR_CTRL_ZON_NO,2,decode(MHS_FUR_CTRL_ZON.AUTO_CTRL_SEL_FLG,'Y',1,0)))/12) AS AUTO_CTRL_CHRG, 
        (Sum(decode(MHS_FUR_CTRL_ZON.FUR_CTRL_ZON_NO,3,decode(MHS_FUR_CTRL_ZON.AUTO_CTRL_SEL_FLG,'Y',1,0)))/12) AS AUTO_CTRL_INT, 
        (Sum(decode(MHS_FUR_CTRL_ZON.FUR_CTRL_ZON_NO,4,decode(MHS_FUR_CTRL_ZON.AUTO_CTRL_SEL_FLG,'Y',1,0)))/12) AS AUTO_CTRL_SOAK, 
        Avg(MHS_FUR.AVG_FLUE_GAS_TC_1_TEMP) AS AVG_FLUE_GAS_TC_1_TEMP, 
        Avg(MHS_FUR.AVG_FLUE_GAS_TC_2_TEMP) AS AVG_FLUE_GAS_TC_2_TEMP, 
        Avg(MHS_FUR.MAX_FLUE_GAS_TC_1_TEMP) AS MAX_FLUE_GAS_TC_1_TEMP, 
        Avg(MHS_FUR.MAX_FLUE_GAS_TC_2_TEMP) AS MAX_FLUE_GAS_TC_2_TEMP, 
        Avg(MHS_FUR.MIN_FLUE_GAS_TC_1_TEMP) AS MIN_FLUE_GAS_TC_1_TEMP, 
        Avg(MHS_FUR.MIN_FLUE_GAS_TC_2_TEMP) AS MIN_FLUE_GAS_TC_2_TEMP, 
        AVG(MHS_FUR.AVG_PREHT_CENTER_FLOOR_TEMP) AS AVG_PREHT_CENTER_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_PREHT_CENTER_ROOF_TEMP) AS AVG_PREHT_CENTER_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_CHRG_CENTER_FLOOR_TEMP) AS AVG_CHRG_CENTER_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_CHRG_CENTER_ROOF_TEMP) AS AVG_CHRG_CENTER_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_CHRG_NORTH_ROOF_TEMP) AS AVG_CHRG_NORTH_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_CHRG_SOUTH_ROOF_TEMP) AS AVG_CHRG_SOUTH_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_CENTER_E_ROOF_TEMP) AS AVG_INTERM_CENTER_E_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_CENTER_W_ROOF_TEMP) AS AVG_INTERM_CENTER_W_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_NORTH_FLOOR_2_TEMP) AS AVG_INTERM_NORTH_FLOOR_2_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_NORTH_FLOOR_TEMP) AS AVG_INTERM_NORTH_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_NORTH_ROOF_TEMP) AS AVG_INTERM_NORTH_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_SOUTH_FLOOR_2_TEMP) AS AVG_INTERM_SOUTH_FLOOR_2_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_SOUTH_FLOOR_TEMP) AS AVG_INTERM_SOUTH_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_INTERM_SOUTH_ROOF_TEMP) AS AVG_INTERM_SOUTH_ROOF_TEMP, 
        AVG(MHS_FUR.AVG_SOAK_NORTH_FLOOR_TEMP) AS AVG_SOAK_NORTH_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_SOAK_SOUTH_FLOOR_TEMP) AS AVG_SOAK_SOUTH_FLOOR_TEMP, 
        AVG(MHS_FUR.AVG_SOAK_CENTER_ROOF_TEMP) AS AVG_SOAK_CENTER_ROOF_TEMP, 
        AVG(MHS_FUR.AIM_PACE_VEL) AS AIM_PACE_VEL, 
        AVG(MHS_FUR.AVG_PACE_VEL) AS AVG_PACE_VEL, 
        AVG(MHS_FUR.LV1_PACE_VEL) AS LV1_PACE_VEL
        FROM MHS_PRD.MHS_FUR MHS_FUR, 
        MHS_PRD.MHS_FUR_COMB_ZON MHS_FUR_COMB_ZON, 
        MHS_PRD.MHS_FUR_CTRL_ZON MHS_FUR_CTRL_ZON
        WHERE MHS_FUR_COMB_ZON.FILE_EST = MHS_FUR_CTRL_ZON.FILE_EST 
        AND MHS_FUR_COMB_ZON.FUR_NO = MHS_FUR_CTRL_ZON.FUR_NO 
        AND MHS_FUR_COMB_ZON.FILE_EST = MHS_FUR.FILE_EST 
        AND MHS_FUR_COMB_ZON.FUR_NO = MHS_FUR.FUR_NO 
        AND ((MHS_FUR_COMB_ZON.FILE_EST BETWEEN TO_DATE('{0}', 'yyyy-mm-dd HH24:MI:SS') And TO_DATE('{1}', 'yyyy-mm-dd HH24:MI:SS')) 
        AND (MHS_FUR_CTRL_ZON.FUR_NO={2}))
        GROUP BY MHS_FUR_COMB_ZON.FILE_EST
        ORDER BY MHS_FUR_COMB_ZON.FILE_EST;
        """

        self.temp_scan_query_string = """
        select * from HMILL_PCE_QUAL_SUMR
        where  
        SCAN_NAME LIKE '%RMX TEMPERATURE - MAX A or B%' AND 
        HMILL_PCE_NO = {0}
        ;
        """
        
        self.piece_data = self._query_data(self.piece_query_string.format(self.start_date, 
                self.end_date, self.page, self.item))
        self.comb_data = self._query_data(self.comb_query_string.format( self.piece_data['SLAB_CHRG_EST'][0].strftime("%Y-%m-%d %H:%M:%S"), 
                self.piece_data['SLAB_EXTR_EST'][0].strftime("%Y-%m-%d %H:%M:%S"), self.piece_data['FUR_NO'][0]))
        self.scan_data = self._query_data(self.temp_scan_query_string.format(self.piece_data['HMILL_PCE_NO'][0]))
        
        self._preheat_idx = (self.comb_data['FILE_EST'] > self.piece_data['SLAB_CHRG_EST'][0]) & (self.comb_data['FILE_EST'] < self.piece_data['PREHT_ZON_EXIT_EST'][0] + timedelta(minutes = 5))
        self._charge_idx = (self.comb_data['FILE_EST'] > self.piece_data['PREHT_ZON_EXIT_EST'][0] - timedelta(minutes = 5)) & (self.comb_data['FILE_EST'] < self.piece_data['CHRG_ZON_EXIT_EST'][0] + timedelta(minutes = 5))
        self._inter_idx = (self.comb_data['FILE_EST'] > self.piece_data['CHRG_ZON_EXIT_EST'][0] - timedelta(minutes = 5)) & (self.comb_data['FILE_EST'] < self.piece_data['INTERM_ZON_EXIT_EST'][0] + timedelta(minutes = 5))
        self._soak_idx = (self.comb_data['FILE_EST'] > self.piece_data['INTERM_ZON_EXIT_EST'][0]) & (self.comb_data['FILE_EST'] < self.piece_data['SLAB_EXTR_EST'][0])
        self.directory = 'C:\\TEMP\\' + str(self.page) + '_' + str(self.item)

        self.do_analysis()

    def _query_data(self, sql_statement):
            try:
                db_conn = pyodbc.connect(DSN='DSSA', uid=self.__user, pwd=self.__password)
                query_results = pd.read_sql(sql_statement, db_conn)
            except pyodbc.Error as err:
                print(err)
            except:
                print("Unexpected error:", sys.exc_info()[0]) 
            finally:
                db_conn.close()
            
            return query_results
    
    def head_body_temp(self):
        head_avg = 0
        for i in range(1, 21):
            head_avg += self.scan_data['SECT_' + str(i) + '_VALUE'][0]
        head_avg = head_avg/20

        body_avg = 0
        for i in range(21, 81):
            body_avg += self.scan_data['SECT_' + str(i) + '_VALUE'][0]
        body_avg = body_avg/60

        head_body_temp_diff = head_avg - body_avg
        if head_body_temp_diff > 20:
            self.head_body_temp_bias = True

        self.email_text += "Head end avg. temp: {0}, body avg. temp: {1}, and head-body temp diff.: {2}. ".format(head_avg, body_avg, head_body_temp_diff)
        if self.head_body_temp_bias:
            self.email_text += "Please check video for gravel/scale and combustion zone north/south temperature. "
        
    def pass_schedule(self):
        if (self.piece_data['PRED_EWK_PASS_CNT'][0] == self.piece_data['EWK_PASS_CNT'][0]) and (self.piece_data['PRED_RGH_PASS_CNT'][0] == self.piece_data['RM_PASS_CNT'][0]):
            self.pass_schedule_change = False
            self.email_text += "No RM pass schedule change. <br />"
        else:
            self.pass_schedule_change = True
            self.email_text += "RM pass schedule changed from {0}-{1} to {2}-{3}. ".format(int(self.piece_data['PRED_EWK_PASS_CNT'][0]), 
            int(self.piece_data['PRED_RGH_PASS_CNT'][0]), int(self.piece_data['EWK_PASS_CNT'][0]), int(self.piece_data['RM_PASS_CNT'][0]))
            
            if (self.piece_data['RM_PASS_CNT'][0] + self.piece_data['EWK_PASS_CNT'][0] - self.piece_data['PRED_EWK_PASS_CNT'][0] - self.piece_data['PRED_RGH_PASS_CNT'][0]) > 0:
                self.email_text += "Extra RM passes made this piece lose more temperature. <br />"

    def fmet_tolerance(self):
        if self.piece_data['FM_FAM_NO'][0] == 2:
            self.fmet_tight = True
            self.email_text += "FMET tolerance is tighter because of FM Family 2. <br />"
        elif self.piece_data['GRT_INDEX_NO'][0] <= 4:
            self.fmet_tight = True
            self.email_text += "FMET tolerance is tighter because of GRT <= 4. <br />"
        else:
            self.email_text += "FMET tolernace is normal. <br />"

    def model_dot_err(self):
        self.email_text += "DOT aim is {0}, predicted DOT is {1}, and back calculated DOT is {2}. ".format(self.piece_data['INTGRT_AIM_FUR_EXTR_TEMP'][0], 
            self.piece_data['CALC_SLAB_EXTR_TEMP'][0], self.piece_data['RM_BCK_CALC_FUR_EXTR_TEMP'][0])
        if self.piece_data['TEMP_ERR'][0] < -30.0:
            self.cold_from_furnace = True
            self.email_text += "This slab is cold from furnace. Temperature error is {0}. ".format(self.piece_data['TEMP_ERR'][0])
            if self.piece_data['MODEL_ERR'][0] < -25:
                self.big_model_err = True
                self.email_text += "Model error is {0}".format(self.piece_data['MODEL_ERR'][0])
        else:
            self.email_text += "Based on back calculated temp, this piece is not too cold compared to DOT aim. <br />"
            self.email_text += "FMET aim is {0}, and initial determined FMET is {1}. It's {2} degrees below the FMET aim. ".format(self.piece_data['AIM_ENT_TEMP'][0], self.piece_data['MEAS_FME_TEMP'][0], self.piece_data['AIM_ENT_TEMP'][0] - self.piece_data['MEAS_FME_TEMP'][0])
            if self.piece_data['MEAS_FME_TEMP'][0] - self.piece_data['AIM_ENT_TEMP'][0] > -40:
                self.fmet_tolerance()
                if not self.fmet_tight:
                    self.email_text += "Please check TAS_FME log for details. <br />"
        
        self.pass_schedule()
        self.head_body_temp()
        

    def do_analysis(self):    
        if not os.path.exists(self.directory):
            os.makedirs(self.directory)
            
        self.plot_temp_scan()
        self.plot_comb_zone()
        self.plot_heating_history()
        self.model_dot_err()

    def plot_temp_scan(self):
        temp_scan = []
        for i in range(1, 201):
            temp_scan.append(self.scan_data['SECT_' + str(i) + '_VALUE'])
        
        fig, ax = plt.subplots()
        ax.plot(temp_scan)
        ax.set_xlabel('Scan Point')
        ax.set_ylabel('RMX Temperature')

        fig.savefig(self.directory + '\\RMXT.png')

    def plot_comb_zone(self):
        fig, axes = plt.subplots(3, 4, sharey='row', figsize=(19, 9.5))
        for i in range(1, 13):
            if i >= 1 and i <= 2 :
                idx = self._preheat_idx
            elif i >= 3 and i <= 4 :
                idx = self._charge_idx
            elif i >= 5 and i <= 8 :
                idx = self._inter_idx
            else:
                idx = self._soak_idx
            row = (i-1)//4
            col = (i-1) % 4
            axes[row][col].plot(self.comb_data.loc[idx, ['FILE_EST']], self.comb_data.loc[idx, ['SETPT_TEMP_' + str(i)]], label='SP')
            axes[row][col].plot(self.comb_data.loc[idx, ['FILE_EST']], self.comb_data.loc[idx, ['TCA_' + str(i)]], '-.', label='TCA')
            axes[row][col].plot(self.comb_data.loc[idx, ['FILE_EST']], self.comb_data.loc[idx, ['TCB_' + str(i)]], '-.', label='TCB')
            axes[row][col].xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%H:%M'))
            axes[row][col].fmt_xdata =  matplotlib.dates.DateFormatter('%H:%M')
            axes[row][col].grid(linestyle='--')
            axes[row][col].text(0.5, 0.5, str(i), fontsize = 20, horizontalalignment='center', 
                verticalalignment='center', transform=axes[row][col].transAxes, alpha=0.2)
            if(row == 1 and col == 1):
                axes[row][col].legend()
        
        plt.tight_layout()
        fig.savefig(self.directory + '\\comb_zone.png')

    def plot_heating_history(self):
        fig, ax = plt.subplots(figsize=(14, 8))
        # Peheat zone
        ln1 = ax.plot(self.comb_data.loc[self._preheat_idx, 'FILE_EST'], 
        self.comb_data.loc[self._preheat_idx, ['SETPT_TEMP_1', 'SETPT_TEMP_2']].mean(axis = 1), 
        color='#1f77b4', label='SP')
        ln2 = ax.plot(self.comb_data.loc[self._preheat_idx, ['FILE_EST']], 
        self.comb_data.loc[self._preheat_idx, ['TCA_1', 'TCA_2']].mean(axis = 1), 
        '-.', color='#ff7f0e', label='TCA')
        ln3 = ax.plot(self.comb_data.loc[self._preheat_idx, ['FILE_EST']], 
        self.comb_data.loc[self._preheat_idx, ['TCB_1', 'TCB_2']].mean(axis = 1), 
        '-.', color='#2ca02c', label='TCB')
        
        # Charge zone
        ax.plot(self.comb_data.loc[self._charge_idx, ['FILE_EST']], 
        self.comb_data.loc[self._charge_idx, ['SETPT_TEMP_3', 'SETPT_TEMP_4']].mean(axis = 1),
        color='#1f77b4')
        ax.plot(self.comb_data.loc[self._charge_idx, ['FILE_EST']], 
        self.comb_data.loc[self._charge_idx, ['TCA_3', 'TCA_4']].mean(axis = 1),
        '-.', color='#ff7f0e')
        ax.plot(self.comb_data.loc[self._charge_idx, ['FILE_EST']], 
        self.comb_data.loc[self._charge_idx, ['TCB_3', 'TCB_4']].mean(axis = 1),
        '-.', color='#2ca02c')
        
        # Intermediate zone
        ax.plot(self.comb_data.loc[self._inter_idx, ['FILE_EST']], 
        self.comb_data.loc[self._inter_idx, ['SETPT_TEMP_5', 'SETPT_TEMP_6', 
                                             'SETPT_TEMP_7', 'SETPT_TEMP_8']].mean(axis = 1), 
                                             color='#1f77b4')
        ax.plot(self.comb_data.loc[self._inter_idx, ['FILE_EST']], 
        self.comb_data.loc[self._inter_idx, ['TCA_5', 'TCA_6', 
                                             'TCA_7', 'TCA_8']].mean(axis = 1),
                                             '-.', color='#ff7f0e')
        ax.plot(self.comb_data.loc[self._inter_idx, ['FILE_EST']], 
        self.comb_data.loc[self._inter_idx, ['TCB_5', 'TCB_6', 
                                             'TCB_7', 'TCB_8']].mean(axis = 1), 
                                             '-.', color='#2ca02c')
        
        # Soak zone
        ax.plot(self.comb_data.loc[self._soak_idx, ['FILE_EST']], 
        self.comb_data.loc[self._soak_idx, ['SETPT_TEMP_9', 'SETPT_TEMP_10', 
                                            'SETPT_TEMP_11', 'SETPT_TEMP_12']].mean(axis = 1),
                                            color='#1f77b4')
        ax.plot(self.comb_data.loc[self._soak_idx, ['FILE_EST']], 
        self.comb_data.loc[self._soak_idx, ['TCA_9', 'TCA_10', 
                                            'TCA_11', 'TCA_12']].mean(axis = 1),
                                            '-.', color='#ff7f0e')
        ax.plot(self.comb_data.loc[self._soak_idx, ['FILE_EST']], 
        self.comb_data.loc[self._soak_idx, ['TCB_9', 'TCB_10', 
                                            'TCB_11', 'TCB_12']].mean(axis = 1),
                                            '-.', color='#2ca02c')

        ax.plot([self.piece_data['PREHT_ZON_EXIT_EST'][0], self.piece_data['PREHT_ZON_EXIT_EST'][0]], [1100, 1400], 'r:', alpha=0.5)
        ax.plot([self.piece_data['CHRG_ZON_EXIT_EST'][0], self.piece_data['CHRG_ZON_EXIT_EST'][0]], [1100, 1400], 'r:', alpha=0.5)
        ax.plot([self.piece_data['INTERM_ZON_EXIT_EST'][0], self.piece_data['INTERM_ZON_EXIT_EST'][0]], [1100, 1400], 'r:', alpha=0.5)
        ax.set_ylim(1100, 1400)
        ax.xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%H:%M'))
        ax.fmt_xdata =  matplotlib.dates.DateFormatter('%H:%M')
        ax.set_ylabel('Temperature/C')
        ax.grid( axis='y', linestyle='--')
        
        ax_s = ax.twinx()
        ln4 = ax_s.plot(self.comb_data['FILE_EST'], self.comb_data['AVG_PACE_VEL'], 
                  ':', color='#9467bd', alpha=0.9, label='Pace')
        ax_s.xaxis.set_major_formatter(matplotlib.dates.DateFormatter('%H:%M'))
        ax_s.fmt_xdata =  matplotlib.dates.DateFormatter('%H:%M')
        ax_s.set_ylabel('Pace m/h')
        
        lns = ln1 + ln2 + ln3 + ln4
        labs = [l.get_label() for l in lns]
        ax.legend(lns, labs, loc=0)
        
        plt.tight_layout()
        fig.savefig(self.directory + '\\heat_his.png')
        
    def send_email(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'kuilin.chen@arcelormittal.com'
        mail.Subject = str(self.page) + '-' + str(self.item) + ' cold'
        attachment1 = mail.Attachments.Add(self.directory + '\\heat_his.png')
        attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
        attachment2 = mail.Attachments.Add(self.directory + '\\comb_zone.png')
        attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId2")
        attachment3 = mail.Attachments.Add(self.directory + '\\RMXT.png')
        attachment3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId3")
        mail.HTMLBody = "<html><body>{0} <br /><br /><br /><img src=""cid:MyId1""><br /><img src=""cid:MyId2""><br /><img src=""cid:MyId3""><br /></body></html>".format(self.email_text)
        mail.send
       
       

if __name__ == "__main__":
    # Close all figures
    plt.close('all')
    
    cb1 = ColdBar(963, 13, '2018-05-18', '2018-05-20')
    
    #cb1.send_email()
