/*****************************************************************************
 * Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd.
 *
 * All rights reserved.
 * ****************************************************************************
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * - Redistributions of source code must retain the above copyright notice,
 * this list of conditions and the disclaimer below.
 *
 * Nations' name may not be used to endorse or promote products derived from
 * this software without specific prior written permission.
 *
 * DISCLAIMER: THIS SOFTWARE IS PROVIDED BY NATIONS "AS IS" AND ANY EXPRESS OR
 * IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT ARE
 * DISCLAIMED. IN NO EVENT SHALL NATIONS BE LIABLE FOR ANY DIRECT, INDIRECT,
 * INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
 * LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA,
 * OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
 * LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
 * EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 * ****************************************************************************/

/**
 * @file ppx_factory.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_FACTORY_H__
#define __PPX_FACTORY_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>
#include "ppx_packet.h"


/* define ppx factory max data length */
#define PPX_FACTORY_DATA_SIZE           ( 128 )


/* define enum, ppx factory response type */
typedef enum
{
    FACTORY_RSP_FAILED      = 0,
    FACTORY_RSP_SUCCESS     = 1,
} factory_rsp_status_t;


/* define enum, ppx imu cali response type */
typedef enum
{
    PPX_IMU_CALI_FAILED       = 0,
    PPX_IMU_CALI_SUCCESS      = 1,
    PPX_IMU_CALI_START        = 2,
    PPX_IMU_CALI_RUNING       = 3,
    PPX_IMU_CALI_SIDE         = 0x80,
} factory_imu_cali_status_t;


/* define enum ppx factory msg type */
typedef enum
{
    FACTORY_RVSD_TYPE       = 0x80,
    FACTORY_MODE_SET_REQ    = 0x81,
    FACTORY_MODE_SET_RSP    = 0x82,
    FACTORY_CCB_SET_REQ     = 0x83,
    FACTORY_CCB_SET_RSP     = 0x84,
    FACTORY_CCB_GET_REQ     = 0x85,
    FACTORY_CCB_GET_RSP     = 0x86,
    FACTORY_MCB_SET_REQ     = 0x87,
    FACTORY_MCB_SET_RSP     = 0x88,
    FACTORY_MCB_GET_REQ     = 0x89,
    FACTORY_MCB_GET_RSP     = 0x8A,
    FACTORY_IMU_CALI_REQ    = 0x8B,
    FACTORY_IMU_CALI_RSP    = 0x8C,
    FACTORY_RESET_REQ       = 0x8D,
    FACTORY_RESET_RSP       = 0x8E,
    FACTORY_PARAMS_SET_REQ  = 0x90,
    FACTORY_PARAMS_SET_RSP  = 0x91,
    FACTORY_VIN_SET_REQ     = 0x92,
    FACTORY_VIN_SET_RSP     = 0x93,
    FACTORY_IOT_GET_REQ     = 0x94,
    FACTORY_IOT_GET_RSP     = 0x95,
    FACTORY_SYS_SLEEP_REQ   = 0x96,
    FACTORY_SYS_SLEEP_RSP   = 0x97,
} factory_msg_type_t;


/* define enum factory imu req type */
typedef enum
{
    IMU_CALI_START = 1,
    IMU_CALI_QUERY = 2,
    ACC_CALI_START = 3,
    ACC_CALI_QUERY = 4,
} factory_imu_cali_type_t;


/* define struct of ppx factory ccb set msg */
typedef struct
{
    uint8_t led_enable;
    uint8_t spk_enable;
    uint8_t light_enable;
    uint8_t sn_write;
    uint8_t serial_num[26];
    uint8_t adc_vref_enable;
    uint8_t charge_enable;
} factory_ccb_req_msg_t;


/* define struct of ppx factory ccb get msg */
typedef struct
{
    uint16_t  batt_voltage;
    uint16_t  dc_voltage;
    uint16_t  dc_current;
    uint16_t  adc_ext_vref;
    uint16_t  handle_bar_val;
    int8_t    key_spk_status;
    int8_t    key_light_status;
    int8_t    key_sos_status;
    int8_t    gear_status;
    int8_t    rs485_status;
    int8_t    handle_bar_status;
    int8_t    sif_status;
    uint8_t   hw_version;
    uint8_t   sw_version[PPX_SW_VER_SIZE];
    uint8_t   serial_num[PPX_SN_SIZE];
    uint8_t   vin_serial_num[PPX_SN_SIZE];
} factory_ccb_rsp_msg_t;


/* define struct of ppx factory mcb set msg */
typedef struct
{
    int16_t speed;
    uint8_t gear;
    uint8_t brake_enable;
    uint8_t reboot;
    uint8_t power_off;
    uint8_t sn_write;
    uint8_t serial_num[PPX_SN_SIZE];
} factory_mcb_req_msg_t;


/* define struct of ppx factory mcb get msg */
typedef struct
{
    int32_t  motor_angle;
    int16_t  speed;
    uint16_t bus_voltage;      /* 0.1V */
    uint16_t bus_current;      /* 0.1A */
    int16_t  angular_speed;
    int16_t  pi_vq;
    int16_t  pi_iq;
    int16_t  phase_current_a;  /* 0.1A */
    int16_t  phase_current_b;  /* 0.1A */
    int16_t  phase_current_c;  /* 0.1A */
    int16_t  imu_pitch;        /* 0.1�� */
    int16_t  imu_roll;         /* 0.1�� */
    uint8_t  imu_acc;          /* 0.01g */
    uint8_t  gear;
    int8_t   rs485_status;
    int8_t   seat_status;
    int8_t   hall_status;
    int8_t   brake_status;
    int8_t   imu_status;
    uint8_t  hw_version;
    uint8_t  sw_version[PPX_SW_VER_SIZE];
    uint8_t  serial_num[PPX_SN_SIZE];
} factory_mcb_rsp_msg_t;


/* define struct of ppx factory iot get msg */
typedef struct
{
    char        imei[16];
    char        imsi[16];
    char        iccid[22];
    uint8_t     sim_card;      // SIM Card status, insert
    uint8_t     reg_state;     // network registration state
    uint8_t     pdp_act;       // network pdpactive state
    int         mcc;           // mobile country code
    int         mnc;           // mobile network code
    int         rssi;          // signal strength
    int         lac;           // local area
    int         cid;           // cell identity
    int         act;           // 1 - GSM, 2 - CDMA, 3 - WCDMA, 4 - TD_SCDMA, 5 - LTE

    uint8_t     gnss_state;    // gps state
    uint16_t    satellites;    // ������
    float       altitude;      // ����
    float       latitude;      // γ��
    float       longitude;     // ����
    float       cog;           // ����
    float       gps_speed;     // gps�ٶ�
} factory_iot_rsp_msg_t;


/* define struct of ppx factory rsp msg */
typedef struct
{
    uint8_t rsp_status;     /* refer factory_rsp_status */
    uint8_t imu_status;     /* refer factory_imu_cali_status */
    
    factory_ccb_rsp_msg_t   ccb_rsp_msg;    /* refer factory_ccb_rsp_msg_t */
    factory_mcb_rsp_msg_t   mcb_rsp_msg;    /* refer factory_mcb_rsp_msg_t */
    factory_iot_rsp_msg_t   iot_rsp_msg;    /* refer factory_iot_rsp_msg_t */
} factory_rsp_msg_t;


/* define struct of ppx factory ccb params msg */
typedef struct
{
    uint8_t  product_type;      /* product type: fs01 */
    uint8_t  battery_type;      /* battery type: li-ion */
    uint8_t  language_type;     /* language, default ZH */
    uint8_t  speed_unit;        /* speed unit, default kmh */
    uint32_t feature_type;      /* feature enable type */
    uint8_t  rsvd_data[16];
} factory_ccb_data_t;


/* define struct of ppx factory msg info */
typedef struct
{
    uint8_t    req_msg;            /* refer factory_msg_type */
    uint8_t    msg_type;           /* refer factory_msg_type */
    uint8_t    mode_enable;        /* 1 enable or 0 disable */
    uint8_t    imu_cali_req;       /* 1 start the other query */
    
    factory_ccb_data_t       ccb_data;       /* ccb params set */
    factory_ccb_req_msg_t    ccb_req_msg;    /* refer factory_ccb_req_msg_t */
    factory_mcb_req_msg_t    mcb_req_msg;    /* refer factory_mcb_req_msg_t */
    factory_rsp_msg_t        rsp_msg;        /* refer factory_rsp_msg_t */
} factory_msg_data_t;


ppx_packet_status_t factory_msg_parse(IN uint8_t* pdata, IN uint8_t data_len, OUT factory_msg_data_t *factory_msg);

uint16_t factory_msg_format(IN ppx_cmd_type_t cmd_type, IN factory_msg_data_t *factory_msg, OUT void* buffer);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_FACTORY_H__ */
