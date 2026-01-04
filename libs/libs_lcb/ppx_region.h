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
 * @file ppx_region.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_REGION_H__
#define __PPX_REGION_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>
#include "ppx_packet.h"


/* define struct of ppx region msg   */
typedef struct
{
    uint8_t parse_status;
    uint8_t cmd_status;
    uint8_t data_status;
} ppx_region_excp_t;


/* define struct of ppx region msg */
typedef struct
{
    uint8_t id;     /* dev id */
    uint8_t cmd;    /* write/read cmd */

    uint8_t msg_type;   /* for master used */
    uint8_t reg_addr;   /* region data addr*/
    uint8_t reg_nums;   /* region data number */

    ppx_region_excp_t reg_excp; /* exception response */
} ppx_region_msg_t;


/* typedef enum rim state */
typedef enum
{
    PPX_RIM_DOWN        = 0x04,         /* 下坡 */
    PPX_RIM_SEAT        = 0x08,         /* 座椅 */
    PPX_RIM_8DEG        = 0x10,         /* 8度 */
    PPX_RIM_DUMP        = 0x20,         /* 碰撞 */
    PPX_RIM_BUMP        = 0x40,         /* 颠簸 */
    PPX_RIM_TURN        = 0x80,         /* 转弯 */
} ppx_rim_state_t;


/* typedef enum data setting */
typedef enum
{
    /* request data type */
    PPX_CHR_CHECK       = (1 << 0),
    PPX_IMU_OPEN        = (1 << 1),
    PPX_IMU_CALI        = (1 << 2),
    PPX_IAP_MODE        = (1 << 3),
    PPX_SN_WRITE        = (1 << 4),
    PPX_TST_MOTO        = (1 << 5),
    PPX_ACC_CALI        = (1 << 6),

    /* response data status */
    PPX_CHR_CHECK_SUCC   = (1 << 16),
    PPX_IMU_OPEN_SUCC    = (1 << 17),
    PPX_IMU_CALI_SUCC    = (1 << 18),
    PPX_IAP_MODE_FALSE   = (1 << 19),
    PPX_ACC_CALI_SIDE    = (1 << 20),
    PPX_ACC_CALI_SUCC    = (1 << 21),
} ppx_data_setting_t;


/* typedef enum rt setting */
typedef enum
{
    /* request rt_setting type */
    PPX_BRAKE_LED_ON    = (1 << 0),
    PPX_TAIL_LED_ON     = (1 << 1),
    PPX_RIGHT_LED_ON    = (1 << 2),
    PPX_LEFT_LED_ON     = (1 << 3),

    PPX_CLR_ERRCODE     = (1 << 15),

    /* response rt_setting status */
} ppx_rt_setting_t;


/* typedef enum  run mode */
typedef enum 
{
    PPX_MODE_IDLE       = 0,
    PPX_MODE_SET        = 1,
    PPX_MODE_RUN        = 2,
    PPX_MODE_LOCK       = 3,
    PPX_MODE_AID        = 4,
    PPX_MODE_BRAKE      = 5,
    PPX_MODE_IAP        = 6,
    PPX_MODE_TST        = 7
} ppx_run_mode_t;


/* define enum of ppx protocl data region reg addr  */
typedef enum
{
    /* PPX_REQ_GET_ID_NUM   */
    PPX_ID_NUM_REG          = 0, /* 0x00 */
    PPX_MODEL_REG           ,
    
    PPX_SERIAL_NUM_REG      = 2, /* 0x02 */
    PPX_HW_VERSION_REG      ,
    PPX_SW_VESRION_REG      ,

    /* PPX_REQ_GET_STATUS */
    PPX_RIM_STATE_REG       = 5, /* 0x05 */
    PPX_MCU_ERRCODE_REG     ,

    PPX_CTRL_MODEL_REG      = 7, /* 0x07 */
    PPX_SPEED_REF_REG       ,
    PPX_MOTOR_SPEED_REG     ,

    /* PPX_REQ_GET_MCB_STS */
    PPX_BUS_VOLTAGE_REG     = 10, /* 0x0A */
    PPX_BUS_CURRENT_REG     ,

    PPX_PHASE_CUR_A_REG     = 12, /* 0x0C */
    PPX_PHASE_CUR_B_REG     ,
    PPX_PHASE_CUR_C_REG     ,

    PPX_HALL_STATE_REG      = 15, /* 0x0F */
    PPX_PI_VQ_REG           ,
    PPX_PI_IQ_REG           ,

    PPX_BRAKE_STATE_REG     = 18, /* 0x12 */
    PPX_IMU_PITCH_REG       ,
    PPX_IMU_ROLL_REG        ,

    PPX_BOARD_TEMP_REG      = 21, /* 0x15 */
    PPX_BRAKE_MILEAGE_REG   ,
    PPX_MOTOR_ANGLE_REG     ,

    /* PPX_REQ_GET_MILEAGE  */
    PPX_SINGLE_MILEAGE_REG  = 24, /* 0x19 */
    PPX_ANGULAR_SPEED_REG   ,

    /* PPX_REQ_RT_SETTING   */
    PPX_RT_SETTING_REG      = 26, /* 0x1A */

    /* PPX_REQ_SET_SPEED */
    PPX_RUN_MODE_REG        = 27, /* 0x1B */
    PPX_GEARS_REG           ,
    PPX_TARGET_SPEED_REG    ,

    /* PPX_REQ_SET_CONFIG */
    PPX_RATED_VOLT_REG      = 30, /* 0x1E */
    PPX_RATED_CUR_REG       ,
    PPX_MAX_VOLTAGE_REG     ,
    PPX_MIN_VOLTAGE_REG     ,
    PPX_ACCERATION_REG      ,
    PPX_DAT_SETTING_REG     ,
    PPX_RVSD_DATA_REG       = 36, /* 0x24 */

    PPX_MAX_REGION_REG
} ppx_region_reg_t;


/* define ppx_region data struct */
#pragma pack (1)
typedef struct
{
    /* ppx_region data struct of read data  */
    uint8_t   id_num;
    uint8_t   model[PPX_MODEL_SIZE];
    
    uint8_t   serial_num[PPX_SN_SIZE];
    uint16_t  hw_version;
    uint8_t   sw_version[PPX_SW_VER_SIZE];
    
    uint8_t   rim_state;
    uint32_t  mcu_errcode;
    
    uint8_t   ctrl_model;
    int16_t   speed_ref;
    int16_t   motor_speed;
    
    uint16_t  bus_voltage;      /* 0.1V */
    uint16_t  bus_current;      /* 0.1A */
    
    int16_t   phase_current_a;  /* 0.1A */
    int16_t   phase_current_b;  /* 0.1A */
    int16_t   phase_current_c;  /* 0.1A */
    
    uint8_t   hall_state;
    int16_t   pi_vq;
    int16_t   pi_iq;
    
    uint8_t   brake_state;
    int16_t   imu_pitch;        /* 0.1deg */
    int16_t   imu_roll;         /* 0.1deg */

    uint8_t   imu_acc;          /* 0.01g */
    //uint8_t   board_temp;
    //uint8_t   motor_temp;     /* NC */
    uint8_t   brake_mileage;    /* dm */
    int32_t   motor_angle;
    
    uint32_t  single_mileage;   /* m */
    //uint16_t  brake_mileage;
    int16_t   angular_speed;    /* 0.1deg */
    
    /* ppx_region data struct of write data */
    uint16_t  rt_setting;
    
    uint8_t   run_mode;
    uint8_t   gear;
    int16_t   target_speed;     /* rpm */
    
    uint16_t  rated_voltage;    /* 0.1V */
    uint16_t  rated_current;    /* 0.1A */
    uint16_t  max_voltage;      /* 0.1V */
    uint16_t  min_voltage;      /* 0.1V */
    
    uint32_t  acceration;
    uint32_t  dat_setting;
    
    uint32_t  rsvd_data;
} ppx_region_data_t;
#pragma pack()


/* region data buffer */
extern ppx_region_data_t g_ppx_region_data;


ppx_packet_status_t ppx_com_region_parse(IN uint8_t *pdata, IN uint8_t data_len, INOUT ppx_region_msg_t *region_msg);

uint16_t ppx_com_region_format(IN ppx_cmd_type_t cmd_type, IN ppx_region_msg_t *region_msg, OUT void *buffer);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_REGION_H__ */
