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
 * PiPiXiong' name may not be used to endorse or promote products derived from
 * this software without specific prior written permission.
 *
 * DISCLAIMER: THIS SOFTWARE IS PROVIDED BY PIPIXIONG "AS IS" AND ANY EXPRESS OR
 * IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NON-INFRINGEMENT ARE
 * DISCLAIMED. IN NO EVENT SHALL PIPIXIONG BE LIABLE FOR ANY DIRECT, INDIRECT,
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


/* typedef enum  run mode */
typedef enum
{
    PPX_MODE_IDLE        = 0,    /* 初始状态 */
    PPX_MODE_SETTING     = 1,    /* 设置模式 */
    PPX_MODE_RUNNING     = 2,    /* 骑行模式  */
    PPX_MODE_LOCK        = 3,    /* 锁车模式  */
    PPX_MODE_PWR_PUSH    = 4,    /* 助力推行 */
    PPX_MODE_BRK_EMER    = 5,    /* 紧急刹车 */
    PPX_MODE_IAP         = 6,    /* 烧录模式  */
    PPX_MODE_TESTING     = 7,    /* 生产测试  */
    PPX_MODE_REHAB_WALK  = 8,    /* 康复助行 */
    PPX_MODE_REHAB_TRAIN = 9,    /* 康复训练 */
    PPX_MODE_BRAKE       = 10,   /* 延迟关闭电磁刹 */
    PPX_MODE_BRK_ECND    = 11,   /* 立刻关闭电磁刹 */
    PPX_MODE_CLUTCH_OPEN = 12,   /* 电磁刹解锁指令 */
} ppx_run_mode_t;


/* typedef enum rim state */
typedef enum
{
    PPX_RIM_FALL        = 0x04,     /* 摔倒 */
    PPX_RIM_SEAT        = 0x08,     /* 座椅 */
    PPX_RIM_8DEG        = 0x10,     /* 8度 */
    PPX_RIM_DUMP        = 0x20,     /* 倾倒 */
    PPX_RIM_BUMP        = 0x40,     /* 颠簸 */
    PPX_RIM_TURN        = 0x80,     /* 转弯 */
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
    PPX_PRODUCT_R3      = (3 << 7),     // bit:7 ~ 9
    PPX_PRODUCT_P3      = (4 << 7),     // bit:7 ~ 9
    PPX_PRODUCT_MAX     = (7 << 7),     // bit:7 ~ 9
    PPX_MOTOR_STOP      = (1 << 10),
    PPX_PARA_LEARN      = (1 << 11),

    /* response data status */
    PPX_CHR_CHECK_SUCC   = (1 << 16),
    PPX_IMU_OPEN_SUCC    = (1 << 17),
    PPX_IMU_CALI_SUCC    = (1 << 18),
    PPX_IAP_MODE_FALSE   = (1 << 19),
    PPX_ACC_CALI_SIDE    = (1 << 20),
    PPX_ACC_CALI_SUCC    = (1 << 21),
    PPX_PRD_SET_SUCC     = (1 << 22),
    PPX_PARA_LEARN_SUCC  = (1 << 23),
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


/* typede enum  brake_state */
typedef enum
{
    PPX_BRAKE_CLOSED     = 0,    /* 刹车闭合 */
    PPX_BRAKE_OPENING    = 1,    /* 刹车解锁中 */
    PPX_BRAKE_OPENED     = 2     /* 刹车已解锁 */
} ppx_brake_state_t;


/* typede enum  road_cond */
typedef enum
{
    PPX_ROAD_LVL        = 0,    /* 平路，小于3deg */
    PPX_ROAD_UP         = 1,    /* 上坡，大于3deg */
    PPX_ROAD_DW         = 2     /* 下坡，大于3deg */
} ppx_road_cond_t;


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
    PPX_MCU_ERRCODE_REG     = 5, /* 0x05 */
    PPX_MOTOR_SPEED_REG     ,

    /* PPX_REQ_GET_MCB_STS */
    PPX_BUS_VOLTAGE_REG     = 7, /* 0x07 */
    PPX_BUS_CURRENT_REG     ,

    PPX_RIM_STATE_REG       = 9,
    PPX_CTRL_MODEL_REG      ,
    PPX_SPEED_REF_REG       ,

    PPX_MOSFET_TEMP_REG     = 12, /* 0x0C */
    PPX_MOTOR_TEMP_REG      ,

    PPX_MOTOR_LIMIT_FLG_REG = 14, /* 0x0E */
    PPX_MOTOR_VQ_REG        ,
    PPX_MOTOR_IQ_REG        ,

    PPX_MOTOR_CALI_STATE_REG = 17, /* 0x11 */
    PPX_MOTOR_CALI_RES_REG  ,
    PPX_MOTOR_CALI_LD_REG   ,
    PPX_MOTOR_CALI_LQ_REG   ,
    PPX_MOTOR_CALI_BEMF_REG ,

    /* PPX_REQ_GET_MILEAGE  */
    PPX_BRAKE_STATE_REG     = 22, /* 0x16 */
    PPX_SINGLE_MILEAGE_REG  ,

    /* PPX_REQ_RT_SETTING   */
    PPX_RT_SETTING_REG      = 24, /* 0x18 */

    /* PPX_REQ_SET_SPEED */
    PPX_RUN_MODE_REG        = 25, /* 0x19 */
    PPX_ROAD_COND_REG       ,
    PPX_TARGET_SPEED_REG    ,
    PPX_TARGET_ACCEL_REG    ,
    PPX_TARGET_CUR_REG      ,

    /* PPX_REQ_SET_CONFIG */
    PPX_RATED_VOLT_REG      = 30, /* 0x1E */
    PPX_RATED_CUR_REG       ,

    PPX_DAT_SETTING_REG     ,
    PPX_RVSD_DATA_REG       = 33, /* 0x21 */

    PPX_MAX_REGION_REG
} ppx_region_reg_t;


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


/* define ppx_region data struct */
#pragma pack (1)
typedef struct
{
    /* ppx_region read data  */
    uint8_t   id_num;
    uint8_t   model[PPX_MODEL_SIZE];
    
    uint8_t   serial_num[PPX_SN_SIZE];
    uint16_t  hw_version;
    uint8_t   sw_version[PPX_SW_VER_SIZE];
    
    uint32_t  mcu_errcode;      /* err code */
    int16_t   motor_speed;      /* fb speed */
    
    uint16_t  bus_voltage;      /* 0.1V */
    uint16_t  bus_current;      /* 0.1A */
    
    uint8_t   rim_state;
    uint8_t   ctrl_model;
    int16_t   speed_ref;        /* rpm */
    
    int16_t   mosfet_temp;      /* deg */
    int16_t   motor_temp;       /* deg */

    uint8_t   motor_limit_flg;
    int32_t   motor_vq;
    int16_t   motor_iq;

    uint8_t   motor_cali_state;
    uint16_t  motor_cali_res;   /* mΩ */
    uint16_t  motor_cali_ld;    /* uH */
    uint16_t  motor_cali_lq;    /* uH */
    uint16_t  motor_cali_bemf;  /* uV/rpm */

    uint8_t   brake_state;      /* ppx_brake_state_t */
    uint32_t  single_mileage;   /* m */
    
    /* ppx_region  write data */
    uint16_t  rt_setting;
    
    uint8_t   run_mode;         /* ppx_run_mode_t */
    uint8_t   road_cond;        /* deg */

    int16_t   target_speed;     /* rpm */
    uint16_t  target_accel;     /* m/s */
    int16_t   target_current;   /* 0.1A */
    
    uint16_t  rated_voltage;    /* 0.1V */
    uint16_t  rated_current;    /* 0.1A */
    
    uint32_t  dat_setting;      /* data cfg */
    
    uint32_t  reserved_data;    /* reserved */
} ppx_region_data_t;
#pragma pack()


/* define ppx_region control struct */
typedef struct
{
    ppx_region_msg_t   msg;     /* message info */
    ppx_region_data_t  data;    /* register data */
} ppx_region_ctrl_t;


/* region data buffer */
//extern ppx_region_data_t g_ppx_region_data;


/* extern function prototypes -----------------------------------------------*/
ppx_packet_status_t ppx_com_region_parse(IN uint8_t *pdata, IN uint16_t data_len, INOUT ppx_region_ctrl_t *region_ctrl);
uint16_t ppx_com_region_format(IN ppx_cmd_type_t cmd_type, IN ppx_region_ctrl_t *region_ctrl, OUT void *buffer);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_REGION_H__ */
