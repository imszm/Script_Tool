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
 * @file ppx_ble.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_BLE_H__
#define __PPX_BLE_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>
#include "ppx_packet.h"

/* define enum status */
typedef enum
{
    PPX_BLE_D1_KEY       = ( 1 << 0 ),  // gear d1 key event
    PPX_BLE_D2_KEY       = ( 1 << 1 ),  // gear d2 key event
    PPX_BLE_D3_KEY       = ( 1 << 2 ),  // gear d3 key event
    PPX_BLE_SPK_KEY      = ( 1 << 3 ),  // speaker key event
    PPX_BLE_BLE_CONN     = ( 1 << 4 ),  // ble connection status

    PPX_BLE_LED_INIT_FAIL   = ( 1 << 16 ), // led init fail
    PPX_BLE_VOICE_INIT_FAIL = ( 1 << 17 ), // voice init fail
    PPX_BLE_JOY_INIT_ERROR  = ( 1 << 18 ), // joystick init error
    PPX_BLE_SN_WRITE_SUCC   = ( 1 << 19 ), // sn write response
    PPX_BLE_SYS_RESET_SUCC  = ( 1 << 20 ), // sys_reset response
    PPX_BLE_JOY_CALI_SUCC   = ( 1 << 21 ), // joy_cali response
} ppx_ble_status_t;


/* typedef enum data setting */
typedef enum
{
    PPX_BLE_SN_WRITE       = (1 << 0),
    PPX_BLE_SYS_RESET      = (1 << 1),
    PPX_BLE_JOY_CALI       = (1 << 2),
} ppx_ble_data_setting_t;


/* define struct of ppx ble msg */
typedef struct
{
    uint8_t id;             /* dev id */
    uint8_t cmd;            /* write/read cmd */

    uint8_t reg_addr;       /* ble data addr*/
    uint8_t reg_nums;       /* ble reg number */
} ppx_ble_msg_t;


/* define struct of ppx led msg 64bit */
typedef struct 
{
    uint32_t screen_on    : 1;  // Display switch: 1 on, 0 off
    uint32_t brightness   : 3;  // Brightness level 0-7
    uint32_t blink_period : 3;  // Blink period: N * 200ms
    uint32_t blink_duty   : 4;  // Blink duty cycle: (N + 1) / 16 * blink_period
    uint32_t blink_en     : 5;  // Blink enable: bit0-bit4 for digital, percent, charge, gear, blink status
    uint32_t color_flag   : 1;  // Color flag 0 white, 1 orange
    uint32_t err_flag     : 2;  // Error flag 0 no error, 1 E, 2 L, 3 R
    uint32_t digital      : 8;  // digital bit0-8: 0-188
    uint32_t percent      : 1;  // percent: 0 off, 1 white
    uint32_t charge       : 1;  // charge: 0 off, 1 green
    uint32_t gear         : 2;  // gear: 0 off, 1 D1, 2 D2, 3 D3
    uint32_t light        : 1;  // light: 0 off, 1 on   
} ppx_led_msg_t;


/* define enum of ppx protocl data ble reg addr  */
typedef enum
{
    PPX_BLE_ID_NUM_REG       = 0,  /* 0x00 */
    PPX_BLE_MODEL_REG           ,
    
    PPX_BLE_SERIAL_NUM_REG   = 2,  /* 0x02 */
    PPX_BLE_HW_VERSION_REG      ,
    PPX_BLE_SW_VESRION_REG      ,

    PPX_BLE_LDR_VALUE_REG     = 5,  /* 0x05 */
    PPX_BLE_STATUS_REG          ,
    PPX_BLE_JOY_X_PCT_REG       ,
    PPX_BLE_JOY_Y_PCT_REG       ,
    PPX_BLE_JOY_X_VAL_REG       ,
    PPX_BLE_JOY_Y_VAL_REG       ,

    PPX_BLE_LED_MSG_REG      = 11,  /* 0x0B */
    PPX_BLE_VOICE_DATA_REG      ,
    PPX_BLE_DAT_SETTING_REG     ,

    PPX_BLE_MAX_REG
} ppx_ble_reg_t;


/* define ppx ble data struct */
#pragma pack (1)
typedef struct
{
    /* ppx_ble read data  */
    uint8_t         id_num;                 // Device ID number
    uint8_t         model[PPX_MODEL_SIZE];  // Model (8 bytes)
    uint8_t         serial_num[PPX_SN_SIZE];// Serial number (26 bytes)
    uint8_t         hw_version;             // Hardware version
    uint8_t         sw_version[PPX_SW_VER_SIZE]; // Software version (20 bytes)

    uint16_t        ldr_value;              // light-dependent resistor brightness
    uint32_t        status;                 // @ref ppx_ble_status_t
    int8_t          joy_x_pct;              // Joystick X value -100 ~ 100
    int8_t          joy_y_pct;              // Joystick Y value -100 ~ 100
    int16_t         joy_x_val;              // Joystick X ADC value (mV)
    int16_t         joy_y_val;              // Joystick Y ADC value (mV)
    
    /* ppx_ble  write data */
    ppx_led_msg_t   led_msg;                // led display message
    uint32_t        voice_data;             // voice data, voice num: bit0-bit15; volume: bit16-bit31

    uint32_t        dat_setting;            // @ref ppx_ble_data_setting_t
} ppx_ble_data_t;
#pragma pack()


/* ble data buffer */
extern ppx_ble_data_t g_ppx_ble_data;


/* extern function prototypes -----------------------------------------------*/
ppx_packet_status_t ppx_com_ble_parse(IN uint8_t *pdata, IN uint8_t data_len, INOUT ppx_ble_msg_t *ble_msg);
uint16_t ppx_com_ble_format(IN ppx_cmd_type_t cmd_type, IN ppx_ble_msg_t *ble_msg, OUT void *buffer);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_BLE_H__ */
