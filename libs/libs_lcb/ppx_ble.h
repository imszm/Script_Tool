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
    PPX_BLE_STS_NFC_LPCD        = ( 1 << 0 ),  //card swipe card
    PPX_BLE_STS_NFC_READ        = ( 1 << 1 ),  //card read card
    PPX_BLE_STS_CARD_VALID      = ( 1 << 2 ),  //card valid status
    PPX_BLE_STS_CARD_INVALID    = ( 1 << 3 ),  //card invalid status
    PPX_BLE_STS_SOS_KEY         = ( 1 << 4 ),  //sos key event
    PPX_BLE_STS_BLE_CONN        = ( 1 << 5 ),  //ble connection status

    PPX_BLE_STS_LED_INIT_FAIL   = ( 1 << 16 ), //led init fail
    PPX_BLE_STS_NFC_INIT_FAIL   = ( 1 << 17 ), //nfc init fail
    PPX_BLE_STS_NFC_READ_FAIL   = ( 1 << 18 ), //nfc read card id fail
} ppx_ble_status_t;


/* typedef enum data setting */
typedef enum
{
    /* request data type */
    PPX_BLE_NFC_BINDING         = (1 << 0),
    PPX_BLE_NFC_UNBIND          = (1 << 1),
    PPX_BLE_NFC_WRITE           = (1 << 2),
    PPX_BLE_SN_WRITE            = (1 << 3),

    /* response data status */
    PPX_BLE_NFC_BINDING_SUCC    = (1 << 16),
    PPX_BLE_NFC_BINDING_FAIL    = (1 << 17),
    PPX_BLE_NFC_UNBIND_SUCC     = (1 << 18),
    PPX_BLE_NFC_UNBIND_FAIL     = (1 << 19),
    PPX_BLE_NFC_WRITE_SUCC      = (1 << 20),
    PPX_BLE_SN_WRITE_SUCC       = (1 << 21),
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
    uint32_t blink_period : 4;  // Blink period: N * 200ms
    uint32_t blink_duty   : 4;  // Blink duty cycle: (N + 1) / 16 * blink_period
    uint32_t blink_en     : 8;  // Blink enable: bit0-bit7 for battery, LOGO, shield, Ready Go, left turn, right turn, light ring, blink status
    uint32_t err_flag     : 2;  // Error code flag 0 no error, 1 error, 2 iap
    uint32_t err_code     : 4;  // Error code: 0-F
    uint32_t digital      : 7;  // Battery SOC: 0-100
    uint32_t logo         : 2;  // LOGO: 0 off, 1 white, 2 red
    uint32_t rim_state    : 2;  // Shield: 0 off, 1 white, 2 green
    uint32_t rdygo        : 2;  // Ready Go: 0 off, 1 white, 2 red
    uint32_t turn_left    : 2;  // Left turn signal: 0 off, 1 white, 2 orange
    uint32_t turn_right   : 2;  // Right turn signal: 0 off, 1 white, 2 orange
    uint32_t ring         : 2;  // Light ring: 0 off, 1 blue, 2 red
    uint32_t rsvd_data    : 19; // Reserved
} ppx_led_msg_t;


/* define enum of ppx protocl data ble reg addr  */
typedef enum
{
    PPX_BLE_ID_NUM_REG       = 0,  /* 0x00 */
    PPX_BLE_MODEL_REG           ,
    
    PPX_BLE_SERIAL_NUM_REG   = 2,  /* 0x02 */
    PPX_BLE_HW_VERSION_REG      ,
    PPX_BLE_SW_VESRION_REG      ,

    PPX_BLE_STATUS_REG       = 5,  /* 0x05 */
    PPX_BLE_LDR_VALUE_REG       ,
    PPX_BLE_IO_STATUS_REG       ,

    PPX_BLE_LED_MSG_REG      = 8,  /* 0x08 */
    PPX_BLE_CARD_ID_REG         ,
    PPX_BLE_DAT_SETTING_REG  = 10, /* 0x0A */

    PPX_BLE_MAX_REG
} ppx_ble_reg_t;


/* define ppx ble data struct */
#pragma pack (1)
typedef struct
{
    uint8_t         id_num;                 // Device ID number
    uint8_t         model[PPX_MODEL_SIZE];  // Model (8 bytes)
    uint8_t         serial_num[PPX_SN_SIZE];// Serial number (26 bytes)
    uint8_t         hw_version;             // Hardware version
    uint8_t         sw_version[PPX_SW_VER_SIZE]; // Software version (20 bytes)
    uint32_t        status;                 // @ref ppx_ble_status_t
    uint16_t        ldr_value;              // light-dependent resistor brightness
    uint16_t        io_status;              // IO pin status (bitfield):
                                            //   bit0~bit11: PA0~PA11
                                            //   bit12: DM
                                            //   bit13: DP
    ppx_led_msg_t   led_msg;                // led display message
    uint32_t        card_id;                // NFC card ID
    uint32_t        dat_setting;            // @ref ppx_ble_data_setting_t
} ppx_ble_data_t;
#pragma pack()


/* ble data buffer */
extern ppx_ble_data_t g_ppx_ble_data;


ppx_packet_status_t ppx_com_ble_parse(IN uint8_t *pdata, IN uint8_t data_len, INOUT ppx_ble_msg_t *ble_msg);

uint16_t ppx_com_ble_format(IN ppx_cmd_type_t cmd_type, IN ppx_ble_msg_t *ble_msg, OUT void *buffer);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_BLE_H__ */
