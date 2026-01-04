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
 * @file ppx_packet.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_PACKET_H__
#define __PPX_PACKET_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>


/* define ppx parameter direction */
#ifndef IN
#define IN 
#endif

#ifndef OUT
#define OUT 
#endif

#ifndef INOUT
#define INOUT 
#endif


/* define ppx protocl max data length */
#define PPX_PACKET_MAX_SIZE         ( 256 )
#define PPX_PACKET_MIN_SIZE         ( 9 )

#define PPX_DATA_HEAD_SIZE          ( 5 )       /* ex : msg type +frame index + crc value*/
#define PPX_DATA_REGION_SIZE        ( 128 + PPX_DATA_HEAD_SIZE )
#define PPX_DATA_BUF_SIZE           ( 192 )     /* -64 */


/* define ppx project information */
#define PPX_SW_VER_SIZE             ( 20 )
#define PPX_VER_MIN_SIZE            ( 12 )

#define PPX_MODEL_SIZE              ( 8 )
#define PPX_SN_SIZE                 ( 26 )

#define PPX_BIN_IAP_VER_OFFSET      (0x0800)    /* BIN IAP offset 2K */
#define PPX_BIN_APP_VER_OFFSET      (0x2800)    /* BIN APP offset 10K */


/* define ppx protocl frame head */
#define PPX_FRAME_HEAD              ( 0xA5 )
#define PPX_FRAME_END               ( 0x55 )
#define PPX_DATA_TAG                ( 0x33 )

#define PPX_DATA_REPHEAD_H          ( 0xAB )  /* 0xA5 -> 0xABBA */
#define PPX_DATA_REPHEAD_L          ( 0xBA )  /* 0xA5 -> 0xABBA */

#define PPX_DATA_REPEND_H           ( 0xCD )  /* 0x55 -> 0xCDDC */
#define PPX_DATA_REPEND_L           ( 0xDC )  /* 0x55 -> 0xCDDC */

#define PPX_DATA_REPHEAD_2          ( 0xBB )  /* 0xABBA -> 0xABBB BA */
#define PPX_DATA_REPEND_2           ( 0xDD )  /* 0xCDDC -> 0xCDDD DC*/


/* define ppx return status */
typedef enum
{
    PPX_ERROR = -1,
    PPX_FALSE = 0,
    PPX_TRUE  = !PPX_FALSE
} ppx_packet_status_t;


/* define ppx packet format type */
typedef enum
{
    PPX_FMT_REGION      = 0x01,
    PPX_FMT_IAP         = 0x02,
} ppx_packet_format_t;


/* define ppx protocl id data type */
typedef enum
{
    PPX_ID_RSVD         = 0x00,
    PPX_ID_CCB          = 0x10,
    PPX_ID_MCB          = 0x20,
    PPX_ID_FCB          = 0x30,
    PPX_ID_BMS          = 0x40,
    PPX_ID_GPRS         = 0x50,
    PPX_ID_BLE          = 0x60,
    PPX_ID_ALARM        = 0x70,
    PPX_ID_VOICE        = 0x80,

    PPX_ID_MAX          = 0x90,
} ppx_packet_id_t;


/* define ppx packet command type */
typedef enum
{
    PPX_CMD_REQ         = 0x00,
    PPX_CMD_RSP         = 0x80,
    PPX_CMD_EXCP        = 0xC0
} ppx_cmd_type_t;


/* define ppx packet command content */
typedef enum
{
    PPX_MSG_RSVD        = 0x00,
    PPX_MSG_READ        = 0x01,
    PPX_MSG_MULTREAD    = 0x02,
    PPX_MSG_WRITE       = 0x03,
    PPX_MSG_MULTWRITE   = 0x04,
    PPX_MSG_COMPARE     = 0x05,
    PPX_MSG_UPGRADE     = 0x06,
    PPX_MSG_NOTIFY      = 0x07,

    PPX_MSG_MASK        = 0x0F,
} ppx_cmd_msg_t;


/* define ppx return status */
#define PPX_CMD_IS_REQ(x)   (((x) > 0) && ((x) <= PPX_MSG_MASK))
#define PPX_CMD_IS_RSP(x)   (((x) & PPX_CMD_RSP) == PPX_CMD_RSP)
#define PPX_CMD_IS_EXCP(x)  (((x) & PPX_CMD_EXCP) == PPX_CMD_EXCP)
#define PPX_CMD_IS_MSG(x)   ((x) & PPX_MSG_MASK)


/* The ppx protocl of packet */
typedef struct
{
    uint8_t id;
    uint8_t cmd;
    uint8_t data_len;
    uint8_t data[PPX_DATA_BUF_SIZE];
} ppx_packet_data_t;


uint16_t ppx_com_packet_crc(IN uint8_t *pdata, IN uint32_t len);

ppx_packet_status_t ppx_com_packet_parse(IN uint8_t *pdata, IN uint16_t data_len, OUT ppx_packet_data_t *ppx_packet);

uint16_t ppx_com_packet_format(IN ppx_cmd_type_t cmd_type, IN ppx_packet_data_t *ppx_packet, OUT uint8_t *buffer);

ppx_packet_status_t ppx_com_packet_verchk(IN uint8_t *new_version, IN uint8_t new_len, IN uint8_t *old_version, IN uint8_t old_len);


#ifdef __cplusplus
}
#endif

#endif /* __PPX_PACKET_H__ */
