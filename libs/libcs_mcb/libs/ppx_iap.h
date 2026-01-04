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
 * @file ppx_iap.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_IAP_H__
#define __PPX_IAP_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>
#include "ppx_packet.h"


/* define ppx iap max data length */
#define PPX_IAP_DATA_SIZE           ( 128 )
#if (( PPX_DATA_REGION_SIZE - PPX_DATA_HEAD_SIZE ) < PPX_IAP_DATA_SIZE)
 #error "PPX_IAP_DATA_SIZE over range, please check define"
#endif


/* define enum, ppx iap from type */
typedef enum
{
    PPX_IAP_FROM_UNKOWN         = 0x0000,
    PPX_IAP_FORM_FLASH          = 0xCCCC,
    PPX_IAP_FORM_PCTOOL         = 0xDDDD,
} ppx_iap_from_type_t;


/* define struct of ppx iap start msg */
typedef struct
{
    uint8_t sw_version[PPX_SW_VER_SIZE];
    uint32_t total_size;
    uint16_t frame_size;
    uint16_t frame_count;
} ppx_iap_start_msg_t;


/* define struct of ppx iap data msg */
typedef struct
{
    uint16_t frame_index;
    uint8_t  data[PPX_IAP_DATA_SIZE];
    uint16_t data_len;
    uint16_t crc_value;
} ppx_iap_data_msg_t;


/* define struct of ppx iap stop msg */
#define PPX_IAP_FIN_SUCCESS         ( 0x1010 )

typedef struct
{
    uint16_t finish_flag;
    uint16_t crc_value;
} ppx_iap_stop_msg_t;


/* define enum, ppx iap board type */
typedef enum
{
    PPX_IAP_RSP_SUCCESS         = 0x0101,
    PPX_IAP_STS_ERROR           = 0x0202,
    PPX_IAP_SW_VERSION_FAILED   = 0x0203,
    PPX_IAP_FRM_SIZE_CNT_FAILED = 0x0204,
    PPX_IAP_TOTAL_SIZE_FAILED   = 0x0205,
    PPX_IAP_DATA_REQ_FAILED     = 0x0206,
    PPX_IAP_FINISH_CRC_FAILED   = 0x0207,
    PPX_IAP_FLASH_RW_FAILED     = 0x0208,
    PPX_IAP_TIMEOUT_FAILED      = 0x0209,
} ppx_iap_rsp_status_t;


/* define enum, ppx iap status type */
typedef enum
{
    PPX_IAP_STS_RSVD            = 0x0800,
    PPX_IAP_STS_READY           = 0x0801,
    PPX_IAP_STS_START           = 0x0802,
    PPX_IAP_STS_UPGRADE         = 0x0803,
    PPX_IAP_STS_CRC             = 0x0804,
} ppx_iap_status_type_t;


/* define struct of ppx iap resp msg */
typedef struct
{
    uint32_t iap_status;    /* refer iap_status_type*/
    
    uint16_t rsp_status;    /* refer iap_rsp_status */
    uint16_t frame_count;   /* slave recv frame count */
} ppx_iap_resp_msg_t;


/* define enum ppx iap msg type */
typedef enum
{
    PPX_IAP_RVSD_TYPE           = 0x60,
    PPX_IAP_QUERY_REQ           = 0x61,
    PPX_IAP_QUERY_RSP           = 0x62,
    PPX_IAP_START_REQ           = 0x63,
    PPX_IAP_START_RSP           = 0x64,
    PPX_IAP_DATA_REQ            = 0x65,
    PPX_IAP_DATA_RSP            = 0x66,
    PPX_IAP_STOP_REQ            = 0x67,
    PPX_IAP_STOP_RSP            = 0x68,
    PPX_IAP_RESET_REQ           = 0x69
} ppx_iap_msg_type_t;


/* define struct of ppx iap msg info */
typedef struct
{
    uint8_t req_msg;    /* refer iap_msg_type */
    uint8_t msg_type;   /* refer iap_msg_type */
    uint8_t hw_type;    /* refer iap_hw_type */

    ppx_iap_start_msg_t  start_msg;   /* iap_start msg */
    ppx_iap_data_msg_t   data_msg;    /* iap_data msg */
    ppx_iap_stop_msg_t   stop_msg;    /* iap_stop msg */

    ppx_iap_resp_msg_t   resp_msg;    /* refer packet_status */
} ppx_iap_data_t;


/* iap data buffer */
extern ppx_iap_data_t g_ppx_iap_data;


ppx_packet_status_t ppx_com_iap_parse(IN uint8_t *pdata, IN uint8_t data_len, OUT ppx_iap_data_t *iap_data);

uint16_t ppx_com_iap_format(IN ppx_cmd_type_t cmd_type, IN ppx_iap_data_t *iap_data, OUT void *buffer);


#ifdef __cplusplus
 }
#endif

#endif /* __PPX_IAP_H__ */
