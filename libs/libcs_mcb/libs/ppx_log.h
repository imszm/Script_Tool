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
 * @file ppx_log.h
 * @author PiPiXiong
 * @version v1.0.0
 *
 * @copyright Copyright (c) 2022, Zhimahuaerkai Technologies Co.,Ltd. All rights reserved.
 */

#ifndef __PPX_LOG_H__
#define __PPX_LOG_H__

#ifdef __cplusplus
extern "C" {
#endif

#include <stdint.h>
#include "ppx_packet.h"


/* define ppx log max data length */
#define PPX_LOG_DATA_SIZE           ( 125 ) //( PPX_DATA_REGION_SIZE - PPX_DATA_HEAD_SIZE )


/* define log putout dir */
typedef enum {
    LOG_DIR_FLASH           = 0,
    LOG_DIR_CONSOLE         = 1,
} ppx_log_dir_t;


/* define enum, ppx log response type */
typedef enum
{
    PPX_LOG_RSP_FAILED      = 0,
    PPX_LOG_RSP_SUCCESS     = 1,
    PPX_LOG_RSP_FINISHED    = 2,
} ppx_log_rsp_status_t;


/* define enum ppx log msg type */
typedef enum
{
    PPX_LOG_RVSD_TYPE       = 0x70,
    PPX_LOG_SET_DIR_REQ     = 0x71,
    PPX_LOG_SET_DIR_RSP     = 0x72,
    PPX_LOG_QUERY_REQ       = 0x73,
    PPX_LOG_QUERY_RSP       = 0x74,
    PPX_LOG_RESET_REQ       = 0x75,
    PPX_LOG_RESET_RSP       = 0x76,
    PPX_LOG_DEV_REPORT      = 0x77,
    PPX_LOG_MEMORY_REQ      = 0x78,
    PPX_LOG_MEMORY_RSP      = 0x79,
} ppx_log_msg_type_t;


/* define struct of ppx log resp msg */
typedef struct
{
    uint8_t rsp_status; /* refer log_rsp_status */
    uint8_t log_type;
    uint16_t memery_offset;
    uint8_t data_len;
    uint8_t data[PPX_LOG_DATA_SIZE];
} ppx_log_resp_t;


/* define struct of ppx log msg info */
typedef struct
{
    uint8_t req_msg;    /* refer msg_type */
    uint8_t msg_type;   /* refer msg_type */
    uint8_t out_dir;    /* refer log_dir */
    
    ppx_log_resp_t resp_msg;  /* refer resp_msg */
} ppx_log_pkt_t;


/* log msg data buffer */
extern ppx_log_pkt_t g_ppx_log_pkt;


ppx_packet_status_t ppx_com_log_parse(IN uint8_t* pdata, IN uint8_t data_len, OUT ppx_log_pkt_t* log_pkt);

uint16_t ppx_com_log_format(IN ppx_cmd_type_t cmd_type, IN ppx_log_pkt_t* log_pkt, OUT void* buffer);


#ifdef __cplusplus
 }
#endif

#endif /* __PPX_LOG_H__ */
