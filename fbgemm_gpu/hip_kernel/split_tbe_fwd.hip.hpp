/*******************************************************************************
 * Copyright (c) 2016 - 2022 Advanced Micro Devices, Inc. All rights reserved.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 *
 ******************************************************************************/
#pragma once

#include <hip/hip_runtime.h>
#include <hip/hip_fp16.h>

#define SPLIT_TBE_FWD_KERNEL(emb_prec, emb_type, bag_prefetch, bag_unroll) \
    extern "C" __global__ void split_tbe_fwd_unweighted_hip_kernel_ ## emb_prec ( \
            float * p_output,              \
            const emb_type * p_emb_table,  \
            const int64_t * p_indices,     \
            const int64_t * p_offsets,     \
            const int64_t pooling_mode,    \
            const int32_t * D_offsets,    \
            const int64_t* weight_offsets, \
            uint32_t emb_dim,              \
            uint32_t batch,                \
            uint32_t num_rows,             \
            uint32_t num_tables);          \
    \
    extern "C" __global__ void split_tbe_fwd_weighted_hip_kernel_ ## emb_prec ( \
            float * p_output,              \
            const emb_type * p_emb_table,  \
            const int64_t * p_indices,     \
            const int64_t * p_offsets,     \
            const int64_t pooling_mode,    \
            const int32_t * D_offsets,    \
            const int64_t* weight_offsets, \
            const float * p_indice_weights,\
            uint32_t emb_dim,              \
            uint32_t batch,                \
            uint32_t num_rows,             \
            uint32_t num_tables)

SPLIT_TBE_FWD_KERNEL(fp16,  half, 2, 16);

SPLIT_TBE_FWD_KERNEL(fp32, float, 2, 16);
