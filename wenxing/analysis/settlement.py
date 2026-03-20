# -*- coding: utf-8 -*-
"""
沉降计算模块
"""

import logging

from ..config import get_alpha_data
from ..utils import interpolate_alpha


def get_under_layers(building, hole_id, buildings, holes, hole_strata, layer_info):
    """获取基底以下地层信息"""
    base_elev = buildings[building]['embed_elev']
    hole_elev = holes[hole_id].get('elev')
    if hole_elev is None:
        return []

    embed_depth = hole_elev - base_elev
    strata = hole_strata.get(hole_id, [])
    under_layers = []
    prev_depth = 0

    for layer_id, bottom in strata:
        if bottom <= embed_depth:
            prev_depth = bottom
            continue

        start = max(embed_depth, prev_depth)
        thickness = bottom - start
        es = layer_info.get(layer_id, {}).get('compression_modulus', 0)
        name = layer_info.get(layer_id, {}).get('name', '')

        under_layers.append({
            'layer_id': layer_id,
            'name': name,
            'start': start,
            'bottom': bottom,
            'thickness': thickness,
            'es': es
        })
        prev_depth = bottom

    return under_layers


def calculate_settlement(building, hole_id, buildings, holes, hole_strata, layer_info):
    """计算沉降量"""
    b_info = buildings[building]
    width = b_info.get('width')
    length = b_info.get('length')
    load = b_info.get('load')
    base_elev = b_info['embed_elev']

    if not all((width, length, load)):
        return None

    hole_elev = holes[hole_id].get('elev')
    if hole_elev is None:
        return None

    embed_depth = hole_elev - base_elev
    if embed_depth < 0:
        return None

    b = width / 2
    l_b = length / width

    under_layers = get_under_layers(building, hole_id, buildings, holes, hole_strata, layer_info)
    if not under_layers:
        return None

    alpha_data, z_b_values, l_b_values = get_alpha_data()
    
    # 计算附加应力 p0
    gamma = 20  # 默认土的重度
    pc = gamma * 1.5  # 基底自重应力
    p0 = load - pc

    settlement = 0.0
    prev_alpha = 0.0
    prev_z = 0.0

    for layer in under_layers:
        z = layer['bottom'] - embed_depth
        z_b = z / b if b != 0 else 0
        
        alpha = interpolate_alpha(alpha_data, z_b_values, l_b_values, z_b, l_b)
        es = layer['es']
        
        if es is None or es == 0:
            continue

        # 分层总和法
        delta_sigma = p0 * (alpha - prev_alpha)
        layer_thickness = layer['thickness']
        
        if delta_sigma > 0:
            layer_settlement = (delta_sigma * layer_thickness) / es
            settlement += layer_settlement

        prev_alpha = alpha
        prev_z = z

    # 应用经验系数 psi_s
    es_avg = sum(l['es'] for l in under_layers if l['es']) / len([l for l in under_layers if l['es']]) if under_layers else 0
    
    if es_avg <= 2.5:
        psi_s = 1.44
    elif es_avg <= 4:
        psi_s = 1.32
    elif es_avg <= 7:
        psi_s = 1.2
    elif es_avg <= 15:
        psi_s = 1.0
    elif es_avg <= 20:
        psi_s = 0.7
    else:
        psi_s = 0.4

    return round(settlement * psi_s * 1000, 2)  # 转换为 mm


def get_tilt_limit(height):
    """根据建筑物高度获取倾斜限值"""
    if height is None:
        return 0.003
    
    if height <= 24:
        return 0.004
    elif height <= 60:
        return 0.003
    elif height <= 100:
        return 0.0025
    else:
        return 0.002
