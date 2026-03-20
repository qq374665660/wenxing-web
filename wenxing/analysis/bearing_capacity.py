# -*- coding: utf-8 -*-
"""
地基承载力评价模块
"""

import logging

from ..utils import is_fill


def get_base_layers(building, buildings, holes, hole_strata, layer_info):
    """获取基底地层信息"""
    build_holes = [h for h, info in holes.items() if building in info.get('builds', [])]
    if not build_holes:
        return set(), False, 0, []

    base_elev = buildings[building]['embed_elev']
    base_layers = set()
    all_above = True
    above_heights = []
    above_info = []

    for hole_id in build_holes:
        hole_elev = holes[hole_id].get('elev')
        if hole_elev is None:
            continue
        embed_depth = hole_elev - base_elev

        if embed_depth < 0:
            above_heights.append(-embed_depth)
            strata = hole_strata.get(hole_id, [])
            if strata:
                first_layer = strata[0][0]
                layer_name = layer_info.get(first_layer, {}).get('name', '')
                fak = layer_info.get(first_layer, {}).get('bearing_capacity', 0)
                load = buildings[building].get('load', 0)

                ind_satisfy = '满足' if fak >= load else '不满足'
                raft_satisfy = '满足' if fak >= load else '不满足'
                full_name = f"{layer_name}{first_layer}"
                above_info.append((full_name, ind_satisfy, raft_satisfy))
        else:
            all_above = False
            strata = hole_strata.get(hole_id, [])
            prev_depth = 0
            for layer_id, bottom in strata:
                if prev_depth <= embed_depth <= bottom:
                    name = layer_info.get(layer_id, {}).get('name', '')
                    full_name = f"{name}{layer_id}"
                    base_layers.add(full_name)
                    break
                prev_depth = bottom

    avg_above = sum(above_heights) / len(above_heights) if above_heights else 0
    return base_layers, all_above, avg_above, above_info


def get_eta_params(layer_name, silt_params, clay_params):
    """根据地层名称获取承载力修正系数"""
    if '粉土' in layer_name:
        return silt_params.get('粉土', (0.3, 1.5)) + (False,)
    elif '粉质黏土' in layer_name or '粉质粘土' in layer_name:
        return clay_params.get('粉质黏土', (0.0, 1.0)) + (True,)
    elif '黏土' in layer_name or '粘土' in layer_name:
        return 0.0, 1.0, True
    elif '砂土' in layer_name or '砂' in layer_name:
        return 0.3, 2.0, False
    elif '碎石' in layer_name or '卵石' in layer_name:
        return 0.3, 3.0, False
    elif '岩' in layer_name:
        return 0.0, 0.0, False
    else:
        return 0.0, 1.0, False


def calculate_fa(fak, eta_b, eta_d, is_clay, b, d, layer_name, gamma, gamma_m):
    """计算修正后的地基承载力特征值"""
    if '中等风化' in layer_name or '中风化' in layer_name:
        return fak  # 中等风化岩不修正

    if is_clay:
        fa = fak + eta_b * gamma * (b - 3) + eta_d * gamma_m * (d - 1)
    else:
        fa = fak + eta_b * gamma * (b - 3) + eta_d * gamma_m * (d - 0.5)
    return round(fa, 2)


def get_weak_underlayers(building, buildings, holes, hole_strata, layer_info):
    """获取软弱下卧层信息"""
    build_holes = [h for h, info in holes.items() if building in info.get('builds', [])]
    if not build_holes:
        return []

    base_elev = buildings[building]['embed_elev']
    load = buildings[building].get('load', 0)
    weak_layers = []

    for hole_id in build_holes:
        hole_elev = holes[hole_id].get('elev')
        if hole_elev is None:
            continue
        embed_depth = hole_elev - base_elev
        if embed_depth < 0:
            continue

        strata = hole_strata.get(hole_id, [])
        base_found = False
        base_layer_fak = None
        prev_depth = 0

        for layer_id, bottom in strata:
            if not base_found:
                if prev_depth <= embed_depth <= bottom:
                    base_found = True
                    base_layer_fak = layer_info.get(layer_id, {}).get('bearing_capacity', 0)
                prev_depth = bottom
                continue

            # 已找到基底，检查下卧层
            layer_fak = layer_info.get(layer_id, {}).get('bearing_capacity', 0)
            layer_es = layer_info.get(layer_id, {}).get('compression_modulus', 0)

            if layer_fak < load and base_layer_fak is not None and layer_fak < base_layer_fak:
                # 计算 Z（软弱层顶到基底的距离）
                Z = prev_depth - embed_depth
                
                # 计算 es1/es2
                es1 = layer_info.get(strata[0][0], {}).get('compression_modulus', 1)
                es2 = layer_es if layer_es else 1
                es1_es2 = es1 / es2 if es2 != 0 else 1

                layer_name = layer_info.get(layer_id, {}).get('name', '')
                weak_layers.append({
                    'layer_id': layer_id,
                    'layer_name': layer_name,
                    'layer_str': f"{layer_name}{layer_id}",
                    'Z': Z,
                    'fak': layer_fak,
                    'es1_es2': es1_es2
                })
                break

            prev_depth = bottom

    return weak_layers


def calculate_theta(z_b_ratio, es_ratio):
    """计算软弱下卧层验算所需的 theta 值（度）"""
    if z_b_ratio <= 0.25:
        if es_ratio <= 3:
            return 23
        elif es_ratio >= 10:
            return 6
        else:
            return 23 - (23 - 6) * (es_ratio - 3) / (10 - 3)
    elif z_b_ratio >= 0.5:
        if es_ratio <= 3:
            return 25
        elif es_ratio >= 10:
            return 10
        else:
            return 25 - (25 - 10) * (es_ratio - 3) / (10 - 3)
    else:
        theta_025 = calculate_theta(0.25, es_ratio)
        theta_050 = calculate_theta(0.50, es_ratio)
        return theta_025 + (theta_050 - theta_025) * (z_b_ratio - 0.25) / (0.5 - 0.25)
