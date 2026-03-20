# -*- coding: utf-8 -*-
"""
地基均匀性评价模块
"""

import logging
from collections import defaultdict, Counter

from ..config import ALPHA_TABLE
from ..utils import is_fill, linear_interpolate, bilinear_interpolate


def find_layer_at_depth(hole_strata, holes, hole_id, depth):
    """在指定深度查找地层"""
    layers = hole_strata.get(hole_id, [])
    if not layers:
        return None
    prev_bottom = 0.0
    for layer_id, bottom in layers:
        if prev_bottom <= depth <= bottom:
            return layer_id
        prev_bottom = bottom
    return layers[-1][0] if layers else None


def compute_desc(building, buildings, holes, layer_info, hole_strata):
    """计算建筑物基底地层描述"""
    logging.info(f"计算建筑物 {building} 的地层描述")
    build_holes_list = [h for h, info in holes.items() if building in info['builds']]
    logging.info(f"建筑物 {building} 关联钻孔列表: {build_holes_list}")
    
    if not build_holes_list:
        logging.warning(f"建筑物 {building} 无关联钻孔")
        return "无勘探点数据"

    hole_elevs = [holes[h]['elev'] for h in build_holes_list if holes[h]['elev'] is not None]
    if not hole_elevs:
        logging.warning(f"建筑物 {building} 无有效钻孔标高数据")
        return "无有效标高数据"
    
    min_elev = min(hole_elevs)
    base_elev = buildings[building]['embed_elev']
    max_above = base_elev - min_elev

    hole_to_layer = {}
    fill_thicks = defaultdict(list)
    first_layers = []
    above_holes = 0

    for hole_id in build_holes_list:
        hole_elev = holes[hole_id]['elev']
        if hole_elev is None:
            logging.warning(f"钻孔 {hole_id} 无标高数据，跳过")
            continue
        embed_depth = hole_elev - base_elev
        strata = hole_strata.get(hole_id, [])
        if not strata:
            logging.warning(f"钻孔 {hole_id} 无地层数据，跳过")
            continue
        first_layer = strata[0][0]
        first_layers.append(first_layer)

        if embed_depth < 0:
            above_holes += 1
            logging.info(f"钻孔 {hole_id} 基底高于地面，嵌入深度 {embed_depth}")
            continue

        layer = find_layer_at_depth(hole_strata, holes, hole_id, embed_depth)
        if layer:
            hole_to_layer[hole_id] = layer
            if is_fill(layer):
                layer_bottom = 0.0
                for l, b in strata:
                    if not is_fill(l):
                        break
                    layer_bottom = b if b is not None else holes[hole_id]['max_depth'] or 0.0
                if embed_depth is not None and layer_bottom is not None:
                    fill_thick = layer_bottom - embed_depth
                    fill_thicks[layer].append(fill_thick)

    desc = ""
    if above_holes > 0 or max_above >= 0:
        max_above = round(max(0, max_above), 2)
        first_counts = Counter(first_layers)
        if first_counts:
            main_first, _ = first_counts.most_common(1)[0]
            main_first_name = layer_info.get(main_first, {}).get('name', '')
            desc += f"基底高于现状地面最大约{max_above}m，现状地面下主要分布{main_first_name}{main_first}"
        else:
            desc += f"基底高于现状地面最大约{max_above}m，现状地面下无地层信息"
        
        if hole_to_layer:
            desc += "；"
        else:
            if first_layers and is_fill(first_layers[0]) and fill_thicks.get(first_layers[0]):
                positive_thicks = [t for t in fill_thicks[first_layers[0]] if t > 0]
                if positive_thicks:
                    fill_thick = round(max(positive_thicks), 2)
                    desc += f"；基底分布地层为最大{fill_thick}m厚{main_first_name}{first_layers[0]}"
                else:
                    desc += f"；基底分布地层为{main_first_name}{first_layers[0]}，无有效填土厚度"
            elif first_layers:
                main_first_name = layer_info.get(first_layers[0], {}).get('name', '')
                desc += f"；基底分布地层为{main_first_name}{first_layers[0]}"
            desc += "。"
            return desc

    if hole_to_layer:
        layer_counts = Counter(hole_to_layer.values())
        if layer_counts:
            main_layer, main_count = layer_counts.most_common(1)[0]
            main_name = layer_info.get(main_layer, {}).get('name', '')
            main_desc = f"{main_name}{main_layer}"
            if is_fill(main_layer) and fill_thicks[main_layer]:
                positive_thicks = [t for t in fill_thicks[main_layer] if t > 0]
                if positive_thicks:
                    mt = round(max(positive_thicks), 2)
                    main_desc = f"最大{mt}m厚{main_desc}"
            if len(layer_counts) == 1:
                desc += f"基底分布地层为{main_desc}"
            else:
                desc += f"基底主要分布地层为{main_desc}"

            others = layer_counts.most_common()[1:]
            if others:
                total_holes = len(hole_to_layer)
                prop_others = (total_holes - main_count) / total_holes
                word = "个别" if prop_others < 0.3 else "部分"
                other_descs = []
                for l, _ in others:
                    n = layer_info.get(l, {}).get('name', '')
                    od = f"{n}{l}"
                    if is_fill(l) and fill_thicks[l]:
                        positive_thicks = [t for t in fill_thicks[l] if t > 0]
                        if positive_thicks:
                            mt = round(max(positive_thicks), 2)
                            od = f"最大{mt}m厚{od}"
                    other_descs.append(od)
                desc += f"，{word}地段分布{'、'.join(other_descs)}"

        desc += "。"

    if not desc:
        desc = "无有效数据"
    return desc


def is_high_rise(building, buildings):
    """判断是否为高层建筑"""
    info = buildings.get(building, {})
    floors, height = info.get('floors'), info.get('height')
    is_res = any(k in building.lower() for k in ['住宅', '公寓', '宿舍'])
    if floors is not None and floors >= 7:
        return True
    if height is not None and ((is_res and height >= 27) or (not is_res and height >= 24)):
        return True
    return False


def calculate_effective_fill_thick(hole_id, hole_strata, holes, base_elev):
    """计算有效填土厚度"""
    hole = holes[hole_id]
    if hole['elev'] is None:
        return 0.0
    embed_depth = hole['elev'] - base_elev
    strata = hole_strata.get(hole_id, [])

    if embed_depth < 0:
        ground_fill_thick = 0.0
        for l_id, bottom in strata:
            if not is_fill(l_id):
                break
            ground_fill_thick = bottom
        return -embed_depth + ground_fill_thick
    else:
        fill_bottom, found_base_in_fill = embed_depth, False
        for l_id, bottom in strata:
            if bottom < embed_depth:
                continue
            if not found_base_in_fill:
                if not is_fill(find_layer_at_depth(hole_strata, holes, hole_id, embed_depth)):
                    return 0.0
                found_base_in_fill = True
            if is_fill(l_id):
                fill_bottom = bottom
            else:
                break
        return fill_bottom - embed_depth


def needs_equivalent_modulus(building, holes, hole_strata, buildings):
    """判断是否需要进行当量模量计算"""
    if not is_high_rise(building, buildings):
        return False, False

    build_holes = [h for h, i in holes.items() if building in i.get('builds', [])]
    if not build_holes:
        return False, False

    base_elev = buildings[building]['embed_elev']
    fill_thicknesses = []
    has_fill = False

    for hole_id in build_holes:
        th = calculate_effective_fill_thick(hole_id, hole_strata, holes, base_elev)
        if th > 0:
            has_fill = True
        fill_thicknesses.append(th)

    if not has_fill:
        return True, False

    c_gt_3 = sum(1 for t in fill_thicknesses if t > 3.0)
    if c_gt_3 > 0:
        return False, False

    if max(fill_thicknesses) < 1.0:
        return True, True

    c_1_to_2 = sum(1 for t in fill_thicknesses if 1.0 <= t <= 2.0)
    c_2_to_3 = sum(1 for t in fill_thicknesses if 2.0 < t <= 3.0)

    if c_1_to_2 <= 2 and c_2_to_3 <= 1:
        return True, True

    return False, False


def calculate_effective_embed_depth(hole_id, hole_strata, holes, base_elev):
    """计算有效嵌入深度"""
    hole = holes[hole_id]
    if hole['elev'] is None:
        return 0.0
    embed_depth = hole['elev'] - base_elev

    base_layer = find_layer_at_depth(hole_strata, holes, hole_id, max(0, embed_depth))
    if not is_fill(base_layer):
        return max(0, embed_depth)

    current_depth = max(0, embed_depth)
    for l_id, bottom in hole_strata.get(hole_id, []):
        if bottom < current_depth:
            continue
        if is_fill(l_id):
            current_depth = bottom
        else:
            break
    return current_depth


def compute_equivalent_modulus(building, holes, hole_strata, buildings, layer_info, is_replace_mode=False):
    """计算当量模量"""
    build_holes_list = [h for h, info in holes.items() if building in info.get('builds', [])]
    if not build_holes_list:
        return None

    b_info = buildings[building]
    width, length = b_info.get('width'), b_info.get('length')
    if not all((width, length)):
        return None

    b, l_b = width / 2, length / width
    total_numerator, total_denominator, es_values = 0.0, 0.0, []
    base_elev = b_info['embed_elev']

    for hole_id in build_holes_list:
        if holes[hole_id].get('elev') is None:
            continue

        hole_elev = holes[hole_id]['elev']
        embed_depth = hole_elev - base_elev

        if is_replace_mode:
            effective_start_depth = calculate_effective_embed_depth(hole_id, hole_strata, holes, base_elev)
        else:
            effective_start_depth = max(0.0, embed_depth)

        hole_depth = holes[hole_id].get('max_depth')
        if hole_depth is None:
            continue

        layers_below, prev_depth = [], 0
        for layer_id, bottom in hole_strata.get(hole_id, []):
            effective_bottom = bottom if bottom is not None else hole_depth
            if effective_bottom is None:
                continue
            if effective_bottom > effective_start_depth:
                start = max(effective_start_depth, prev_depth)
                if is_replace_mode and is_fill(layer_id):
                    pass
                else:
                    if start < effective_bottom:
                        layers_below.append((layer_id, start, effective_bottom))
            prev_depth = effective_bottom

        if not layers_below:
            continue

        hole_numerator, hole_denominator, alpha_prev, zi_prev = 0.0, 0.0, 0.0, 0.0

        for layer_id, start, bottom in layers_below:
            zi = bottom - effective_start_depth
            z_b = zi / b if b != 0 else float('inf')
            alpha = bilinear_interpolate(ALPHA_TABLE, z_b, l_b)

            delta_sigma = alpha * zi - alpha_prev * zi_prev
            comp_mod = layer_info.get(layer_id, {}).get('compression_modulus')
            if comp_mod is None or comp_mod == 0:
                continue

            if delta_sigma != 0:
                hole_numerator += delta_sigma
                hole_denominator += delta_sigma / comp_mod

            alpha_prev, zi_prev = alpha, zi

        if hole_denominator != 0:
            es_hole = hole_numerator / hole_denominator
            es_values.append(es_hole)
            total_numerator += hole_numerator
            total_denominator += hole_denominator

    if not es_values or total_denominator == 0:
        return None

    es_building = total_numerator / total_denominator
    es_values.append(es_building)
    es_max, es_min = max(es_values), min(es_values)
    es_avg = sum(es_values) / len(es_values)
    es_ratio = es_max / es_min if es_min != 0 else float('inf')

    if es_avg <= 4:
        k = 1.3
    elif es_avg <= 7.5:
        k = linear_interpolate(4, 1.3, 7.5, 1.5, es_avg)
    elif es_avg <= 15:
        k = linear_interpolate(7.5, 1.5, 15, 1.8, es_avg)
    elif es_avg < 20:
        k = linear_interpolate(15, 1.8, 20, 2.5, es_avg)
    else:
        k = 2.5

    return {
        'es_max': round(es_max, 2),
        'es_min': round(es_min, 2),
        'es_avg': round(es_avg, 2),
        'es_ratio': round(es_ratio, 2),
        'k': round(k, 2),
        'uniformity': '均匀' if es_ratio <= k else '不均匀'
    }
