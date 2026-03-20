# -*- coding: utf-8 -*-
"""
Excel 数据解析模块：建筑物、钻孔、地层信息解析
"""

import re
import logging
from collections import defaultdict

from .config import (
    COL_BUILDING_NAME, COL_BUILDING_FLOORS, COL_BUILDING_HEIGHT,
    COL_BUILDING_EMBED_ELEV, COL_BUILDING_LOAD, COL_BUILDING_WIDTH, COL_BUILDING_LENGTH
)
from .utils import parse_optional_float, safe_float


def parse_buildings(sheet):
    """解析建筑物数据（1.1工作表）"""
    logging.info("开始解析建筑物数据（1.1工作表）")

    buildings = {}
    buildings_list = []

    for row in sheet.iter_rows(min_row=6, values_only=True):
        if not row or len(row) <= COL_BUILDING_NAME or not row[COL_BUILDING_NAME]:
            continue

        name = str(row[COL_BUILDING_NAME]).strip()

        embed_elev_val = parse_optional_float(row[COL_BUILDING_EMBED_ELEV] if len(row) > COL_BUILDING_EMBED_ELEV else None)
        if embed_elev_val is None:
            logging.warning(f"建筑物 {name} 基础埋深标高无效，跳过")
            continue

        # 解析层数（处理范围值和空值）
        floors_raw = row[COL_BUILDING_FLOORS] if len(row) > COL_BUILDING_FLOORS else None
        floors = 1  # 默认值为 1

        if floors_raw is not None:
            raw_str = str(floors_raw).strip()
            if raw_str in {'', '/', '-', '—', 'None'}:
                floors = 1
            else:
                try:
                    if '~' in raw_str:
                        parts = raw_str.split('~')
                        floors = max([float(p) for p in parts if p.strip()])
                    elif '-' in raw_str:
                        parts = raw_str.split('-')
                        floors = max([float(p) for p in parts if p.strip()])
                    else:
                        floors = float(raw_str)
                    floors = int(floors)
                    if floors <= 0:
                        floors = 1
                except (ValueError, TypeError):
                    floors = 1

        height = parse_optional_float(row[COL_BUILDING_HEIGHT] if len(row) > COL_BUILDING_HEIGHT else None)

        if height is None or height <= 0:
            if floors is not None and floors > 0:
                height = floors * 3.5
                logging.info(f"建筑物 {name} 高度缺失，使用层数 {floors} × 3.5m = {height}m 自动补全")
            else:
                logging.info(f"建筑物 {name} 高度和层数均无效，高度保持 None")

        width = parse_optional_float(row[COL_BUILDING_WIDTH] if len(row) > COL_BUILDING_WIDTH else None)
        length = parse_optional_float(row[COL_BUILDING_LENGTH] if len(row) > COL_BUILDING_LENGTH else None)
        load = parse_optional_float(row[COL_BUILDING_LOAD] if len(row) > COL_BUILDING_LOAD else None)

        buildings[name] = {
            'embed_elev': embed_elev_val,
            'width': width,
            'length': length,
            'height': height,
            'floors': floors,
            'load': load
        }
        buildings_list.append(name)

    return buildings, buildings_list


def parse_holes(sheet, buildings_list):
    """解析钻孔数据（1.5单孔工作表）"""
    logging.info("开始解析钻孔数据（1.5单孔工作表）")
    holes = {}
    unmatched_buildings = set()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            hole_id = str(row[0]).strip()

            builds_str = row[12] if len(row) > 12 else ''
            assoc_builds = [b.strip() for b in re.split(r'[,、]', str(builds_str or '')) if b.strip()]

            for b in assoc_builds:
                if b and b not in buildings_list:
                    unmatched_buildings.add(b)

            holes[hole_id] = {
                'elev': parse_optional_float(row[1]),
                'max_depth': parse_optional_float(row[2]),
                'x': parse_optional_float(row[7]),
                'y': parse_optional_float(row[8]),
                'builds': assoc_builds
            }

    if unmatched_buildings:
        logging.warning(f"以下建筑物在 1.5单孔 中未匹配到 1.1 的建筑物: {unmatched_buildings}")

    return holes


def parse_layer_info(sheet):
    """解析地层信息（1.6地层信息工作表）"""
    logging.info("开始解析地层信息（1.6地层信息工作表）")
    layer_info = {}

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            layer_id = str(row[0]).strip()

            # 遇到 END 则跳过
            if layer_id.upper() == 'END':
                continue

            try:
                name = str(row[7]).strip().replace(' ', '').replace('\u3000', '') if len(row) > 7 and row[7] else ''
                state = str(row[8]).strip().replace(' ', '').replace('\u3000', '') if len(row) > 8 and row[8] else '/'

                layer_info[layer_id] = {
                    'name': name,
                    'state': state,
                    'bearing_capacity': safe_float(row[2]),
                    'compression_modulus': safe_float(row[3]),
                    'density': safe_float(row[9])
                }
                logging.info(f"Parsed layer {layer_id}: name={repr(name)}, fak={layer_info[layer_id]['bearing_capacity']}")
            except Exception as e:
                logging.warning(f"地层 {layer_id} 数据转换失败: {e}")

    return layer_info


def parse_hole_strata(sheet, holes):
    """解析各孔地层数据（2.4各孔地层工作表）"""
    logging.info("开始解析各孔地层数据（2.4各孔地层工作表）")
    hole_strata = defaultdict(list)
    current_hole = None

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row and str(row[0]).strip().upper() == 'END':
            break

        if row and row[0]:
            current_hole = str(row[0]).strip()

        if current_hole and row and row[1]:
            final_depth = None
            try:
                layer_id = str(row[1]).strip()
                raw_depth = row[2] if len(row) > 2 else None

                if raw_depth is not None and str(raw_depth).strip() != '':
                    parsed_depth = float(raw_depth)
                    if parsed_depth > 0:
                        final_depth = parsed_depth
                    else:
                        logging.warning(f"钻孔 {current_hole} 层 {layer_id} 深度值 {parsed_depth} 无效。")

                if final_depth is None:
                    max_depth = holes.get(current_hole, {}).get('max_depth')
                    if max_depth is not None and max_depth > 0:
                        final_depth = max_depth
                        logging.info(f"钻孔 {current_hole} 层 {layer_id} 深度为空或无效，使用孔深 {final_depth}m。")
                    else:
                        logging.warning(f"钻孔 {current_hole} 层 {layer_id} 无有效层底深度或孔深，跳过。")

                if final_depth is not None:
                    hole_strata[current_hole].append((layer_id, final_depth))

            except (ValueError, TypeError) as e:
                logging.warning(f"钻孔 {current_hole} 地层数据转换失败: {e}，跳过该行。")

    return dict(hole_strata)


def parse_chengdu_params(sheet):
    """解析成都地区地层参数表"""
    param_dict = {}
    for row in sheet.iter_rows(min_row=5, values_only=True):
        if row[3]:
            name = str(row[3]).replace(' ', '').replace('\u3000', '').strip()
            state = str(row[4] or '/').replace(' ', '').replace('\u3000', '').strip()
            key = (name, state)
            param_dict[key] = {
                'cfg_qsia': row[6] if row[6] is not None else '/',
                'cfg_qpa': row[7] if row[7] is not None else '/',
                'jet_qsia': row[8] if row[8] is not None else '/',
                'jet_qpa': row[9] if row[9] is not None else '/',
                'name': row[3] if row[3] is not None else '/',
                'state': row[4] if row[4] is not None else '/',
                'dry_qsik': row[10] if row[10] is not None else '/',
                'dry_qpk': row[11] if row[11] is not None else '/',
                'mud_qsik': row[12] if row[12] is not None else '/',
                'mud_qpk': row[13] if row[13] is not None else '/',
                'pre_qsik': row[14] if row[14] is not None else '/',
                'pre_qpk': row[15] if row[15] is not None else '/',
                'rc_natural': row[16] if row[16] is not None else '/',
                'rc_saturated': row[17] if row[17] is not None else '/',
                'm_coeff': row[18] if row[18] is not None else '/',
                'k_coeff': row[19] if row[19] is not None else '/'
            }
    return param_dict
