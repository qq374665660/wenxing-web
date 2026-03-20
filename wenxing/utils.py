# -*- coding: utf-8 -*-
"""
通用工具函数：插值、类型转换、填土判断等
"""

from .config import FILL_LAYER_PREFIX


def linear_interpolate(x1, y1, x2, y2, x):
    """一维线性插值"""
    if x1 == x2:
        return y1
    return y1 + (y2 - y1) * (x - x1) / (x2 - x1)


def bilinear_interpolate(table, z_b, l_b):
    """二维线性插值函数，用于 ALPHA_TABLE"""
    z_values = [float(row[0]) for row in table[1:] if row[0] is not None]
    l_values = [float(val) for val in table[0][1:] if val is not None]
    alpha_values = [[float(cell) if cell is not None else 0.0 for cell in row[1:]] for row in table[1:]]

    i = 0
    while i < len(z_values) - 1 and z_values[i] < z_b:
        i += 1
    i = max(0, min(i, len(z_values) - 2))

    j = 0
    while j < len(l_values) - 1 and l_values[j] < l_b:
        j += 1
    j = max(0, min(j, len(l_values) - 2))

    alpha11, alpha12 = alpha_values[i][j], alpha_values[i][j + 1]
    alpha21, alpha22 = alpha_values[i + 1][j], alpha_values[i + 1][j + 1]

    alpha1 = linear_interpolate(l_values[j], alpha11, l_values[j + 1], alpha12, l_b)
    alpha2 = linear_interpolate(l_values[j], alpha21, l_values[j + 1], alpha22, l_b)

    return linear_interpolate(z_values[i], alpha1, z_values[i + 1], alpha2, z_b)


def interpolate_alpha(alpha_data, z_b_values, l_b_values, z_b, l_b):
    """沉降计算用的 alpha 插值（基于字典格式的 alpha_data）"""
    z_b_sorted = sorted(z_b_values)
    l_b_sorted = sorted(l_b_values)

    # Find indices for z_b
    if z_b <= z_b_sorted[0]:
        z1 = z2 = z_b_sorted[0]
    elif z_b >= z_b_sorted[-1]:
        z1 = z2 = z_b_sorted[-1]
    else:
        for i in range(len(z_b_sorted) - 1):
            if z_b_sorted[i] <= z_b <= z_b_sorted[i + 1]:
                z1 = z_b_sorted[i]
                z2 = z_b_sorted[i + 1]
                break

    # Find indices for l_b
    if l_b <= l_b_sorted[0]:
        l1 = l2 = l_b_sorted[0]
    elif l_b >= l_b_sorted[-1]:
        l1 = l2 = l_b_sorted[-1]
    else:
        for i in range(len(l_b_sorted) - 1):
            if l_b_sorted[i] <= l_b <= l_b_sorted[i + 1]:
                l1 = l_b_sorted[i]
                l2 = l_b_sorted[i + 1]
                break

    if z1 == z2 and l1 == l2:
        return alpha_data[z1][l1]

    # Get the four points
    a11 = alpha_data[z1].get(l1, 0)
    a12 = alpha_data[z1].get(l2, 0)
    a21 = alpha_data[z2].get(l1, 0)
    a22 = alpha_data[z2].get(l2, 0)

    # Bilinear interpolation
    if z1 != z2 and l1 != l2:
        r1 = a11 + (a12 - a11) * (l_b - l1) / (l2 - l1)
        r2 = a21 + (a22 - a21) * (l_b - l1) / (l2 - l1)
        return r1 + (r2 - r1) * (z_b - z1) / (z2 - z1)
    elif z1 != z2:
        return a11 + (a21 - a11) * (z_b - z1) / (z2 - z1)
    elif l1 != l2:
        return a11 + (a12 - a11) * (l_b - l1) / (l2 - l1)
    else:
        return a11


def is_fill(layer_id):
    """判断地层是否为填土"""
    return layer_id and str(layer_id).startswith(FILL_LAYER_PREFIX)


def parse_optional_float(val, default=None):
    """安全地将值转换为浮点数"""
    if val is None or val == '' or str(val).strip() in {'/', '-', '—', '–', 'N/A', '无', '暂无', 'None', 'null'}:
        return default
    try:
        return float(str(val).strip().replace(',', ''))
    except (ValueError, TypeError):
        return default


def safe_float(val, default=0.0):
    """安全转换为浮点数，带默认值"""
    if val is None or val == '' or str(val).strip() in {'-', '/', '—', '－', '无', '暂无'}:
        return default
    try:
        return float(str(val).strip().replace(',', ''))
    except (ValueError, TypeError):
        return default
