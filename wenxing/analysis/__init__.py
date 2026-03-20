# -*- coding: utf-8 -*-
"""
分析模块：地基均匀性、承载力、沉降等计算
"""

from .uniformity import (
    find_layer_at_depth,
    compute_desc,
    is_high_rise,
    calculate_effective_fill_thick,
    needs_equivalent_modulus,
    calculate_effective_embed_depth,
    compute_equivalent_modulus
)

from .bearing_capacity import (
    get_base_layers,
    get_eta_params,
    calculate_fa,
    get_weak_underlayers,
    calculate_theta
)

from .settlement import (
    get_under_layers,
    calculate_settlement,
    get_tilt_limit
)

__all__ = [
    # uniformity
    'find_layer_at_depth', 'compute_desc', 'is_high_rise',
    'calculate_effective_fill_thick', 'needs_equivalent_modulus',
    'calculate_effective_embed_depth', 'compute_equivalent_modulus',
    # bearing_capacity
    'get_base_layers', 'get_eta_params', 'calculate_fa',
    'get_weak_underlayers', 'calculate_theta',
    # settlement
    'get_under_layers', 'calculate_settlement', 'get_tilt_limit'
]
