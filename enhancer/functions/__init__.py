"""
Domain UDF functions package for Excel Enhancer.
"""

from enhancer.functions.array_ops import (
    xf_fit_table,
    xl_concat,
    xl_expand_down,
    xl_expand_right,
    xl_expand_table,
)
from enhancer.functions.combinatorics import (
    xf_permutations,
    xl_combinations,
)
from enhancer.functions.expressions import (
    xl_format_by,
    xlv_apply_expr,
    xlv_apply_str,
    xlv_map_by,
)
from enhancer.functions.similarity import (
    custom_find_similar,
    mean_validation,
    rmse_validation,
    similar_std_instance,
    xl_find_similar,
    xv_find_similar,
)
from enhancer.functions.table_ops import (
    build_group_df,
    custom_group_by,
    operation_resume,
    xl_filter_by,
    xl_group_by,
    xl_rename,
    xl_select_by,
    xl_sort_by,
)
from enhancer.functions.utilities import (
    xf_collect_unittest,
    xf_delete_library,
    xf_delete_unittest,
    xf_evaluate,
    xf_get_caller_address,
    xf_get_selection_range,
    xf_load_library,
    xf_make_array,
    xf_print_unittest,
    xf_recalculate_udfs,
    xf_show_formula,
    xf_test_data,
)
from enhancer.functions.vector_ops import (
    xv_concat_h,
    xv_concat_v,
    xv_index,
    xv_normalize_accents,
    xv_normalize_to_string,
)

__all__ = [
    # Array operations
    "xl_expand_down",
    "xl_expand_right",
    "xl_expand_table",
    "xf_fit_table",
    "xl_concat",
    # Vector operations
    "xv_concat_h",
    "xv_concat_v",
    "xv_index",
    "xv_normalize_accents",
    "xv_normalize_to_string",
    # Table operations
    "build_group_df",
    "custom_group_by",
    "operation_resume",
    "xl_group_by",
    "xl_filter_by",
    "xl_sort_by",
    "xl_select_by",
    "xl_rename",
    # Expressions
    "xlv_map_by",
    "xl_format_by",
    "xlv_apply_expr",
    "xlv_apply_str",
    # Similarity
    "similar_std_instance",
    "custom_find_similar",
    "mean_validation",
    "rmse_validation",
    "xv_find_similar",
    "xl_find_similar",
    # Combinatorics
    "xl_combinations",
    "xf_permutations",
    # Utilities
    "xf_test_data",
    "xf_print_unittest",
    "xf_delete_unittest",
    "xf_collect_unittest",
    "xf_show_formula",
    "xf_evaluate",
    "xf_make_array",
    "xf_get_selection_range",
    "xf_get_caller_address",
    "xf_recalculate_udfs",
    "xf_load_library",
    "xf_delete_library",
]

