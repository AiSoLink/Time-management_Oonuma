"""
分割休息モジュール（オーケストレータ）
①→②の順で実行。既存の run_split_rest 呼び出しを維持。
"""

import os
from typing import List, Optional

from split_rest_1 import process_split_rest_step1
from split_rest_2 import process_split_rest_step2


def run_split_rest(
    monthly_data: List[List],
    work_detail_path: Optional[str] = None,
) -> List[List]:
    """
    分割休息①→②を適用した結果を返す。

    Args:
        monthly_data: [ヘッダー, データ行...] の統合データ
        work_detail_path: 作業明細ファイルパス（Noneなら②はスキップ）

    Returns:
        更新後の行リスト
    """
    rows = [list(r) for r in monthly_data]
    process_split_rest_step1(rows)
    if work_detail_path and os.path.isfile(work_detail_path):
        process_split_rest_step2(rows, work_detail_path)
    return rows
