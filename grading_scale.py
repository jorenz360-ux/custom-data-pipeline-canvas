"""
The 0-100 -> 1.0-5.0 grade-equivalent lookup table, copied exactly from the
'Grading' sheet in your original workbook (VLOOKUP(percent, Grading!A:B, 2, TRUE)).
Index = whole-number percent (0-100), value = grade equivalent.
"""

TABLE = [
    5.0, 4.9, 4.9, 4.9, 4.8, 4.8, 4.7, 4.7, 4.6, 4.6,  # 0-9
    4.5, 4.5, 4.4, 4.4, 4.3, 4.3, 4.2, 4.2, 4.1, 4.1,  # 10-19
    4.0, 4.0, 4.0, 3.9, 3.9, 3.9, 3.8, 3.8, 3.8, 3.7,  # 20-29
    3.7, 3.7, 3.6, 3.6, 3.6, 3.5, 3.5, 3.5, 3.4, 3.4,  # 30-39
    3.4, 3.3, 3.3, 3.3, 3.2, 3.2, 3.1, 3.1, 3.0, 3.0,  # 40-49
    3.0, 3.0, 2.9, 2.9, 2.8, 2.8, 2.7, 2.7, 2.6, 2.6,  # 50-59
    2.5, 2.5, 2.4, 2.4, 2.3, 2.3, 2.2, 2.2, 2.1, 2.1,  # 60-69
    2.0, 2.0, 2.0, 1.9, 1.9, 1.9, 1.8, 1.8, 1.8, 1.7,  # 70-79
    1.7, 1.7, 1.6, 1.6, 1.6, 1.5, 1.5, 1.5, 1.4, 1.4,  # 80-89
    1.4, 1.3, 1.3, 1.2, 1.2, 1.2, 1.1, 1.1, 1.1, 1.0,  # 90-99
    1.0,  # 100-100
]
assert len(TABLE) == 101, f"expected 101 entries, got {len(TABLE)}"


def grade_equivalent(percent):
    """VLOOKUP(percent, table, TRUE) behavior: clamps into [0, 100] and rounds
    to the nearest whole percent (callers already hand this a whole number from
    ROUNDUP/ROUNDDOWN, matching the original sheet)."""
    if percent is None:
        return None
    p = int(round(percent))
    p = max(0, min(100, p))
    return TABLE[p]
