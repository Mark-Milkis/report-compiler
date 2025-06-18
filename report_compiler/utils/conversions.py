"""
Utility functions for unit conversions.
"""

def points_to_inches(points: float) -> float:
    """Converts points to inches."""
    return points / 72.0

def emu_to_points(emu: int) -> float:
    """Converts English Metric Units (EMU) to points."""
    # 1 point = 12700 EMU
    return emu / 12700.0
