"""Notify dashboard of stage completion (silent no-op if dashboard unavailable)."""

def notify(stage_id, status="completed", date_str=None, extra=None):
    """Update dashboard progress state. Safe to call from any context."""
    try:
        import sys
        from pathlib import Path
        root = Path(__file__).resolve().parent.parent
        if str(root) not in sys.path:
            sys.path.insert(0, str(root))
        from dashboard.utils import mark_stage
        mark_stage(stage_id, status, extra=extra, date_str=date_str)
    except Exception:
        pass
