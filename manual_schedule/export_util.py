import pandas as pd
from .manual_state import ManualSession

def export_full(session: ManualSession, path: str):
    return session.export_excel(path)
