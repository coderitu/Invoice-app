import pandas as pd

REQUIRED_COLUMNS = [
    "Member",
    "Name",
    "Email",
    "Renewal",
    "Invoice no.",
    "Pending 2025",
    "Pending 2024",
    "Pending 2023"
]

def load_excel(path):
    df = pd.read_excel(path)

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {', '.join(missing)}")

    return df.fillna(0)
