"""Normalize employee labels without bundling an organization's staff directory."""
import re


def canonical_employee(value):
    raw = str(value or '').strip()
    parts = re.split(r'[,;\n/&+]+|\s+and\s+', raw, flags=re.IGNORECASE)
    return ', '.join(dict.fromkeys(part.strip() for part in parts if part.strip()))
