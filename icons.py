"""Lucide icons (https://lucide.dev), lucide-static v0.469.0, ISC License.

Vendored as static SVG markup so the app needs no extra dependency or network
access. Only trusted, hard-coded markup lives here; never interpolate user input.
"""

LUCIDE_VERSION = "0.469.0"

# Inner SVG elements keyed by Lucide icon name.
_PATHS = {
    "calendar-days": (
        '<path d="M8 2v4" />'
        '<path d="M16 2v4" />'
        '<rect width="18" height="18" x="3" y="4" rx="2" />'
        '<path d="M3 10h18" />'
        '<path d="M8 14h.01" />'
        '<path d="M12 14h.01" />'
        '<path d="M16 14h.01" />'
        '<path d="M8 18h.01" />'
        '<path d="M12 18h.01" />'
        '<path d="M16 18h.01" />'
    ),
    "chart-bar": (
        '<path d="M3 3v16a2 2 0 0 0 2 2h16" />'
        '<path d="M7 16h8" />'
        '<path d="M7 11h12" />'
        '<path d="M7 6h3" />'
    ),
    "chart-column": (
        '<path d="M3 3v16a2 2 0 0 0 2 2h16" />'
        '<path d="M18 17V9" />'
        '<path d="M13 17V5" />'
        '<path d="M8 17v-3" />'
    ),
    "chart-no-axes-combined": (
        '<path d="M12 16v5" />'
        '<path d="M16 14v7" />'
        '<path d="M20 10v11" />'
        '<path d="m22 3-8.646 8.646a.5.5 0 0 1-.708 0L9.354 8.354a.5.5 0 0 0-.707 0L2 15" />'
        '<path d="M4 18v3" />'
        '<path d="M8 14v7" />'
    ),
    "circle-alert": (
        '<circle cx="12" cy="12" r="10" />'
        '<line x1="12" x2="12" y1="8" y2="12" />'
        '<line x1="12" x2="12.01" y1="16" y2="16" />'
    ),
    "circle-check": (
        '<circle cx="12" cy="12" r="10" />'
        '<path d="m9 12 2 2 4-4" />'
    ),
    "database": (
        '<ellipse cx="12" cy="5" rx="9" ry="3" />'
        '<path d="M3 5V19A9 3 0 0 0 21 19V5" />'
        '<path d="M3 12A9 3 0 0 0 21 12" />'
    ),
    "info": (
        '<circle cx="12" cy="12" r="10" />'
        '<path d="M12 16v-4" />'
        '<path d="M12 8h.01" />'
    ),
    "layers": (
        '<path d="M12.83 2.18a2 2 0 0 0-1.66 0L2.6 6.08a1 1 0 0 0 0 1.83l8.58 3.91a2 2 0 0 0 1.66 0l8.58-3.9a1 1 0 0 0 0-1.83z" />'
        '<path d="M2 12a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 12" />'
        '<path d="M2 17a1 1 0 0 0 .58.91l8.6 3.91a2 2 0 0 0 1.65 0l8.58-3.9A1 1 0 0 0 22 17" />'
    ),
    "layout-dashboard": (
        '<rect width="7" height="9" x="3" y="3" rx="1" />'
        '<rect width="7" height="5" x="14" y="3" rx="1" />'
        '<rect width="7" height="9" x="14" y="12" rx="1" />'
        '<rect width="7" height="5" x="3" y="16" rx="1" />'
    ),
    "list-checks": (
        '<path d="m3 17 2 2 4-4" />'
        '<path d="m3 7 2 2 4-4" />'
        '<path d="M13 6h8" />'
        '<path d="M13 12h8" />'
        '<path d="M13 18h8" />'
    ),
    "map-pin": (
        '<path d="M20 10c0 4.993-5.539 10.193-7.399 11.799a1 1 0 0 1-1.202 0C9.539 20.193 4 14.993 4 10a8 8 0 0 1 16 0" />'
        '<circle cx="12" cy="10" r="3" />'
    ),
    "rotate-ccw": (
        '<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />'
        '<path d="M3 3v5h5" />'
    ),
    "sliders-horizontal": (
        '<line x1="21" x2="14" y1="4" y2="4" />'
        '<line x1="10" x2="3" y1="4" y2="4" />'
        '<line x1="21" x2="12" y1="12" y2="12" />'
        '<line x1="8" x2="3" y1="12" y2="12" />'
        '<line x1="21" x2="16" y1="20" y2="20" />'
        '<line x1="12" x2="3" y1="20" y2="20" />'
        '<line x1="14" x2="14" y1="2" y2="6" />'
        '<line x1="8" x2="8" y1="10" y2="14" />'
        '<line x1="16" x2="16" y1="18" y2="22" />'
    ),
    "table": (
        '<path d="M12 3v18" />'
        '<rect width="18" height="18" x="3" y="3" rx="2" />'
        '<path d="M3 9h18" />'
        '<path d="M3 15h18" />'
    ),
    "triangle-alert": (
        '<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3" />'
        '<path d="M12 9v4" />'
        '<path d="M12 17h.01" />'
    ),
    "trophy": (
        '<path d="M6 9H4.5a2.5 2.5 0 0 1 0-5H6" />'
        '<path d="M18 9h1.5a2.5 2.5 0 0 0 0-5H18" />'
        '<path d="M4 22h16" />'
        '<path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22" />'
        '<path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22" />'
        '<path d="M18 2H6v7a6 6 0 0 0 12 0V2Z" />'
    ),
    "users": (
        '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />'
        '<circle cx="9" cy="7" r="4" />'
        '<path d="M22 21v-2a4 4 0 0 0-3-3.87" />'
        '<path d="M16 3.13a4 4 0 0 1 0 7.75" />'
    ),
    "wallet": (
        '<path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1" />'
        '<path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4" />'
    ),
}


def icon(name: str, size: int = 18, color: str = "currentColor", stroke_width: float = 2) -> str:
    """Return inline SVG markup for a vendored Lucide icon (KeyError if unknown)."""
    return (
        f'<svg class="lc-icon" xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" '
        f'viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="{stroke_width}" '
        f'stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{_PATHS[name]}</svg>'
    )
