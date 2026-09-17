import html
import os
import shutil
from pathlib import Path


def build():
    key = os.environ.get("CARTO_PUBLIC_BASEMAP_KEY", "").strip()
    if not key:
        raise SystemExit("Set CARTO_PUBLIC_BASEMAP_KEY in the hosting provider's build environment to a deliberately public browser token.")
    if os.environ.get("CARTO_KEY_PUBLIC_CONFIRMED", "").strip().lower() != "yes":
        raise SystemExit(
            "Refusing to publish CARTO_PUBLIC_BASEMAP_KEY: it becomes public in served HTML. "
            "Verify basemap-only access and provider-side restrictions for the deployed origins, "
            "then set CARTO_KEY_PUBLIC_CONFIRMED=yes. Never use a confidential credential."
        )

    root = Path(__file__).resolve().parent.parent
    dashboard = root / "hand_sharepoint_dashboard"
    marker = '<meta name="carto-api-key" content="" />'
    content = (dashboard / "index.html").read_text(encoding="utf-8")
    if content.count(marker) != 1:
        raise SystemExit("Expected exactly one CARTO configuration placeholder.")
    content = content.replace(marker, '<meta name="carto-api-key" content="' + html.escape(key, quote=True) + '" />')

    output = root / "_site"
    if output.exists():
        shutil.rmtree(output)
    output.mkdir()
    shutil.copy2(root / "index.html", output / "index.html")
    shutil.copytree(dashboard, output / "hand_sharepoint_dashboard")
    (output / "hand_sharepoint_dashboard/index.html").write_text(content, encoding="utf-8")
    print("Public site prepared in _site.")


if __name__ == "__main__":
    build()
