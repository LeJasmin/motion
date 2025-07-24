import os
from pathlib import Path
from datetime import datetime
import hashlib
import shutil
import re

# Constants
INBOX = Path('Inbox-tris-docs')
DEST_ROOT = Path('Documents triés')

CATEGORY_PREFIXES = {
    'ADM': 'Administratif',
    'THS': 'Thèse',
    'ENS': 'Enseignement',
}

MARKDOWN_FILES = {
    'Administratif': 'administratif.md',
    'Thèse': 'these.md',
    'Enseignement': 'enseignement.md',
}

VALID_EXTENSIONS = {'.pdf', '.doc', '.docx', '.jpg', '.jpeg', '.png', '.mp4'}


def is_date(text: str) -> bool:
    """Return True if text matches YYYY-MM-DD."""
    return bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", text))


def file_hash(path: Path) -> str:
    """Compute md5 hash of a file."""
    h = hashlib.md5()
    with path.open('rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def detect_category(filename: str):
    """Return category name and prefix based on filename."""
    upper = filename.upper()
    for prefix, cat in CATEGORY_PREFIXES.items():
        if upper.startswith(prefix):
            return cat, prefix
    if 'THESE' in upper or 'THS' in upper:
        return 'Thèse', 'THS'
    if 'ENSEIGNEMENT' in upper or 'ENS' in upper:
        return 'Enseignement', 'ENS'
    return 'Administratif', 'ADM'


def extract_subcategory(stem: str, prefix: str) -> str:
    """Extract subcategory from filename."""
    parts = stem.split('_')
    if parts and parts[0].upper() == prefix:
        parts = parts[1:]
    if parts and not is_date(parts[0]):
        return parts[0]
    return 'Divers'


def extract_date(stem: str, src: Path) -> str:
    """Extract date from filename or file metadata."""
    match = re.search(r"\d{4}-\d{2}-\d{2}", stem)
    if match:
        return match.group(0)
    dt = datetime.fromtimestamp(src.stat().st_mtime)
    return dt.strftime('%Y-%m-%d')


def sanitize_name(stem: str, prefix: str) -> str:
    """Remove prefix and dates from stem."""
    parts = [p for p in stem.split('_') if p]
    if parts and parts[0].upper() == prefix:
        parts = parts[1:]
    parts = [p for p in parts if not is_date(p)]
    return '_'.join(parts)


def move_file(src: Path):
    """Move a single file to the organized directory."""
    if src.suffix.lower() not in VALID_EXTENSIONS:
        return
    if src.stat().st_size == 0:
        src.unlink()
        return

    category, prefix = detect_category(src.name)
    stem = src.stem
    subcategory = extract_subcategory(stem, prefix)
    date = extract_date(stem, src)
    year = date.split('-')[0]
    original = sanitize_name(stem, prefix)

    new_name = f"{prefix}_{date}_{original}{src.suffix.lower()}"

    dest_dir = DEST_ROOT / category / subcategory / year
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / new_name

    if dest.exists() and file_hash(dest) == file_hash(src):
        src.unlink()
        return

    counter = 1
    base_name = dest.stem
    while dest.exists():
        dest = dest_dir / f"{base_name}-{counter}{src.suffix.lower()}"
        counter += 1

    shutil.move(str(src), str(dest))


def organize_inbox():
    """Process all files in the inbox directory."""
    if not INBOX.exists():
        return
    for item in INBOX.iterdir():
        if item.is_file():
            move_file(item)


def generate_markdown():
    """Generate markdown index files for each category."""
    for category, md_name in MARKDOWN_FILES.items():
        cat_dir = DEST_ROOT / category
        if not cat_dir.exists():
            continue
        lines = [f"# {category}\n\n"]
        for subcat in sorted(p for p in cat_dir.iterdir() if p.is_dir()):
            years = sorted([y for y in subcat.iterdir() if y.is_dir()])
            for year in years:
                lines.append(f"## {subcat.name} - {year.name}\n")
                files = sorted(f for f in year.iterdir() if f.is_file())
                for fpath in files:
                    rel = fpath.relative_to(cat_dir)
                    lines.append(f"- [{fpath.name}](./{rel.as_posix()})\n")
                lines.append("\n")
        (cat_dir / md_name).write_text(''.join(lines))


def main():
    organize_inbox()
    generate_markdown()


if __name__ == '__main__':
    main()
