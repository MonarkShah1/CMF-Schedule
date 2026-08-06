#!/usr/bin/env python3
"""
cmf_dedupe_media.py  —  Shrink the saved workbook by de-duplicating screenshots.

Why this exists
---------------
openpyxl writes one media part per Image object. The builder places the same
screenshot on the MAIN row, on every CALENDAR card and on every PURCHASING card
for that part, so a single photo is stored dozens of times — measured at 920
media parts for 139 unique images (85% redundant, 18.5 MB instead of 4.1 MB).

A workbook that size is slow to open, slow to sync to OneDrive, and prone to
sync conflicts when two people touch it the same day.

This pass runs after save: it hashes every media part, keeps one copy of each
distinct image, repoints all drawing relationships at the survivor, and drops
the redundant parts. Pixels are untouched — only storage changes.

Usage:
    python3 cmf_dedupe_media.py                    # dedupe the working file
    python3 cmf_dedupe_media.py path/to/file.xlsx
"""

import os
import re
import shutil
import sys
import zipfile
from collections import defaultdict
from hashlib import md5

FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CMF WIP - Schedule.xlsx')

_MEDIA_RE = re.compile(r"^xl/media/")


def _canonical_media_map(zf):
    """{duplicate_part_name: canonical_part_name} for byte-identical media."""
    by_digest = defaultdict(list)
    for name in zf.namelist():
        if _MEDIA_RE.match(name):
            by_digest[md5(zf.read(name)).hexdigest()].append(name)

    mapping = {}
    for names in by_digest.values():
        # Keep the lowest-numbered part so output ordering stays deterministic.
        canonical = min(names, key=lambda n: (len(n), n))
        for name in names:
            if name != canonical:
                mapping[name] = canonical
    return mapping, by_digest


def _rewrite_rels(data, mapping):
    """Repoint ../media/imageN.ext targets at the canonical part."""
    text = data.decode("utf-8")
    for dup, canonical in mapping.items():
        dup_base = dup.split("/")[-1]
        canon_base = canonical.split("/")[-1]
        if dup_base == canon_base:
            continue
        # Targets appear as Target="../media/image12.png"
        text = text.replace(f'/media/{dup_base}"', f'/media/{canon_base}"')
    return text.encode("utf-8")


def dedupe(path=None, verbose=True):
    path = path or FILE
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    with zipfile.ZipFile(path) as zf:
        mapping, by_digest = _canonical_media_map(zf)
        total_media = sum(len(v) for v in by_digest.values())

        if not mapping:
            if verbose:
                print(f"  MEDIA-DEDUPE: {total_media} image(s), no duplicates — nothing to do")
            return False

        entries = []
        for info in zf.infolist():
            if info.filename in mapping:
                continue  # drop the redundant copy
            data = zf.read(info.filename)
            if info.filename.endswith(".rels"):
                data = _rewrite_rels(data, mapping)
            entries.append((info, data))

    tmp = path + ".dedupe.tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            # Preserve the original timestamp/compression metadata.
            new_info = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            new_info.compress_type = zipfile.ZIP_DEFLATED
            new_info.external_attr = info.external_attr
            out.writestr(new_info, data)

    before = os.path.getsize(path)
    os.replace(tmp, path)
    after = os.path.getsize(path)

    if verbose:
        print(f"  MEDIA-DEDUPE: {total_media} → {len(by_digest)} image part(s) "
              f"({len(mapping)} duplicate(s) removed)")
        print(f"  MEDIA-DEDUPE: {before/1e6:.1f} MB → {after/1e6:.1f} MB "
              f"({100 * (before - after) / before:.0f}% smaller)")
    return True


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else FILE
    backup = target + ".pre-dedupe"
    shutil.copy2(target, backup)
    try:
        dedupe(target)
    except Exception:
        shutil.copy2(backup, target)
        raise
    finally:
        if os.path.exists(backup):
            os.remove(backup)
