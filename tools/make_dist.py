r"""配布用最小パッケージを作るスクリプト.

実行: .venv\Scripts\python.exe tools\make_dist.py

作られるもの:
  dist\pdf-split-unified\            … 動作に必要なファイルだけを集めた受け取り用フォルダー
  dist\pdf-split-unified-v<版>-min.zip … 上のフォルダーをそのまま収めたzip（メール等での配布用）

含める8ファイル（これだけで動作する。tools\samples\docs\refs\テストは配布しない）:
  pdf_split_unified.py / split_core.py / modes.py / preview.py
  起動.bat / README.md / LICENSE / LICENSE.ja.md
.gitignore・.venv・__pycache__ は含めない（.venvはこのPC専用で、受け取り側では 起動.bat がライブラリを入れる）。
"""

from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

VERSION = "1.0.0"
PACKAGE_NAME = "pdf-split-unified"
DIST_DIR = ROOT / "dist"
STAGING = DIST_DIR / PACKAGE_NAME
ZIP_PATH = DIST_DIR / f"{PACKAGE_NAME}-v{VERSION}-min.zip"

# 動作に必須の4ファイル + 起動用の bat + 案内とライセンス
PACKAGE_FILES = [
    "pdf_split_unified.py",
    "split_core.py",
    "modes.py",
    "preview.py",
    "起動.bat",
    "README.md",
    "LICENSE",
    "LICENSE.ja.md",
]


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> int:
    missing = [name for name in PACKAGE_FILES if not (ROOT / name).is_file()]
    if missing:
        print(f"FAIL: 元ファイルが見つかりません: {' / '.join(missing)}")
        return 1

    if STAGING.exists():
        shutil.rmtree(STAGING)
    STAGING.mkdir(parents=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()

    for name in PACKAGE_FILES:
        shutil.copy2(ROOT / name, STAGING / name)
    shutil.make_archive(str(DIST_DIR / f"{PACKAGE_NAME}-v{VERSION}-min"), "zip",
                        root_dir=str(DIST_DIR), base_dir=PACKAGE_NAME)

    with zipfile.ZipFile(ZIP_PATH) as archive:
        names = sorted(name for name in archive.namelist() if not name.endswith("/"))
        broken = archive.testzip()
    expected = sorted(f"{PACKAGE_NAME}/{name}" for name in PACKAGE_FILES)
    print(f"created: {STAGING.relative_to(ROOT)}\\（{len(PACKAGE_FILES)}ファイル / "
          f"{sum(p.stat().st_size for p in STAGING.iterdir()):,} バイト）")
    print(f"created: {ZIP_PATH.relative_to(ROOT)}（{ZIP_PATH.stat().st_size:,} バイト / "
          f"SHA-256 {sha256(ZIP_PATH)[:16]}…）")
    for name in names:
        if not name.endswith("/"):
            print(f"      - {name}")
    ok = broken is None and names == expected
    print(("PASS: " if ok else "FAIL: ") + f"zip内は8ファイルで一致・破損なし（{len(names)} エントリ）")
    if not ok:
        print(f"      期待: {' / '.join(expected)}")
        print(f"      実際: {' / '.join(names)}")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
