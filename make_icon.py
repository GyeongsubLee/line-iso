# -*- coding: utf-8 -*-
"""
프로그램 안에 들어 있는 아이콘(_V33_ICON_BASE64)을 exe 파일 아이콘(.ico)으로 저장한다.

사용법:
    py make_icon.py line_iso_desktop_tool_V1.40.py app.ico
"""

import base64
import io
import re
import sys
from pathlib import Path

from PIL import Image


def main():
    if len(sys.argv) != 3:
        raise SystemExit("사용법: py make_icon.py <프로그램.py> <저장할.ico>")

    source = Path(sys.argv[1]).read_text(encoding="utf-8-sig")
    m = re.search(r'_V33_ICON_BASE64\s*=\s*"""(.*?)"""', source, flags=re.S)
    if not m:
        raise SystemExit("[중단] 프로그램에서 _V33_ICON_BASE64 아이콘을 찾지 못했습니다.")

    image = Image.open(io.BytesIO(base64.b64decode("".join(m.group(1).split())))).convert("RGBA")

    # 정사각형 투명 캔버스 가운데에 놓은 뒤 256px로 키워 여러 크기의 ico로 저장한다.
    side = max(image.size)
    square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    square.paste(image, ((side - image.width) // 2, (side - image.height) // 2))
    square = square.resize((256, 256), Image.LANCZOS)
    square.save(sys.argv[2], format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"[완료] 아이콘 저장: {sys.argv[2]}")


if __name__ == "__main__":
    main()
