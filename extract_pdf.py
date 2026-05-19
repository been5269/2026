import sys
from pathlib import Path


def main() -> int:
    pdf_path = Path(r"c:\Users\been5\OneDrive\바탕 화면\앱\영어기능지도_개발계획서(5월6일 수정).pdf")
    out_path = Path(r"c:\Users\been5\OneDrive\바탕 화면\앱\영어기능지도_개발계획서(5월6일 수정).txt")

    try:
        import pypdf  # type: ignore

        reader = pypdf.PdfReader(str(pdf_path))
        chunks: list[str] = []
        for i, page in enumerate(reader.pages):
            try:
                txt = page.extract_text() or ""
            except Exception as e:  # noqa: BLE001
                txt = f"[page {i+1} extract error: {e}]"
            chunks.append(txt)
        out_path.write_text("\n\n".join(chunks), encoding="utf-8")
        return 0
    except Exception:
        # Fallback: dump raw bytes as a last resort (for debugging)
        out_path.write_bytes(pdf_path.read_bytes())
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

