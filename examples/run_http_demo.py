"""Exercise path and upload inputs against an already running local HTTP service."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from urllib.request import Request, urlopen
from uuid import uuid4

from run_multipage_demo import demo_payload


def post(url: str, body: bytes, content_type: str) -> str:
    request = Request(url, data=body, headers={"Content-Type": content_type}, method="POST")
    with urlopen(request, timeout=120) as response:
        return json.load(response)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", default="http://127.0.0.1:8000/api/v1/presentations")
    parser.add_argument(
        "--template", type=Path, default=root / "tests/fixtures/test-multipage.pptx"
    )
    args = parser.parse_args()
    data = json.dumps(demo_payload(), ensure_ascii=False)
    path = post(
        args.url,
        json.dumps({"template": str(args.template.resolve()), "data": data}).encode(),
        "application/json",
    )
    boundary = uuid4().hex
    upload = b"".join(
        (
            f'--{boundary}\r\nContent-Disposition: form-data; name="data"\r\n\r\n'.encode(),
            data.encode(),
            f'\r\n--{boundary}\r\nContent-Disposition: form-data; name="template"; '
            'filename="report.pptx"\r\nContent-Type: application/vnd.openxmlformats-officedocument.'
            "presentationml.presentation\r\n\r\n".encode(),
            args.template.read_bytes(),
            f"\r\n--{boundary}--\r\n".encode(),
        )
    )
    file = post(args.url, upload, f"multipart/form-data; boundary={boundary}")
    print(json.dumps({"path_request": path, "upload_request": file}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
