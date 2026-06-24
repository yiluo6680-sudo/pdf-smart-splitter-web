"""PDF 智能拆分工具的高精度 OCR 服务。"""

from __future__ import annotations

import io
import threading
from pathlib import Path

import numpy as np
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from PIL import Image, UnidentifiedImageError
from paddleocr import PaddleOCR


ROOT = Path(__file__).resolve().parent
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_IMAGE_PIXELS = 30_000_000

app = FastAPI(title="PDF 智能拆分 OCR", version="2.0.0")
_ocr: PaddleOCR | None = None
_ocr_lock = threading.Lock()


def get_ocr() -> PaddleOCR:
    """延迟加载模型，并确保并发请求不会重复初始化。"""
    global _ocr
    if _ocr is None:
        with _ocr_lock:
            if _ocr is None:
                _ocr = PaddleOCR(
                    use_doc_orientation_classify=False,
                    use_doc_unwarping=False,
                    use_textline_orientation=False,
                    engine="onnxruntime",
                    engine_config={
                        "providers": ["CPUExecutionProvider"],
                        "intra_op_num_threads": 4,
                    },
                )
    return _ocr


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok", "engine": "PP-OCRv6_medium"}


@app.post("/api/ocr")
async def recognize(image: UploadFile = File(...)) -> dict:
    if image.content_type not in {"image/jpeg", "image/png", "image/webp"}:
        raise HTTPException(status_code=415, detail="仅支持 JPEG、PNG 或 WebP 图片")
    payload = await image.read(MAX_UPLOAD_BYTES + 1)
    if len(payload) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="分界页图片超过 15 MB")

    try:
        with Image.open(io.BytesIO(payload)) as source:
            source.load()
            if source.width * source.height > MAX_IMAGE_PIXELS:
                raise HTTPException(status_code=413, detail="分界页图片尺寸过大")
            rgb = np.asarray(source.convert("RGB"))
    except UnidentifiedImageError as error:
        raise HTTPException(status_code=400, detail="无法读取上传的图片") from error

    try:
        result = next(iter(get_ocr().predict(rgb))).json["res"]
    except Exception as error:
        raise HTTPException(status_code=500, detail="高精度 OCR 识别失败") from error

    texts = [str(value) for value in result.get("rec_texts", [])]
    scores = [float(value) for value in result.get("rec_scores", [])]
    boxes = result.get("rec_boxes", [])
    lines = []
    for index, text in enumerate(texts):
        if not text.strip():
            continue
        if index < len(boxes):
            box = boxes[index].tolist() if hasattr(boxes[index], "tolist") else list(boxes[index])
        else:
            box = [0, 0, 0, 0]
        lines.append(
            {
                "text": text,
                "confidence": round((scores[index] if index < len(scores) else 0) * 100, 2),
                "bbox": {"x0": box[0], "y0": box[1], "x1": box[2], "y1": box[3]},
            }
        )

    return {
        "engine": "PP-OCRv6_medium",
        "text": "\n".join(line["text"] for line in lines),
        "lines": lines,
        "averageConfidence": round(sum(scores) / len(scores) * 100, 2) if scores else 0,
    }


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(ROOT / "index.html")


@app.get("/{asset_name}", include_in_schema=False)
def web_asset(asset_name: str) -> FileResponse:
    if asset_name not in {"app.js", "app.css"}:
        raise HTTPException(status_code=404, detail="页面不存在")
    return FileResponse(ROOT / asset_name)
