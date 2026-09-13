"""FastAPI 主入口：上传课表并解析无课时间。"""

from pathlib import Path
import logging
import traceback
from typing import Any, Optional

import jinja2
import pandas as pd
import uvicorn
from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.exceptions import HTTPException as StarletteHTTPException

from model import calculate_free_schedule

PORT = 8001
DEFAULT_TOTAL_WEEKS = 16
MIN_TOTAL_WEEKS = 1
MAX_TOTAL_WEEKS = 30
MAX_UPLOAD_SIZE_BYTES = 10 * 1024 * 1024

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("schedule_tool")

app = FastAPI(title="无课表生成工具", version="1.0")

templates = Jinja2Templates(
    env=jinja2.Environment(
        loader=jinja2.FileSystemLoader(str(Path(__file__).resolve().parent / "templates")),
        autoescape=True,
    )
)


def parse_form_int(value: Any) -> int:
    """安全解析总周次，限制 1~30 周，非法值用默认 16。"""
    try:
        number = int(value)
    except Exception:
        return DEFAULT_TOTAL_WEEKS
    return max(MIN_TOTAL_WEEKS, min(MAX_TOTAL_WEEKS, number))


def wants_json_response(request: Request) -> bool:
    requested_with = request.headers.get("x-requested-with", "")
    accept = request.headers.get("accept", "")
    return requested_with == "XMLHttpRequest" or "application/json" in accept.lower()


def render_index(
    request: Request,
    free_schedule: Optional[list[dict[str, Any]]] = None,
    total_count: int = 0,
    total_weeks: int = DEFAULT_TOTAL_WEEKS,
    file_name: str = "",
    has_result: bool = False,
    msg: str = "",
):
    context = {
        "free_schedule": free_schedule or [],
        "total_count": total_count,
        "total_weeks": total_weeks,
        "file_name": file_name,
        "has_result": has_result,
        "msg": msg,
    }
    return templates.TemplateResponse(request=request, name="index.html", context=context)


def error_response(request: Request, message: str, status_code: int = 400):
    """页面请求返回模板，接口请求返回 JSON；两者都使用真实 HTTP 状态码。"""
    if not wants_json_response(request):
        response = render_index(request=request, msg=message)
        response.status_code = status_code
        return response
    return JSONResponse(status_code=status_code, content={"code": status_code, "msg": message})


def validate_upload_file(upload_file: UploadFile) -> Optional[str]:
    """校验上传文件格式与大小。"""
    ext = Path(upload_file.filename or "").suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return f"不支持的文件格式：{ext or '未知'}，仅支持 {', '.join(sorted(ALLOWED_EXTENSIONS))}"

    upload_file.file.seek(0, 2)
    file_size = upload_file.file.tell()
    upload_file.file.seek(0)

    if file_size == 0:
        return "上传的文件为空，请重新选择有效文件"
    if file_size > MAX_UPLOAD_SIZE_BYTES:
        return "文件大小超过10MB限制，请上传更小的课表文件"
    return None


def read_excel_file(upload_file: UploadFile) -> pd.DataFrame:
    """读取 Excel，优先使用 Sheet1，失败后回退首个 sheet。"""
    ext = Path(upload_file.filename or "").suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    upload_file.file.seek(0)
    try:
        return pd.read_excel(upload_file.file, sheet_name="Sheet1", header=None, engine=engine)
    except Exception:
        upload_file.file.seek(0)
        return pd.read_excel(upload_file.file, header=None, engine=engine)


@app.get("/")
async def index(request: Request):
    return render_index(request=request)


@app.post("/process_schedule")
async def process_schedule(request: Request):
    try:
        form = await request.form()
    except Exception as exc:
        return error_response(request, f"表单解析失败：{exc}")

    upload = form.get("file")
    total_weeks = parse_form_int(form.get("total_weeks", DEFAULT_TOTAL_WEEKS))

    if upload is None or not hasattr(upload, "file"):
        return error_response(request, "未检测到上传文件，请重新选择 Excel 文件。")

    filename = getattr(upload, "filename", "") or "未命名文件"
    logger.info("处理课表文件：%s，总周数：%s", filename, total_weeks)

    file_err = validate_upload_file(upload)
    if file_err:
        return error_response(request, file_err)

    try:
        df = read_excel_file(upload)
    except Exception as exc:
        logger.error("Excel 读取失败：%s", exc, exc_info=True)
        return error_response(request, f"Excel 读取失败：{exc}")

    try:
        free_data = calculate_free_schedule(df, total_weeks)
    except Exception as exc:
        return error_response(request, f"课表解析失败：{exc}")

    if not wants_json_response(request):
        return render_index(
            request=request,
            free_schedule=free_data,
            total_count=len(free_data),
            total_weeks=total_weeks,
            file_name=filename,
            has_result=True,
            msg=f"解析成功！共找到 {len(free_data)} 个空闲时间段",
        )

    return JSONResponse(
        status_code=200,
        content={
            "code": 200,
            "msg": "ok",
            "file_name": filename,
            "total_weeks": total_weeks,
            "total_count": len(free_data),
            "data": free_data,
        },
    )


@app.post("/api/preview_schedule")
async def preview_schedule(file: UploadFile = File(...)):
    file_err = validate_upload_file(file)
    if file_err:
        return JSONResponse(status_code=400, content={"code": 400, "msg": file_err})

    try:
        df = read_excel_file(file)
        preview_data = df.head(10).iloc[:, :10].fillna("").to_dict()
        return JSONResponse(status_code=200, content={"code": 200, "msg": "ok", "data": preview_data})
    except Exception as exc:
        logger.error("预览失败：%s", exc, exc_info=True)
        return JSONResponse(status_code=400, content={"code": 400, "msg": f"预览失败：{exc}"})


@app.exception_handler(StarletteHTTPException)
async def http_exception_handler(request: Request, exc: StarletteHTTPException):
    message = exc.detail if exc.detail else str(exc)
    if wants_json_response(request):
        return JSONResponse(
            status_code=exc.status_code,
            content={"code": exc.status_code, "error": "HTTPException", "message": str(message)},
        )
    return error_response(request, f"请求异常：{message}", status_code=exc.status_code)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    # ponytail: 本机工具，直接把异常原文返回给页面；若日后对外部署，需换成通用错误文案
    logger.error("Unhandled exception: %s\n%s", exc, traceback.format_exc())

    if wants_json_response(request):
        return JSONResponse(
            status_code=500,
            content={"code": 500, "error": exc.__class__.__name__, "message": str(exc)},
        )
    return error_response(request, f"系统异常：{exc}", status_code=500)


def main() -> None:
    """python app.py 直接启动。"""
    logger.info("=" * 50)
    logger.info("无课表工具启动成功")
    logger.info("浏览器访问: http://127.0.0.1:%s", PORT)
    logger.info("=" * 50)

    uvicorn.run(app, host="127.0.0.1", port=PORT, log_level="info", access_log=False)


if __name__ == "__main__":
    main()
