"""FastAPI 主入口：上传课表并解析无课时间。"""

from pathlib import Path
import logging
import traceback
import sys
from typing import Any, Optional
import jinja2
import pandas as pd
import uvicorn
from fastapi import FastAPI, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.exceptions import HTTPException as StarletteHTTPException

from model import calculate_free_schedule

DEFAULT_TOTAL_WEEKS = 16
MIN_TOTAL_WEEKS = 1
MAX_TOTAL_WEEKS = 30
MAX_UPLOAD_SIZE_BYTES = 10 * 1024 * 1024

DEPENDENCY_MESSAGES = {
    "multipart": "缺少 python-multipart，无法处理表单上传。请先执行：pip install python-multipart",
    "openpyxl": "缺少 openpyxl，无法读取 Excel。请先执行：pip install openpyxl",
    "xlrd": "缺少 xlrd，无法读取老版本 Excel (.xls)。请先执行：pip install xlrd==1.2.0",
}

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm"}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("schedule_tool")


def get_resource_path(relative_path: str) -> str:
    """获取资源绝对路径，兼容开发模式与 PyInstaller EXE。"""
    if hasattr(sys, "_MEIPASS"):
        base_path = Path(getattr(sys, "_MEIPASS"))
    else:
        base_path = Path(__file__).resolve().parent
    return str((base_path / relative_path).resolve())


def has_module(module_name: str) -> bool:
    """判断模块是否已安装。"""
    try:
        __import__(module_name)
        return True
    except Exception:
        return False


app = FastAPI(title="无课表生成工具", version="1.1")

static_path = Path(get_resource_path("static"))
templates_path = Path(get_resource_path("templates"))

if static_path.exists() and static_path.is_dir():
    app.mount("/static", StaticFiles(directory=str(static_path)), name="static")

templates: Optional[Jinja2Templates] = None
if templates_path.exists() and templates_path.is_dir():
    env = jinja2.Environment(
        loader=jinja2.FileSystemLoader(str(templates_path)),
        cache_size=0,
    )
    templates = Jinja2Templates(env=env)

def parse_form_int(value: Any, default: int = 16) -> int:
    """安全解析总周次，限制 1~30 周。"""
    try:
        number = int(value)
    except Exception:
        return default
    return max(MIN_TOTAL_WEEKS, min(MAX_TOTAL_WEEKS, number))


def wants_json_response(request: Request) -> bool:
    requested_with = request.headers.get("x-requested-with", "")
    accept = request.headers.get("accept", "")
    return requested_with == "XMLHttpRequest" or "application/json" in accept.lower()


def render_index(
    request: Request,
    free_schedule: Optional[list[dict[str, Any]]] = None,
    total_count: int = 0,
    total_weeks: int = 16,
    file_name: str = "",
    has_result: bool = False,
    msg: str = "",
):
    """统一渲染首页，模板不可用时回退到简易 HTML。"""
    if not templates:
        return HTMLResponse("<h2>未找到 templates/index.html，请检查模板路径配置。</h2>", status_code=200)

    context = {
        "free_schedule": free_schedule or [],
        "total_count": total_count,
        "total_weeks": total_weeks,
        "file_name": file_name,
        "has_result": has_result,
        "msg": msg,
    }
    return templates.TemplateResponse(request=request, name="index.html", context=context)


def error_response(request: Request, message: str, status_code: int = 200):
    """页面请求返回模板，接口请求返回 JSON。"""
    if templates and not wants_json_response(request):
        return render_index(
            request=request,
            free_schedule=[],
            total_count=0,
            total_weeks=DEFAULT_TOTAL_WEEKS,
            file_name="",
            has_result=False,
            msg=message,
        )
    return JSONResponse(status_code=status_code, content={"code": status_code, "msg": message})


def ensure_form_dependency() -> Optional[str]:
    """检查 python-multipart 依赖（Form/File 上传必需）。"""
    if not has_module("multipart"):
        return DEPENDENCY_MESSAGES["multipart"]
    return None


def ensure_excel_dependency() -> Optional[str]:
    """检查 openpyxl 依赖（pandas 读取 xlsx 必需）。"""
    if not has_module("openpyxl"):
        return DEPENDENCY_MESSAGES["openpyxl"]
    return None


def ensure_xls_dependency() -> Optional[str]:
    """检查 xlrd 依赖（读取 xls 必需）。"""
    if not has_module("xlrd"):
        return DEPENDENCY_MESSAGES["xlrd"]
    return None


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
    dep_err = ensure_form_dependency()
    if dep_err:
        return error_response(request, dep_err)

    dep_err = ensure_excel_dependency()
    if dep_err:
        return error_response(request, dep_err)

    try:
        form = await request.form()
    except Exception as exc:
        return error_response(request, f"表单解析失败：{exc}")

    upload = form.get("file")
    total_weeks = parse_form_int(form.get("total_weeks", DEFAULT_TOTAL_WEEKS), default=DEFAULT_TOTAL_WEEKS)

    if upload is None or not hasattr(upload, "file"):
        return error_response(request, "未检测到上传文件，请重新选择 Excel 文件。")

    filename = getattr(upload, "filename", "") or "未命名文件"
    logger.info("处理课表文件：%s，总周数：%s", filename, total_weeks)

    file_err = validate_upload_file(upload)
    if file_err:
        return error_response(request, file_err)

    ext = Path(filename).suffix.lower()
    if ext == ".xls":
        dep_err = ensure_xls_dependency()
        if dep_err:
            return error_response(request, dep_err)

    try:
        df = read_excel_file(upload)
    except Exception as exc:
        logger.error("Excel 读取失败：%s", exc, exc_info=True)
        return error_response(request, f"Excel 读取失败：{exc}")

    try:
        free_data = calculate_free_schedule(df, total_weeks)
    except Exception as exc:
        return error_response(request, f"课表解析失败：{exc}")

    if templates and not wants_json_response(request):
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


@app.post("/api/process_schedule")
async def api_process_schedule(
    file: UploadFile = File(...),
    total_weeks: int = Form(DEFAULT_TOTAL_WEEKS),
):
    dep_err = ensure_form_dependency()
    if dep_err:
        return JSONResponse(status_code=200, content={"code": 500, "msg": dep_err})

    dep_err = ensure_excel_dependency()
    if dep_err:
        return JSONResponse(status_code=200, content={"code": 500, "msg": dep_err})

    total_weeks = parse_form_int(total_weeks, default=DEFAULT_TOTAL_WEEKS)
    filename = file.filename or "未命名文件"
    logger.info("处理课表文件：%s，总周数：%s", filename, total_weeks)

    file_err = validate_upload_file(file)
    if file_err:
        return JSONResponse(status_code=200, content={"code": 400, "msg": file_err})

    ext = Path(filename).suffix.lower()
    if ext == ".xls":
        dep_err = ensure_xls_dependency()
        if dep_err:
            return JSONResponse(status_code=200, content={"code": 500, "msg": dep_err})

    try:
        df = read_excel_file(file)
    except Exception as exc:
        logger.error("Excel 读取失败：%s", exc, exc_info=True)
        return JSONResponse(status_code=200, content={"code": 400, "msg": f"Excel 读取失败：{exc}"})

    try:
        free_data = calculate_free_schedule(df, total_weeks)
    except Exception as exc:
        return JSONResponse(status_code=200, content={"code": 400, "msg": f"课表解析失败：{exc}"})

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
    dep_err = ensure_form_dependency()
    if dep_err:
        return JSONResponse(status_code=500, content={"code": 500, "msg": dep_err})

    dep_err = ensure_excel_dependency()
    if dep_err:
        return JSONResponse(status_code=500, content={"code": 500, "msg": dep_err})

    file_err = validate_upload_file(file)
    if file_err:
        return JSONResponse(status_code=400, content={"code": 400, "msg": file_err})

    ext = Path(file.filename or "").suffix.lower()
    if ext == ".xls":
        dep_err = ensure_xls_dependency()
        if dep_err:
            return JSONResponse(status_code=500, content={"code": 500, "msg": dep_err})

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
            status_code=200,
            content={"code": exc.status_code, "error": "HTTPException", "message": str(message)},
        )
    return error_response(request, f"请求异常：{message}", status_code=200)


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    trace = traceback.format_exc()
    logger.error("Unhandled exception: %s\n%s", exc, trace)

    url_str = str(request.url)
    is_dev = "127.0.0.1" in url_str or "localhost" in url_str
    error_msg = str(exc) if is_dev else "系统异常，请稍后重试（开发者可查看日志获取详情）"

    if wants_json_response(request):
        return JSONResponse(
            status_code=200,
            content={"code": 500, "error": exc.__class__.__name__, "message": error_msg},
        )
    return error_response(request, f"系统异常：{error_msg}", status_code=200)


def main() -> None:
    """python app.py 直接启动，默认端口 8001。"""
    logger.info("=" * 50)
    logger.info("无课表工具启动成功")
    logger.info("浏览器访问: http://127.0.0.1:8001")
    logger.info("=" * 50)

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8001,
        log_level="info",
        access_log=False,
    )


if __name__ == "__main__":
    main()
