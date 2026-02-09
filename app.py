# app.py - FastAPI 主文件（整合上传+解析+聊天）
from fastapi import FastAPI, Request, File, UploadFile, Form, HTTPException
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import pandas as pd
import os
import sys
import uvicorn
from model import calculate_free_schedule

# ========== 适配EXE的路径处理 ==========
def get_resource_path(relative_path):
    """获取打包后的资源路径（兼容开发/EXE模式）"""
    if hasattr(sys, '_MEIPASS'):
        # 打包后EXE的临时目录
        base_path = sys._MEIPASS
    else:
        # 开发模式的当前目录
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ========== 初始化FastAPI ==========
app = FastAPI(title="无课表生成工具", version="1.0")

# 配置静态文件和模板路径（适配EXE）
static_path = get_resource_path("static")
templates_path = get_resource_path("templates")

# 挂载静态文件
if os.path.exists(static_path):
    app.mount("/static", StaticFiles(directory=static_path), name="static")

# 模板文件夹
if os.path.exists(templates_path):
    templates = Jinja2Templates(directory=templates_path)
else:
    templates = None

# 全局存储当前解析的无课表结果（仅内存中临时存储，用于聊天交互）
current_free_data = []

# 主页面（整合上传+结果+聊天）
@app.get("/")
async def index(request: Request):
    if templates:
        return templates.TemplateResponse("index.html", {
            "request": request,
            "free_schedule": [],
            "total_count": 0,
            "total_weeks": 16,
            "file_name": "",
            "has_result": False,
            "msg": ""
        })
    else:
        from fastapi.responses import HTMLResponse
        return HTMLResponse(content="<h1>请创建 templates 文件夹和 index.html 文件</h1>")

# 处理课表上传、解析（不存储文件）
@app.post("/process_schedule")
async def process_schedule(
    request: Request,
    file: UploadFile = File(...),
    total_weeks: int = Form(..., ge=1, le=20)  # 限制总周次范围
):
    global current_free_data
    try:
        # 调试信息
        print(f"[DEBUG] 接收到文件：{file.filename}")
        print(f"[DEBUG] 总周次：{total_weeks}")
        
        # 直接从内存读取Excel（不保存到磁盘）
        try:
            df = pd.read_excel(file.file, sheet_name="Sheet1", header=None)
        except:
            # 如果 Sheet1 不存在，尝试读取第一个工作表
            file.file.seek(0)  # 重置文件指针
            df = pd.read_excel(file.file, header=None)
        
        print(f"[DEBUG] Excel 读取成功，形状：{df.shape}")
        print(f"[DEBUG] 前5行数据：")
        print(df.head())
        
        current_free_data = calculate_free_schedule(df, total_weeks)
        print(f"[DEBUG] 解析完成，共找到 {len(current_free_data)} 个空闲时间段")
        
        # 调试：打印前几条结果和列信息
        print(f"[DEBUG] Excel列数：{df.shape[1]}，行数：{df.shape[0]}")
        print(f"[DEBUG] 第2行（索引2）的前几列数据：")
        if df.shape[0] > 2:
            for col_idx in range(min(10, df.shape[1])):
                cell_val = df.iloc[2, col_idx]
                print(f"  列{col_idx}：{cell_val}")
        
        # 调试：打印前几条结果
        if current_free_data:
            print(f"[DEBUG] 前5条结果示例：")
            for item in current_free_data[:5]:
                print(f"  {item['weekday']} {item['section']}节：{item['free_desc']}")
        else:
            print(f"[DEBUG] 警告：未找到任何空闲时间段！")
        
        if templates:
            return templates.TemplateResponse("index.html", {
                "request": request,
                "free_schedule": current_free_data,
                "total_count": len(current_free_data),
                "total_weeks": total_weeks,
                "file_name": file.filename,
                "has_result": True,
                "msg": f"解析成功！共找到 {len(current_free_data)} 个空闲时间段"
            })
        else:
            from fastapi.responses import JSONResponse
            return JSONResponse(content={"code": 200, "data": current_free_data})
    
    except Exception as e:
        import traceback
        error_trace = traceback.format_exc()
        print(f"[ERROR] 处理文件时出错：{str(e)}")
        print(f"[ERROR] 错误堆栈：\n{error_trace}")
        
        if templates:
            return templates.TemplateResponse("index.html", {
                "request": request,
                "free_schedule": [],
                "total_count": 0,
                "total_weeks": 16,
                "file_name": "",
                "has_result": False,
                "msg": f"解析失败：{str(e)[:80]}"  # 限制错误信息长度
            })
        else:
            from fastapi.responses import JSONResponse
            return JSONResponse(
                status_code=500,
                content={"code": 500, "msg": str(e)}
            )

# 聊天交互接口（支持查询星期/节次）
@app.post("/chat_query")
async def chat_query(request: Request, message: str = Form(...)):
    global current_free_data
    message = message.strip()
    response = ""
    
    # 支持的查询规则
    weekday_list = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
    section_list = ["1-2", "3-4", "5-6", "7-8", "9-10", "11-12"]
    
    # 1. 查询某星期的空闲（如"星期一有空吗"）
    for weekday in weekday_list:
        if weekday in message:
            weekday_free = [f"- {item['section']}节：{item['free_desc']}" 
                           for item in current_free_data if item["weekday"] == weekday]
            response = f"{weekday}空闲时段：\n" + "\n".join(weekday_free) if weekday_free else f"{weekday}无空闲"
            break
    
    # 2. 查询某节次的空闲（如"3-4节什么时候有空"）
    if not response:
        for section in section_list:
            if section in message:
                section_free = [f"- {item['weekday']}：{item['free_desc']}" 
                               for item in current_free_data if item["section"] == section]
                response = f"{section}节空闲时段：\n" + "\n".join(section_free) if section_free else f"{section}节无空闲"
                break
    
    # 3. 默认回复（提示支持的查询类型）
    if not response:
        response = "支持查询：\n1. 某星期空闲（例：'星期二有空吗'）\n2. 某节次空闲（例：'9-10节什么时候有空'）"
    
    return {"user_msg": message, "bot_response": response}

# ========== 添加EXE启动入口 ==========
def main():
    """启动入口函数（适配EXE打包，禁用日志避免终端依赖）"""
    import logging
    
    # 配置日志：禁用 uvicorn 的详细日志，避免终端依赖
    logging.getLogger("uvicorn").setLevel(logging.ERROR)
    logging.getLogger("uvicorn.access").setLevel(logging.ERROR)
    logging.getLogger("uvicorn.error").setLevel(logging.ERROR)
    
    print("=" * 50)
    print("课表无空闲查询工具已启动！")
    print("=" * 50)
    print("请打开浏览器访问：http://127.0.0.1:8001")
    print("关闭此窗口即可停止服务")
    print("=" * 50)
    
    # 启动UVICORN服务（使用 critical 日志级别，最小化日志输出）
    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8001,
        log_level="critical",  # 仅记录严重错误，避免终端依赖
        access_log=False  # 禁用访问日志
    )

if __name__ == "__main__":
    main()
