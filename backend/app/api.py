from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

try:
    from app.service import download_report, export_email_workbook, generate_report, parse_email_table
except ModuleNotFoundError:
    from service import download_report, export_email_workbook, generate_report, parse_email_table

router = APIRouter(prefix="/api")


@router.post("/generate")
async def generate(
    health_file: UploadFile = File(...),
    strategy_file: UploadFile = File(...),
):
    if not health_file.filename:
        raise HTTPException(status_code=400, detail="请上传健康度源数据文件")
    if not strategy_file.filename:
        raise HTTPException(status_code=400, detail="请上传战略文件")
    return await generate_report(health_file, strategy_file)


@router.get("/download/{filename}")
def download(filename: str):
    file_path = download_report(filename)
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@router.post("/email-data")
async def upload_email_table(email_file: UploadFile = File(...)):
    if not email_file.filename:
        raise HTTPException(status_code=400, detail="请上传邮件表格文件")
    return await parse_email_table(email_file)


@router.post("/email-export")
async def export_email_sheet(
    email_file: UploadFile = File(...),
):
    if not email_file.filename:
        raise HTTPException(status_code=400, detail="请上传邮件表格文件")
    return await export_email_workbook(email_file)
