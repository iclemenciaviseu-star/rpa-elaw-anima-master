"""
Backend FastAPI - serve a interface e orquestra a execucao do rpa.py.
"""
from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from pydantic import BaseModel

BASE = Path(__file__).resolve().parent
UPLOADS = BASE / "uploads"
LOGS = BASE / "logs"
UPLOADS.mkdir(exist_ok=True)
LOGS.mkdir(exist_ok=True)

app = FastAPI(title="RPA eLaw Anima")


class Job:
    def __init__(self, job_id, process):
        self.id = job_id
        self.process = process
        self.queue = asyncio.Queue()
        self.done = False
        self.return_code = None
        self.started_at = datetime.now().isoformat()


JOBS = {}


class FolderReq(BaseModel):
    path: str


class PreviewReq(BaseModel):
    spreadsheet: str


class RunReq(BaseModel):
    usuario: str
    senha: str
    microsoft_email: str | None = None
    spreadsheet: str
    pasta_pdfs: str
    inicio: int = 0
    fim: int | None = None
    headless: bool = False
    slow_mo: int = 200
    preservar_campos: bool = True


@app.get("/")
async def root():
    return FileResponse(BASE / "index.html")


@app.get("/favicon.ico")
async def favicon():
    return JSONResponse(status_code=204, content=None)


@app.post("/api/upload-spreadsheet")
async def upload_spreadsheet(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(400, "Arquivo sem nome")
    ext = Path(file.filename).suffix.lower()
    if ext not in (".xlsx", ".xls", ".csv"):
        raise HTTPException(400, f"Extensao {ext} nao suportada")

    safe = file.filename.replace("/", "_").replace("\\", "_")
    dest = UPLOADS / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{safe}"
    with open(dest, "wb") as f:
        while chunk := await file.read(1024 * 1024):
            f.write(chunk)

    try:
        df = _read_df(str(dest))
    except Exception as e:
        dest.unlink(missing_ok=True)
        raise HTTPException(400, f"Nao consegui ler a planilha: {e}")

    return {
        "path": str(dest),
        "filename": file.filename,
        "rows": len(df),
        "columns": list(df.columns),
    }


@app.post("/api/upload-pdfs")
async def upload_pdfs(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(400, "Arquivo sem nome")
    if Path(file.filename).suffix.lower() != ".zip":
        raise HTTPException(400, "Envie um arquivo .zip com os PDFs")

    session_id = uuid.uuid4().hex[:10]
    zip_path = UPLOADS / f"pdfs_{session_id}.zip"
    extract_dir = UPLOADS / f"pdfs_{session_id}"

    contents = await file.read()
    await asyncio.to_thread(zip_path.write_bytes, contents)

    try:
        await asyncio.to_thread(_extract_zip, zip_path, extract_dir)
    except Exception as e:
        zip_path.unlink(missing_ok=True)
        shutil.rmtree(extract_dir, ignore_errors=True)
        raise HTTPException(400, f"Erro ao extrair ZIP: {e}")

    pdfs = list(extract_dir.rglob("*.pdf")) + list(extract_dir.rglob("*.PDF"))
    zips = list(extract_dir.rglob("*.zip")) + list(extract_dir.rglob("*.ZIP"))
    return {
        "path": str(extract_dir),
        "pdf_count": len(pdfs),
        "zip_count": len(zips),
        "total": len(pdfs) + len(zips),
        "session_id": session_id,
    }


def _extract_zip(zip_path: Path, extract_dir: Path):
    extract_dir.mkdir(exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as z:
        for member in z.infolist():
            if ".." in member.filename or member.filename.startswith("/"):
                continue
            z.extract(member, extract_dir)
    zip_path.unlink(missing_ok=True)


AUTH_STATE = BASE / "auth_state.json"


@app.post("/api/upload-auth")
async def upload_auth(file: UploadFile = File(...)):
    if not file.filename:
        raise HTTPException(400, "Arquivo sem nome")
    if not file.filename.endswith(".json"):
        raise HTTPException(400, "Envie o arquivo auth_state.json")

    with open(AUTH_STATE, "wb") as f:
        while chunk := await file.read(1024 * 1024):
            f.write(chunk)

    try:
        import json as _json
        data = _json.loads(AUTH_STATE.read_text(encoding="utf-8"))
        cookies = len(data.get("cookies", []))
    except Exception:
        cookies = 0

    return {"ok": True, "cookies": cookies}


@app.get("/api/auth-status")
async def auth_status():
    if AUTH_STATE.exists():
        try:
            import json as _json
            data = _json.loads(AUTH_STATE.read_text(encoding="utf-8"))
            cookies = len(data.get("cookies", []))
            return {"exists": True, "cookies": cookies}
        except Exception:
            return {"exists": True, "cookies": 0}
    return {"exists": False, "cookies": 0}


@app.delete("/api/upload-auth")
async def delete_auth():
    AUTH_STATE.unlink(missing_ok=True)
    return {"ok": True}


@app.post("/api/validate-folder")
async def validate_folder(req: FolderReq):
    p = Path(req.path)
    if not p.exists():
        raise HTTPException(400, "Pasta nao existe")
    if not p.is_dir():
        raise HTTPException(400, "Caminho nao e uma pasta")
    pdfs = list(p.glob("*.pdf")) + list(p.glob("*.PDF"))
    zips = list(p.glob("*.zip")) + list(p.glob("*.ZIP"))
    return {
        "path": str(p),
        "pdf_count": len(pdfs),
        "zip_count": len(zips),
        "total": len(pdfs) + len(zips),
    }


@app.post("/api/preview")
async def preview(req: PreviewReq):
    df = _read_df(req.spreadsheet)
    required = [
        "(Processo) Número", "Tipo de Ação", "Procedimento",
        "Fase processual", "Resumo da Ação", "Pedidos",
        "Valor da Causa", "Prazo para Defesa",
        "Deseja solicitar subsídios?",
    ]
    missing = [c for c in required if c not in df.columns]
    sample = df.head(5).fillna("").to_dict(orient="records")
    for row in sample:
        for k, v in list(row.items()):
            if isinstance(v, pd.Timestamp):
                row[k] = v.strftime("%d/%m/%Y")
            elif not isinstance(v, (str, int, float, bool, type(None))):
                row[k] = str(v)
    return {
        "rows": len(df),
        "columns": list(df.columns),
        "missing": missing,
        "sample": sample,
    }


def _read_df(caminho):
    c = caminho.lower()
    if c.endswith(".xls"):
        return _strip_cols(pd.read_excel(caminho, engine="xlrd"))
    if c.endswith(".csv"):
        return _strip_cols(pd.read_csv(caminho))
    return _strip_cols(pd.read_excel(caminho))


def _strip_cols(df):
    df.columns = df.columns.str.strip()
    return df


@app.post("/api/run")
async def run(req: RunReq):
    if not Path(req.spreadsheet).exists():
        raise HTTPException(400, "Planilha nao encontrada")
    if not Path(req.pasta_pdfs).exists():
        raise HTTPException(400, "Pasta de PDFs nao encontrada")
    if not req.usuario or not req.senha:
        raise HTTPException(400, "Usuario e senha sao obrigatorios")

    job_id = uuid.uuid4().hex[:12]
    log_file = LOGS / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{job_id}.log"
    cfg_file = LOGS / f"cfg_{job_id}.json"

    # no Render nao ha display — force headless
    headless = req.headless or bool(os.environ.get("RENDER"))

    cfg = {
        "usuario": req.usuario,
        "senha": req.senha,
        "microsoft_email": req.microsoft_email or req.usuario,
        "planilha": req.spreadsheet,
        "pasta_pdfs": req.pasta_pdfs,
        "inicio": req.inicio,
        "fim": req.fim,
        "headless": headless,
        "slow_mo": req.slow_mo,
        "preservar_campos": req.preservar_campos,
        "log_file": str(log_file),
    }
    cfg_file.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    proc = await asyncio.create_subprocess_exec(
        sys.executable, "-u",
        str(BASE / "rpa.py"), str(cfg_file),
        stdout=asyncio.subprocess.PIPE,
        stderr=asyncio.subprocess.STDOUT,
        cwd=str(BASE),
        env=env,
    )

    job = Job(job_id, proc)
    JOBS[job_id] = job
    # pasta extraida do ZIP (para limpeza apos o job)
    pdf_dir = Path(req.pasta_pdfs) if req.pasta_pdfs else None

    asyncio.create_task(_pump_output(job, cfg_file, pdf_dir))
    return {"job_id": job_id, "log_file": str(log_file)}


async def _pump_output(job, cfg_file, pdf_dir=None):
    try:
        async for line in job.process.stdout:
            decoded = line.decode("utf-8", errors="replace").rstrip()
            if not decoded:
                continue
            try:
                json.loads(decoded)
                await job.queue.put(decoded)
            except Exception:
                payload = {
                    "ts": datetime.now().isoformat(),
                    "level": "info",
                    "msg": decoded,
                }
                await job.queue.put(json.dumps(payload, ensure_ascii=False))
    finally:
        rc = await job.process.wait()
        job.return_code = rc
        job.done = True
        final = {
            "ts": datetime.now().isoformat(),
            "level": "success" if rc == 0 else "error",
            "msg": f"Processo finalizado (exit code {rc})",
            "step": "finished",
            "return_code": rc,
        }
        await job.queue.put(json.dumps(final, ensure_ascii=False))
        await job.queue.put("__END__")
        try:
            cfg_file.unlink(missing_ok=True)
        except Exception:
            pass
        # limpa pasta de PDFs extraida do ZIP (evita acumulo no servidor)
        if pdf_dir and pdf_dir.exists() and pdf_dir.name.startswith("pdfs_"):
            try:
                shutil.rmtree(pdf_dir, ignore_errors=True)
            except Exception:
                pass


@app.get("/api/stream/{job_id}")
async def stream(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job nao encontrado")

    async def gen():
        while True:
            item = await job.queue.get()
            if item == "__END__":
                yield "event: end\ndata: {}\n\n"
                break
            yield f"data: {item}\n\n"

    return StreamingResponse(
        gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.get("/api/status/{job_id}")
async def status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job nao encontrado")
    return {
        "id": job.id,
        "done": job.done,
        "return_code": job.return_code,
        "started_at": job.started_at,
    }


@app.post("/api/stop/{job_id}")
async def stop(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, "Job nao encontrado")
    if job.done:
        return {"ok": True, "already_done": True}
    try:
        job.process.terminate()
    except Exception:
        try:
            job.process.kill()
        except Exception:
            pass
    return {"ok": True}


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8765))
    host = "0.0.0.0" if os.environ.get("RENDER") else "127.0.0.1"
    print(f"RPA eLaw Anima - http://{host}:{port}")
    uvicorn.run(app, host=host, port=port, log_level="warning")
