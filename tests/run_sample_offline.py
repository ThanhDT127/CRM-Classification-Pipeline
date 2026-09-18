"""Chay thu tren du lieu mau, KHONG cham SharePoint va KHONG gui thu ket qua.

    python tests/run_sample_offline.py                 # chay mot lan
    python tests/run_sample_offline.py --lan-hai       # chay lai tren checkpoint da co

VI SAO CAN FILE NAY

`src/pipeline.py` khong co co chay kho: bat len la tai SharePoint that, ghi de tep
Excel that, va gui thu cho nguoi that. Nen khong the dung no de nghiem thu hai dieu
con lai cua change -- "khong dem doi" va "khong mat dong khi Gateway chet".

File nay dung DUNG hai thu ma pipeline dung:
  · `call_llm_batch`        tang goi LLM, ke ca vong thu lai 3 lan ben trong
  · `save_history_db_atomic` co che checkpoint, ghi nguyen tu ra output/

Va bo qua moi thu khac. Do la co y: no nghiem thu tang goi va tang checkpoint, chu
khong nghiem thu ca pipeline.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

SAMPLE = ROOT / "sample_data" / "CRM_merge_sample.xlsx"
CKPT = ROOT / "output" / "sample_offline_history.json"
GHI_SO = ROOT / "output" / "sample_offline_writes.jsonl"
BATCH_SIZE = 5


def doc_mau(so_dong: int) -> list[dict]:
    """Doc file mau, tra ve dang ma `call_llm_batch` nhan."""
    import pandas as pd
    df = pd.read_excel(SAMPLE, header=[0, 1])
    df.columns = [" | ".join(str(x).strip() for x in c) for c in df.columns]
    cot = [c for c in df.columns if df[c].notna().any()][:4]
    ra = []
    for i, (_, hang) in enumerate(df.head(so_dong).iterrows(), 1):
        ra.append({"row_idx": str(i),
                   **{c.split(" | ")[-1][:30]: str(hang[c])[:120] for c in cot}})
    return ra


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--lan-hai", action="store_true",
                   help="Chay lai tren checkpoint da co, de do viec dem doi")
    p.add_argument("--so-dong", type=int, default=20)
    args = p.parse_args()

    os.environ.setdefault("GEMINI_BACKEND", "gateway")
    os.environ.setdefault("GEMINI_GATEWAY_BASE_URL", "http://127.0.0.1:4000")
    os.environ.setdefault("GEMINI_GATEWAY_USER", "svc.crm-feedback")
    logging.basicConfig(level=logging.WARNING, format="%(levelname)s %(message)s")
    for ten in ("httpx", "google_genai", "google_genai.models"):
        logging.getLogger(ten).setLevel(logging.ERROR)

    from llm import init_llm_client, call_llm_batch
    from pipeline import save_history_db_atomic

    CKPT.parent.mkdir(parents=True, exist_ok=True)
    if not args.lan_hai:
        CKPT.unlink(missing_ok=True)
        GHI_SO.unlink(missing_ok=True)

    history = json.loads(CKPT.read_text(encoding="utf-8")) if CKPT.exists() else {}
    print(f"checkpoint dau vao: {len(history)} dong")

    dong = doc_mau(args.so_dong)
    con_lai = [d for d in dong if d["row_idx"] not in history]
    print(f"doc {len(dong)} dong tu file mau, {len(con_lai)} dong chua co trong checkpoint")

    if not con_lai:
        print("khong con gi de lam -- dung cai ta muon o lan chay thu hai")

    client, model = init_llm_client()
    prompt = ('Ban phan loai du lieu CRM. Voi MOI dong dau vao, tra ve mot phan tu '
              '{"row_idx": "<row_idx dau vao>", "fills": {"nhan": "<mot tu>"}}. '
              'Tra ve DUNG mot JSON array, khong giai thich gi them.')

    bo = 0
    with GHI_SO.open("a", encoding="utf-8") as so:
        for i in range(0, len(con_lai), BATCH_SIZE):
            lo = con_lai[i:i + BATCH_SIZE]
            so_lo = i // BATCH_SIZE + 1
            try:
                ket_qua = call_llm_batch(client, model, prompt, lo, max_retry=3)
            except Exception as e:
                bo += len(lo)
                print(f"  lo {so_lo}: BO {len(lo)} dong -- {str(e)[:80]}")
                continue
            for m in ket_qua:
                rid = str(m.get("row_idx") or "")
                if not rid:
                    continue
                history[rid] = m.get("fills") or {}
                # MOI lan ghi so mot dong deu duoc ghi lai o day. Phep kiem dem doi
                # phai dem SO LAN GHI, khong phai dem so ban ghi trong `history` --
                # `history` la mot dict nen moi khoa chi co mot ban ghi BANG CAU TRUC,
                # va dem no thi khong chung minh duoc gi.
                so.write(json.dumps({"row_idx": rid}, ensure_ascii=False) + "\n")
            save_history_db_atomic(history)
            print(f"  lo {so_lo}: OK {len(ket_qua)} dong")

    CKPT.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")

    lan_ghi = Counter()
    if GHI_SO.exists():
        for line in GHI_SO.read_text(encoding="utf-8").splitlines():
            if line.strip():
                lan_ghi[json.loads(line)["row_idx"]] += 1
    ghi_hai_lan = {k: v for k, v in lan_ghi.items() if v > 1}

    print(f"\ncheckpoint dau ra : {len(history)} dong")
    print(f"dong bi bo        : {bo}")
    print(f"dong ghi so 2+ lan: {len(ghi_hai_lan)} {list(ghi_hai_lan.items())[:5]}")
    dat = bo == 0 and not ghi_hai_lan
    print("DAT" if dat else "KHONG DAT")
    return 0 if dat else 1


if __name__ == "__main__":
    sys.exit(main())
