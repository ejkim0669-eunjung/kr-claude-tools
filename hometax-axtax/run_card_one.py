"""카드 화면으로 이동 후 한 거래처만 card-sum 실행."""
import sys, time, argparse
sys.path.insert(0, r"C:\hometax-axtax")
import hometax_axtax as h
from playwright.sync_api import sync_playwright

BIZ_NO = "8488703021"
YEAR   = "2026"
FROM_Q = "3분기"
TO_Q   = "4분기"

deadline = time.time() + 300  # 5분 대기

h.navigate_to_screen(h.CardScreen(), deadline)

ns = argparse.Namespace(
    biz_no=BIZ_NO, name=None, list=None,
    only=None, limit=None, redo=True, redo_before=None,
    dry_run=False, out=None, folder_name="name", allow_cloud_folder=False,
    year=YEAR, from_q=FROM_Q, to_q=TO_Q, excel=True,
)
h.run_card_sum(ns)
