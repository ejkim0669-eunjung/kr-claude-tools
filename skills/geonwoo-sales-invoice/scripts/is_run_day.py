"""오늘이 '매월 26일, 휴무면 다음 영업일' 발행일인지 판정한다.

사용: python is_run_day.py [YYYY-MM-DD]   (날짜 생략 시 오늘)
출력 첫 줄: RUN 또는 SKIP, 이어서 판정 근거.

휴무 = 토·일 + 아래 공휴일 목록(대체공휴일 포함).
임시공휴일은 미리 알 수 없으므로 지정되면 HOLIDAYS에 추가할 것.
"""
import sys
from datetime import date, timedelta

RUN_DAY = 26

HOLIDAYS = {
    # 2026
    "2026-01-01", "2026-02-16", "2026-02-17", "2026-02-18", "2026-03-02",
    "2026-05-05", "2026-05-25", "2026-06-03", "2026-08-17",
    "2026-09-24", "2026-09-25", "2026-09-28", "2026-10-05", "2026-10-09",
    "2026-12-25",
    # 2027
    "2027-01-01", "2027-02-08", "2027-02-09", "2027-03-01", "2027-05-05",
    "2027-05-13", "2027-08-16", "2027-09-14", "2027-09-15", "2027-09-16",
    "2027-10-04", "2027-10-11", "2027-12-27",
    # 2028
    "2028-01-25", "2028-01-26", "2028-01-27", "2028-03-01", "2028-04-12",
    "2028-05-02", "2028-05-05", "2028-06-06", "2028-08-15",
    "2028-10-02", "2028-10-03", "2028-10-04", "2028-10-05", "2028-10-09",
    "2028-12-25",
}
LAST_LISTED_YEAR = 2028


def is_business_day(d: date) -> bool:
    return d.weekday() < 5 and d.isoformat() not in HOLIDAYS


def target_for_month(year: int, month: int) -> date:
    d = date(year, month, RUN_DAY)
    while not is_business_day(d):
        d += timedelta(days=1)
    return d


def main() -> None:
    today = date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else date.today()
    if today.year > LAST_LISTED_YEAR:
        print(f"SKIP\n공휴일 목록이 {LAST_LISTED_YEAR}년까지만 있음 - 목록 갱신 필요 (사용자에게 알릴 것)")
        return
    # 1~5일 실행이면 전월 26일 기준 발행일이 이월된 경우일 수 있다
    if today.day >= RUN_DAY:
        y, m = today.year, today.month
    else:
        first = today.replace(day=1) - timedelta(days=1)
        y, m = first.year, first.month
    target = target_for_month(y, m)
    if today == target:
        print(f"RUN\n{y}년 {m}월분 발행일: {target} (오늘)")
    else:
        print(f"SKIP\n{y}년 {m}월분 발행일은 {target}, 오늘({today})은 아님")


if __name__ == "__main__":
    main()
