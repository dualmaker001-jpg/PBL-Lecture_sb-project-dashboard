from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

DATA_DIR = Path(__file__).parent / "data"
PROJECT_PATH = DATA_DIR / "SB_project_demo.csv"
RESOURCE_PATH = DATA_DIR / "SB_resource_demo.csv"
STATUSES = ["접수", "자료확인", "시험진행", "성적서검토", "보완요청", "완료"]
STATUS_COLORS = ["#9ec5fe", "#6ea8fe", "#3d8bfd", "#1a56db", "#f59e0b", "#0b1f44"]
COLUMNS = ["의뢰번호", "고객사", "제품명", "모델명", "시험분야", "적용규격",
           "담당엔지니어", "접수일", "완료예정일", "진행상태", "시료상태", "메모"]
FILTER_COLS = ["시험분야", "적용규격", "담당엔지니어", "진행상태", "시료상태"]
INIT_FIELDS = ["EMC", "무선(RF)", "인증서류", "SAR"]
DELAY = "지연 후보(일정 확인 필요)"
EXCEL_NOTICE = "교육용 검토 파일 — 공식 보고서 아님"
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
RES_COLORS = {"사용 가능": "#1a56db", "사용 중": "#0b1f44", "점검 중": "#b45309"}

st.set_page_config(page_title="인증 프로젝트 대시보드", layout="wide")
st.markdown("""
<style>
.stApp { background: #ffffff; color: #0b1f44; }
h1, h2, h3 { color: #0b1f44; }
.notice { background: #e8f0fe; color: #0b3d91; border-left: 5px solid #1a56db;
          padding: 8px 14px; border-radius: 4px; font-weight: 600; margin-bottom: 12px; }
div[class*="st-key-card"] button { min-height: 90px; border-radius: 10px; background: #f3f6fc;
                   color: #0b1f44; border: 1px solid #c9d6f0; }
div[class*="st-key-card"] button:hover { border-color: #1a56db; color: #1a56db; }
div[class*="st-key-card"] button strong { font-size: 1.8rem; color: #1a56db; }
div[class*="st-key-card"] button[data-testid="stBaseButton-primary"] { background: #0b1f44; color: #fff; border-color: #0b1f44; }
div[class*="st-key-card"] button[data-testid="stBaseButton-primary"] strong { color: #fff; }
.res { background: #f3f6fc; border: 1px solid #c9d6f0; border-radius: 10px; padding: 12px 14px; }
.res .name { font-weight: 700; }
.res .state { font-size: 1.4rem; font-weight: 700; }
.res .note { font-size: 0.8rem; opacity: 0.75; }
</style>
""", unsafe_allow_html=True)


# ---------- 데이터 로드 ----------

def load_csv(src):
    """CSV를 UTF-8(BOM) → CP949 순서로 시도해 읽는다 (한글 CSV 인코딩 대응)."""
    for enc in ("utf-8-sig", "cp949"):
        try:
            if hasattr(src, "seek"):
                src.seek(0)
            return pd.read_csv(src, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSV 인코딩을 읽을 수 없습니다 (UTF-8 또는 CP949로 저장해 주세요).")


def load_projects():
    """의뢰 CSV(없으면 업로드)를 읽고 필수 필드·날짜를 정리해 반환. 실패하면 안내 후 화면을 멈춘다."""
    if PROJECT_PATH.exists():
        source = PROJECT_PATH
    else:
        st.info(f"기본 데이터 파일이 없습니다: `data/{PROJECT_PATH.name}`\n\n아래에서 CSV 파일을 업로드해 주세요.")
        source = st.file_uploader(f"{PROJECT_PATH.name} 업로드", type="csv")
        if source is None:
            st.stop()
    try:
        df = load_csv(source)
    except Exception as e:
        st.error(f"파일을 읽지 못했습니다: {e}")
        st.stop()
    missing = [c for c in COLUMNS if c not in df.columns]
    if missing:
        st.error(f"필수 필드가 없습니다: {', '.join(missing)}")
        st.stop()
    df = df[COLUMNS].copy()
    for c in ("접수일", "완료예정일"):
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    df["메모"] = df["메모"].fillna("").astype(str)
    return df


def init_state():
    """세션 최초 1회: 의뢰 데이터와 시험분야 목록을 session_state에 올린다 (이후 수정은 이 세션 안에서만 유지)."""
    if "projects" not in st.session_state:
        df = load_projects()
        st.session_state.projects = df
        st.session_state.fields = INIT_FIELDS + sorted(set(df["시험분야"].dropna()) - set(INIT_FIELDS))


# ---------- 일정 계산 ----------

def add_schedule(df, base):
    """잔여일수(완료예정일 - 기준일)와 일정상태 열을 붙인 사본을 반환."""
    df = df.copy()
    days = (pd.to_datetime(df["완료예정일"]) - pd.Timestamp(base)).dt.days
    # 순서가 곧 우선순위: 완료 > 날짜 없음/오류 > 지연 > 오늘 마감 > 그 외
    df["일정상태"] = np.select(
        [df["진행상태"].eq("완료"), days.isna(), days < 0, days == 0],
        ["완료", "일정 확인 필요", DELAY, "오늘 마감"], "진행 예정")
    df["잔여일수"] = days.astype("Int64")
    return df[COLUMNS[:9] + ["잔여일수", "일정상태"] + COLUMNS[9:]]


def received_by(df, base):
    """기준일까지 접수된 의뢰만 남긴다 (접수일이 비어 있으면 포함)."""
    rec = pd.to_datetime(df["접수일"])
    return df[rec.isna() | (rec <= pd.Timestamp(base))]


# ---------- 사이드바: 필터 ----------

def keep_valid(key, options):
    """옵션이 바뀐 뒤(분야 이름 변경·삭제 등) 위젯에 남은 무효 선택값을 정리한다."""
    v = st.session_state.get(key)
    if isinstance(v, list):
        st.session_state[key] = [x for x in v if x in options]
    elif v is not None and v not in options:
        del st.session_state[key]


def sidebar_filters(df):
    """사이드바 '필터' 영역을 그리고 {필드: 선택값 목록}을 반환 (빈 목록 = 전체)."""
    st.sidebar.header("필터")
    chosen = {}
    for col in FILTER_COLS:
        if col == "시험분야":
            opts = st.session_state.fields
        elif col == "진행상태":
            opts = STATUSES
        else:
            opts = sorted(df[col].dropna().unique().tolist())
        keep_valid(f"f_{col}", opts)
        chosen[col] = st.sidebar.multiselect(col, opts, key=f"f_{col}", placeholder="전체")
    return chosen


def apply_filters(df, chosen):
    """선택된 필터 조건을 모두 만족하는 행만 남긴다."""
    for col, vals in chosen.items():
        if vals:
            df = df[df[col].isin(vals)]
    return df


# ---------- 사이드바: 카테고리 관리 (버튼 콜백) ----------

def set_msg(kind, text):
    """카테고리 관리 결과 메시지를 다음 화면 갱신 때 보여주도록 저장한다."""
    st.session_state.cat_msg = (kind, text)


def cat_add():
    """시험분야 등록: 빈 이름·중복 이름은 거부."""
    name = st.session_state.cat_new.strip()
    if not name or name in st.session_state.fields:
        return set_msg("error", "비어 있거나 이미 있는 이름입니다.")
    st.session_state.fields.append(name)
    st.session_state.cat_new = ""
    set_msg("success", f"'{name}'을(를) 등록했습니다.")


def cat_rename():
    """시험분야 이름 변경: 목록·의뢰 데이터·필터 선택값을 함께 갱신."""
    old, new = st.session_state.cat_ren_from, st.session_state.cat_ren_to.strip()
    fields, p = st.session_state.fields, st.session_state.projects
    if not new or new in fields:
        return set_msg("error", "비어 있거나 이미 있는 이름입니다.")
    fields[fields.index(old)] = new
    mask = p["시험분야"] == old
    p.loc[mask, "시험분야"] = new
    st.session_state["f_시험분야"] = [new if x == old else x for x in st.session_state.get("f_시험분야", [])]
    st.session_state.cat_ren_to = ""
    set_msg("success", f"'{old}' → '{new}' 변경, 의뢰 {int(mask.sum())}건 갱신")


def cat_reclassify():
    """기존 분야의 의뢰를 모두 다른 분야로 옮긴다 (사용 중인 분야를 삭제하기 전 단계)."""
    src, dst = st.session_state.cat_rc_from, st.session_state.cat_rc_to
    if src == dst:
        return set_msg("error", "서로 다른 분야를 선택하세요.")
    p = st.session_state.projects
    mask = p["시험분야"] == src
    p.loc[mask, "시험분야"] = dst
    set_msg("success", f"의뢰 {int(mask.sum())}건을 '{src}' → '{dst}'로 재분류했습니다.")


def cat_delete():
    """시험분야 삭제: 의뢰에서 사용 중이면 거부하고 재분류를 안내."""
    name = st.session_state.cat_del
    used = int((st.session_state.projects["시험분야"] == name).sum())
    if used:
        return set_msg("error", f"'{name}'은(는) 의뢰 {used}건에서 사용 중이라 삭제할 수 없습니다. "
                                "'재분류' 탭에서 다른 분야로 옮긴 뒤 삭제하세요.")
    st.session_state.fields.remove(name)
    set_msg("success", f"'{name}'을(를) 삭제했습니다.")


def sidebar_categories():
    """사이드바 '카테고리 관리' 영역: 시험분야 등록 / 이름 변경 / 재분류 / 삭제."""
    sb = st.sidebar
    sb.divider()
    sb.header("카테고리 관리")
    fields = st.session_state.fields
    used = st.session_state.projects["시험분야"].value_counts()
    sb.caption("시험분야(사용 건수): " + ", ".join(f"{f}({used.get(f, 0)})" for f in fields))
    if msg := st.session_state.pop("cat_msg", None):
        getattr(sb, msg[0])(msg[1])
    for k in ("cat_ren_from", "cat_rc_from", "cat_rc_to", "cat_del"):
        keep_valid(k, fields)
    add, ren, rc, rm = sb.tabs(["등록", "이름 변경", "재분류", "삭제"])
    add.text_input("새 시험분야", key="cat_new")
    add.button("등록", on_click=cat_add)
    ren.selectbox("변경할 분야", fields, key="cat_ren_from")
    ren.text_input("새 이름", key="cat_ren_to")
    ren.button("이름 변경", on_click=cat_rename)
    rc.selectbox("기존 분야", fields, key="cat_rc_from")
    rc.selectbox("옮길 분야", fields, key="cat_rc_to")
    rc.button("재분류", on_click=cat_reclassify)
    rm.selectbox("삭제할 분야", fields, key="cat_del")
    rm.button("삭제", on_click=cat_delete)


# ---------- 본문 구성 요소 ----------

def render_resources():
    """시험 자원 상태 카드(읽기 전용). CSV 상태값을 그대로 보여주며 코드가 상태를 바꾸지 않는다."""
    st.subheader("시험 자원 현황")
    if not RESOURCE_PATH.exists():
        st.info(f"자원 데이터 파일이 없습니다: `data/{RESOURCE_PATH.name}`")
        return
    try:
        res = load_csv(RESOURCE_PATH)
    except Exception as e:
        st.error(f"자원 파일을 읽지 못했습니다: {e}")
        return
    if not {"자원", "상태"} <= set(res.columns) or res.empty:
        st.error("자원 파일에 '자원', '상태' 필드가 필요합니다.")
        return
    for col, r in zip(st.columns(len(res)), res.to_dict("records")):
        color = RES_COLORS.get(r["상태"], "#6b7280")
        col.markdown(f'<div class="res" style="border-left:6px solid {color}">'
                     f'<div class="name">{r["자원"]}</div>'
                     f'<div class="state" style="color:{color}">{r["상태"]}</div>'
                     f'<div class="note">현황 조회용이며 예약·시험 재개를 승인하는 기능이 아님</div></div>',
                     unsafe_allow_html=True)


def render_status_cards(view):
    """상태별 요약 카드(클릭하면 표가 해당 상태로 좁혀짐). (선택 카드 이름, 해당 행)을 반환."""
    groups = {"전체": view, **{s: view[view["진행상태"] == s] for s in STATUSES},
              "지연 후보": view[view["일정상태"] == DELAY]}
    if st.session_state.get("sel") not in groups:
        st.session_state.sel = "전체"
    sel = st.session_state.sel
    for i, (col, (label, rows)) in enumerate(zip(st.columns(len(groups)), groups.items())):
        col.button(f"{label}\n\n**{len(rows)}**", key=f"card{i}", width="stretch",
                   type="primary" if label == sel else "secondary",
                   on_click=st.session_state.__setitem__, args=("sel", label))
    return sel, groups[sel]


def highlight_delay(row):
    """지연 후보 행을 빨간색 계열로 칠한다 (Styler.apply 용)."""
    style = "background-color:#fde2e2;color:#9b1c1c" if row["일정상태"] == DELAY else ""
    return [style] * len(row)


def count_by(df, col, name):
    """col 값별 건수를 [col, name] 두 열 표로 만든다 (차트·엑셀 집계 공용)."""
    return df[col].value_counts().rename_axis(col).reset_index(name=name)


def render_charts(df):
    """시험분야별 건수 막대그래프와 진행상태 비율 도넛 차트 (빈 데이터는 호출 전에 걸러짐)."""
    left, right = st.columns(2)
    left.subheader("시험분야별 건수")
    left.altair_chart(alt.Chart(count_by(df, "시험분야", "건수")).mark_bar(color="#1a56db").encode(
        x=alt.X("시험분야:N", sort="-y", axis=alt.Axis(labelAngle=0)), y="건수:Q",
        tooltip=["시험분야", "건수"]), width="stretch")
    right.subheader("진행상태 비율")
    right.altair_chart(alt.Chart(count_by(df, "진행상태", "건수")).mark_arc(innerRadius=60).encode(
        theta="건수:Q",
        color=alt.Color("진행상태:N", sort=STATUSES, scale=alt.Scale(domain=STATUSES, range=STATUS_COLORS)),
        tooltip=["진행상태", "건수"]), width="stretch")


def to_xlsx(df, conditions):
    """프로젝트목록 / 상태별집계 / 담당자별집계 / 분석조건 시트를 담은 xlsx 바이트를 만든다."""
    status = df["진행상태"].value_counts().reindex(STATUSES, fill_value=0).rename_axis("진행상태").reset_index(name="건수")
    owner = df.groupby("담당엔지니어").agg(
        건수=("의뢰번호", "size"), 지연후보=("일정상태", lambda s: int((s == DELAY).sum()))).reset_index()
    sheets = {"프로젝트목록": df, "상태별집계": status, "담당자별집계": owner,
              "분석조건": pd.DataFrame(list(conditions.items()), columns=["항목", "값"])}
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, data in sheets.items():
            data.to_excel(xw, sheet_name=name, index=False, startrow=2)
            xw.sheets[name]["A1"] = EXCEL_NOTICE
    return buf.getvalue()


def conditions(base, scope, chosen, sel):
    """분석조건 시트에 기록할 기준일·다운로드 범위·필터 값을 dict로 만든다."""
    c = {"기준일": str(base), "다운로드 범위": scope, "생성일시": datetime.now().strftime("%Y-%m-%d %H:%M")}
    c.update({f"필터: {k}": ", ".join(v) or "전체" for k, v in chosen.items()})
    c["상태 카드 선택"] = sel
    return c


def save_memo(rid):
    """상세 영역에서 편집한 메모를 세션 데이터에 저장한다."""
    p = st.session_state.projects
    p.loc[p["의뢰번호"] == rid, "메모"] = st.session_state[f"memo_{rid}"]
    st.session_state.memo_msg = f"{rid} 메모를 저장했습니다 (현재 세션에만 보관)."


def render_detail(df):
    """의뢰번호를 고르면 상세 정보를 보여주고 담당자 메모를 수정할 수 있게 한다."""
    st.subheader("의뢰 상세조회")
    ids = df["의뢰번호"].tolist()
    if not ids:
        st.info("조회할 의뢰가 없습니다.")
        return
    keep_valid("detail_id", ids)
    rid = st.selectbox("의뢰번호", ids, key="detail_id")
    row = df[df["의뢰번호"] == rid].iloc[0]
    v = lambda c: "-" if pd.isna(row[c]) or row[c] == "" else row[c]
    left, right = st.columns(2)
    left.markdown(f"""
- **고객사**: {v('고객사')}
- **제품·모델**: {v('제품명')} / {v('모델명')}
- **시험분야**: {v('시험분야')}
- **적용규격**: {v('적용규격')}
- **담당엔지니어**: {v('담당엔지니어')}
- **시료상태**: {v('시료상태')}
- **완료예정일**: {v('완료예정일')} ({v('일정상태')})
""")
    right.text_area("담당자 메모", value=row["메모"], key=f"memo_{rid}")
    right.button("메모 저장", on_click=save_memo, args=(rid,))
    if msg := st.session_state.pop("memo_msg", None):
        right.success(msg)


# ---------- 화면 ----------

st.markdown('<div class="notice">교육용 가상 데이터 기준 — 실제 업무 시스템 아님</div>',
            unsafe_allow_html=True)
st.title("스탠다드뱅크 인증 프로젝트 접수·진행현황")

init_state()
projects = st.session_state.projects
st.caption("※ 시험분야·메모 변경 내용은 현재 브라우저 세션(session_state)에만 보관되며, "
           "새로고침하거나 세션이 종료되면 사라질 수 있습니다. 원본 CSV는 바뀌지 않습니다.")

chosen = sidebar_filters(projects)
sidebar_categories()

base = st.date_input("기준일", value=date(2026, 10, 23))
render_resources()

data = add_schedule(projects, base)
view = apply_filters(received_by(data, base), chosen)

st.subheader("진행현황 요약")
sel, shown = render_status_cards(view)

if view.empty:
    st.warning("조건에 맞는 인증 프로젝트가 없습니다")
else:
    st.subheader(f"{'전체 의뢰' if sel == '전체' else sel} 목록 ({base} 기준, {len(shown)}건)")
    st.dataframe(shown.style.apply(highlight_delay, axis=1), width="stretch", hide_index=True)
    render_charts(view)

today = date.today().strftime("%Y%m%d")
d1, d2 = st.columns(2)
d1.download_button("전체 목록 다운로드 (.xlsx)", on_click="ignore", mime=XLSX,
                   data=to_xlsx(data, conditions(base, "전체 목록 (필터 미적용)", {}, "전체")),
                   file_name=f"standardbank_project_status_{today}.xlsx")
d2.download_button("현재 필터 결과 다운로드 (.xlsx)", on_click="ignore", mime=XLSX,
                   data=to_xlsx(shown, conditions(base, "현재 필터 결과 (기준일까지 접수분)", chosen, sel)),
                   file_name=f"standardbank_project_status_filtered_{today}.xlsx")

render_detail(data)
