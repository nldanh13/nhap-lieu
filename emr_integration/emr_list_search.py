# -*- coding: utf-8 -*-
"""Luồng EMR tối giản: đăng nhập -> D/S Thủ thuật hoặc D/S Phẫu thuật
-> chọn thời gian 3 tháng -> tìm người bệnh -> trả dữ liệu đang hiển thị.

Mô-đun dừng tại danh sách kết quả, không mở hồ sơ và không nhập/chỉnh sửa dữ liệu.
"""
from __future__ import annotations

import json
import re
import sys
import time
import unicodedata
from calendar import monthrange
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    _HAS_SELENIUM = True
except ModuleNotFoundError:  # pragma: no cover
    webdriver = ChromeOptions = ChromeService = By = Keys = EC = WebDriverWait = None  # type: ignore
    _HAS_SELENIUM = False


LIST_TYPES: Dict[str, Dict[str, str]] = {
    "procedure": {
        "wpid": "danhsachthuthuatdraw",
        "label": "D/s Thủ thuật",
    },
    "surgery": {
        "wpid": "danhsachphauthuatdraw",
        "label": "D/s Phẫu thuật",
    },
}

SEARCH_INPUT_IDS: Tuple[str, ...] = (
    "txtTimKiem",
    "txtSearch",
    "txtKeyword",
    "txtTuKhoa",
    "txtSearchAll",
    "txtMaBN",
    "txtHoTen",
    "txtTukhoa",
    "txtSearchString",
)

LogFn = Callable[[str], None]


def _default_log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


def _norm(value: Any) -> str:
    text = unicodedata.normalize("NFD", str(value or ""))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", text.replace("đ", "d").replace("Đ", "D").lower()).strip()


def _safe_key(value: Any, fallback: str) -> str:
    key = _norm(value)
    key = re.sub(r"[^a-z0-9]+", "_", key).strip("_")
    return key or fallback


def _subtract_months(value: datetime, months: int) -> datetime:
    total = value.year * 12 + (value.month - 1) - max(0, int(months))
    year, month0 = divmod(total, 12)
    month = month0 + 1
    day = min(value.day, monthrange(year, month)[1])
    return value.replace(year=year, month=month, day=day)


def load_config(path: str | Path) -> Dict[str, Any]:
    config_path = Path(path).expanduser().resolve()
    if not config_path.is_file():
        raise FileNotFoundError(f"Không tìm thấy config: {config_path}")
    with config_path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("Config phải là một JSON object")
    return data


@dataclass
class SearchResult:
    list_type: str
    list_label: str
    patient_query: str
    period: str
    current_url: str
    rows: List[Dict[str, Any]]
    table_headers: List[str]
    search_input_id: str
    search_trigger: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            "ok": True,
            "list_type": self.list_type,
            "list_label": self.list_label,
            "patient_query": self.patient_query,
            "period": self.period,
            "current_url": self.current_url,
            "count": len(self.rows),
            "table_headers": self.table_headers,
            "search_input_id": self.search_input_id,
            "search_trigger": self.search_trigger,
            "rows": self.rows,
        }


class EmrListSearcher:
    """Có thể dùng trực tiếp trong ứng dụng khác.

    Ví dụ:
        emr = EmrListSearcher(config, headless=False)
        emr.start()
        result = emr.search("procedure", "26069905", months=3)
        # Browser vẫn đứng tại danh sách kết quả.
        emr.close()
    """

    def __init__(
        self,
        config: Dict[str, Any],
        *,
        headless: bool = False,
        timeout: int = 30,
        log: Optional[LogFn] = None,
    ) -> None:
        self.config = dict(config or {})
        self.headless = bool(headless)
        self.timeout = max(10, int(timeout or 30))
        self.log = log or _default_log
        self.driver: Any = None
        self.wait: Any = None

    def __enter__(self) -> "EmrListSearcher":
        self.start()
        return self

    def __exit__(self, exc_type: Any, exc: Any, tb: Any) -> None:
        self.close()

    def start(self) -> "EmrListSearcher":
        if self.driver is not None:
            return self
        if not _HAS_SELENIUM:
            raise RuntimeError("Thiếu Selenium. Chạy: pip install -r requirements.txt")

        options = ChromeOptions()
        try:
            options.page_load_strategy = str(
                self.config.get("page_load_strategy") or "eager"
            ).lower()
        except Exception:
            pass
        if self.headless:
            options.add_argument(str(self.config.get("headless_arg") or "--headless"))
            options.add_argument("--window-position=-32000,-32000")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--lang=vi-VN")
        options.add_argument("--window-size=1366,900")
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        self.driver = webdriver.Chrome(service=ChromeService(), options=options)
        self.driver.set_page_load_timeout(int(self.config.get("page_load_timeout") or 30))
        self.driver.set_script_timeout(int(self.config.get("script_timeout") or 30))
        self.wait = WebDriverWait(self.driver, self.timeout)
        if not self.headless:
            try:
                self.driver.maximize_window()
            except Exception:
                pass
        return self

    def close(self) -> None:
        if self.driver is not None:
            try:
                self.driver.quit()
            finally:
                self.driver = None
                self.wait = None

    def login(self) -> None:
        self.start()
        url = str(self.config.get("url_login") or "").strip()
        username = str(self.config.get("username") or "").strip()
        password = str(self.config.get("password") or "").strip()
        if not url:
            raise RuntimeError("Thiếu url_login trong config")
        if not username or not password:
            raise RuntimeError("Thiếu username hoặc password trong config")

        user_id = str(self.config.get("login_user_field") or "txtLoginName")
        pass_id = str(self.config.get("login_pass_field") or "txtPassword")
        button_id = str(self.config.get("login_button_field") or "btnLogin")

        self.log(f"[LOGIN] {url}")
        self.driver.get(url)
        if self._is_logged_in():
            return

        last_error: Optional[Exception] = None
        for attempt in range(1, 4):
            try:
                user_el = self._find_first(
                    [
                        (By.ID, user_id),
                        (By.NAME, user_id),
                        (By.XPATH, "//input[@type='text' or @type='email']"),
                    ],
                    visible=True,
                )
                pass_el = self._find_first(
                    [
                        (By.ID, pass_id),
                        (By.NAME, pass_id),
                        (By.XPATH, "//input[@type='password']"),
                    ],
                    visible=True,
                )
                self._replace_text(user_el, username)
                self._replace_text(pass_el, password)

                clicked = False
                for by, value in (
                    (By.ID, button_id),
                    (By.NAME, button_id),
                    (By.XPATH, "//button[contains(normalize-space(),'Đăng nhập') or contains(normalize-space(),'Login')]"),
                    (By.XPATH, "//input[@type='submit']"),
                ):
                    try:
                        button = self.driver.find_element(by, value)
                        if button.is_displayed() and button.is_enabled():
                            self.driver.execute_script("arguments[0].click();", button)
                            clicked = True
                            break
                    except Exception:
                        continue
                if not clicked:
                    pass_el.send_keys(Keys.ENTER)

                WebDriverWait(self.driver, 12).until(lambda _: self._is_logged_in())
                self.log(f"[LOGIN] Thành công: {self.driver.current_url}")
                return
            except Exception as exc:
                last_error = exc
                self.log(f"[LOGIN] Lần {attempt} chưa thành công")
                time.sleep(1)

        raise RuntimeError(f"Đăng nhập EMR không thành công: {last_error}")

    def search(self, list_type: str, patient_query: str, *, months: int = 3) -> SearchResult:
        kind = str(list_type or "").strip().lower()
        if kind not in LIST_TYPES:
            raise ValueError("list_type chỉ nhận 'procedure' hoặc 'surgery'")
        query = str(patient_query or "").strip()
        if not query:
            raise ValueError("Thiếu mã hoặc tên người bệnh")

        if not self._is_logged_in():
            self.login()

        meta = LIST_TYPES[kind]
        self._open_list_page(meta["wpid"], meta["label"])
        period = self._select_recent_months(months)
        field_id = self._set_search_query(query)
        before_signature = self._table_signature()
        trigger = self._click_search_button()
        self._wait_for_results(query, before_signature)
        headers, rows = self._extract_visible_rows(query, kind)

        self.log(f"[RESULT] {meta['label']}: {len(rows)} dòng")
        return SearchResult(
            list_type=kind,
            list_label=meta["label"],
            patient_query=query,
            period=period,
            current_url=str(self.driver.current_url or ""),
            rows=rows,
            table_headers=headers,
            search_input_id=field_id,
            search_trigger=trigger,
        )

    def _is_logged_in(self) -> bool:
        if self.driver is None:
            return False
        current = str(self.driver.current_url or "").lower()
        html = str(self.driver.page_source or "").lower()
        return (
            "login.aspx" not in current
            and (
                "home.aspx" in current
                or "wpid=" in current
                or "đăng xuất" in html
                or "logout" in html
            )
        )

    def _find_first(
        self,
        selectors: Sequence[Tuple[str, str]],
        *,
        visible: bool = False,
        timeout: Optional[int] = None,
    ) -> Any:
        wait = WebDriverWait(self.driver, timeout or self.timeout)
        last_error: Optional[Exception] = None
        for by, value in selectors:
            try:
                condition = (
                    EC.visibility_of_element_located((by, value))
                    if visible
                    else EC.presence_of_element_located((by, value))
                )
                return wait.until(condition)
            except Exception as exc:
                last_error = exc
        raise last_error or RuntimeError("Không tìm thấy phần tử")

    def _replace_text(self, element: Any, value: str) -> None:
        try:
            element.click()
            element.send_keys(Keys.CONTROL, "a")
            element.send_keys(value)
        except Exception:
            self.driver.execute_script(
                "arguments[0].value=arguments[1];"
                "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
                element,
                value,
            )

    def _open_list_page(self, wpid: str, label: str) -> None:
        current = str(self.driver.current_url or "")
        target = self._find_sidebar_url(wpid) or self._build_wpid_url(current, wpid)
        self.log(f"[NAV] {label}: {target}")

        clicked = False
        try:
            clicked = bool(
                self.driver.execute_script(
                    """
                    const wpid = arguments[0].toLowerCase();
                    const links = Array.from(document.querySelectorAll('a[href]'));
                    const link = links.find(a => String(a.getAttribute('href') || '').toLowerCase().includes('wpid=' + wpid));
                    if (!link) return false;
                    try {
                      const parent = link.closest('ul.collapse');
                      if (parent) { parent.style.height='auto'; parent.classList.add('in'); }
                    } catch(e) {}
                    link.scrollIntoView({block:'center'});
                    link.click();
                    return true;
                    """,
                    wpid,
                )
            )
        except Exception:
            clicked = False

        if clicked:
            time.sleep(0.8)
        if wpid.lower() not in str(self.driver.current_url or "").lower():
            self.driver.get(target)

        try:
            WebDriverWait(self.driver, 15).until(
                lambda d: wpid.lower() in str(d.current_url or "").lower()
                or d.find_elements(By.ID, "txtTimKiem")
                or d.find_elements(By.ID, "cbbLoai")
            )
        except Exception:
            pass

        if "login.aspx" in str(self.driver.current_url or "").lower():
            self.login()
            self.driver.get(self._build_wpid_url(str(self.driver.current_url or ""), wpid))
        time.sleep(0.5)

    def _find_sidebar_url(self, wpid: str) -> str:
        try:
            links = self.driver.find_elements(By.CSS_SELECTOR, "a[href]")
            for link in links:
                href = str(link.get_attribute("href") or "")
                if f"wpid={wpid}".lower() in href.lower():
                    return href
        except Exception:
            pass
        return ""

    def _build_wpid_url(self, current_url: str, wpid: str) -> str:
        configured_key = (
            "url_procedure_list"
            if wpid == LIST_TYPES["procedure"]["wpid"]
            else "url_surgery_list"
        )
        configured = str(self.config.get(configured_key) or "").strip()
        login_url = str(self.config.get("url_login") or "").strip()

        current_parsed = urlparse(current_url or login_url)
        configured_parsed = urlparse(configured) if configured else current_parsed
        current_query = dict(parse_qsl(current_parsed.query, keep_blank_values=True))
        configured_query = (
            dict(parse_qsl(configured_parsed.query, keep_blank_values=True))
            if configured
            else {}
        )

        # Giữ token phiên/role từ URL hiện tại; config chỉ override khi có giá trị rõ ràng.
        clean_query: Dict[str, str] = {}
        for key in ("scope", "lang", "role", "usid", "st"):
            value = configured_query.get(key) or current_query.get(key)
            if value:
                clean_query[key] = value
        clean_query["wpid"] = wpid

        scheme = configured_parsed.scheme or current_parsed.scheme or "http"
        netloc = configured_parsed.netloc or current_parsed.netloc
        path = configured_parsed.path or current_parsed.path or "/home.aspx"
        if path.lower().endswith("login.aspx"):
            path = "/home.aspx"
        return urlunparse((scheme, netloc, path, "", urlencode(clean_query), ""))

    def _select_recent_months(self, months: int) -> str:
        months = max(1, int(months or 3))
        now = datetime.now()
        start = _subtract_months(now, months)
        start_dmy = start.strftime("%d/%m/%Y")
        end_dmy = now.strftime("%d/%m/%Y")

        result = self.driver.execute_script(
            r"""
            const months = Number(arguments[0] || 3);
            const startDmy = arguments[1];
            const endDmy = arguments[2];
            const norm = s => String(s || '').normalize('NFD')
              .replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D')
              .toLowerCase().replace(/\s+/g,' ').trim();
            const fire = (el, name) => { try { el.dispatchEvent(new Event(name,{bubbles:true})); } catch(e) {} };
            const setValue = (el, value) => {
              if (!el) return false;
              try {
                const proto = Object.getPrototypeOf(el);
                const desc = proto && Object.getOwnPropertyDescriptor(proto, 'value');
                if (desc && desc.set) desc.set.call(el, value); else el.value = value;
              } catch(e) { el.value = value; }
              el.setAttribute('value', value);
              fire(el,'input'); fire(el,'change'); fire(el,'blur');
              try { if (window.jQuery) window.jQuery(el).val(value).trigger('input').trigger('change').trigger('dp.change'); } catch(e) {}
              return true;
            };

            const sel = document.getElementById('cbbLoai');
            if (!sel) return {mode:'missing', label:'', fallback:false};
            const opts = Array.from(sel.options || []);
            const monthText = String(months) + ' thang';
            let target = opts.find(o => norm(o.textContent || o.innerText).includes(monthText));
            let fallback = false;
            if (!target) {
              target = opts.find(o => norm(o.textContent || o.innerText).includes('khoang'))
                    || opts.find(o => String(o.value || '').trim() === '7');
              fallback = true;
            }
            if (!target) return {mode:'not_found', label:'', fallback:false};

            sel.value = target.value;
            for (const opt of opts) opt.selected = opt === target;
            fire(sel,'change');
            try { if (typeof ThoiGianValueChange === 'function') ThoiGianValueChange(sel); } catch(e) {}
            try { if (window.jQuery) window.jQuery(sel).val(target.value).trigger('change').trigger('change.select2'); } catch(e) {}

            const label = String(target.textContent || target.innerText || '').trim();
            const select2 = document.getElementById('select2-cbbLoai-container');
            if (select2) { select2.textContent = label; select2.setAttribute('title', label); }

            if (fallback) {
              const box = document.getElementById('data_5');
              if (box) { box.style.display='block'; box.style.visibility='visible'; }
              setValue(document.getElementById('dtTuNgay'), '00:00 ' + startDmy);
              setValue(document.getElementById('dtDenNgay'), '23:59 ' + endDmy);
            }
            return {mode: fallback ? 'range_fallback' : 'preset', label, fallback};
            """,
            months,
            start_dmy,
            end_dmy,
        ) or {}

        mode = str(result.get("mode") or "")
        label = str(result.get("label") or "").strip()
        if mode == "preset":
            period = label or f"{months} tháng"
        elif mode == "range_fallback":
            period = f"Khoảng {start_dmy} - {end_dmy}"
        else:
            # Một số bản EMR đã mặc định 3 tháng và không render select ngay.
            period = label or f"{months} tháng (giữ bộ lọc hiện tại)"
        self.log(f"[FILTER] {period}")
        time.sleep(0.3)
        return period

    def _set_search_query(self, query: str) -> str:
        for field_id in SEARCH_INPUT_IDS:
            try:
                element = self.driver.find_element(By.ID, field_id)
                if element.is_displayed() and element.is_enabled():
                    self._replace_text(element, query)
                    self.log(f"[SEARCH] Nhập '{query}' vào #{field_id}")
                    return field_id
            except Exception:
                continue

        selectors = (
            "input[placeholder*='Tìm kiếm']",
            "input[placeholder*='tìm kiếm']",
            "input[type='search']",
        )
        for selector in selectors:
            try:
                element = self.driver.find_element(By.CSS_SELECTOR, selector)
                if element.is_displayed() and element.is_enabled():
                    self._replace_text(element, query)
                    return selector
            except Exception:
                continue
        raise RuntimeError("Không tìm thấy ô tìm kiếm người bệnh")

    def _click_search_button(self) -> str:
        candidates = (
            (By.ID, "btnTimKiem", "#btnTimKiem"),
            (By.XPATH, "//button[contains(normalize-space(),'Tìm kiếm') or normalize-space()='Tìm']", "button-text"),
            (By.XPATH, "//a[contains(normalize-space(),'Tìm kiếm') or normalize-space()='Tìm']", "link-text"),
            (By.CSS_SELECTOR, "button[id*='Tim'],button[id*='Search'],input[id*='Tim'],input[id*='Search']", "search-like"),
        )
        for by, selector, label in candidates:
            try:
                button = self.driver.find_element(by, selector)
                if button.is_displayed() and button.is_enabled():
                    self.driver.execute_script(
                        "arguments[0].scrollIntoView({block:'center'}); arguments[0].click();",
                        button,
                    )
                    self.log(f"[SEARCH] Đã bấm Tìm kiếm bằng {label}")
                    return label
            except Exception:
                continue

        try:
            ok = bool(
                self.driver.execute_script(
                    "if (typeof FilterChange === 'function') { FilterChange(); return true; } return false;"
                )
            )
            if ok:
                self.log("[SEARCH] Đã gọi FilterChange()")
                return "FilterChange"
        except Exception:
            pass
        raise RuntimeError("Không tìm thấy nút Tìm kiếm và không có FilterChange()")

    def _table_signature(self) -> str:
        try:
            return str(
                self.driver.execute_script(
                    """
                    return Array.from(document.querySelectorAll('table tbody tr'))
                      .filter(r => r.offsetParent !== null)
                      .map(r => (r.innerText || '').trim()).join('\n');
                    """
                )
                or ""
            )
        except Exception:
            return ""

    def _wait_for_results(self, query: str, before_signature: str) -> None:
        query_norm = _norm(query)

        def ready(driver: Any) -> bool:
            try:
                state = driver.execute_script(
                    r"""
                    const q = arguments[0];
                    const norm = s => String(s || '').normalize('NFD')
                      .replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D')
                      .toLowerCase().replace(/\s+/g,' ').trim();
                    const rows = Array.from(document.querySelectorAll('table tbody tr, table tr'))
                      .filter(r => r.offsetParent !== null && r.querySelectorAll('td').length > 0);
                    const texts = rows.map(r => norm(r.innerText || ''));
                    const loading = Array.from(document.querySelectorAll('.loading,.dataTables_processing,[id*=loading],[class*=loading]'))
                      .some(el => el.offsetParent !== null && norm(el.innerText || el.textContent || '').includes('dang'));
                    return {match:texts.some(t => t.includes(q)), count:rows.length, loading};
                    """,
                    query_norm,
                ) or {}
                if state.get("loading"):
                    return False
                if state.get("match"):
                    return True
                current = self._table_signature()
                return bool(current != before_signature and state.get("count", 0) >= 0)
            except Exception:
                return False

        try:
            WebDriverWait(self.driver, 18).until(ready)
        except Exception:
            # Danh sách không có kết quả vẫn là trạng thái hợp lệ.
            time.sleep(1.0)

    def _extract_visible_rows(self, query: str, list_type: str) -> Tuple[List[str], List[Dict[str, Any]]]:
        raw_tables = self.driver.execute_script(
            r"""
            const q = arguments[0];
            const norm = s => String(s || '').normalize('NFD')
              .replace(/[\u0300-\u036f]/g,'').replace(/đ/g,'d').replace(/Đ/g,'D')
              .toLowerCase().replace(/\s+/g,' ').trim();
            return Array.from(document.querySelectorAll('table')).map((table, tableIndex) => {
              const headers = Array.from(table.querySelectorAll('thead th')).map(x => (x.innerText || x.textContent || '').trim());
              if (!headers.length) {
                const first = table.querySelector('tr');
                if (first && first.querySelectorAll('th').length) {
                  headers.push(...Array.from(first.querySelectorAll('th')).map(x => (x.innerText || x.textContent || '').trim()));
                }
              }
              const rows = Array.from(table.querySelectorAll('tbody tr, tr')).filter(row => {
                return row.offsetParent !== null && row.querySelectorAll('td').length > 0;
              }).map(row => ({
                cells: Array.from(row.querySelectorAll('td')).map(td => (td.innerText || td.textContent || '').replace(/\s+/g,' ').trim()),
                text: (row.innerText || row.textContent || '').replace(/\s+/g,' ').trim(),
                access_id: row.getAttribute('access_id') || '',
                links: Array.from(row.querySelectorAll('a')).map(a => ({
                  text: (a.innerText || a.textContent || '').replace(/\s+/g,' ').trim(),
                  href: a.href || a.getAttribute('href') || '',
                  onclick: a.getAttribute('onclick') || ''
                }))
              }));
              const matchCount = rows.filter(r => norm(r.text).includes(q)).length;
              return {tableIndex, headers, rows, matchCount};
            });
            """,
            _norm(query),
        ) or []

        tables = [t for t in raw_tables if isinstance(t, dict) and t.get("rows")]
        if not tables:
            return [], []
        tables.sort(key=lambda t: (int(t.get("matchCount") or 0), len(t.get("rows") or [])), reverse=True)
        chosen = tables[0]
        headers = [str(x or "").strip() for x in (chosen.get("headers") or [])]
        all_rows = chosen.get("rows") or []
        query_norm = _norm(query)
        matched = [row for row in all_rows if query_norm in _norm(row.get("text"))]
        selected_rows = matched if matched else all_rows

        output: List[Dict[str, Any]] = []
        for index, row in enumerate(selected_rows, start=1):
            cells = [str(x or "").strip() for x in (row.get("cells") or [])]
            if not cells or _norm(" ".join(cells)) in {"khong co du lieu", "no data available in table"}:
                continue

            values: Dict[str, Any] = {}
            used_keys: Dict[str, int] = {}
            for cell_index, cell in enumerate(cells):
                header = headers[cell_index] if cell_index < len(headers) else ""
                base_key = _safe_key(header, f"col_{cell_index + 1}")
                count = used_keys.get(base_key, 0) + 1
                used_keys[base_key] = count
                key = base_key if count == 1 else f"{base_key}_{count}"
                values[key] = cell

            record: Dict[str, Any] = {
                "row_number": index,
                "access_id": str(row.get("access_id") or ""),
                "raw_text": str(row.get("text") or ""),
                "cells": cells,
                "values": values,
                "links": row.get("links") or [],
            }
            if list_type == "surgery":
                known = (
                    "stt", "ma_bn", "ho_ten", "gioi_tinh", "tuoi", "phong_mo",
                    "noi_chuyen_mo", "noi_dung_phau_thuat", "tinh_trang", "thoi_gian",
                    "trang_thai", "doi_tuong", "tam_ung",
                )
                record["known_fields"] = {
                    name: cells[pos] if pos < len(cells) else ""
                    for pos, name in enumerate(known)
                }
            elif list_type == "procedure":
                known = (
                    "stt", "ngay_thu_thuat", "ma_bn", "ho_ten", "gioi_tinh", "tuoi",
                    "doi_tuong", "noi_chi_dinh", "chi_dinh_thu_thuat", "tam_ung",
                    "phai_thanh_toan", "tinh_trang", "thoi_gian_chi_dinh", "trang_thai",
                )
                record["known_fields"] = {
                    name: cells[pos] if pos < len(cells) else ""
                    for pos, name in enumerate(known)
                }
            output.append(record)
        return headers, output

    def open_detail(self, list_type: str, row: Dict[str, Any]) -> Dict[str, Any]:
        """Mở chi tiết một dòng và lấy bác sĩ chính, thời gian, kỹ thuật.

        D/s Phẫu thuật có hai kiểu mở hồ sơ:
        - liên kết trực tiếp chứa ``phauthuatid=``;
        - ``javascript:void(0)`` gọi ``checkBanGiaoPT(access_id)`` rồi
          ``window.open(url)``. Trường hợp thứ hai được bắt URL trước khi mở.
        """
        kind = str(list_type or "").strip().lower()
        if kind not in LIST_TYPES:
            raise ValueError("list_type chỉ nhận 'procedure' hoặc 'surgery'")

        direct_url = ""
        onclick_text = ""
        for link in row.get("links") or []:
            href = str((link or {}).get("href") or "").strip()
            onclick = str((link or {}).get("onclick") or "").strip()
            href_lower = href.lower()
            if kind == "procedure" and "thuthuatid=" in href_lower:
                direct_url = href
                break
            if kind == "surgery" and "phauthuatid=" in href_lower:
                direct_url = href
                break
            if kind == "surgery" and "checkbangiaopt" in onclick.lower():
                onclick_text = onclick

        source_url = str(self.driver.current_url or "")
        access_id = str(row.get("access_id") or "").strip()
        self.log(
            f"[DETAIL] kind={kind}; access_id={access_id or '-'}; "
            f"direct_url={'yes' if direct_url else 'no'}; onclick={'yes' if onclick_text else 'no'}"
        )

        if direct_url:
            detail_url = urljoin(source_url, direct_url)
            self.log(f"[DETAIL] Mở liên kết trực tiếp: {detail_url}")
            self.driver.get(detail_url)
        elif kind == "surgery" and access_id:
            # checkBanGiaoPT dùng Ajax đồng bộ rồi window.open(AjaxOut.HtmlContent).
            # Tạm thay window.open để lấy URL mà không tạo tab ngoài kiểm soát.
            capture = self.driver.execute_script(
                r"""
                const accessId = String(arguments[0] || '');
                const result = {
                  access_id: accessId,
                  row_found: false,
                  anchor_found: false,
                  onclick: '',
                  captured_url: '',
                  function_found: typeof checkBanGiaoPT === 'function',
                  error: ''
                };
                const rows = Array.from(document.querySelectorAll('tr[access_id]'));
                const row = rows.find(r => String(r.getAttribute('access_id') || '') === accessId) || null;
                result.row_found = !!row;
                const anchor = row ? (row.querySelector('a[onclick*="checkBanGiaoPT"],a')) : null;
                result.anchor_found = !!anchor;
                result.onclick = anchor ? String(anchor.getAttribute('onclick') || '') : '';

                const oldOpen = window.open;
                window.open = function(url) {
                  result.captured_url = String(url || '');
                  return { closed: false, close: function(){}, focus: function(){} };
                };
                try {
                  if (typeof checkBanGiaoPT === 'function') {
                    checkBanGiaoPT(accessId);
                  } else if (anchor) {
                    anchor.click();
                  } else {
                    result.error = 'Không tìm thấy checkBanGiaoPT hoặc anchor trong dòng';
                  }
                } catch (e) {
                  result.error = String((e && (e.stack || e.message)) || e || 'unknown error');
                } finally {
                  window.open = oldOpen;
                }
                return result;
                """,
                access_id,
            ) or {}
            self.log(f"[DETAIL] Kết quả gọi checkBanGiaoPT: {json.dumps(capture, ensure_ascii=False)}")

            captured_url = str(capture.get("captured_url") or "").strip()
            if captured_url:
                detail_url = urljoin(source_url, captured_url)
                self.log(f"[DETAIL] URL chi tiết bắt được: {detail_url}")
                self.driver.get(detail_url)
            else:
                # Dự phòng: bấm đúng anchor và theo dõi tab mới/current URL.
                before_url = str(self.driver.current_url or "")
                before_handles = list(self.driver.window_handles)
                click_result = self.driver.execute_script(
                    r"""
                    const accessId = String(arguments[0] || '');
                    const row = Array.from(document.querySelectorAll('tr[access_id]'))
                      .find(r => String(r.getAttribute('access_id') || '') === accessId);
                    if (!row) return {ok:false, reason:'row_not_found'};
                    const anchor = row.querySelector('a[onclick*="checkBanGiaoPT"],a');
                    if (!anchor) return {ok:false, reason:'anchor_not_found'};
                    anchor.scrollIntoView({block:'center'});
                    anchor.click();
                    return {ok:true, onclick:String(anchor.getAttribute('onclick') || '')};
                    """,
                    access_id,
                ) or {}
                self.log(f"[DETAIL] Dự phòng click anchor: {json.dumps(click_result, ensure_ascii=False)}")
                if not click_result.get("ok"):
                    raise RuntimeError(
                        "Đã thấy dòng phẫu thuật nhưng không gọi được checkBanGiaoPT; "
                        f"capture={capture}; click={click_result}"
                    )

                def detail_opened(driver: Any) -> bool:
                    try:
                        if len(driver.window_handles) > len(before_handles):
                            return True
                        if str(driver.current_url or "") != before_url:
                            return True
                        return bool(
                            driver.find_elements(By.ID, "divPhauThuatContent")
                            or driver.find_elements(By.ID, "cbbBacSiPT")
                        )
                    except Exception:
                        return False

                try:
                    WebDriverWait(self.driver, 15).until(detail_opened)
                except Exception:
                    pass
                after_handles = list(self.driver.window_handles)
                if len(after_handles) > len(before_handles):
                    new_handles = [h for h in after_handles if h not in before_handles]
                    self.driver.switch_to.window(new_handles[-1])
                    self.log("[DETAIL] Đã chuyển sang tab chi tiết mới")
                if (
                    str(self.driver.current_url or "") == before_url
                    and not self.driver.find_elements(By.ID, "divPhauThuatContent")
                    and not self.driver.find_elements(By.ID, "cbbBacSiPT")
                ):
                    raise RuntimeError(
                        "Đã click dòng phẫu thuật nhưng trang chi tiết không mở; "
                        f"capture={capture}; click={click_result}; current_url={self.driver.current_url}"
                    )
        else:
            raise RuntimeError(
                "Không tìm thấy cách mở hồ sơ chi tiết; "
                f"kind={kind}; access_id={access_id}; onclick={onclick_text}; links={row.get('links') or []}"
            )

        root_id = "divThuThuatContent" if kind == "procedure" else "divPhauThuatContent"
        # Với D/s Thủ thuật phải lấy Bác sĩ chỉ định, không lấy Thủ thuật viên chính.
        # EMR hiện dùng id txtBacSyChiDinh; giữ thêm biến thể txtBacSiChiDinh
        # để tương thích nếu giao diện được chuẩn hóa lại tên id.
        doctor_ids = (
            ["txtBacSyChiDinh", "txtBacSiChiDinh"]
            if kind == "procedure"
            else ["cbbBacSiPT"]
        )
        # Thời gian đối chiếu của sheet tieuphau là thời gian CHỈ ĐỊNH.
        # Vì vậy ở chi tiết thủ thuật phải ưu tiên txtThoiGianChiDinh, không dùng
        # txtTgBatDau làm nguồn chính. D/s phẫu thuật vẫn dùng giờ bắt đầu mổ.
        start_ids = (
            ["txtThoiGianChiDinh", "txtTgBatDau"]
            if kind == "procedure"
            else ["txtBatDauPT"]
        )
        # Tên kỹ thuật dùng để đối chiếu phải đúng trường nghiệp vụ:
        # - Thủ thuật: Chỉ định (txtChiDinh)
        # - Phẫu thuật: Nội dung/Phương pháp phẫu thuật
        method_ids = (
            ["txtChiDinh", "cbbPhuongPhapTT"]
            if kind == "procedure"
            else ["txtNoiDungPT", "cbbPhuongPhapPT", "cbbChiDinhMoPT"]
        )

        try:
            WebDriverWait(self.driver, 18).until(
                lambda d: d.find_elements(By.ID, root_id)
                or any(d.find_elements(By.ID, field_id) for field_id in doctor_ids)
            )
        except Exception as exc:
            raise RuntimeError(
                f"Trang đã mở nhưng không thấy {root_id}/{'|'.join(doctor_ids)}; "
                f"url={self.driver.current_url}; lỗi={exc}"
            ) from exc

        # Các input ở trang chi tiết được EMR đổ dữ liệu bằng Ajax sau khi phần tử
        # đã xuất hiện. Chờ đến khi Bác sĩ chỉ định/BS mổ chính có GIÁ TRỊ thật,
        # tránh đọc quá sớm rồi báo trống.
        def detail_values_ready(driver: Any) -> bool:
            try:
                state = driver.execute_script(
                    r"""
                    const ids = arguments[0] || [];
                    const read = id => {
                      const el = document.getElementById(id);
                      if (!el) return '';
                      if (el.tagName === 'SELECT') {
                        const opt = el.options && el.selectedIndex >= 0 ? el.options[el.selectedIndex] : null;
                        const text = opt ? (opt.textContent || opt.innerText || '') : '';
                        if (String(text || '').trim()) return String(text).replace(/\s+/g,' ').trim();
                        const rendered = document.getElementById('select2-' + id + '-container');
                        return rendered ? String(rendered.textContent || '').replace(/\s+/g,' ').trim() : '';
                      }
                      return String(el.value || el.textContent || el.innerText || '').replace(/\s+/g,' ').trim();
                    };
                    return {
                      ready: ids.some(id => !!read(id)),
                      values: ids.map(id => ({id, value: read(id)})),
                      ajax_active: (window.jQuery && typeof window.jQuery.active === 'number') ? window.jQuery.active : null
                    };
                    """,
                    doctor_ids,
                ) or {}
                return bool(state.get("ready"))
            except Exception:
                return False

        try:
            WebDriverWait(self.driver, 20).until(detail_values_ready)
        except Exception:
            # Vẫn đọc toàn bộ trạng thái phía dưới để log rõ trường nào còn trống.
            self.log("[DETAIL] Hết thời gian chờ trường bác sĩ được EMR nạp dữ liệu")

        data = self.driver.execute_script(
            r"""
            const doctorIds = arguments[0] || [];
            const startIds = arguments[1] || [];
            const methodIds = arguments[2] || [];
            const read = id => {
              const el = document.getElementById(id);
              if (!el) return '';
              if (el.tagName === 'SELECT') {
                const opt = el.options && el.selectedIndex >= 0 ? el.options[el.selectedIndex] : null;
                const text = opt ? (opt.textContent || opt.innerText || '') : '';
                if (String(text || '').trim()) return String(text).replace(/\s+/g,' ').trim();
                const rendered = document.getElementById('select2-' + id + '-container');
                return rendered ? String(rendered.textContent || '').replace(/\s+/g,' ').trim() : '';
              }
              return String(el.value || el.textContent || el.innerText || '').replace(/\s+/g,' ').trim();
            };
            const readFirst = ids => {
              for (const id of ids || []) {
                const value = read(id);
                if (value) return {id, value};
              }
              return {id:'', value:''};
            };
            const doctorResult = readFirst(doctorIds);
            const startResult = readFirst(startIds);
            const methodResult = readFirst(methodIds);
            return {
              doctor: doctorResult.value,
              doctor_field: doctorResult.id,
              start_time: startResult.value,
              start_time_field: startResult.id,
              method: methodResult.value,
              method_field: methodResult.id,
              patient_info: read('txtTTBN'),
              complete_info: String((document.getElementById('txtTTHoanTat') || {}).textContent || '').replace(/\s+/g,' ').trim(),
              current_url: String(location.href || ''),
              root_found: !!document.getElementById(arguments[3]),
              doctor_found: doctorIds.some(id => !!document.getElementById(id)),
              doctor_candidates: doctorIds.map(id => ({id, value: read(id)})),
              start_candidates: startIds.map(id => ({id, value: read(id)})),
              method_candidates: methodIds.map(id => ({id, value: read(id)})),
              ajax_active: (window.jQuery && typeof window.jQuery.active === 'number') ? window.jQuery.active : null
            };
            """,
            doctor_ids,
            start_ids,
            method_ids,
            root_id,
        ) or {}
        self.log(f"[DETAIL] Dữ liệu đọc được: {json.dumps(data, ensure_ascii=False)}")

        doctor = str(data.get("doctor") or "").strip()
        doctor_field = str(data.get("doctor_field") or "").strip()
        if not doctor:
            field_label = "Bác sĩ chỉ định" if kind == "procedure" else "BS mổ chính"
            raise RuntimeError(
                f"Đã mở chi tiết nhưng trường {field_label} đang trống; "
                f"doctor_ids={'|'.join(doctor_ids)}; url={data.get('current_url') or self.driver.current_url}"
            )
        self.log(
            f"[DETAIL] Nguồn bác sĩ: "
            f"{'Bác sĩ chỉ định' if kind == 'procedure' else 'BS mổ chính'}; "
            f"field={doctor_field or '-'}; value={doctor}"
        )

        return {
            "list_type": kind,
            "doctor": doctor,
            "doctor_field": doctor_field,
            "doctor_source": "Bác sĩ chỉ định" if kind == "procedure" else "BS mổ chính",
            "start_time": str(data.get("start_time") or "").strip(),
            "start_time_field": str(data.get("start_time_field") or "").strip(),
            "method": str(data.get("method") or "").strip(),
            "method_field": str(data.get("method_field") or "").strip(),
            "patient_info": str(data.get("patient_info") or "").strip(),
            "complete_info": str(data.get("complete_info") or "").strip(),
            "current_url": str(data.get("current_url") or self.driver.current_url),
            "source_row": row,
        }


def search_emr_list(
    config: Dict[str, Any],
    list_type: str,
    patient_query: str,
    *,
    months: int = 3,
    headless: bool = True,
    keep_browser: bool = False,
    log: Optional[LogFn] = None,
) -> Dict[str, Any]:
    """Hàm tiện lợi cho ứng dụng khác.

    Khi keep_browser=False, browser tự đóng sau khi lấy kết quả.
    Khi cần giữ trang kết quả để thao tác tiếp, dùng trực tiếp lớp EmrListSearcher.
    """
    client = EmrListSearcher(config, headless=headless, log=log)
    try:
        client.start()
        result = client.search(list_type, patient_query, months=months).to_dict()
        if keep_browser:
            result["_client"] = client
            return result
        return result
    except Exception:
        client.close()
        raise
    finally:
        if not keep_browser:
            client.close()
