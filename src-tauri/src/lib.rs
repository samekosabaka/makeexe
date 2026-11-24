use calamine::{Data, Reader, Xls, Xlsx};
use chrono::{Duration, Local, NaiveDate, NaiveDateTime, NaiveTime, Timelike};
use regex::Regex;
use rust_xlsxwriter::{
    Color, DataValidation, Format, FormatAlign, FormatBorder, Formula, Workbook,
};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::io::Cursor;
use std::sync::Mutex;
use tauri::State;

const CC_COLUMNS: [&str; 15] = [
    "MSNV", "Name", "Date", "Timetable", "On duty", "Off duty", "Clock In", "Clock Out",
    "Late", "Early", "OT Time", "Work Time", "Kết Quả Công", "Nghỉ phép", "Tăng ca",
];

const XP_COLUMNS: [&str; 6] = [
    "MSNV", "Ngày", "*Loại hình xin phép",
    "*Thời gian bắt đầu xin phép", "*Thời gian kết thúc xin phép", "Lý do nghỉ phép",
];

const LEAVE_TYPES: [&str; 9] = [
    "ACTR VIỆC RIÊNG",
    "ACTR NGHỈ BÙ",
    "ACTR PHÉP NĂM",
    "ACTR PHÉP BỆNH",
    "ACTR NGHỈ KHÔNG LƯƠNG",
    "ACTR VIỆC RIÊNG LUẬT",
    "ACTR PHÉP TANG",
    "ACTR KẾT HÔN",
    "ACTR THAI SẢN NAM",
];

macro_rules! row_map {
    ($($k:expr => $v:expr),* $(,)?) => {{
        let mut m = HashMap::new();
        $(m.insert($k.to_string(), $v.to_string());)*
        m
    }};
}

// =========================================================================
// --- CẤU TRÚC DỮ LIỆU & PAYLOAD KHỚP 100% VỚI INDEX.HTML ---
// =========================================================================
#[derive(Clone, Serialize, Deserialize)]
pub struct ShiftConfig {
    pub hc_start: String,
    pub hc_end: String,
    pub ca1_start: String,
    pub ca1_end: String,
    pub ca2_start: String,
    pub ca2_end: String,
    pub ca3_start: String,
    pub ca3_end: String,
}

impl Default for ShiftConfig {
    fn default() -> Self {
        Self {
            hc_start: "07:30".into(),
            hc_end: "17:00".into(),
            ca1_start: "07:00".into(),
            ca1_end: "15:00".into(),
            ca2_start: "15:00".into(),
            ca2_end: "23:00".into(),
            ca3_start: "23:00".into(),
            ca3_end: "07:00".into(),
        }
    }
}

#[derive(Deserialize)]
pub struct DeleteTarget {
    pub base_idx: usize,
    pub part: u8,
}

#[derive(Clone, Default)]
struct NghiBuSummary {
    msnv: String,
    ho_ten: String,
    to_chuc: String,
    vi_tri: String,
    tong_so_du: f64,
    tong_khau_tru: f64,
    con_lai_ban_dau: f64,
    da_dung_xin_phep: f64,
    con_lai_sau_tru: f64,
    don_vi: String,
}

#[derive(Clone)]
struct BaseXpItem {
    msnv: String,
    ngay: String,
    orig_loai: String,
    start: String,
    end: String,
    lydo: String,
    is_extra: bool,
    has_lunch: bool,
    skip_nb: bool,
    user_choice: u8,  // 0 = Tự động, 1 = Phép năm, 2 = Nghỉ bù, 3 = Việc riêng (giữ mọi phép)
    partial_mode: u8, // 0 = TH1 giữ 1 dòng, 1 = TH2 chia 2 dòng, 2 = TH3 đổi dòng trên, 3 = Việc riêng
    manual_part0: Option<String>,
    manual_part1: Option<String>,
    deleted_part0: bool,
    deleted_part1: bool,
}

impl BaseXpItem {
    fn is_fully_deleted(&self) -> bool {
        if self.partial_mode == 1 {
            self.deleted_part0 && self.deleted_part1
        } else {
            self.deleted_part0
        }
    }
}

#[derive(Clone, Default)]
struct NaturalRowInfo {
    is_upper_nb: bool,
    linked_partial_idx: Option<usize>,
    is_partial: bool,
    linked_upper_idx: Option<usize>,
    can_swap_upper: bool,
}

#[derive(Serialize)]
pub struct UiState {
    pub file1: String,
    pub file2: String,
    pub file_xp: String,
    pub file_nb: String,
    pub file_pn: String,
    pub shift_settings: ShiftConfig,
    pub all_dates: Vec<String>,
    pub all_msnvs: Vec<String>,
    pub all_timetables: Vec<String>,
    pub cc_rows: Vec<HashMap<String, String>>,
    pub xp_rows: Vec<HashMap<String, String>>,
}

#[derive(Serialize)]
pub struct CmdResponse {
    pub state: UiState,
    pub message: Option<String>,
    pub is_error: bool,
}

#[derive(Serialize)]
pub struct ExportPayload {
    pub filename: String,
    pub data_base64: String,
}

#[derive(Default)]
struct FileSlot {
    name: String,
    bytes: Vec<u8>,
}

#[derive(Default)]
struct TimekeepApp {
    file1: FileSlot,
    file2: FileSlot,
    file_xp: FileSlot,
    file_nb: FileSlot,
    file_pn: FileSlot,
    data1_clean: Vec<HashMap<String, String>>,
    data2_clean_map: HashMap<String, Vec<NaiveDateTime>>,
    data_result: Vec<HashMap<String, String>>,
    base_xp_items: Vec<BaseXpItem>,
    base_nb_map: BTreeMap<String, NghiBuSummary>,
    base_pn_map: BTreeMap<String, f64>, // Đơn vị: Ngày (1 ngày = 8H)
    data_xp_result: Vec<HashMap<String, String>>,
    nghi_bu_summary: BTreeMap<String, NghiBuSummary>,
    all_dates: Vec<String>,
    all_msnvs: Vec<String>,
    all_timetables: Vec<String>,
    shift_settings: ShiftConfig,
}

impl TimekeepApp {
    fn to_response(&self, message: Option<String>, is_error: bool) -> CmdResponse {
        CmdResponse {
            state: UiState {
                file1: self.file1.name.clone(),
                file2: self.file2.name.clone(),
                file_xp: self.file_xp.name.clone(),
                file_nb: self.file_nb.name.clone(),
                file_pn: self.file_pn.name.clone(),
                shift_settings: self.shift_settings.clone(),
                all_dates: self.all_dates.clone(),
                all_msnvs: self.all_msnvs.clone(),
                all_timetables: self.all_timetables.clone(),
                cc_rows: self.data_result.clone(),
                xp_rows: self.data_xp_result.clone(),
            },
            message,
            is_error,
        }
    }

    // =========================================================================
    // --- LOGIC TAB 1: CHẤM CÔNG ---
    // =========================================================================
    fn process_cham_cong(&mut self) -> Result<String, String> {
        if self.file1.bytes.is_empty() || self.file2.bytes.is_empty() {
            return Err("Vui lòng chọn đủ cả 2 file (File ZKT và Dữ liệu HR)!".into());
        }
        let raw_df1 = read_bytes_to_dicts(&self.file1.bytes)?;
        let raw_df2 = read_bytes_to_dicts(&self.file2.bytes)?;

        if raw_df1.is_empty() || !raw_df1[0].contains_key("AC-No.") {
            return Err("Không tìm thấy cột 'AC-No.' trong File 1!".into());
        }
        let time_cols = [
            "Thời gian quẹt thẻ",
            "Thời gian quẹt thẻ lên ca",
            "Thời gian quẹt thẻ xuống ca",
        ];
        if raw_df2.is_empty()
            || !raw_df2[0].contains_key("Mã số nhân viên")
            || !time_cols.iter().any(|&c| raw_df2[0].contains_key(c))
        {
            return Err("File 2 thiếu cột 'Mã số nhân viên' hoặc cột thời gian quẹt thẻ!".into());
        }

        self.data1_clean.clear();
        for row in &raw_df1 {
            let mut nr = row_map!("MSNV" => normalize_msnv(get_val(row, "AC-No.")));
            for col in [
                "Name", "Date", "Timetable", "On duty", "Off duty", "Clock In", "Clock Out",
                "Late", "Early", "OT Time", "Work Time",
            ] {
                let v = get_val(row, col);
                let clean = if col == "Date" && !v.is_empty() {
                    convert_to_dmy(v)
                } else if col == "Timetable" && !v.is_empty() {
                    clean_shift_spaces(v)
                } else if matches!(col, "On duty" | "Off duty" | "Clock In" | "Clock Out")
                    && !v.is_empty()
                {
                    normalize_time_str(v)
                } else {
                    v.to_string()
                };
                nr.insert(col.into(), clean);
            }
            self.data1_clean.push(nr);
        }

        self.data2_clean_map.clear();
        for row in &raw_df2 {
            let msnv = normalize_msnv(get_val(row, "Mã số nhân viên"));
            if !msnv.is_empty() {
                for &tc in &time_cols {
                    if let Some(dt) = parse_datetime_flexible(get_val(row, tc)) {
                        self.data2_clean_map.entry(msnv.clone()).or_default().push(dt);
                    }
                }
            }
        }

        self.data_result = self.data1_clean.clone();
        for i in 0..self.data_result.len() {
            self.process_single_row(i, false);
        }

        let mut dates: Vec<(NaiveDate, String)> = self
            .data_result
            .iter()
            .filter_map(|r| {
                NaiveDate::parse_from_str(get_val(r, "Date"), "%d/%m/%Y")
                    .ok()
                    .map(|d| (d, get_val(r, "Date").to_string()))
            })
            .collect();
        dates.sort_by_key(|k| k.0);
        dates.dedup_by(|a, b| a.1 == b.1);
        self.all_dates = dates.into_iter().map(|x| x.1).collect();

        self.all_msnvs = self
            .data_result
            .iter()
            .map(|r| get_val(r, "MSNV").to_string())
            .filter(|s| !s.is_empty())
            .collect::<BTreeSet<_>>()
            .into_iter()
            .collect();
        self.refresh_timetable_options();

        Ok(format!("Xử lý xong {} dòng chấm công!", self.data_result.len()))
    }

    fn refresh_timetable_options(&mut self) {
        self.all_timetables = self
            .data1_clean
            .iter()
            .chain(self.data_result.iter())
            .map(|r| get_val(r, "Timetable").to_string())
            .filter(|s| !s.is_empty())
            .collect::<BTreeSet<_>>()
            .into_iter()
            .collect();
    }

    fn process_single_row(&mut self, idx: usize, force_recalc: bool) {
        let orig = self.data1_clean[idx].clone();
        let msnv = get_val(&orig, "MSNV").to_string();
        let cur_tt = get_val(&self.data_result[idx], "Timetable").to_string();
        let tt_up = clean_shift_spaces(&cur_tt.to_uppercase());
        let (mut is_hc, mut is_ca1, mut is_ca2, mut is_ca3) = classify_shift(&tt_up, "", "");

        let (mut on_d, mut off_d) = if !force_recalc {
            (
                normalize_time_str(get_val(&orig, "On duty")),
                normalize_time_str(get_val(&orig, "Off duty")),
            )
        } else if let Some(s) = self.data1_clean.iter().find(|r| {
            get_val(r, "Timetable").eq_ignore_ascii_case(&cur_tt)
                && !get_val(r, "On duty").is_empty()
                && !get_val(r, "Off duty").is_empty()
        }) {
            (
                normalize_time_str(get_val(s, "On duty")),
                normalize_time_str(get_val(s, "Off duty")),
            )
        } else {
            (String::new(), String::new())
        };

        if on_d.is_empty() || off_d.is_empty() {
            (on_d, off_d) = if is_hc {
                if tt_up.contains("15H30") || tt_up.contains("15:30") {
                    ("07:30".into(), "15:30".into())
                } else {
                    (self.shift_settings.hc_start.clone(), self.shift_settings.hc_end.clone())
                }
            } else if is_ca1 {
                (self.shift_settings.ca1_start.clone(), self.shift_settings.ca1_end.clone())
            } else if is_ca2 {
                (self.shift_settings.ca2_start.clone(), self.shift_settings.ca2_end.clone())
            } else if is_ca3 {
                (self.shift_settings.ca3_start.clone(), self.shift_settings.ca3_end.clone())
            } else {
                (on_d, off_d)
            };
        }
        (is_hc, is_ca1, is_ca2, is_ca3) = classify_shift(&tt_up, &on_d, &off_d);

        let mut avail: Vec<String> = [
            normalize_time_str(get_val(&orig, "Clock In")),
            normalize_time_str(get_val(&orig, "Clock Out")),
        ]
        .into_iter()
        .filter(|s| !s.is_empty() && s != "Bù vân tay")
        .collect();

        let (mut c_in, mut c_out) = (String::new(), String::new());
        if is_ca3 {
            for t in &avail {
                if let Some(hm) = parse_hm(t) {
                    if hm.hour() >= 15 && c_in.is_empty() {
                        c_in = t.clone();
                    } else if hm.hour() < 14 {
                        c_out = t.clone();
                    }
                }
            }
        }
        if c_in.is_empty() && c_out.is_empty() {
            if avail.len() == 2 {
                avail.sort();
                (c_in, c_out) = if is_ca3 {
                    (avail[1].clone(), avail[0].clone())
                } else {
                    (avail[0].clone(), avail[1].clone())
                };
            } else if avail.len() == 1 {
                let t = avail[0].clone();
                if let (Some(t_dt), Some(on_dt), Some(off_dt)) =
                    (parse_hm(&t), parse_hm(&on_d), parse_hm(&off_d))
                {
                    if min_dist_seconds(t_dt, on_dt) <= min_dist_seconds(t_dt, off_dt) {
                        c_in = t;
                    } else {
                        c_out = t;
                    }
                } else {
                    c_in = t;
                }
            }
        }

        let mut filled_hr = false;
        if c_in.is_empty() || c_out.is_empty() {
            if let (Ok(d_val), Some(on_t), Some(off_t), Some(swipes)) = (
                NaiveDate::parse_from_str(get_val(&orig, "Date"), "%d/%m/%Y"),
                parse_hm(&on_d),
                parse_hm(&off_d),
                self.data2_clean_map.get(&msnv),
            ) {
                if is_ca3 {
                    let t_in = (d_val - Duration::days(1)).and_time(on_t);
                    let mut t_out = d_val.and_time(off_t);
                    if t_out <= t_in {
                        t_out += Duration::days(1);
                    }

                    if c_in.is_empty() {
                        let w_start = t_in - Duration::hours(5);
                        let w_end = parse_hm(&c_out)
                            .map(|hm| {
                                if hm.hour() < 18 {
                                    d_val.and_time(hm)
                                } else {
                                    (d_val - Duration::days(1)).and_time(hm)
                                }
                            })
                            .unwrap_or(t_out);
                        if let Some(m_in) = swipes
                            .iter()
                            .filter(|&&s| {
                                if !c_out.is_empty() {
                                    s >= w_start && s < w_end
                                } else {
                                    s >= w_start && s <= t_in + Duration::hours(5)
                                }
                            })
                            .min()
                        {
                            let cand = m_in.format("%H:%M").to_string();
                            if cand != c_out {
                                c_in = cand;
                                filled_hr = true;
                            }
                        }
                    }
                    if c_out.is_empty() {
                        let w_start = parse_hm(&c_in)
                            .map(|hm| {
                                if hm.hour() >= 12 {
                                    (d_val - Duration::days(1)).and_time(hm)
                                } else {
                                    d_val.and_time(hm)
                                }
                            })
                            .unwrap_or(t_in);
                        let w_end = t_out + Duration::hours(3);
                        if let Some(m_out) = swipes
                            .iter()
                            .filter(|&&s| {
                                if !c_in.is_empty() {
                                    s > w_start && s <= w_end
                                } else {
                                    s >= t_out - Duration::hours(4) && s <= w_end
                                }
                            })
                            .max()
                        {
                            let cand = m_out.format("%H:%M").to_string();
                            if cand != c_in {
                                c_out = cand;
                                filled_hr = true;
                            }
                        }
                    }
                } else {
                    let t_in = d_val.and_time(on_t);
                    let mut t_out = d_val.and_time(off_t);
                    if t_out <= t_in {
                        t_out += Duration::days(1);
                    }

                    if c_in.is_empty() {
                        let min_b = std::cmp::max(
                            t_in - Duration::hours(4),
                            d_val.and_hms_opt(0, 0, 0).unwrap(),
                        );
                        let max_b = parse_hm(&c_out)
                            .map(|hm| d_val.and_time(hm))
                            .unwrap_or(t_out - Duration::hours(1));
                        if let Some(m_in) = swipes.iter().filter(|&&s| s >= min_b && s < max_b).min() {
                            let cand = m_in.format("%H:%M").to_string();
                            if cand != c_out {
                                c_in = cand;
                                filled_hr = true;
                            }
                        }
                    }
                    if c_out.is_empty() {
                        let min_b = parse_hm(&c_in)
                            .map(|hm| d_val.and_time(hm))
                            .unwrap_or(t_in + Duration::hours(1));
                        let max_b = std::cmp::min(
                            t_out + Duration::hours(8),
                            d_val.and_hms_opt(23, 59, 59).unwrap(),
                        );
                        if let Some(m_out) = swipes.iter().filter(|&&s| s > min_b && s <= max_b).max() {
                            let cand = m_out.format("%H:%M").to_string();
                            if cand != c_in {
                                c_out = cand;
                                filled_hr = true;
                            }
                        }
                    }
                }
            }
        }

        let res = &mut self.data_result[idx];
        res.insert("On duty".into(), on_d.clone());
        res.insert("Off duty".into(), off_d.clone());
        res.insert("Clock In".into(), c_in.clone());
        res.insert("Clock Out".into(), c_out.clone());
        for k in ["Work Time", "Late", "Early", "OT Time"] {
            res.insert(
                k.into(),
                if force_recalc || filled_hr {
                    String::new()
                } else {
                    get_val(&orig, k).to_string()
                },
            );
        }
        for k in ["Kết Quả Công", "Nghỉ phép", "Tăng ca"] {
            res.insert(k.into(), String::new());
        }

        if c_in.is_empty() && c_out.is_empty() {
            res.insert("Nghỉ phép".into(), "PLP".into());
        } else if c_in.is_empty() {
            res.insert("Clock In".into(), "Bù vân tay".into());
            res.insert("Work Time".into(), "Bù vân tay".into());
        } else if c_out.is_empty() {
            res.insert("Clock Out".into(), "Bù vân tay".into());
            res.insert("Work Time".into(), "Bù vân tay".into());
        }

        let mut wt = get_val(res, "Work Time").to_string();
        let mut recalc = force_recalc || filled_hr || wt.is_empty() || wt == "00:00";
        if !recalc
            && is_hc
            && !c_in.is_empty()
            && !c_out.is_empty()
            && c_in != "Bù vân tay"
            && c_out != "Bù vân tay"
        {
            if let (Some(tin), Some(tout)) = (parse_hm(&c_in), parse_hm(&c_out)) {
                let d0 = NaiveDate::from_ymd_opt(2000, 1, 2).unwrap();
                let (dt_in, mut dt_out) = (d0.and_time(tin), d0.and_time(tout));
                if dt_out < dt_in {
                    dt_out += Duration::days(1);
                }
                if (dt_out - dt_in).num_seconds() <= (5.5 * 3600.0) as i64 {
                    recalc = true;
                }
            }
        }

        if recalc
            && !c_in.is_empty()
            && !c_out.is_empty()
            && c_in != "Bù vân tay"
            && c_out != "Bù vân tay"
        {
            if let (Some(tin), Some(tout), Some(on_hm), Some(off_hm)) = (
                parse_hm(&c_in),
                parse_hm(&c_out),
                parse_hm(&on_d),
                parse_hm(&off_d),
            ) {
                let d0 = NaiveDate::from_ymd_opt(2000, 1, 2).unwrap();
                let (dt_on, dt_off, dt_in, dt_out) = if is_ca3 {
                    let on_dt = (d0 - Duration::days(1)).and_time(on_hm);
                    let mut off_dt = d0.and_time(off_hm);
                    if off_dt <= on_dt {
                        off_dt += Duration::days(1);
                    }
                    let in_dt = if tin.hour() >= 12 {
                        (d0 - Duration::days(1)).and_time(tin)
                    } else {
                        d0.and_time(tin)
                    };
                    let mut out_dt = if tout.hour() < 18 {
                        d0.and_time(tout)
                    } else {
                        (d0 - Duration::days(1)).and_time(tout)
                    };
                    if out_dt <= in_dt {
                        out_dt += Duration::days(1);
                    }
                    (on_dt, off_dt, in_dt, out_dt)
                } else {
                    let on_dt = d0.and_time(on_hm);
                    let mut off_dt = d0.and_time(off_hm);
                    if off_dt <= on_dt {
                        off_dt += Duration::days(1);
                    }
                    let (mut in_dt, mut out_dt) = (d0.and_time(tin), d0.and_time(tout));
                    if in_dt < on_dt && (on_dt - in_dt).num_seconds() > 12 * 3600 {
                        in_dt += Duration::days(1);
                    } else if in_dt > on_dt && (in_dt - on_dt).num_seconds() > 12 * 3600 {
                        in_dt -= Duration::days(1);
                    }
                    if out_dt <= in_dt {
                        out_dt += Duration::days(1);
                    }
                    (on_dt, off_dt, in_dt, out_dt)
                };

                let (eff_in, eff_out) = (dt_in.max(dt_on), dt_out.min(dt_off));
                let mut total_sec = (eff_out - eff_in).num_seconds().max(0);
                let has_lunch = is_hc && (off_d.contains("17:00") || off_d.contains("17H00"));
                let (ls, le) = (
                    dt_on.date().and_hms_opt(11, 30, 0).unwrap(),
                    dt_on.date().and_hms_opt(13, 0, 0).unwrap(),
                );

                if has_lunch {
                    total_sec -= overlap_secs(eff_in, eff_out, ls, le);
                }
                total_sec = total_sec.min(8 * 3600);
                res.insert("Work Time".into(), fmt_hm_sec(total_sec, true));

                let mut late_s = (dt_in - dt_on).num_seconds().max(0);
                if has_lunch && late_s > 0 {
                    late_s -= overlap_secs(dt_on, dt_in, ls, le);
                }
                res.insert("Late".into(), fmt_hm_sec(late_s, false));

                let mut early_s = (dt_off - dt_out).num_seconds().max(0);
                if has_lunch && early_s > 0 {
                    early_s -= overlap_secs(dt_out, dt_off, ls, le);
                }
                res.insert("Early".into(), fmt_hm_sec(early_s, false));
            }
        }

        wt = get_val(res, "Work Time").to_string();
        let mut kq = String::new();
        if !wt.is_empty() && wt != "Bù vân tay" {
            let hm = if wt.contains(':') {
                let p: Vec<&str> = wt.split(':').collect();
                p.first()
                    .and_then(|h| h.parse::<i64>().ok())
                    .zip(p.get(1).and_then(|m| m.parse::<i64>().ok()))
            } else {
                wt.parse::<f64>()
                    .ok()
                    .map(|f| (f.trunc() as i64, ((f - f.trunc()) * 60.0).round() as i64))
            };
            if let Some((h, m)) = hm {
                let suf = if on_d == self.shift_settings.ca1_start && off_d == self.shift_settings.ca1_end {
                    "C1"
                } else if on_d == self.shift_settings.ca2_start && off_d == self.shift_settings.ca2_end {
                    "C2"
                } else if on_d == self.shift_settings.ca3_start && off_d == self.shift_settings.ca3_end {
                    "C3"
                } else if on_d == self.shift_settings.hc_start {
                    ""
                } else if is_ca1 {
                    "C1"
                } else if is_ca2 {
                    "C2"
                } else if is_ca3 {
                    "C3"
                } else {
                    ""
                };

                let num_w = h as f64 + if m >= 30 { 0.5 } else { 0.0 };
                kq = if suf.is_empty() {
                    format_num_clean(num_w)
                } else {
                    format!("{}{}{}", h, if m >= 30 { ".5" } else { "" }, suf)
                };
                if get_val(res, "Nghỉ phép") != "PLP" && num_w < 8.0 {
                    res.insert("Nghỉ phép".into(), format!("{}VR", format_num_clean(8.0 - num_w)));
                }
            }
        } else if wt == "Bù vân tay" {
            kq = "Bù vân tay".into();
        }
        res.insert("Kết Quả Công".into(), kq);
    }

    // =========================================================================
    // --- LOGIC TAB 2: XIN PHÉP, NGHỈ BÙ, PHÉP NĂM, CHỌN TAY & XÓA DÒNG ---
    // =========================================================================
    fn load_nghi_bu_map(&self, filter: Option<&HashSet<String>>) -> Result<BTreeMap<String, NghiBuSummary>, String> {
        let mut map = BTreeMap::new();
        if self.file_nb.bytes.is_empty() {
            return Ok(map);
        }
        let raw = read_bytes_to_dicts(&self.file_nb.bytes)?;
        if raw.is_empty() {
            return Ok(map);
        }
        if !raw[0].contains_key("Mã số nhân viên") || !raw[0].contains_key("Còn lại") {
            return Err("File số dư nghỉ bù thiếu cột 'Mã số nhân viên' hoặc cột 'Còn lại'!".into());
        }
        for row in &raw {
            let msnv = normalize_msnv(get_val(row, "Mã số nhân viên"));
            if msnv.is_empty() || filter.map_or(false, |s| !s.contains(&msnv)) {
                continue;
            }
            let e = map.entry(msnv.clone()).or_insert_with(|| NghiBuSummary {
                msnv,
                ho_ten: get_val(row, "Họ tên").into(),
                to_chuc: get_val(row, "Tổ chức").into(),
                vi_tri: get_val(row, "Vị trí").into(),
                don_vi: {
                    let d = get_val(row, "Đơn vị");
                    if d.is_empty() { "Giờ".into() } else { d.into() }
                },
                ..Default::default()
            });
            e.tong_so_du += parse_f64(get_val(row, "Tổng số dư"));
            e.tong_khau_tru += parse_f64(get_val(row, "Tổng số khấu trừ"));
            e.con_lai_ban_dau += parse_f64(get_val(row, "Còn lại"));
            e.con_lai_sau_tru = e.con_lai_ban_dau;
        }
        Ok(map)
    }

    fn load_phep_nam_map(&self, filter: Option<&HashSet<String>>) -> Result<BTreeMap<String, f64>, String> {
        let mut map = BTreeMap::new();
        if self.file_pn.bytes.is_empty() {
            return Ok(map);
        }
        let raw = read_bytes_to_dicts(&self.file_pn.bytes)?;
        if raw.is_empty() {
            return Ok(map);
        }

        let has_msnv = raw[0].contains_key("MSNV") || raw[0].contains_key("Mã số nhân viên");
        if !has_msnv || !raw[0].contains_key("Tổng số dư") {
            return Err("File phép năm thiếu cột 'MSNV' hoặc cột 'Tổng số dư'!".into());
        }

        for row in &raw {
            let raw_id = {
                let m = get_val(row, "MSNV");
                if m.is_empty() { get_val(row, "Mã số nhân viên") } else { m }
            };
            let msnv = normalize_msnv(raw_id);
            if msnv.is_empty() || filter.map_or(false, |s| !s.contains(&msnv)) {
                continue;
            }
            *map.entry(msnv).or_insert(0.0) += parse_f64(get_val(row, "Tổng số dư"));
        }
        Ok(map)
    }

    fn process_xin_phep(&mut self) -> Result<String, String> {
        if self.file_xp.bytes.is_empty() {
            return Err("Vui lòng chọn file dữ liệu đầu vào trước!".into());
        }
        let raw_data = read_bytes_to_dicts(&self.file_xp.bytes)?;
        if raw_data.is_empty() {
            return Err("File dữ liệu đầu vào trống!".into());
        }

        let req_cols = [
            "Bất thường chấm công",
            "Mã số nhân viên",
            "Ngày",
            "Thời gian quẹt thẻ lên ca",
            "Thời gian quẹt thẻ xuống ca",
            "Số ca làm việc trong ngày",
        ];
        let missing: Vec<&str> = req_cols
            .into_iter()
            .filter(|&c| !raw_data[0].contains_key(c))
            .collect();
        if !missing.is_empty() {
            return Err(format!("File đầu vào thiếu các cột:\n{}", missing.join(", ")));
        }

        let mut shift_dict = HashMap::new();
        let mut filtered = Vec::new();
        for row in &raw_data {
            let m = normalize_msnv(get_val(row, "Mã số nhân viên"));
            let s = clean_shift_spaces(&get_val(row, "Số ca làm việc trong ngày").to_uppercase());
            if !m.is_empty() {
                if let Some(d) = parse_date_only_flexible(get_val(row, "Ngày")) {
                    shift_dict.insert((m, d), s);
                }
            }
            if get_val(row, "Bất thường chấm công").eq_ignore_ascii_case("phải") {
                filtered.push(row.clone());
            }
        }

        self.base_xp_items.clear();
        for row in &filtered {
            let msnv = normalize_msnv(get_val(row, "Mã số nhân viên"));
            let raw_d = get_val(row, "Ngày");
            let s_name = clean_shift_spaces(&get_val(row, "Số ca làm việc trong ngày").to_uppercase());
            let t_in_opt = parse_datetime_flexible(get_val(row, "Thời gian quẹt thẻ lên ca"));
            let t_out_opt = parse_datetime_flexible(get_val(row, "Thời gian quẹt thẻ xuống ca"));
            let t_in_s = t_in_opt.map(|d| d.format("%Y-%m-%d %H:%M").to_string()).unwrap_or_default();
            let t_out_s = t_out_opt.map(|d| d.format("%Y-%m-%d %H:%M").to_string()).unwrap_or_default();
            let date_out = raw_d.replace(" 00:00:00", "").trim().to_string();
            let is_night = s_name.contains("ĐÊM")
                || s_name.contains("DEM")
                || s_name.contains("CA 3")
                || s_name.contains("CA3");
            let raw_d_opt = parse_date_only_flexible(raw_d);

            let start_d = if is_night {
                if let Some(tin) = t_in_opt {
                    Some(if tin.hour() >= 12 { tin.date() } else { tin.date() - Duration::days(1) })
                } else if let Some(tout) = t_out_opt {
                    Some(if tout.hour() < 14 { tout.date() - Duration::days(1) } else { tout.date() })
                } else {
                    raw_d_opt.map(|d| d - Duration::days(1))
                }
            } else {
                t_in_opt.map(|d| d.date()).or_else(|| t_out_opt.map(|d| d.date())).or(raw_d_opt)
            };

            let Some(base_d) = start_d else {
                self.base_xp_items.push(new_xp_item(
                    msnv, date_out, "LỖI DỮ LIỆU: KHÔNG ĐỌC ĐƯỢC NGÀY".into(),
                    String::new(), String::new(), "Kiểm tra lại dữ liệu gốc".into(),
                    false, false,
                ));
                continue;
            };
            let Some((s_start, s_end)) = self.get_shift_times_xp(base_d, &s_name) else {
                self.base_xp_items.push(new_xp_item(
                    msnv, date_out, "LỖI DỮ LIỆU: KHÔNG NHẬN DIỆN ĐƯỢC CA".into(),
                    String::new(), String::new(), format!("Ca bị lỗi: {}", s_name),
                    false, false,
                ));
                continue;
            };

            let is_hc = s_name.contains("HC")
                || s_name.contains("HÀNH CHÍNH")
                || s_name.contains("HANH CHINH")
                || s_name.contains("7H30-17");
            let has_lunch = is_hc && (s_end - s_start).num_minutes() > 480;
            let mut recs: Vec<(String, String, String, bool)> = match (t_in_opt, t_out_opt) {
                (None, None) => vec![(
                    "ACTR PHÉP NĂM".into(),
                    s_start.format("%Y-%m-%d %H:%M").to_string(),
                    s_end.format("%Y-%m-%d %H:%M").to_string(),
                    false,
                )],
                (None, Some(_)) => vec![(
                    "ACTR VIỆC RIÊNG".into(),
                    "Kiểm tra xếp ca/Bù vân tay".into(),
                    t_out_s,
                    false,
                )],
                (Some(_), None) => vec![(
                    "ACTR VIỆC RIÊNG".into(),
                    t_in_s,
                    "Kiểm tra xếp ca/Bù vân tay".into(),
                    false,
                )],
                (Some(tin), Some(tout)) => {
                    let late_end = |dt: NaiveDateTime| {
                        let (ls, le) = (
                            dt.date().and_hms_opt(11, 30, 0).unwrap(),
                            dt.date().and_hms_opt(13, 0, 0).unwrap(),
                        );
                        if has_lunch && dt >= ls && dt <= le { ls } else { ceil30(dt) }
                    };
                    let early_start = |dt: NaiveDateTime| {
                        let (ls, le) = (
                            dt.date().and_hms_opt(11, 30, 0).unwrap(),
                            dt.date().and_hms_opt(13, 0, 0).unwrap(),
                        );
                        if has_lunch && dt >= ls && dt <= le { le } else { floor30(dt) }
                    };
                    match (tin > s_start, tout < s_end) {
                        (true, true) => vec![
                            (
                                "ACTR VIỆC RIÊNG".into(),
                                s_start.format("%Y-%m-%d %H:%M").to_string(),
                                late_end(tin).format("%Y-%m-%d %H:%M").to_string(),
                                false,
                            ),
                            (
                                "ACTR VIỆC RIÊNG".into(),
                                early_start(tout).format("%Y-%m-%d %H:%M").to_string(),
                                s_end.format("%Y-%m-%d %H:%M").to_string(),
                                true,
                            ),
                        ],
                        (true, false) => vec![(
                            "ACTR VIỆC RIÊNG".into(),
                            s_start.format("%Y-%m-%d %H:%M").to_string(),
                            late_end(tin).format("%Y-%m-%d %H:%M").to_string(),
                            false,
                        )],
                        (false, true) => vec![(
                            "ACTR VIỆC RIÊNG".into(),
                            early_start(tout).format("%Y-%m-%d %H:%M").to_string(),
                            s_end.format("%Y-%m-%d %H:%M").to_string(),
                            false,
                        )],
                        (false, false) => vec![("ACTR VIỆC RIÊNG".into(), t_in_s, t_out_s, false)],
                    }
                }
            };

            let next_d = raw_d_opt.unwrap_or(base_d) + Duration::days(1);
            let next_s = shift_dict.get(&(msnv.clone(), next_d)).cloned().unwrap_or_default();
            if (t_in_opt.is_none() || t_out_opt.is_none())
                && is_hc
                && (next_s.contains("CA 3")
                    || next_s.contains("CA3")
                    || next_s.contains("ĐÊM")
                    || next_s.contains("DEM"))
            {
                recs = vec![(
                    "ACTR PHÉP NĂM".into(),
                    s_start.format("%Y-%m-%d %H:%M").to_string(),
                    s_end.format("%Y-%m-%d %H:%M").to_string(),
                    false,
                )];
            }

            for (mut loai, st, end, extra) in recs {
                if calc_leave_mins(&st, &end, has_lunch).map_or(false, |m| m >= 480) {
                    loai = "ACTR PHÉP NĂM".into();
                }
                self.base_xp_items.push(new_xp_item(
                    msnv.clone(),
                    date_out.clone(),
                    loai.clone(),
                    st,
                    end,
                    loai,
                    extra,
                    has_lunch,
                ));
            }
        }

        let xp_msnvs: HashSet<String> = self
            .base_xp_items
            .iter()
            .map(|r| r.msnv.clone())
            .filter(|s| !s.is_empty())
            .collect();

        self.base_nb_map.clear();
        if !self.file_nb.bytes.is_empty() {
            let mut nb_map = self.load_nghi_bu_map(Some(&xp_msnvs))?;
            for m in &xp_msnvs {
                nb_map.entry(m.clone()).or_insert_with(|| NghiBuSummary {
                    msnv: m.clone(),
                    don_vi: "Giờ".into(),
                    ..Default::default()
                });
            }
            self.base_nb_map = nb_map;
        }

        self.base_pn_map.clear();
        if !self.file_pn.bytes.is_empty() {
            self.base_pn_map = self.load_phep_nam_map(Some(&xp_msnvs))?;
        }

        let (full_nb, splittable) = self.rebuild_xp_display_rows();
        let extra = if !self.file_nb.bytes.is_empty() || !self.file_pn.bytes.is_empty() {
            format!(
                " (Đã chuyển {} dòng Nghỉ bù; {} trường hợp thiếu giờ có thể chọn Dropdown)",
                full_nb, splittable
            )
        } else {
            String::new()
        };
        Ok(format!(
            "Đã xử lý xong {} dòng xin phép!{}",
            self.data_xp_result.len(),
            extra
        ))
    }

    fn analyze_natural_nb_links(&self) -> Vec<NaturalRowInfo> {
        let mut info = vec![NaturalRowInfo::default(); self.base_xp_items.len()];
        if self.file_nb.bytes.is_empty() || self.base_nb_map.is_empty() {
            return info;
        }

        let mut rem_map: HashMap<String, i64> = self
            .base_nb_map
            .iter()
            .map(|(k, v)| (k.clone(), ((v.con_lai_ban_dau * 60.0).round() as i64 / 30) * 30))
            .collect();
        let has_pn = !self.file_pn.bytes.is_empty();
        let mut pn_rem_map: HashMap<String, i64> = self
            .base_pn_map
            .iter()
            .map(|(k, &days)| (k.clone(), ((days * 480.0).round() as i64 / 30) * 30))
            .collect();

        let mut upper_map: HashMap<String, Vec<(usize, i64)>> = HashMap::new();
        let mut done: HashSet<String> = HashSet::new();

        for (idx, it) in self.base_xp_items.iter().enumerate() {
            if it.is_fully_deleted() {
                continue;
            }
            let Some(req) = calc_leave_mins(&it.start, &it.end, it.has_lunch) else {
                continue;
            };
            let is_4_or_8 = req == 240 || req == 480;

            if it.partial_mode != 1 {
                if let Some(ref m_loai) = it.manual_part0 {
                    if m_loai == "ACTR NGHỈ BÙ" {
                        if let Some(rem) = rem_map.get_mut(&it.msnv) {
                            if *rem >= req {
                                *rem -= req;
                            }
                        }
                    } else if m_loai == "ACTR PHÉP NĂM" && has_pn && is_4_or_8 {
                        if let Some(p) = pn_rem_map.get_mut(&it.msnv) {
                            if *p >= req {
                                *p -= req;
                            }
                        }
                    }
                    continue;
                }
            }

            if it.user_choice == 3 && !it.skip_nb {
                continue;
            }

            if is_4_or_8 {
                let pn_avail = pn_rem_map.get(&it.msnv).copied().unwrap_or(0);
                if has_pn && pn_avail >= req && it.user_choice != 2 {
                    if let Some(p) = pn_rem_map.get_mut(&it.msnv) {
                        *p -= req;
                    }
                    continue;
                }
                if !has_pn && it.user_choice == 1 {
                    continue;
                }
            }

            let Some(rem) = rem_map.get_mut(&it.msnv) else {
                continue;
            };

            if *rem >= req {
                *rem -= req;
                if !done.contains(&it.msnv) {
                    upper_map.entry(it.msnv.clone()).or_default().push((idx, req));
                }
            } else if *rem > 0 && *rem < req && !done.contains(&it.msnv) {
                done.insert(it.msnv.clone());
                info[idx].is_partial = true;
                let nat_rem = *rem;
                if let Some(uppers) = upper_map.get(&it.msnv) {
                    if !uppers.is_empty() {
                        let swapped = uppers
                            .iter()
                            .rev()
                            .find(|&&(_, u_m)| nat_rem + u_m >= req)
                            .map(|&(u_i, _)| u_i);
                        let u_idx = swapped.unwrap_or(uppers.last().unwrap().0);
                        let can_swap = swapped.is_some();
                        info[idx].linked_upper_idx = Some(u_idx);
                        info[idx].can_swap_upper = can_swap;
                        info[u_idx].is_upper_nb = true;
                        info[u_idx].linked_partial_idx = Some(idx);
                        info[u_idx].can_swap_upper = can_swap;
                    }
                }
            }
        }
        info
    }

    fn rebuild_xp_display_rows(&mut self) -> (usize, usize) {
        self.data_xp_result.clear();
        let mut nb_map = self.base_nb_map.clone();
        let has_nb = !self.file_nb.bytes.is_empty() && !nb_map.is_empty();
        let has_pn = !self.file_pn.bytes.is_empty();
        let mut pn_rem_map: HashMap<String, i64> = self
            .base_pn_map
            .iter()
            .map(|(k, &days)| (k.clone(), ((days * 480.0).round() as i64 / 30) * 30))
            .collect();

        let nat_infos = self.analyze_natural_nb_links();
        let (mut full_cnt, mut split_cnt) = (0, 0);

        for (idx, it) in self.base_xp_items.iter().enumerate() {
            if it.is_fully_deleted() {
                continue;
            }
            let nat = &nat_infos[idx];
            let req_opt = calc_leave_mins(&it.start, &it.end, it.has_lunch);
            let is_four_or_eight = req_opt.map_or(false, |m| m == 240 || m == 480);

            let pn_avail_m = if has_pn {
                pn_rem_map.get(&it.msnv).copied().unwrap_or(0)
            } else {
                999_999
            };
            let can_use_pn = is_four_or_eight && req_opt.map_or(false, |req_m| pn_avail_m >= req_m);

            let fallback_loai = if has_pn && is_four_or_eight && !can_use_pn && it.orig_loai == "ACTR PHÉP NĂM" {
                "ACTR VIỆC RIÊNG".to_string()
            } else {
                it.orig_loai.clone()
            };

            let mut row = row_map!(
                "MSNV" => &it.msnv, "Ngày" => &it.ngay, "*Loại hình xin phép" => &fallback_loai,
                "*Thời gian bắt đầu xin phép" => &it.start, "*Thời gian kết thúc xin phép" => &it.end,
                "Lý do nghỉ phép" => &fallback_loai, "IsExtra" => if it.is_extra { "True" } else { "False" },
                "MenuKind" => "", "BaseIdx" => idx, "RowPart" => "0"
            );

            // Nếu người dùng đã chọn trực tiếp loại hình (Manual) trên dòng không tách
            if it.partial_mode != 1 {
                if let Some(ref m_loai) = it.manual_part0 {
                    if let Some(req_m) = req_opt {
                        if m_loai == "ACTR NGHỈ BÙ" && has_nb {
                            if let Some(sum) = nb_map.get_mut(&it.msnv) {
                                let avail_m = ((sum.con_lai_sau_tru * 60.0).round() as i64 / 30) * 30;
                                if avail_m >= req_m {
                                    let req_h = req_m as f64 / 60.0;
                                    sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round() / 100.0;
                                    sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                                    full_cnt += 1;
                                }
                            }
                        } else if m_loai == "ACTR PHÉP NĂM" && has_pn && can_use_pn {
                            if let Some(p) = pn_rem_map.get_mut(&it.msnv) {
                                *p -= req_m;
                            }
                        }
                    }
                    row.insert("*Loại hình xin phép".into(), m_loai.clone());
                    row.insert("Lý do nghỉ phép".into(), m_loai.clone());
                    row.insert("CurVal".into(), format!("MANUAL:{}", m_loai));
                    self.data_xp_result.push(row);
                    continue;
                }
            }

            // KHÔNG NẠP FILE NGHỈ BÙ (CHỈ DÙNG FILE PHÉP NĂM)
            if !has_nb {
                if let Some(req_m) = req_opt {
                    if is_four_or_eight && has_pn {
                        let req_h = req_m as f64 / 60.0;
                        let req_days = req_m as f64 / 480.0;
                        if can_use_pn {
                            let opts = vec![
                                ("1", format!("TH1: ACTR PHÉP NĂM ({}h = {} ngày)", format_num_clean(req_h), format_num_clean(req_days))),
                                ("3", "TH2: ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()),
                            ];
                            if it.user_choice == 3 {
                                set_menu_opts(&mut row, "CHOICE", "3", &opts);
                                row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                                row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                            } else {
                                if let Some(p) = pn_rem_map.get_mut(&it.msnv) { *p -= req_m; }
                                set_menu_opts(&mut row, "CHOICE", "1", &opts);
                                row.insert("*Loại hình xin phép".into(), "ACTR PHÉP NĂM".into());
                                row.insert("Lý do nghỉ phép".into(), "ACTR PHÉP NĂM".into());
                            }
                        } else {
                            row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                        }
                    }
                }
                self.data_xp_result.push(row);
                continue;
            }

            if let (Some(req_m), Some(sum)) = (req_opt, nb_map.get_mut(&it.msnv)) {
                let avail_m = ((sum.con_lai_sau_tru * 60.0).round() as i64 / 30) * 30;
                let req_h = req_m as f64 / 60.0;
                let req_days = req_m as f64 / 480.0;

                if nat.is_upper_nb {
                    row.insert("LinkedPartialIdx".into(), nat.linked_partial_idx.unwrap_or(idx).to_string());
                    row.insert("CanSwap".into(), if nat.can_swap_upper { "True" } else { "False" }.into());

                    let vr_lbl = if nat.can_swap_upper {
                        "ACTR VIỆC RIÊNG (Giữ mọi phép & Dòng dưới thành NGHỈ BÙ)".to_string()
                    } else {
                        "ACTR VIỆC RIÊNG (Giữ mọi phép & Nhường nghỉ bù xuống dưới)".to_string()
                    };

                    if can_use_pn && has_pn {
                        let opts = vec![
                            ("1", format!("TH1: ACTR PHÉP NĂM ({}h = {} ngày)", format_num_clean(req_h), format_num_clean(req_days))),
                            ("2", format!("TH2: ACTR NGHỈ BÙ ({}h)", format_num_clean(req_h))),
                            ("3", format!("TH3: {}", vr_lbl)),
                        ];
                        if it.user_choice == 3 || it.skip_nb {
                            set_menu_opts(&mut row, "UPPER", "3", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("IsExtra".into(), "True".into());
                        } else if it.user_choice == 2 && avail_m >= req_m {
                            sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round() / 100.0;
                            sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                            set_menu_opts(&mut row, "UPPER", "2", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR NGHỈ BÙ".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR NGHỈ BÙ".into());
                            full_cnt += 1;
                        } else {
                            if let Some(p) = pn_rem_map.get_mut(&it.msnv) { *p -= req_m; }
                            set_menu_opts(&mut row, "UPPER", "1", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR PHÉP NĂM".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR PHÉP NĂM".into());
                        }
                    } else {
                        let mut opts = vec![
                            ("2", format!("TH1 (Giữ nguyên): ACTR NGHỈ BÙ ({}h)", format_num_clean(req_h))),
                            ("3", format!("TH2: {}", vr_lbl)),
                        ];
                        if can_use_pn {
                            opts.push(("1", format!("TH3: ACTR PHÉP NĂM ({}h - Giữ nghỉ bù cho dòng sau)", format_num_clean(req_h))));
                        }

                        if it.user_choice == 1 && can_use_pn {
                            set_menu_opts(&mut row, "UPPER", "1", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR PHÉP NĂM".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR PHÉP NĂM".into());
                        } else if it.user_choice == 3 || it.skip_nb {
                            set_menu_opts(&mut row, "UPPER", "3", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("IsExtra".into(), "True".into());
                        } else if avail_m >= req_m {
                            sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round() / 100.0;
                            sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                            set_menu_opts(&mut row, "UPPER", "2", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR NGHỈ BÙ".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR NGHỈ BÙ".into());
                            full_cnt += 1;
                        }
                    }
                    self.data_xp_result.push(row);
                    continue;
                }

                if nat.is_partial {
                    split_cnt += 1;
                    let th1_loai = if is_four_or_eight {
                        if can_use_pn { "ACTR PHÉP NĂM".to_string() } else { "ACTR VIỆC RIÊNG".to_string() }
                    } else {
                        fallback_loai.clone()
                    };

                    if let Some(u_i) = nat.linked_upper_idx {
                        row.insert("LinkedUpperIdx".into(), u_i.to_string());
                    }

                    let th1_lbl = if is_four_or_eight && can_use_pn && has_pn {
                        format!("TH1 (Xin phép năm - 1 dòng): ACTR PHÉP NĂM ({}h = {} ngày)", format_num_clean(req_h), format_num_clean(req_days))
                    } else if th1_loai == "ACTR VIỆC RIÊNG" {
                        "TH1 (Giữ mọi phép - 1 dòng): ACTR VIỆC RIÊNG".to_string()
                    } else {
                        format!("TH1 (Giữ ban đầu - 1 dòng): {}", th1_loai)
                    };

                    let split_m = avail_m.min(req_m - 30).max(30);
                    let (nb_h, vr_h) = (split_m as f64 / 60.0, (req_m - split_m) as f64 / 60.0);
                    let th2_lbl = if it.partial_mode == 2 && avail_m >= req_m {
                        "TH2 (Chia 2 dòng): ACTR NGHỈ BÙ + ACTR VIỆC RIÊNG".to_string()
                    } else {
                        format!("TH2 (Chia 2 dòng): NGHỈ BÙ ({}h) + VIỆC RIÊNG ({}h)", format_num_clean(nb_h), format_num_clean(vr_h))
                    };

                    let mut opts = vec![("0", th1_lbl), ("1", th2_lbl)];
                    if nat.can_swap_upper {
                        opts.push(("2", format!("TH3: Đổi dòng trên thành VIỆC RIÊNG & Dòng này thành NGHỈ BÙ ({}h)", format_num_clean(req_h))));
                    }
                    if th1_loai == "ACTR PHÉP NĂM" {
                        opts.push(("3", "ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()));
                    }

                    if it.partial_mode == 2 && avail_m >= req_m {
                        sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round().max(0.0) / 100.0;
                        sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                        set_menu_opts(&mut row, "PARTIAL", "2", &opts);
                        row.insert("*Loại hình xin phép".into(), "ACTR NGHỈ BÙ".into());
                        row.insert("Lý do nghỉ phép".into(), "ACTR NGHỈ BÙ".into());
                        row.insert("IsExtra".into(), "True".into());
                        full_cnt += 1;
                        self.data_xp_result.push(row);
                        continue;
                    }

                    if it.partial_mode == 1 && avail_m > 0 {
                        if let Some(((r1_s, r1_e), (r2_s, r2_e))) = split_leave_interval(&it.start, &it.end, split_m, it.has_lunch) {
                            set_menu_opts(&mut row, "PARTIAL", "1", &opts);

                            if !it.deleted_part0 {
                                let loai1 = it.manual_part0.clone().unwrap_or_else(|| "ACTR NGHỈ BÙ".into());
                                if loai1 == "ACTR NGHỈ BÙ" {
                                    sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - nb_h) * 100.0).round().max(0.0) / 100.0;
                                    sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + nb_h) * 100.0).round() / 100.0;
                                }
                                let mut r1 = row.clone();
                                r1.insert("RowPart".into(), "0".into());
                                r1.insert("*Loại hình xin phép".into(), loai1.clone());
                                r1.insert("*Thời gian bắt đầu xin phép".into(), r1_s);
                                r1.insert("*Thời gian kết thúc xin phép".into(), r1_e);
                                r1.insert("Lý do nghỉ phép".into(), loai1.clone());
                                r1.insert("IsExtra".into(), "True".into());
                                if it.manual_part0.is_some() {
                                    r1.insert("CurVal".into(), format!("MANUAL:{}", loai1));
                                }
                                self.data_xp_result.push(r1);
                            }

                            if !it.deleted_part1 {
                                let loai2 = it.manual_part1.clone().unwrap_or_else(|| "ACTR VIỆC RIÊNG".into());
                                let mut r2 = row;
                                r2.insert("RowPart".into(), "1".into());
                                r2.insert("*Loại hình xin phép".into(), loai2.clone());
                                r2.insert("*Thời gian bắt đầu xin phép".into(), r2_s);
                                r2.insert("*Thời gian kết thúc xin phép".into(), r2_e);
                                r2.insert("Lý do nghỉ phép".into(), loai2.clone());
                                r2.insert("IsExtra".into(), "True".into());
                                if it.manual_part1.is_some() {
                                    r2.insert("CurVal".into(), format!("MANUAL:{}", loai2));
                                }
                                self.data_xp_result.push(r2);
                            }
                            continue;
                        }
                    }

                    if it.partial_mode == 3 && th1_loai == "ACTR PHÉP NĂM" {
                        set_menu_opts(&mut row, "PARTIAL", "3", &opts);
                        row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                        row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                        self.data_xp_result.push(row);
                        continue;
                    }

                    if th1_loai == "ACTR PHÉP NĂM" && has_pn && can_use_pn {
                        if let Some(p) = pn_rem_map.get_mut(&it.msnv) { *p -= req_m; }
                    }
                    set_menu_opts(&mut row, "PARTIAL", "0", &opts);
                    row.insert("*Loại hình xin phép".into(), th1_loai.clone());
                    row.insert("Lý do nghỉ phép".into(), th1_loai);
                    self.data_xp_result.push(row);
                    continue;
                }

                if avail_m >= req_m {
                    if is_four_or_eight && can_use_pn {
                        let (opts, eff_choice) = if has_pn {
                            let o = vec![
                                ("1", format!("TH1: ACTR PHÉP NĂM ({}h = {} ngày)", format_num_clean(req_h), format_num_clean(req_days))),
                                ("2", format!("TH2: ACTR NGHỈ BÙ ({}h)", format_num_clean(req_h))),
                                ("3", "TH3: ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()),
                            ];
                            (o, if it.user_choice == 0 { 1 } else { it.user_choice })
                        } else {
                            let o = vec![
                                ("2", format!("TH1: ACTR NGHỈ BÙ ({}h)", format_num_clean(req_h))),
                                ("1", format!("TH2: ACTR PHÉP NĂM ({}h - Giữ nghỉ bù lại)", format_num_clean(req_h))),
                                ("3", "TH3: ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()),
                            ];
                            (o, if it.user_choice == 0 { 2 } else { it.user_choice })
                        };

                        match eff_choice {
                            1 => {
                                if has_pn { if let Some(p) = pn_rem_map.get_mut(&it.msnv) { *p -= req_m; } }
                                set_menu_opts(&mut row, "CHOICE", "1", &opts);
                                row.insert("*Loại hình xin phép".into(), "ACTR PHÉP NĂM".into());
                                row.insert("Lý do nghỉ phép".into(), "ACTR PHÉP NĂM".into());
                            }
                            3 => {
                                set_menu_opts(&mut row, "CHOICE", "3", &opts);
                                row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                                row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                            }
                            _ => {
                                sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round() / 100.0;
                                sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                                set_menu_opts(&mut row, "CHOICE", "2", &opts);
                                row.insert("*Loại hình xin phép".into(), "ACTR NGHỈ BÙ".into());
                                row.insert("Lý do nghỉ phép".into(), "ACTR NGHỈ BÙ".into());
                                full_cnt += 1;
                            }
                        }
                    } else {
                        let opts = vec![
                            ("2", format!("TH1: ACTR NGHỈ BÙ ({}h)", format_num_clean(req_h))),
                            ("3", "TH2: ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()),
                        ];
                        if it.user_choice == 3 {
                            set_menu_opts(&mut row, "CHOICE", "3", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                        } else {
                            sum.con_lai_sau_tru = ((sum.con_lai_sau_tru - req_h) * 100.0).round() / 100.0;
                            sum.da_dung_xin_phep = ((sum.da_dung_xin_phep + req_h) * 100.0).round() / 100.0;
                            set_menu_opts(&mut row, "CHOICE", "2", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR NGHỈ BÙ".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR NGHỈ BÙ".into());
                            full_cnt += 1;
                        }
                    }
                } else if is_four_or_eight && has_pn {
                    if can_use_pn {
                        let opts = vec![
                            ("1", format!("TH1: ACTR PHÉP NĂM ({}h = {} ngày)", format_num_clean(req_h), format_num_clean(req_days))),
                            ("3", "TH2: ACTR VIỆC RIÊNG (Giữ nguyên mọi phép)".into()),
                        ];
                        if it.user_choice == 3 {
                            set_menu_opts(&mut row, "CHOICE", "3", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                        } else {
                            if let Some(p) = pn_rem_map.get_mut(&it.msnv) { *p -= req_m; }
                            set_menu_opts(&mut row, "CHOICE", "1", &opts);
                            row.insert("*Loại hình xin phép".into(), "ACTR PHÉP NĂM".into());
                            row.insert("Lý do nghỉ phép".into(), "ACTR PHÉP NĂM".into());
                        }
                    } else {
                        row.insert("*Loại hình xin phép".into(), "ACTR VIỆC RIÊNG".into());
                        row.insert("Lý do nghỉ phép".into(), "ACTR VIỆC RIÊNG".into());
                    }
                }
            }
            self.data_xp_result.push(row);
        }

        self.nghi_bu_summary = nb_map;
        (full_cnt, split_cnt)
    }

    fn get_shift_times_xp(&self, base_d: NaiveDate, shift_name: &str) -> Option<(NaiveDateTime, NaiveDateTime)> {
        let sn = clean_shift_spaces(&shift_name.to_uppercase());
        if sn.contains("7H30-15H30") || sn.contains("7:30-15:30") {
            return Some((base_d.and_hms_opt(7, 30, 0)?, base_d.and_hms_opt(15, 30, 0)?));
        }
        let pair = if sn.contains("HC") || sn.contains("HÀNH CHÍNH") || sn.contains("HANH CHINH") || sn.contains("7H30-17") {
            Some((&self.shift_settings.hc_start, &self.shift_settings.hc_end))
        } else if sn.contains("SÁNG") || sn.contains("SANG") || sn.contains("CA 1") || sn.contains("CA1") {
            Some((&self.shift_settings.ca1_start, &self.shift_settings.ca1_end))
        } else if sn.contains("TRƯA") || sn.contains("TRUA") || sn.contains("CHIỀU") || sn.contains("CHIEU") || sn.contains("CA 2") || sn.contains("CA2") {
            Some((&self.shift_settings.ca2_start, &self.shift_settings.ca2_end))
        } else if sn.contains("ĐÊM") || sn.contains("DEM") || sn.contains("CA 3") || sn.contains("CA3") {
            Some((&self.shift_settings.ca3_start, &self.shift_settings.ca3_end))
        } else {
            None
        };

        if let Some((s, e)) = pair {
            if let (Some(st), Some(et)) = (parse_hm(s), parse_hm(e)) {
                let s_dt = base_d.and_time(st);
                let mut e_dt = base_d.and_time(et);
                if e_dt <= s_dt { e_dt += Duration::days(1); }
                return Some((s_dt, e_dt));
            }
        }
        let c = Regex::new(r"(\d{1,2})[H:](\d{2})\s*-\s*(\d{1,2})[H:](\d{2})").unwrap().captures(&sn)?;
        let s_dt = base_d.and_hms_opt(c[1].parse().ok()?, c[2].parse().ok()?, 0)?;
        let mut e_dt = base_d.and_hms_opt(c[3].parse().ok()?, c[4].parse().ok()?, 0)?;
        if e_dt <= s_dt { e_dt += Duration::days(1); }
        Some((s_dt, e_dt))
    }

    // =========================================================================
    // --- XUẤT 4 BÁO CÁO EXCEL RA BỘ NHỚ (ANDROID 16 & DESKTOP) ---
    // =========================================================================
    fn build_excel_cc(&self) -> Result<ExportPayload, String> {
        if self.data_result.is_empty() {
            return Err("Không có dữ liệu chấm công để xuất!".into());
        }
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let (fmt_h, fmt_n, fmt_y) = excel_formats();

        for (c, &name) in CC_COLUMNS.iter().enumerate() {
            let _ = ws.write_with_format(0, c as u16, name, &fmt_h);
        }
        for (r, row) in self.data_result.iter().enumerate() {
            let r_idx = (r + 1) as u32;
            let is_bvt = ["Work Time", "Clock In", "Clock Out", "Kết Quả Công"]
                .iter()
                .any(|&k| get_val(row, k) == "Bù vân tay");
            for (c, &col) in CC_COLUMNS.iter().enumerate() {
                let val = get_val(row, col);
                let hl = is_bvt
                    || (!val.is_empty()
                        && ((col == "Work Time" && val != "08:00")
                            || matches!(col, "Late" | "Early" | "OT Time")
                            || (col == "Nghỉ phép" && val != "PLP")));
                let fmt = if hl { &fmt_y } else { &fmt_n };
                if col == "Kết Quả Công" {
                    if let Ok(n) = val.parse::<f64>() {
                        let _ = ws.write_number_with_format(r_idx, c as u16, n, fmt);
                        continue;
                    }
                }
                let _ = ws.write_string_with_format(r_idx, c as u16, val, fmt);
            }
        }
        ws.autofit();
        let buf = wb.save_to_buffer().map_err(|e| format!("Lỗi tạo Excel: {}", e))?;
        Ok(ExportPayload {
            filename: format!("Bao_Cao_Cham_Cong_{}.xlsx", Local::now().format("%d_%m_%Y")),
            data_base64: b64_encode(&buf),
        })
    }

    fn build_excel_fixed(&self) -> Result<ExportPayload, String> {
        if self.data_result.is_empty() {
            return Err("Bạn chưa xử lý dữ liệu chấm công!".into());
        }
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let _ = ws.set_name("Báo Cáo");
        let (fmt_b, fmt_n, fmt_y) = excel_formats();
        let fmt_cb = Format::new().set_bold().set_align(FormatAlign::Center).set_border(FormatBorder::Thin);

        let mut dates: Vec<NaiveDate> = self
            .data_result
            .iter()
            .filter_map(|r| NaiveDate::parse_from_str(get_val(r, "Date"), "%d/%m/%Y").ok())
            .collect();
        dates.sort();
        dates.dedup();

        let (mut emps, mut seen) = (Vec::new(), HashSet::new());
        for r in &self.data_result {
            let m = get_val(r, "MSNV").to_string();
            if !m.is_empty() && seen.insert(m.clone()) {
                emps.push((m, get_val(r, "Name").to_string()));
            }
        }

        let _ = ws.write_string_with_format(0, 0, "MSNV", &fmt_b);
        let _ = ws.write_string_with_format(0, 1, "Họ Tên", &fmt_b);
        let _ = ws.write_string_with_format(1, 0, "", &fmt_b);
        let _ = ws.write_string_with_format(1, 1, "", &fmt_b);

        for (i, dt) in dates.iter().enumerate() {
            let c_idx = (2 + i * 2) as u16;
            let _ = ws.merge_range(0, c_idx, 0, c_idx + 1, dt.format("%d").to_string().trim_start_matches('0'), &fmt_cb);
            let _ = ws.write_string_with_format(1, c_idx, "Kết Quả Công", &fmt_b);
            let _ = ws.write_string_with_format(1, c_idx + 1, "Nghỉ phép", &fmt_b);
        }

        for (i, (msnv, name)) in emps.iter().enumerate() {
            let r_idx = (i + 2) as u32;
            let _ = ws.write_string_with_format(r_idx, 0, msnv, &fmt_n);
            let _ = ws.write_string_with_format(r_idx, 1, name, &fmt_n);
            for (j, dt) in dates.iter().enumerate() {
                let c_idx = (2 + j * 2) as u16;
                let d_str = dt.format("%d/%m/%Y").to_string();
                if let Some(rec) = self.data_result.iter().find(|r| get_val(r, "MSNV") == msnv && get_val(r, "Date") == d_str) {
                    let (kq, np) = (get_val(rec, "Kết Quả Công"), get_val(rec, "Nghỉ phép"));
                    let fmt_kq = if kq == "Bù vân tay" { &fmt_y } else { &fmt_n };
                    let fmt_np = if !np.is_empty() && np != "PLP" { &fmt_y } else { &fmt_n };
                    if let Ok(num) = kq.parse::<f64>() {
                        let _ = ws.write_number_with_format(r_idx, c_idx, num, fmt_kq);
                    } else {
                        let _ = ws.write_string_with_format(r_idx, c_idx, kq, fmt_kq);
                    }
                    let _ = ws.write_string_with_format(r_idx, c_idx + 1, np, fmt_np);
                } else {
                    let _ = ws.write_string_with_format(r_idx, c_idx, "", &fmt_n);
                    let _ = ws.write_string_with_format(r_idx, c_idx + 1, "", &fmt_n);
                }
            }
        }
        ws.autofit();
        let buf = wb.save_to_buffer().map_err(|e| format!("Lỗi tạo Form Excel: {}", e))?;
        Ok(ExportPayload {
            filename: format!("Form_Bao_Cao_Tong_Hop_{}.xlsx", Local::now().format("%d_%m_%Y")),
            data_base64: b64_encode(&buf),
        })
    }

    fn build_excel_xp(&self) -> Result<ExportPayload, String> {
        if self.data_xp_result.is_empty() {
            return Err("Bạn chưa xử lý dữ liệu xin phép!".into());
        }
        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let _ = ws.set_name("Báo Cáo");
        let (fmt_h, fmt_n, fmt_y) = excel_formats();

        for (c, &name) in XP_COLUMNS.iter().enumerate() {
            let _ = ws.write_string_with_format(0, c as u16, name, &fmt_h);
        }
        for (r, row) in self.data_xp_result.iter().enumerate() {
            let r_idx = (r + 1) as u32;
            let is_ex = get_val(row, "IsExtra") == "True";
            for (c, &col) in XP_COLUMNS.iter().enumerate() {
                let val = get_val(row, col);
                let hl = val.contains("Kiểm tra xếp ca/Bù vân tay") || val.contains("LỖI DỮ LIỆU") || is_ex;
                let _ = ws.write_string_with_format(r_idx, c as u16, val, if hl { &fmt_y } else { &fmt_n });
            }
        }

        let last_r = self.data_xp_result.len() as u32;
        if last_r >= 1 {
            let dv = DataValidation::new()
                .allow_list_formula(Formula::new("='请假项目_XIN PHÉP'!$A$1:$A$9"));
            let _ = ws.add_data_validation(1, 2, last_r, 2, &dv);
            let _ = ws.add_data_validation(1, 5, last_r, 5, &dv);
        }
        ws.autofit();

        let ws_ref = wb.add_worksheet();
        let _ = ws_ref.set_name("请假项目_XIN PHÉP");
        for (i, &lt) in LEAVE_TYPES.iter().enumerate() {
            let _ = ws_ref.write_string(i as u32, 0, lt);
        }
        ws_ref.autofit();

        let buf = wb.save_to_buffer().map_err(|e| format!("Lỗi tạo file Xin Phép: {}", e))?;
        Ok(ExportPayload {
            filename: format!("Ket_Qua_Xin_Phep_{}.xlsx", Local::now().format("%d_%m_%Y")),
            data_base64: b64_encode(&buf),
        })
    }

    fn build_excel_nb(&mut self) -> Result<ExportPayload, String> {
        if self.file_nb.bytes.is_empty() {
            return Err("Vui lòng chọn File số dư nghỉ bù trước!".into());
        }
        let map = if !self.nghi_bu_summary.is_empty() {
            self.nghi_bu_summary.clone()
        } else {
            let filter = (!self.data_xp_result.is_empty()).then(|| {
                self.data_xp_result
                    .iter()
                    .map(|r| normalize_msnv(get_val(r, "MSNV")))
                    .filter(|s| !s.is_empty())
                    .collect::<HashSet<_>>()
            });
            self.load_nghi_bu_map(filter.as_ref())?
        };
        if map.is_empty() {
            return Err("Không tìm thấy dữ liệu số dư nghỉ bù tương ứng!".into());
        }

        let mut wb = Workbook::new();
        let ws = wb.add_worksheet();
        let _ = ws.set_name("Tổng Hợp Số Dư Nghỉ Bù");
        let (fmt_h, fmt_n, _) = excel_formats();
        let fmt_ok = Format::new()
            .set_bold()
            .set_border(FormatBorder::Thin)
            .set_background_color(Color::RGB(0xC6EFCE));

        let headers = [
            "Mã số nhân viên", "Họ tên", "Tổ chức", "Vị trí",
            "Tổng số dư", "Tổng số khấu trừ", "Tổng còn lại (Ban đầu)",
            "Đã dùng chuyển Nghỉ bù", "Số dư còn lại (Sau trừ)", "Đơn vị",
        ];
        for (c, &h) in headers.iter().enumerate() {
            let _ = ws.write_string_with_format(0, c as u16, h, &fmt_h);
        }
        for (i, it) in map.values().enumerate() {
            let r = (i + 1) as u32;
            for (c, s) in [(0, &it.msnv), (1, &it.ho_ten), (2, &it.to_chuc), (3, &it.vi_tri), (9, &it.don_vi)] {
                let _ = ws.write_string_with_format(r, c, s, &fmt_n);
            }
            for (c, v) in [(4, it.tong_so_du), (5, it.tong_khau_tru), (6, it.con_lai_ban_dau), (7, it.da_dung_xin_phep)] {
                let _ = ws.write_number_with_format(r, c, v, &fmt_n);
            }
            let _ = ws.write_number_with_format(r, 8, it.con_lai_sau_tru, if it.con_lai_sau_tru > 0.0 { &fmt_ok } else { &fmt_n });
        }
        ws.autofit();

        let buf = wb.save_to_buffer().map_err(|e| format!("Lỗi tạo file Nghỉ bù: {}", e))?;
        Ok(ExportPayload {
            filename: format!("Ket_Qua_So_Du_Nghi_Bu_{}.xlsx", Local::now().format("%d_%m_%Y")),
            data_base64: b64_encode(&buf),
        })
    }
}

// =========================================================================
// --- TAURI V2 COMMANDS (KHỚP TỪNG TÊN LỆNH & THAM SỐ VỚI INDEX.HTML) ---
// =========================================================================
#[tauri::command(rename_all = "snake_case")]
fn upload_file(
    state: State<'_, Mutex<TimekeepApp>>,
    target: String,
    name: String,
    data_base64: String,
) -> Result<CmdResponse, String> {
    let clean_b64 = data_base64
        .split_once("base64,")
        .map(|(_, b)| b)
        .unwrap_or(&data_base64);
    let bytes = b64_decode(clean_b64)?;
    if bytes.is_empty() {
        return Err("File được chọn bị rỗng (0 byte)!".into());
    }
    let mut app = state.lock().map_err(|e| e.to_string())?;
    let disp_name = name.clone();
    let slot = FileSlot { name, bytes };
    match target.as_str() {
        "f1" => app.file1 = slot,
        "f2" => app.file2 = slot,
        "fxp" => app.file_xp = slot,
        "fnb" => app.file_nb = slot,
        "fpn" => app.file_pn = slot,
        _ => return Err(format!("Target file không hợp lệ: {}", target)),
    }
    Ok(app.to_response(Some(format!("Đã nạp file: {}", disp_name)), false))
}

#[tauri::command(rename_all = "snake_case")]
fn clear_opt_file(state: State<'_, Mutex<TimekeepApp>>, kind: String) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    match kind.as_str() {
        "nb" | "fnb" => {
            app.file_nb = FileSlot::default();
            app.base_nb_map.clear();
            app.nghi_bu_summary.clear();
            app.rebuild_xp_display_rows();
        }
        "pn" | "fpn" => {
            app.file_pn = FileSlot::default();
            app.base_pn_map.clear();
            app.rebuild_xp_display_rows();
        }
        _ => {}
    }
    Ok(app.to_response(Some("Đã bỏ chọn file tùy chọn.".into()), false))
}

#[tauri::command(rename_all = "snake_case")]
fn save_settings(state: State<'_, Mutex<TimekeepApp>>, settings: ShiftConfig) -> Result<CmdResponse, String> {
    let pairs = [
        (&settings.hc_start, &settings.hc_end),
        (&settings.ca1_start, &settings.ca1_end),
        (&settings.ca2_start, &settings.ca2_end),
        (&settings.ca3_start, &settings.ca3_end),
    ];
    if !pairs.iter().all(|(s, e)| parse_hm(s).is_some() && parse_hm(e).is_some()) {
        return Err("Vui lòng nhập đúng định dạng 24h HH:MM (VD: 07:30, 15:00)!".into());
    }
    let mut app = state.lock().map_err(|e| e.to_string())?;
    app.shift_settings = ShiftConfig {
        hc_start: normalize_time_str(&settings.hc_start),
        hc_end: normalize_time_str(&settings.hc_end),
        ca1_start: normalize_time_str(&settings.ca1_start),
        ca1_end: normalize_time_str(&settings.ca1_end),
        ca2_start: normalize_time_str(&settings.ca2_start),
        ca2_end: normalize_time_str(&settings.ca2_end),
        ca3_start: normalize_time_str(&settings.ca3_start),
        ca3_end: normalize_time_str(&settings.ca3_end),
    };
    Ok(app.to_response(Some("Đã cập nhật cấu hình giờ ca!".into()), false))
}

#[tauri::command(rename_all = "snake_case")]
fn process_cc(state: State<'_, Mutex<TimekeepApp>>) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    match app.process_cham_cong() {
        Ok(msg) => Ok(app.to_response(Some(msg), false)),
        Err(err) => Ok(app.to_response(Some(err), true)),
    }
}

#[tauri::command(rename_all = "snake_case")]
fn change_tt(state: State<'_, Mutex<TimekeepApp>>, idx: usize, val: String) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    if idx < app.data_result.len() {
        app.data_result[idx].insert("Timetable".into(), val);
        app.process_single_row(idx, true);
        app.refresh_timetable_options();
    }
    Ok(app.to_response(None, false))
}

#[tauri::command(rename_all = "snake_case")]
fn process_xp(state: State<'_, Mutex<TimekeepApp>>) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    match app.process_xin_phep() {
        Ok(msg) => Ok(app.to_response(Some(msg), false)),
        Err(err) => Ok(app.to_response(Some(err), true)),
    }
}

#[tauri::command(rename_all = "snake_case")]
fn xp_set_manual(
    state: State<'_, Mutex<TimekeepApp>>,
    base_idx: usize,
    part: u8,
    loai: String,
) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    if let Some(it) = app.base_xp_items.get_mut(base_idx) {
        if part == 1 {
            it.manual_part1 = Some(loai);
        } else {
            it.manual_part0 = Some(loai);
        }
    }
    app.rebuild_xp_display_rows();
    Ok(app.to_response(None, false))
}

#[tauri::command(rename_all = "snake_case")]
fn xp_set_choice(
    state: State<'_, Mutex<TimekeepApp>>,
    base_idx: usize,
    choice: u8,
) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    if let Some(it) = app.base_xp_items.get_mut(base_idx) {
        it.user_choice = choice;
        it.skip_nb = false;
        it.manual_part0 = None;
        it.manual_part1 = None;
    }
    app.rebuild_xp_display_rows();
    Ok(app.to_response(None, false))
}

#[tauri::command(rename_all = "snake_case")]
fn xp_upper_choice(
    state: State<'_, Mutex<TimekeepApp>>,
    upper_idx: usize,
    partial_idx: usize,
    choice: u8,
    can_swap: bool,
) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    if let Some(u) = app.base_xp_items.get_mut(upper_idx) {
        u.user_choice = choice;
        u.skip_nb = choice == 3;
        u.manual_part0 = None;
    }
    if let Some(p) = app.base_xp_items.get_mut(partial_idx) {
        if choice == 3 && can_swap {
            p.partial_mode = 2;
            p.manual_part0 = None;
            p.manual_part1 = None;
        } else if choice != 3 && p.partial_mode == 2 {
            p.partial_mode = 0;
        }
    }
    app.rebuild_xp_display_rows();
    Ok(app.to_response(None, false))
}

#[tauri::command(rename_all = "snake_case")]
fn xp_partial_mode(
    state: State<'_, Mutex<TimekeepApp>>,
    partial_idx: usize,
    upper_idx: Option<usize>,
    mode: u8,
) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    if let Some(p) = app.base_xp_items.get_mut(partial_idx) {
        p.partial_mode = mode;
        p.manual_part0 = None;
        p.manual_part1 = None;
    }
    if let Some(u_idx) = upper_idx {
        if let Some(u) = app.base_xp_items.get_mut(u_idx) {
            u.skip_nb = mode == 2;
            if mode == 2 {
                u.user_choice = 3;
                u.manual_part0 = None;
            } else if u.user_choice == 3 {
                u.user_choice = 0;
            }
        }
    }
    app.rebuild_xp_display_rows();
    Ok(app.to_response(None, false))
}

#[tauri::command(rename_all = "snake_case")]
fn xp_delete_rows(
    state: State<'_, Mutex<TimekeepApp>>,
    targets: Vec<DeleteTarget>,
) -> Result<CmdResponse, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    let mut deleted_cnt = 0;
    for t in targets {
        if let Some(it) = app.base_xp_items.get_mut(t.base_idx) {
            if t.part == 1 {
                if !it.deleted_part1 {
                    it.deleted_part1 = true;
                    deleted_cnt += 1;
                }
            } else if !it.deleted_part0 {
                it.deleted_part0 = true;
                deleted_cnt += 1;
            }
        }
    }
    app.rebuild_xp_display_rows();
    Ok(app.to_response(Some(format!("Đã xóa {} dòng xin phép!", deleted_cnt)), false))
}

#[tauri::command(rename_all = "snake_case")]
fn export_excel(state: State<'_, Mutex<TimekeepApp>>, kind: String) -> Result<ExportPayload, String> {
    let mut app = state.lock().map_err(|e| e.to_string())?;
    match kind.as_str() {
        "cc" => app.build_excel_cc(),
        "fixed" => app.build_excel_fixed(),
        "xp" => app.build_excel_xp(),
        "nb" => app.build_excel_nb(),
        _ => Err("Loại xuất file không hợp lệ!".into()),
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .manage(Mutex::new(TimekeepApp::default()))
        .invoke_handler(tauri::generate_handler![
            upload_file,
            clear_opt_file,
            save_settings,
            process_cc,
            change_tt,
            process_xp,
            xp_set_manual,
            xp_set_choice,
            xp_upper_choice,
            xp_partial_mode,
            xp_delete_rows,
            export_excel
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

// =========================================================================
// --- HÀM HỖ TRỢ MENU, EXCEL, BASE64 & PARSER ĐỌC TRỰC TIẾP TỪ BYTES ---
// =========================================================================
#[allow(clippy::too_many_arguments)]
fn new_xp_item(
    msnv: String,
    ngay: String,
    orig_loai: String,
    start: String,
    end: String,
    lydo: String,
    is_extra: bool,
    has_lunch: bool,
) -> BaseXpItem {
    BaseXpItem {
        msnv,
        ngay,
        orig_loai,
        start,
        end,
        lydo,
        is_extra,
        has_lunch,
        skip_nb: false,
        user_choice: 0,
        partial_mode: 0,
        manual_part0: None,
        manual_part1: None,
        deleted_part0: false,
        deleted_part1: false,
    }
}

fn set_menu_opts(row: &mut HashMap<String, String>, kind: &str, cur_val: &str, opts: &[(&str, String)]) {
    row.insert("MenuKind".into(), kind.into());
    row.insert("CurVal".into(), cur_val.into());
    for (i, (v, lbl)) in opts.iter().enumerate() {
        row.insert(format!("Opt{}Val", i + 1), (*v).to_string());
        row.insert(format!("Opt{}Label", i + 1), lbl.clone());
    }
}

fn excel_formats() -> (Format, Format, Format) {
    (
        Format::new().set_bold().set_border(FormatBorder::Thin),
        Format::new().set_border(FormatBorder::Thin),
        Format::new().set_border(FormatBorder::Thin).set_background_color(Color::Yellow),
    )
}

fn b64_encode(input: &[u8]) -> String {
    const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity((input.len() + 2) / 3 * 4);
    for c in input.chunks(3) {
        let b0 = c[0] as u32;
        let b1 = if c.len() > 1 { c[1] as u32 } else { 0 };
        let b2 = if c.len() > 2 { c[2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        out.push(T[((n >> 18) & 63) as usize] as char);
        out.push(T[((n >> 12) & 63) as usize] as char);
        out.push(if c.len() > 1 { T[((n >> 6) & 63) as usize] as char } else { '=' });
        out.push(if c.len() > 2 { T[(n & 63) as usize] as char } else { '=' });
    }
    out
}

fn b64_decode(input: &str) -> Result<Vec<u8>, String> {
    let mut out = Vec::with_capacity(input.len() * 3 / 4);
    let (mut buf, mut bits): (u32, u32) = (0, 0);
    for &b in input.as_bytes() {
        let v = match b {
            b'A'..=b'Z' => b - b'A',
            b'a'..=b'z' => b - b'a' + 26,
            b'0'..=b'9' => b - b'0' + 52,
            b'+' => 62,
            b'/' => 63,
            b'=' | b'\r' | b'\n' | b' ' | b'\t' => continue,
            _ => return Err("Dữ liệu file Base64 không hợp lệ".into()),
        };
        buf = (buf << 6) | (v as u32);
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push(((buf >> bits) & 0xFF) as u8);
        }
    }
    Ok(out)
}

fn read_bytes_to_dicts(raw: &[u8]) -> Result<Vec<HashMap<String, String>>, String> {
    if raw.is_empty() {
        return Ok(Vec::new());
    }

    if raw.len() > 8 && raw[0] == 0x09 && matches!(raw[1], 0x08 | 0x00 | 0x02 | 0x04) {
        let rows = parse_raw_biff_stream(raw);
        if rows.len() > 1 {
            return Ok(convert_matrix_to_dicts(rows));
        }
    }

    if raw.len() > 4 && &raw[0..4] == b"PK\x03\x04" {
        if let Ok(mut wb) = Xlsx::new(Cursor::new(raw)) {
            let names = wb.sheet_names().to_vec();
            if !names.is_empty() {
                let sheet = names.iter().find(|s| *s == "Dữ liệu chấm công").unwrap_or(&names[0]).clone();
                if let Ok(range) = wb.worksheet_range(&sheet) {
                    let matrix = range.rows().map(|r| r.iter().map(format_calamine_cell).collect()).collect();
                    return Ok(convert_matrix_to_dicts(matrix));
                }
            }
        }
    }

    if raw.len() > 4 && &raw[0..4] == b"\xD0\xCF\x11\xE0" {
        if let Ok(mut wb) = Xls::new(Cursor::new(raw)) {
            let names = wb.sheet_names().to_vec();
            if !names.is_empty() {
                let sheet = names.iter().find(|s| *s == "Dữ liệu chấm công").unwrap_or(&names[0]).clone();
                if let Ok(range) = wb.worksheet_range(&sheet) {
                    let matrix = range.rows().map(|r| r.iter().map(format_calamine_cell).collect()).collect();
                    return Ok(convert_matrix_to_dicts(matrix));
                }
            }
        }
    }

    let decoded = decode_zkt_bytes(raw);
    if decoded.trim().is_empty() {
        return Err("Không thể giải mã nội dung file.".into());
    }

    if decoded.to_lowercase().contains("<table") {
        let (tr_re, td_re, tag_re) = (
            Regex::new(r"(?is)<tr[^>]*>(.*?)</tr>").unwrap(),
            Regex::new(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>").unwrap(),
            Regex::new(r"<[^>]+>").unwrap(),
        );
        let html_rows: Vec<Vec<String>> = tr_re
            .captures_iter(&decoded)
            .map(|tr| {
                td_re
                    .captures_iter(&tr[1])
                    .map(|td| tag_re.replace_all(&td[1], "").replace("&nbsp;", " ").trim().to_string())
                    .collect::<Vec<_>>()
            })
            .filter(|cols| cols.iter().any(|c| !c.is_empty()))
            .collect();
        if html_rows.len() > 1 {
            return Ok(convert_matrix_to_dicts(html_rows));
        }
    }

    let lines: Vec<&str> = decoded.lines().filter(|l| !l.trim().is_empty()).collect();
    let start_idx = lines
        .iter()
        .position(|l| ["AC-No.", "Mã số nhân viên", "MSNV", "Emp No.", "Bất thường chấm công"].iter().any(|k| l.contains(k)))
        .unwrap_or(0);
    let valid = &lines[start_idx..];

    let mut best: Option<(f64, Vec<Vec<String>>)> = None;
    for sep in ['\t', ',', ';'] {
        let parsed: Vec<Vec<String>> = valid
            .iter()
            .map(|l| parse_csv_line(l, sep))
            .filter(|r| r.iter().any(|c| !c.is_empty()))
            .collect();
        if !parsed.is_empty() {
            let avg = parsed.iter().map(|r| r.len()).sum::<usize>() as f64 / parsed.len() as f64;
            if best.as_ref().map_or(true, |(m, _)| avg > *m) {
                best = Some((avg, parsed));
            }
        }
    }
    if let Some((avg, m)) = best {
        if avg >= 2.0 {
            return Ok(convert_matrix_to_dicts(m));
        }
    }
    Err("Định dạng file không hợp lệ hoặc không tìm thấy bảng dữ liệu.".into())
}

fn parse_raw_biff_stream(data: &[u8]) -> Vec<Vec<String>> {
    let (mut cells, mut max_r, mut max_c, mut pos) = (BTreeMap::new(), 0, 0, 0);
    while pos + 4 <= data.len() {
        let (rec_id, len) = (
            u16::from_le_bytes([data[pos], data[pos + 1]]),
            u16::from_le_bytes([data[pos + 2], data[pos + 3]]) as usize,
        );
        let d = pos + 4;
        if d + len > data.len() || rec_id == 0x000A {
            break;
        }

        if (rec_id == 0x0004 || rec_id == 0x0204) && len >= 8 {
            let (r, c) = (
                u16::from_le_bytes([data[d], data[d + 1]]) as usize,
                u16::from_le_bytes([data[d + 2], data[d + 3]]) as usize,
            );
            let s_len = if rec_id == 0x0004 {
                data[d + 7] as usize
            } else {
                u16::from_le_bytes([data[d + 6], data[d + 7]]) as usize
            }
            .min(len - 8);
            let slice = &data[d + 8..d + 8 + s_len];
            let txt = std::str::from_utf8(slice)
                .map(|s| s.trim().to_string())
                .unwrap_or_else(|_| encoding_rs::WINDOWS_1258.decode(slice).0.trim().to_string());
            cells.insert((r, c), txt);
            max_r = max_r.max(r);
            max_c = max_c.max(c);
        } else if rec_id == 0x0002 && len >= 9 {
            let (r, c) = (
                u16::from_le_bytes([data[d], data[d + 1]]) as usize,
                u16::from_le_bytes([data[d + 2], data[d + 3]]) as usize,
            );
            cells.insert((r, c), u16::from_le_bytes([data[d + 7], data[d + 8]]).to_string());
            max_r = max_r.max(r);
            max_c = max_c.max(c);
        } else if rec_id == 0x0003 && len >= 15 {
            let (r, c) = (
                u16::from_le_bytes([data[d], data[d + 1]]) as usize,
                u16::from_le_bytes([data[d + 2], data[d + 3]]) as usize,
            );
            let mut b = [0u8; 8];
            b.copy_from_slice(&data[d + 7..d + 15]);
            cells.insert((r, c), format_num_clean(f64::from_le_bytes(b)));
            max_r = max_r.max(r);
            max_c = max_c.max(c);
        }
        pos += 4 + len;
    }
    if cells.is_empty() {
        Vec::new()
    } else {
        (0..=max_r)
            .map(|r| (0..=max_c).map(|c| cells.get(&(r, c)).cloned().unwrap_or_default()).collect())
            .collect()
    }
}

fn format_calamine_cell(cell: &Data) -> String {
    match cell {
        Data::Empty | Data::Error(_) => String::new(),
        Data::String(s) | Data::DateTimeIso(s) | Data::DurationIso(s) => s.trim().to_string(),
        Data::Int(i) => i.to_string(),
        Data::Float(f) => format_num_clean(*f),
        Data::Bool(b) => b.to_string(),
        Data::DateTime(dt) => {
            let s = dt.as_f64();
            let (days, secs) = (s.floor() as i64, ((s - s.floor()) * 86400.0).round() as i64);
            let ndt = NaiveDate::from_ymd_opt(1899, 12, 30)
                .unwrap()
                .and_hms_opt(0, 0, 0)
                .unwrap()
                + Duration::days(days)
                + Duration::seconds(secs);
            if days <= 1 {
                ndt.format("%H:%M").to_string()
            } else if secs == 0 {
                ndt.format("%Y-%m-%d").to_string()
            } else {
                ndt.format("%Y-%m-%d %H:%M").to_string()
            }
        }
    }
}

fn decode_zkt_bytes(raw: &[u8]) -> String {
    let s = if raw.len() >= 2 && raw[0] == 0xFF && raw[1] == 0xFE {
        encoding_rs::UTF_16LE.decode(&raw[2..]).0.into_owned()
    } else if raw.len() >= 2 && raw[0] == 0xFE && raw[1] == 0xFF {
        encoding_rs::UTF_16BE.decode(&raw[2..]).0.into_owned()
    } else if let Ok(utf8) = std::str::from_utf8(raw) {
        utf8.to_string()
    } else {
        encoding_rs::WINDOWS_1258.decode(raw).0.into_owned()
    };
    s.chars()
        .filter(|&c| c != '\0' && (!c.is_control() || matches!(c, '\r' | '\n' | '\t')))
        .collect()
}

fn parse_csv_line(line: &str, sep: char) -> Vec<String> {
    let (mut res, mut cur, mut in_q) = (Vec::new(), String::new(), false);
    for c in line.chars() {
        if c == '"' {
            in_q = !in_q;
        } else if c == sep && !in_q {
            res.push(cur.trim().to_string());
            cur.clear();
        } else {
            cur.push(c);
        }
    }
    res.push(cur.trim().to_string());
    res
}

fn convert_matrix_to_dicts(rows: Vec<Vec<String>>) -> Vec<HashMap<String, String>> {
    if rows.is_empty() {
        return Vec::new();
    }
    let h_idx = rows
        .iter()
        .take(10)
        .position(|r| {
            let j = r.join(" ");
            ["AC-No.", "Mã số nhân viên", "MSNV", "Emp No.", "Bất thường chấm công"]
                .iter()
                .any(|k| j.contains(k))
        })
        .unwrap_or(0);
    let headers: Vec<String> = rows[h_idx]
        .iter()
        .enumerate()
        .map(|(i, h)| if h.trim().is_empty() { format!("Col_{}", i) } else { h.trim().to_string() })
        .collect();
    rows.into_iter()
        .skip(h_idx + 1)
        .filter(|r| !r.iter().all(|c| c.trim().is_empty()))
        .map(|row| {
            headers
                .iter()
                .enumerate()
                .map(|(c, h)| {
                    let v = row.get(c).cloned().unwrap_or_default().trim().to_string();
                    (h.clone(), if matches!(v.as_str(), "nan" | "NaT" | "None") { String::new() } else { v })
                })
                .collect()
        })
        .collect()
}

fn overlap_secs(a: NaiveDateTime, b: NaiveDateTime, c: NaiveDateTime, d: NaiveDateTime) -> i64 {
    (b.min(d) - a.max(c)).num_seconds().max(0)
}

fn fmt_hm_sec(sec: i64, zero_show: bool) -> String {
    if sec > 0 || zero_show {
        format!("{:02}:{:02}", sec / 3600, (sec % 3600) / 60)
    } else {
        String::new()
    }
}

fn calc_leave_mins(start_str: &str, end_str: &str, has_lunch: bool) -> Option<i64> {
    let st = NaiveDateTime::parse_from_str(start_str.trim(), "%Y-%m-%d %H:%M").ok()?;
    let end = NaiveDateTime::parse_from_str(end_str.trim(), "%Y-%m-%d %H:%M").ok()?;
    if end <= st {
        return None;
    }
    let mut mins = (end - st).num_minutes();
    if has_lunch {
        let (ls, le) = (st.date().and_hms_opt(11, 30, 0)?, st.date().and_hms_opt(13, 0, 0)?);
        mins -= overlap_secs(st, end, ls, le) / 60;
    }
    (mins > 0).then_some(mins)
}

fn split_leave_interval(
    start_str: &str,
    end_str: &str,
    nb_mins: i64,
    has_lunch: bool,
) -> Option<((String, String), (String, String))> {
    let st = NaiveDateTime::parse_from_str(start_str.trim(), "%Y-%m-%d %H:%M").ok()?;
    let end = NaiveDateTime::parse_from_str(end_str.trim(), "%Y-%m-%d %H:%M").ok()?;
    if end <= st || nb_mins <= 0 {
        return None;
    }

    let fmt = "%Y-%m-%d %H:%M";
    let (ls, le) = (st.date().and_hms_opt(11, 30, 0)?, st.date().and_hms_opt(13, 0, 0)?);
    if !has_lunch || end <= ls || st >= le {
        let mid = st + Duration::minutes(nb_mins);
        return Some((
            (st.format(fmt).to_string(), mid.format(fmt).to_string()),
            (mid.format(fmt).to_string(), end.format(fmt).to_string()),
        ));
    }

    let morning_m = (ls - st).num_minutes();
    let (r1_e, r2_s) = if nb_mins < morning_m {
        let m = st + Duration::minutes(nb_mins);
        (m, m)
    } else if nb_mins == morning_m {
        (ls, le)
    } else {
        let m = le + Duration::minutes(nb_mins - morning_m);
        (m, m)
    };
    Some((
        (st.format(fmt).to_string(), r1_e.format(fmt).to_string()),
        (r2_s.format(fmt).to_string(), end.format(fmt).to_string()),
    ))
}

fn get_val<'a>(map: &'a HashMap<String, String>, key: &str) -> &'a str {
    map.get(key).map(|s| s.as_str()).unwrap_or("")
}
fn parse_f64(s: &str) -> f64 {
    s.trim().replace(',', ".").parse::<f64>().unwrap_or(0.0)
}
fn clean_shift_spaces(s: &str) -> String {
    Regex::new(r"\s+\(").unwrap().replace_all(s.trim(), " (").to_string()
}

fn normalize_msnv(raw: &str) -> String {
    let s = raw.trim().to_uppercase();
    if s.is_empty() {
        return String::new();
    }
    if let Some(st) = s.strip_prefix('S') {
        return format!("S{:0>6}", st.trim());
    }
    if s.len() >= 8 && (s.starts_with("99") || s.starts_with("88")) {
        return format!("S{}", &s[2..]);
    }
    format!("S{:0>6}", s)
}

fn normalize_time_str(t: &str) -> String {
    let up = t.trim().to_uppercase();
    if up.is_empty() {
        return String::new();
    }
    if let Some(c) = Regex::new(r"(\d{1,2})[H:](\d{2})").unwrap().captures(&up) {
        if let (Ok(mut h), Ok(m)) = (c[1].parse::<u32>(), c[2].parse::<u32>()) {
            if (up.contains("PM") || up.contains("CH")) && h < 12 {
                h += 12;
            }
            if (up.contains("AM") || up.contains("SA")) && h == 12 {
                h = 0;
            }
            return format!("{:02}:{:02}", h, m);
        }
    }
    t.trim().to_string()
}

fn parse_hm(s: &str) -> Option<NaiveTime> {
    NaiveTime::parse_from_str(&normalize_time_str(s), "%H:%M").ok()
}

fn min_dist_seconds(t1: NaiveTime, t2: NaiveTime) -> i64 {
    let d0 = NaiveDate::from_ymd_opt(2000, 1, 2).unwrap();
    let (dt1, dt2) = (d0.and_time(t1), d0.and_time(t2));
    (dt1 - dt2)
        .num_seconds()
        .abs()
        .min(((dt1 + Duration::days(1)) - dt2).num_seconds().abs())
        .min(((dt1 - Duration::days(1)) - dt2).num_seconds().abs())
}

fn classify_shift(tt: &str, on_d: &str, off_d: &str) -> (bool, bool, bool, bool) {
    let is_hc = ["HC", "HÀNH CHÍNH", "HANH CHINH", "7H30-17"].iter().any(|k| tt.contains(k));
    let is_ca1 = ["CA 1", "CA1", "SÁNG", "SANG"].iter().any(|k| tt.contains(k));
    let is_ca2 = ["CA 2", "CA2", "TRƯA", "TRUA", "CHIỀU", "CHIEU"].iter().any(|k| tt.contains(k));
    let is_ca3 = ["CA 3", "CA3", "ĐÊM", "DEM"].iter().any(|k| tt.contains(k))
        || parse_hm(on_d)
            .zip(parse_hm(off_d))
            .map_or(false, |(on_t, off_t)| off_t <= on_t);
    (is_hc, is_ca1, is_ca2, is_ca3)
}

fn ceil30(dt: NaiveDateTime) -> NaiveDateTime {
    match dt.minute() {
        0 => dt.with_second(0).unwrap(),
        1..=30 => dt.with_minute(30).unwrap().with_second(0).unwrap(),
        _ => (dt + Duration::hours(1)).with_minute(0).unwrap().with_second(0).unwrap(),
    }
}
fn floor30(dt: NaiveDateTime) -> NaiveDateTime {
    dt.with_minute(if dt.minute() < 30 { 0 } else { 30 })
        .unwrap()
        .with_second(0)
        .unwrap()
}

fn parse_date_only_flexible(input: &str) -> Option<NaiveDate> {
    let d_part = input.trim().split_whitespace().next()?;
    let sep = if d_part.contains('/') { '/' } else if d_part.contains('-') { '-' } else { return None };
    let p: Vec<&str> = d_part.split(sep).collect();
    if p.len() != 3 {
        return None;
    }
    if p[0].len() == 4 {
        NaiveDate::from_ymd_opt(p[0].parse().ok()?, p[1].parse().ok()?, p[2].parse().ok()?)
    } else {
        let (mut m, mut d, y): (u32, u32, i32) = (p[0].parse().ok()?, p[1].parse().ok()?, p[2].parse().ok()?);
        if m > 12 {
            std::mem::swap(&mut m, &mut d);
        }
        NaiveDate::from_ymd_opt(y, m, d)
    }
}

fn parse_datetime_flexible(input: &str) -> Option<NaiveDateTime> {
    let p: Vec<&str> = input.trim().split_whitespace().collect();
    (p.len() >= 2)
        .then(|| parse_date_only_flexible(p[0]).zip(parse_hm(&p[1..].join(" "))))?
        .map(|(d, t)| d.and_time(t))
}

fn convert_to_dmy(s: &str) -> String {
    parse_date_only_flexible(s)
        .map(|d| d.format("%d/%m/%Y").to_string())
        .unwrap_or_else(|| s.trim().to_string())
}

fn format_num_clean(v: f64) -> String {
    if (v - v.round()).abs() < 1e-6 {
        format!("{}", v.round() as i64)
    } else {
        format!("{}", (v * 100.0).round() / 100.0)
    }
}
