//! Record encoders of the worksheet-feature editor (record 0327). Every
//! encoder returns complete records (headers included) and refuses values
//! outside their BIFF8 bounds instead of truncating them.

use super::Area;
use super::stream::{
    self, AUTO_FILTER_INFO, CF, CF12, COND_FMT, COND_FMT12, DV, DVAL, HLINK, HLINK_TOOLTIP, NOTE,
    record,
};
use crate::{Error, Result};

const HLINK_CLSID: [u8; 16] = [
    0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];
const URL_MONIKER_CLSID: [u8; 16] = [
    0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b,
];

fn invalid(message: &'static str) -> Error {
    Error::InvalidData(message.into())
}

fn ref8(out: &mut Vec<u8>, area: Area) {
    for value in [
        area.first_row,
        area.last_row,
        area.first_column,
        area.last_column,
    ] {
        out.extend_from_slice(&value.to_le_bytes());
    }
}

fn utf16z(out: &mut Vec<u8>, text: &str) {
    for unit in text.encode_utf16() {
        out.extend_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&[0, 0]);
}

/// A `HyperlinkString`: its length in characters (with the NUL) and the
/// NUL-terminated UTF-16 text.
fn hyperlink_string(out: &mut Vec<u8>, text: &str) -> Result<()> {
    if text.contains('\0') {
        return Err(invalid("a hyperlink string contains NUL"));
    }
    let units = u32::try_from(text.encode_utf16().count() + 1)
        .map_err(|_| invalid("a hyperlink string is too long"))?;
    out.extend_from_slice(&units.to_le_bytes());
    utf16z(out, text);
    Ok(())
}

/// One hyperlink to encode.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HyperlinkSpec {
    /// The cells the link covers.
    pub area: Area,
    /// An external target (URL).
    pub url: Option<String>,
    /// A location in the workbook (`Sheet1!A1` or a name).
    pub location: Option<String>,
    /// The display text.
    pub display: Option<String>,
    /// The screen tip.
    pub tooltip: Option<String>,
}

/// `HLink` and, with a tooltip, `HLinkTooltip`.
pub(super) fn hyperlink(spec: &HyperlinkSpec) -> Result<Vec<u8>> {
    if spec.url.is_none() && spec.location.is_none() {
        return Err(invalid("a hyperlink needs a URL or a location"));
    }
    if spec.area.last_column > 255 {
        return Err(invalid("a hyperlink column is outside the BIFF8 grid"));
    }
    let mut payload = Vec::with_capacity(128);
    ref8(&mut payload, spec.area);
    payload.extend_from_slice(&HLINK_CLSID);
    payload.extend_from_slice(&2u32.to_le_bytes());
    let mut flags = 0u32;
    if spec.url.is_some() {
        flags |= 0x03;
    }
    if spec.location.is_some() {
        flags |= 0x08;
    }
    if spec.display.is_some() {
        flags |= 0x14;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    if let Some(display) = &spec.display {
        hyperlink_string(&mut payload, display)?;
    }
    if let Some(url) = &spec.url {
        if url.contains('\0') {
            return Err(invalid("a hyperlink URL contains NUL"));
        }
        payload.extend_from_slice(&URL_MONIKER_CLSID);
        let bytes = u32::try_from((url.encode_utf16().count() + 1) * 2)
            .map_err(|_| invalid("a hyperlink URL is too long"))?;
        payload.extend_from_slice(&bytes.to_le_bytes());
        utf16z(&mut payload, url);
    }
    if let Some(location) = &spec.location {
        hyperlink_string(&mut payload, location)?;
    }
    let mut out = record(HLINK, &payload)?;
    if let Some(tooltip) = &spec.tooltip {
        let units = tooltip.encode_utf16().count() + 1;
        if units > 256 || tooltip.contains('\0') {
            return Err(invalid("a hyperlink tooltip is longer than 255 characters"));
        }
        let mut payload = Vec::with_capacity(10 + units * 2);
        payload.extend_from_slice(&HLINK_TOOLTIP.to_le_bytes());
        ref8(&mut payload, spec.area);
        utf16z(&mut payload, tooltip);
        out.extend(record(HLINK_TOOLTIP, &payload)?);
    }
    Ok(out)
}

/// A `NoteSh` record.
pub(super) fn note(
    row: u16,
    column: u16,
    visible: bool,
    object_id: u16,
    author: &str,
) -> Result<Vec<u8>> {
    let units: Vec<u16> = author.encode_utf16().collect();
    if units.len() > 54 || author.contains('\0') {
        return Err(invalid("a note author is longer than 54 characters"));
    }
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    let mut payload = Vec::with_capacity(12 + units.len() * 2);
    payload.extend_from_slice(&row.to_le_bytes());
    payload.extend_from_slice(&column.to_le_bytes());
    payload.extend_from_slice(&(if visible { 2u16 } else { 0 }).to_le_bytes());
    payload.extend_from_slice(&object_id.to_le_bytes());
    payload.extend_from_slice(&u16::try_from(units.len()).unwrap_or(0).to_le_bytes());
    payload.push(u8::from(!compressed));
    for unit in units {
        if compressed {
            payload.push(unit as u8);
        } else {
            payload.extend_from_slice(&unit.to_le_bytes());
        }
    }
    payload.push(0);
    record(NOTE, &payload)
}

/// A `TXO` record and its text and formatting-run `Continue` records, with
/// one run in font 0.
pub(super) fn txo(text: &str) -> Result<Vec<u8>> {
    let units: Vec<u16> = text.encode_utf16().collect();
    let count = u16::try_from(units.len())
        .map_err(|_| invalid("text is longer than 65,535 UTF-16 units"))?;
    let run_bytes: u16 = if units.is_empty() { 0 } else { 16 };
    let mut payload = Vec::with_capacity(18);
    payload.extend_from_slice(&0x0212u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&[0; 6]);
    payload.extend_from_slice(&count.to_le_bytes());
    payload.extend_from_slice(&run_bytes.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    let mut out = record(stream::TXO, &payload)?;
    let mut start = 0;
    while start < units.len() {
        let mut end = (start + 4_111).min(units.len());
        if end < units.len() && (0xd800..=0xdbff).contains(&units[end - 1]) {
            end -= 1;
        }
        let mut segment = Vec::with_capacity(1 + (end - start) * 2);
        segment.push(1);
        for unit in &units[start..end] {
            segment.extend_from_slice(&unit.to_le_bytes());
        }
        out.extend(record(stream::CONTINUE, &segment)?);
        start = end;
    }
    if !units.is_empty() {
        let mut runs = Vec::with_capacity(16);
        runs.extend_from_slice(&[0; 8]);
        runs.extend_from_slice(&count.to_le_bytes());
        runs.extend_from_slice(&[0; 6]);
        out.extend(record(stream::CONTINUE, &runs)?);
    }
    Ok(out)
}

/// `HorizontalPageBreaks` or `VerticalPageBreaks`.
pub(super) fn page_breaks(kind: u16, breaks: &[super::PageBreak]) -> Result<Vec<u8>> {
    let count = u16::try_from(breaks.len())
        .ok()
        .filter(|count| *count <= 1_026)
        .ok_or_else(|| invalid("a worksheet holds at most 1,026 page breaks per axis"))?;
    let mut payload = Vec::with_capacity(2 + breaks.len() * 6);
    payload.extend_from_slice(&count.to_le_bytes());
    for item in breaks {
        payload.extend_from_slice(&item.index.to_le_bytes());
        payload.extend_from_slice(&item.min.to_le_bytes());
        payload.extend_from_slice(&item.max.to_le_bytes());
    }
    record(kind, &payload)
}

/// `AutoFilterInfo`.
pub(super) fn auto_filter_info(columns: u16) -> Result<Vec<u8>> {
    record(AUTO_FILTER_INFO, &columns.to_le_bytes())
}

/// A `PtgArea3d` over one sheet's rectangle.
pub(super) fn area3d(xti: u16, area: Area) -> Vec<u8> {
    let mut out = Vec::with_capacity(11);
    out.push(0x3b);
    out.extend_from_slice(&xti.to_le_bytes());
    out.extend_from_slice(&area.first_row.to_le_bytes());
    out.extend_from_slice(&area.last_row.to_le_bytes());
    out.extend_from_slice(&area.first_column.to_le_bytes());
    out.extend_from_slice(&area.last_column.to_le_bytes());
    out
}

/// A hidden, built-in `_FilterDatabase` `Lbl` for the sheet at `sheet`.
pub(super) fn filter_database(sheet: usize, xti: u16, area: Area) -> Result<Vec<u8>> {
    let tokens = area3d(xti, area);
    let itab = u16::try_from(sheet + 1).map_err(|_| invalid("a sheet position exceeds 16 bits"))?;
    let mut payload = Vec::with_capacity(32);
    // fHidden | fBuiltin.
    payload.extend_from_slice(&0x0021u16.to_le_bytes());
    payload.push(0);
    payload.push(1);
    payload.extend_from_slice(&u16::try_from(tokens.len()).unwrap_or(0).to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&itab.to_le_bytes());
    payload.extend_from_slice(&[0; 4]);
    payload.push(0);
    payload.push(0x0d);
    payload.extend_from_slice(&tokens);
    record(stream::LBL, &payload)
}

// ---------------------------------------------------------------------------
// Conditional formats

/// The formatting a rule applies, as palette indexes.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct StyleSpec {
    /// The solid fill's palette index.
    pub fill: Option<u8>,
    /// The font color's palette index.
    pub font_color: Option<u16>,
    /// Bold.
    pub bold: Option<bool>,
    /// Italic.
    pub italic: Option<bool>,
    /// Strikethrough.
    pub strike: Option<bool>,
    /// `uls`.
    pub underline: Option<u8>,
    /// The font height in twips.
    pub size_twips: Option<u32>,
    /// The font name.
    pub font_name: Option<String>,
}

impl StyleSpec {
    fn has_font(&self) -> bool {
        self.font_color.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.strike.is_some()
            || self.underline.is_some()
            || self.size_twips.is_some()
            || self.font_name.is_some()
    }

    /// Whether the style sets nothing.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.fill.is_none() && !self.has_font()
    }
}

/// A `DXFN`.
pub(super) fn dxfn(style: &StyleSpec) -> Result<Vec<u8>> {
    // Every attribute ignored (ninch), reading order ignored.
    let mut options: u32 = 0x001f_ffff | 0x8000_0000;
    if style.has_font() {
        options |= 0x0400_0000;
    }
    if style.fill.is_some() {
        options &= !0x0007_0000;
        options |= 0x2000_0000;
    }
    let mut out = Vec::with_capacity(128);
    out.extend_from_slice(&options.to_le_bytes());
    out.extend_from_slice(&0x8002u16.to_le_bytes());
    if style.has_font() {
        let mut font = [0u8; 118];
        if let Some(name) = &style.font_name {
            let units: Vec<u16> = name.encode_utf16().collect();
            if units.is_empty() || units.len() > 31 {
                return Err(invalid(
                    "a conditional font name is longer than 31 characters",
                ));
            }
            font[0] = units.len() as u8;
            font[1] = 1;
            for (index, unit) in units.iter().enumerate() {
                font[2 + index * 2..4 + index * 2].copy_from_slice(&unit.to_le_bytes());
            }
        }
        let height = match style.size_twips {
            Some(twips) if (20..=8_191).contains(&twips) => twips,
            Some(_) => {
                return Err(invalid(
                    "a conditional font size is outside 1–409.55 points",
                ));
            },
            None => u32::MAX,
        };
        font[64..68].copy_from_slice(&height.to_le_bytes());
        let mut ts = 0u32;
        let mut ts_ninch = 0u32;
        match style.italic {
            Some(true) => ts |= 2,
            Some(false) => {},
            None => ts_ninch |= 2,
        }
        match style.strike {
            Some(true) => ts |= 0x80,
            Some(false) => {},
            None => ts_ninch |= 0x80,
        }
        font[68..72].copy_from_slice(&ts.to_le_bytes());
        let weight: u16 = if style.bold == Some(true) { 700 } else { 400 };
        font[72..74].copy_from_slice(&weight.to_le_bytes());
        font[76] = style.underline.unwrap_or(0);
        font[78] = 1;
        let color: u32 = style.font_color.map_or(u32::MAX, u32::from);
        font[80..84].copy_from_slice(&color.to_le_bytes());
        font[88..92].copy_from_slice(&ts_ninch.to_le_bytes());
        font[92..96].copy_from_slice(&1u32.to_le_bytes());
        font[96..100].copy_from_slice(&u32::from(style.underline.is_none()).to_le_bytes());
        font[100..104].copy_from_slice(&u32::from(style.bold.is_none()).to_le_bytes());
        font[104..108].copy_from_slice(&1u32.to_le_bytes());
        font[108..112].copy_from_slice(&0u32.to_le_bytes());
        font[112..116].copy_from_slice(&0x7fff_ffffu32.to_le_bytes());
        font[116..118].copy_from_slice(&1u16.to_le_bytes());
        out.extend_from_slice(&font);
    }
    if let Some(icv) = style.fill {
        // A solid pattern whose foreground and background are the fill.
        let icv = u16::from(icv & 0x7f);
        out.extend_from_slice(&(1u16 << 10).to_le_bytes());
        out.extend_from_slice(&(icv | (icv << 7)).to_le_bytes());
    }
    Ok(out)
}

fn sqref(out: &mut Vec<u8>, ranges: &[Area]) -> Result<()> {
    let count = u16::try_from(ranges.len())
        .ok()
        .filter(|count| (1..=1_026).contains(count))
        .ok_or_else(|| invalid("a conditional format needs 1–1,026 ranges"))?;
    out.extend_from_slice(&count.to_le_bytes());
    for range in ranges {
        ref8(out, *range);
    }
    Ok(())
}

fn bound(ranges: &[Area]) -> Area {
    let mut out = ranges[0];
    for range in &ranges[1..] {
        out.first_row = out.first_row.min(range.first_row);
        out.last_row = out.last_row.max(range.last_row);
        out.first_column = out.first_column.min(range.first_column);
        out.last_column = out.last_column.max(range.last_column);
    }
    out
}

/// The `CondFmt` structure body: rule count, identifier, bound and ranges.
fn cond_fmt_body(ranges: &[Area], rules: u16, identifier: u16) -> Result<Vec<u8>> {
    cond_fmt_header(ranges, rules, identifier, false)
}

/// The `CondFmt` structure body with its `fToughRecalc` flag.
pub(super) fn cond_fmt_header(
    ranges: &[Area],
    rules: u16,
    identifier: u16,
    tough: bool,
) -> Result<Vec<u8>> {
    if ranges.is_empty() || ranges.iter().any(|range| range.last_column > 255) {
        return Err(invalid(
            "a conditional format range is outside the BIFF8 grid",
        ));
    }
    if identifier > 0x7fff {
        return Err(invalid("a conditional-format identifier exceeds 15 bits"));
    }
    let mut out = Vec::with_capacity(14 + ranges.len() * 8);
    out.extend_from_slice(&rules.to_le_bytes());
    out.extend_from_slice(&((identifier << 1) | u16::from(tough)).to_le_bytes());
    ref8(&mut out, bound(ranges));
    sqref(&mut out, ranges)?;
    Ok(out)
}

/// A classic `CondFmt` with one `CF` rule (`ct` 1 cell value or 2 formula).
pub(super) fn classic_format(
    ranges: &[Area],
    identifier: u16,
    ct: u8,
    cp: u8,
    formulas: [&[u8]; 2],
    style: &StyleSpec,
) -> Result<Vec<u8>> {
    let mut out = record(COND_FMT, &cond_fmt_body(ranges, 1, identifier)?)?;
    out.extend(classic_rule(ct, cp, formulas, style)?);
    Ok(out)
}

/// One `CF` record.
pub(super) fn classic_rule(
    ct: u8,
    cp: u8,
    formulas: [&[u8]; 2],
    style: &StyleSpec,
) -> Result<Vec<u8>> {
    let mut payload = vec![ct, cp];
    for formula in formulas {
        payload.extend_from_slice(
            &u16::try_from(formula.len())
                .map_err(|_| invalid("a conditional-format formula is too long"))?
                .to_le_bytes(),
        );
    }
    payload.extend(dxfn(style)?);
    payload.extend_from_slice(formulas[0]);
    payload.extend_from_slice(formulas[1]);
    record(CF, &payload)
}

/// One Office 2007 rule.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Rule12 {
    pub(super) ct: u8,
    pub(super) cp: u8,
    pub(super) formulas: [Vec<u8>; 2],
    pub(super) style: StyleSpec,
    pub(super) stop_if_true: bool,
    pub(super) priority: u16,
    pub(super) template: u16,
    pub(super) params: [u8; 16],
    pub(super) payload: Vec<u8>,
}

/// A `CondFmt12` with one `CF12` rule.
pub(super) fn future_format(ranges: &[Area], identifier: u16, rule: &Rule12) -> Result<Vec<u8>> {
    let body = cond_fmt_body(ranges, 1, identifier)?;
    let mut payload = Vec::with_capacity(12 + body.len());
    payload.extend_from_slice(&COND_FMT12.to_le_bytes());
    payload.extend_from_slice(&1u16.to_le_bytes());
    ref8(&mut payload, bound(ranges));
    payload.extend(body);
    let mut out = record(COND_FMT12, &payload)?;
    out.extend(future_rule(rule)?);
    Ok(out)
}

/// One `CF12` record.
pub(super) fn future_rule(rule: &Rule12) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(96);
    payload.extend_from_slice(&CF12.to_le_bytes());
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&[0; 8]);
    payload.push(rule.ct);
    payload.push(rule.cp);
    for formula in &rule.formulas {
        payload.extend_from_slice(
            &u16::try_from(formula.len())
                .map_err(|_| invalid("a conditional-format formula is too long"))?
                .to_le_bytes(),
        );
    }
    if rule.style.is_empty() {
        payload.extend_from_slice(&0u32.to_le_bytes());
        payload.extend_from_slice(&0u16.to_le_bytes());
    } else {
        let dxfn = dxfn(&rule.style)?;
        payload.extend_from_slice(&u32::try_from(dxfn.len()).unwrap_or(0).to_le_bytes());
        payload.extend(dxfn);
    }
    payload.extend_from_slice(&rule.formulas[0]);
    payload.extend_from_slice(&rule.formulas[1]);
    // fmlaActive: empty.
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.push(if rule.stop_if_true { 2 } else { 0 });
    payload.extend_from_slice(&rule.priority.to_le_bytes());
    payload.extend_from_slice(&rule.template.to_le_bytes());
    payload.push(16);
    payload.extend_from_slice(&rule.params);
    payload.extend_from_slice(&rule.payload);
    record(CF12, &payload)
}

/// A `CFVO` with a number or a formula.
pub(super) fn cfvo(out: &mut Vec<u8>, kind: u8, number: Option<f64>, formula: Option<&[u8]>) {
    out.push(kind);
    match formula {
        Some(tokens) => {
            out.extend_from_slice(&u16::try_from(tokens.len()).unwrap_or(0).to_le_bytes());
            out.extend_from_slice(tokens);
        },
        None => {
            out.extend_from_slice(&0u16.to_le_bytes());
            if !matches!(kind, 2 | 3) {
                out.extend_from_slice(&number.unwrap_or(0.0).to_le_bytes());
            }
        },
    }
}

/// A `CFColor` of an RGB value (`0xRRGGBB`).
pub(super) fn cf_color(out: &mut Vec<u8>, rgb: u32) {
    out.extend_from_slice(&2u32.to_le_bytes());
    let red = (rgb >> 16) & 0xff;
    let green = (rgb >> 8) & 0xff;
    let blue = rgb & 0xff;
    out.extend_from_slice(&(red | (green << 8) | (blue << 16) | 0xff00_0000).to_le_bytes());
    out.extend_from_slice(&0f64.to_le_bytes());
}

// ---------------------------------------------------------------------------
// Data validations

/// One data validation to encode.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ValidationSpec {
    /// The ranges it applies to.
    pub ranges: Vec<Area>,
    /// `valType`.
    pub kind: u8,
    /// `typOperator`.
    pub operator: u8,
    /// `errStyle`.
    pub error_style: u8,
    /// The first formula's tokens.
    pub formula1: Vec<u8>,
    /// The second formula's tokens.
    pub formula2: Vec<u8>,
    /// An explicit list (`fStrLookup`).
    pub list: bool,
    /// `fAllowBlank`.
    pub allow_blank: bool,
    /// `fSuppressCombo`.
    pub suppress_dropdown: bool,
    /// `fShowInputMsg`.
    pub show_input: bool,
    /// `fShowErrorMsg`.
    pub show_error: bool,
    /// The input title.
    pub prompt_title: Option<String>,
    /// The input message.
    pub prompt: Option<String>,
    /// The error title.
    pub error_title: Option<String>,
    /// The error message.
    pub error: Option<String>,
}

fn xl_unicode(out: &mut Vec<u8>, text: Option<&str>, limit: usize) -> Result<()> {
    let text = text.filter(|text| !text.is_empty()).unwrap_or("\0");
    let units: Vec<u16> = text.encode_utf16().collect();
    if units.len() > limit {
        return Err(invalid("a validation message is too long"));
    }
    out.extend_from_slice(&u16::try_from(units.len()).unwrap_or(0).to_le_bytes());
    let compressed = units.iter().all(|unit| *unit <= 0xff);
    out.push(u8::from(!compressed));
    for unit in units {
        if compressed {
            out.push(unit as u8);
        } else {
            out.extend_from_slice(&unit.to_le_bytes());
        }
    }
    Ok(())
}

/// A `Dv` record.
pub(super) fn validation(spec: &ValidationSpec) -> Result<Vec<u8>> {
    if spec.ranges.is_empty() || spec.ranges.iter().any(|range| range.last_column > 255) {
        return Err(invalid("a validation range is outside the BIFF8 grid"));
    }
    let mut flags = u32::from(spec.kind & 0x0f) | (u32::from(spec.error_style & 7) << 4);
    if spec.list {
        flags |= 0x80;
    }
    if spec.allow_blank {
        flags |= 0x100;
    }
    if spec.suppress_dropdown {
        flags |= 0x200;
    }
    if spec.show_input {
        flags |= 0x0004_0000;
    }
    if spec.show_error {
        flags |= 0x0008_0000;
    }
    flags |= u32::from(spec.operator & 0x0f) << 20;
    let mut payload = Vec::with_capacity(64);
    payload.extend_from_slice(&flags.to_le_bytes());
    xl_unicode(&mut payload, spec.prompt_title.as_deref(), 32)?;
    xl_unicode(&mut payload, spec.error_title.as_deref(), 32)?;
    xl_unicode(&mut payload, spec.prompt.as_deref(), 255)?;
    xl_unicode(&mut payload, spec.error.as_deref(), 225)?;
    for formula in [&spec.formula1, &spec.formula2] {
        payload.extend_from_slice(
            &u16::try_from(formula.len())
                .map_err(|_| invalid("a validation formula is too long"))?
                .to_le_bytes(),
        );
        payload.extend_from_slice(&0u16.to_le_bytes());
        payload.extend_from_slice(formula);
    }
    sqref(&mut payload, &spec.ranges)?;
    record(DV, &payload)
}

/// A `DVal` record with no drop-down object.
pub(super) fn dval(count: u32) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(18);
    payload.extend_from_slice(&0u16.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&u32::MAX.to_le_bytes());
    payload.extend_from_slice(&count.to_le_bytes());
    record(DVAL, &payload)
}
