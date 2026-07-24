//! Record-walking parser for the XLSB Chart Sheet stream (MS-XLSB 2.1.7.7).
//!
//! The parser is strict about record payloads it fully understands and
//! tolerant about everything else: unknown record types are ignored, and
//! known begin/end record pairs that carry no modelled data (FRT wrappers,
//! header/footer blocks, ...) are skipped as balanced collections.

use crate::xlsb::chartsheet::model::*;
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::{XlsbRecord, XlsbRecordIter, record_types as rt, wide_str_with_len};
use crate::xlsb::worksheet::XlsbStrongProtection;

// `BrtCsProp` flags word (MS-XLSB 2.4.344).
const CS_PROP_PUBLISHED: u16 = 1 << 0;

// `BrtBeginCsView` flags word (MS-XLSB 2.4.38).
const CS_VIEW_SELECTED: u16 = 1 << 0;

/// Minimum and maximum `wScale` zoom percentages, and the "no zoom" value
/// (MS-XLSB 2.4.38).
const CS_VIEW_MIN_SCALE: u32 = 10;
const CS_VIEW_MAX_SCALE: u32 = 400;
const CS_VIEW_NO_ZOOM: u32 = 0;

// `BrtCsPageSetup` flags word (MS-XLSB 2.4.343).
const PAGE_SETUP_LANDSCAPE: u16 = 1 << 0;
const PAGE_SETUP_NO_COLOR: u16 = 1 << 2;
const PAGE_SETUP_NO_ORIENT: u16 = 1 << 3;
const PAGE_SETUP_USE_PAGE: u16 = 1 << 4;
const PAGE_SETUP_DRAFT: u16 = 1 << 5;

/// Maximum `iCopies` value (MS-XLSB 2.4.343).
const PAGE_SETUP_MAX_COPIES: u32 = 32_767;

/// Byte length of a `BrtColor` payload (MS-XLSB 2.4.337).
const BRT_COLOR_LEN: usize = 8;

// `BrtColor` first byte (MS-XLSB 2.4.337).
const COLOR_VALID_RGB: u8 = 1 << 0;
const COLOR_TYPE_SHIFT: u8 = 1;
const COLOR_TYPE_AUTOMATIC: u8 = 0;
const COLOR_TYPE_INDEXED: u8 = 1;
const COLOR_TYPE_RGB: u8 = 2;
const COLOR_TYPE_THEME: u8 = 3;

/// Maximum theme color index (MS-XLSB 2.5.77 `Icv` theme range).
const MAX_THEME_COLOR_INDEX: u8 = 0x0B;

/// Maximum `dwSpinCount` value (MS-XLSB 2.4.346).
const MAX_SPIN_COUNT: u32 = 10_000_000;

/// Parse a Chart Sheet part (`xl/chartsheets/sheet*.bin`) into a typed
/// [`XlsbChartSheet`].
///
/// `name` and `state` come from the sheet's `BrtBundleSh` record in the
/// workbook part. The stream must start with `BrtBeginSheet` and end with
/// `BrtEndSheet`. Records after `BrtEndSheet` are ignored. Unknown record
/// types anywhere in the stream are skipped without failing.
pub fn parse_chart_sheet_part(data: &[u8], name: String, state: u32) -> XlsbResult<XlsbChartSheet> {
    let state = match state {
        0 => XlsbChartSheetState::Visible,
        1 => XlsbChartSheetState::Hidden,
        2 => XlsbChartSheetState::VeryHidden,
        other => {
            return Err(XlsbError::Unrecognized {
                typ: "BrtBundleSh hsState".to_string(),
                val: other.to_string(),
            });
        },
    };
    let mut sheet = XlsbChartSheet {
        name,
        state,
        code_name: String::new(),
        published: false,
        tab_color: XlsbChartSheetColor::automatic(),
        views: Vec::new(),
        protection: None,
        strong_protection: None,
        page_setup: None,
        drawing_rel_id: None,
        legacy_drawing_rel_id: None,
        legacy_drawing_header_footer_rel_id: None,
    };
    let mut walker = RecordWalker::new(data);
    let first = walker.required("BrtBeginSheet")?;
    if first.header.record_type != rt::BEGIN_SHEET {
        return Err(XlsbError::UnexpectedRecord {
            expected: rt::BEGIN_SHEET,
            found: first.header.record_type,
        });
    }
    let mut seen_cs_prop = false;
    // Strong protection parsed from `BrtCsProtectionIso`, pending the classic
    // `BrtCsProtection` record that MS-XLSB 2.4.346 requires immediately after.
    let mut pending_iso_protection: Option<(XlsbStrongProtection, bool, bool)> = None;
    while let Some(record) = walker.next()? {
        if pending_iso_protection.is_some() && record.header.record_type != rt::CS_PROTECTION {
            return Err(malformed(
                "BrtCsProtectionIso",
                "not immediately followed by BrtCsProtection",
            ));
        }
        match record.header.record_type {
            rt::END_SHEET => {
                if pending_iso_protection.is_some() {
                    return Err(malformed(
                        "BrtCsProtectionIso",
                        "not immediately followed by BrtCsProtection",
                    ));
                }
                return Ok(sheet);
            },
            rt::CS_PROP => {
                if seen_cs_prop {
                    return Err(malformed("BrtCsProp", "duplicate record"));
                }
                seen_cs_prop = true;
                parse_cs_prop(&record.data, &mut sheet)?;
            },
            rt::CS_PAGE_SETUP => {
                if sheet.page_setup.is_some() {
                    return Err(malformed("BrtCsPageSetup", "duplicate record"));
                }
                sheet.page_setup = Some(parse_cs_page_setup(&record.data)?);
            },
            rt::CS_PROTECTION => {
                if sheet.protection.is_some() {
                    return Err(malformed("BrtCsProtection", "duplicate record"));
                }
                let protection = parse_cs_protection(&record.data)?;
                if let Some((strong, iso_locked, iso_objects)) = pending_iso_protection.take() {
                    // MS-XLSB 2.4.346: the ISO record must be immediately
                    // followed by a classic record with a zero verifier and
                    // matching flags.
                    if protection.password_verifier != 0
                        || protection.locked != iso_locked
                        || protection.objects != iso_objects
                    {
                        return Err(malformed(
                            "BrtCsProtectionIso",
                            "following BrtCsProtection record does not match",
                        ));
                    }
                    sheet.strong_protection = Some(strong);
                }
                sheet.protection = Some(protection);
            },
            rt::CS_PROTECTION_ISO => {
                if sheet.protection.is_some() || pending_iso_protection.is_some() {
                    return Err(malformed(
                        "BrtCsProtectionIso",
                        "duplicate protection record",
                    ));
                }
                pending_iso_protection = Some(parse_cs_protection_iso(&record.data)?);
            },
            rt::BEGIN_CS_VIEWS => parse_cs_views(&mut walker, &record.data, &mut sheet.views)?,
            rt::DRAWING => set_rel_id(&record.data, "BrtDrawing", &mut sheet.drawing_rel_id)?,
            rt::LEGACY_DRAWING => set_rel_id(
                &record.data,
                "BrtLegacyDrawing",
                &mut sheet.legacy_drawing_rel_id,
            )?,
            rt::LEGACY_DRAWING_HF => set_rel_id(
                &record.data,
                "BrtLegacyDrawingHF",
                &mut sheet.legacy_drawing_header_footer_rel_id,
            )?,
            other => walker.skip_unhandled(other, "Chart Sheet stream")?,
        }
    }
    Err(XlsbError::UnexpectedEndOfStream("BrtEndSheet".to_string()))
}

/// `BrtCsProp` payload (MS-XLSB 2.4.344).
fn parse_cs_prop(data: &[u8], sheet: &mut XlsbChartSheet) -> XlsbResult<()> {
    let mut cursor = PayloadCursor::new(data, "BrtCsProp");
    let flags = cursor.read_u16()?;
    sheet.published = flags & CS_PROP_PUBLISHED != 0;
    sheet.tab_color = parse_brt_color(cursor.read_bytes(BRT_COLOR_LEN)?)?;
    sheet.code_name = cursor.read_wide_string()?;
    cursor.finish()?;
    Ok(())
}

/// `BrtColor` structure (MS-XLSB 2.4.337).
fn parse_brt_color(data: &[u8]) -> XlsbResult<XlsbChartSheetColor> {
    debug_assert_eq!(data.len(), BRT_COLOR_LEN);
    let valid_rgb = data[0] & COLOR_VALID_RGB != 0;
    let color_type = match data[0] >> COLOR_TYPE_SHIFT {
        COLOR_TYPE_AUTOMATIC => XlsbChartSheetColorType::Automatic,
        COLOR_TYPE_INDEXED => XlsbChartSheetColorType::Indexed,
        COLOR_TYPE_RGB => XlsbChartSheetColorType::Rgb,
        COLOR_TYPE_THEME => XlsbChartSheetColorType::Theme,
        other => {
            return Err(malformed("BrtColor", format!("color type {other}")));
        },
    };
    if color_type == XlsbChartSheetColorType::Rgb && !valid_rgb {
        return Err(malformed("BrtColor", "direct color is not marked valid"));
    }
    if color_type == XlsbChartSheetColorType::Theme && data[1] > MAX_THEME_COLOR_INDEX {
        return Err(malformed(
            "BrtColor",
            format!("theme color index {}", data[1]),
        ));
    }
    Ok(XlsbChartSheetColor {
        valid_rgb,
        color_type,
        index: data[1],
        tint: i16::from_le_bytes([data[2], data[3]]),
        rgba: [data[4], data[5], data[6], data[7]],
    })
}

/// `BrtCsPageSetup` payload (MS-XLSB 2.4.343).
fn parse_cs_page_setup(data: &[u8]) -> XlsbResult<XlsbChartSheetPageSetup> {
    let mut cursor = PayloadCursor::new(data, "BrtCsPageSetup");
    let paper_size = cursor.read_u32()?;
    let horizontal_resolution = cursor.read_u32()?;
    let vertical_resolution = cursor.read_u32()?;
    let copies = cursor.read_u32()?;
    if copies == 0 || copies > PAGE_SETUP_MAX_COPIES {
        return Err(malformed("BrtCsPageSetup", format!("iCopies {copies}")));
    }
    let page_start = cursor.read_i16()?;
    let flags = cursor.read_u16()?;
    let printer_settings_rel_id = cursor.read_wide_string()?;
    if printer_settings_rel_id.is_empty() {
        return Err(malformed("BrtCsPageSetup", "empty szRelID"));
    }
    cursor.finish()?;
    Ok(XlsbChartSheetPageSetup {
        paper_size,
        horizontal_resolution,
        vertical_resolution,
        copies,
        page_start,
        landscape: flags & PAGE_SETUP_LANDSCAPE != 0,
        black_and_white: flags & PAGE_SETUP_NO_COLOR != 0,
        use_default_orientation: flags & PAGE_SETUP_NO_ORIENT != 0,
        use_page_start: flags & PAGE_SETUP_USE_PAGE != 0,
        draft: flags & PAGE_SETUP_DRAFT != 0,
        printer_settings_rel_id,
    })
}

/// `BrtCsProtection` payload (MS-XLSB 2.4.345).
fn parse_cs_protection(data: &[u8]) -> XlsbResult<XlsbChartSheetProtection> {
    let mut cursor = PayloadCursor::new(data, "BrtCsProtection");
    let protection = XlsbChartSheetProtection {
        password_verifier: cursor.read_u16()?,
        locked: cursor.read_bool32()? != 0,
        objects: cursor.read_bool32()? != 0,
    };
    cursor.finish()?;
    Ok(protection)
}

/// `BrtCsProtectionIso` payload (MS-XLSB 2.4.346).
///
/// Returns the strong protection data plus the `fLocked` and `fObjects`
/// flags, which the immediately following `BrtCsProtection` record must
/// repeat.
fn parse_cs_protection_iso(data: &[u8]) -> XlsbResult<(XlsbStrongProtection, bool, bool)> {
    let mut cursor = PayloadCursor::new(data, "BrtCsProtectionIso");
    let spin_count = cursor.read_u32()?;
    if spin_count > MAX_SPIN_COUNT {
        return Err(malformed(
            "BrtCsProtectionIso",
            format!("dwSpinCount {spin_count}"),
        ));
    }
    let locked = cursor.read_bool32()? != 0;
    let objects = cursor.read_bool32()? != 0;
    // IsoPasswordData (MS-XLSB 2.5.80).
    let hash = cursor.read_blob()?;
    if hash.is_empty() {
        return Err(malformed("BrtCsProtectionIso", "empty rgbHash"));
    }
    let salt = cursor.read_blob()?;
    let algorithm = cursor
        .read_nullable_wide_string()?
        .filter(|value| !value.is_empty())
        .ok_or_else(|| malformed("BrtCsProtectionIso", "null or empty szAlgName"))?;
    cursor.finish()?;
    Ok((
        XlsbStrongProtection {
            spin_count,
            hash,
            salt,
            algorithm,
        },
        locked,
        objects,
    ))
}

/// `BrtBeginCsViews` collection (MS-XLSB 2.4.38, 2.4.39).
fn parse_cs_views(
    walker: &mut RecordWalker<'_>,
    data: &[u8],
    views: &mut Vec<XlsbChartSheetView>,
) -> XlsbResult<()> {
    if !views.is_empty() {
        return Err(malformed("BrtBeginCsViews", "duplicate collection"));
    }
    PayloadCursor::new(data, "BrtBeginCsViews").finish()?;
    loop {
        let record = walker.required("BrtEndCsViews")?;
        match record.header.record_type {
            rt::END_CS_VIEWS => {
                if views.is_empty() {
                    return Err(malformed("BrtBeginCsViews", "collection without a view"));
                }
                PayloadCursor::new(&record.data, "BrtEndCsViews").finish()?;
                return Ok(());
            },
            rt::BEGIN_CS_VIEW => {
                views.push(parse_cs_view(&record.data)?);
                let end = walker.required("BrtEndCsView")?;
                if end.header.record_type != rt::END_CS_VIEW {
                    return Err(XlsbError::UnexpectedRecord {
                        expected: rt::END_CS_VIEW,
                        found: end.header.record_type,
                    });
                }
                PayloadCursor::new(&end.data, "BrtEndCsView").finish()?;
            },
            other => walker.skip_unhandled(other, "BrtBeginCsViews collection")?,
        }
    }
}

/// `BrtBeginCsView` payload (MS-XLSB 2.4.38).
fn parse_cs_view(data: &[u8]) -> XlsbResult<XlsbChartSheetView> {
    let mut cursor = PayloadCursor::new(data, "BrtBeginCsView");
    let flags = cursor.read_u16()?;
    let scale = cursor.read_u32()?;
    if scale != CS_VIEW_NO_ZOOM && !(CS_VIEW_MIN_SCALE..=CS_VIEW_MAX_SCALE).contains(&scale) {
        return Err(malformed("BrtBeginCsView", format!("wScale {scale}")));
    }
    let workbook_view_index = cursor.read_u32()?;
    cursor.finish()?;
    Ok(XlsbChartSheetView {
        selected: flags & CS_VIEW_SELECTED != 0,
        scale,
        workbook_view_index,
    })
}

/// Shared payload shape of `BrtDrawing` / `BrtLegacyDrawing` /
/// `BrtLegacyDrawingHF`: a single `RelID` (MS-XLSB 2.5.115).
fn set_rel_id(data: &[u8], context: &'static str, slot: &mut Option<String>) -> XlsbResult<()> {
    if slot.is_some() {
        return Err(malformed(context, "duplicate record"));
    }
    let mut cursor = PayloadCursor::new(data, context);
    let rel_id = cursor.read_wide_string()?;
    if rel_id.is_empty() {
        return Err(malformed(context, "empty relationship ID"));
    }
    cursor.finish()?;
    *slot = Some(rel_id);
    Ok(())
}

/// Wraps the shared record iterator with the collection helpers this parser needs.
struct RecordWalker<'a> {
    iter: XlsbRecordIter<&'a [u8]>,
}

impl<'a> RecordWalker<'a> {
    fn new(data: &'a [u8]) -> Self {
        RecordWalker {
            iter: XlsbRecordIter::new(data),
        }
    }

    fn next(&mut self) -> XlsbResult<Option<XlsbRecord>> {
        self.iter.next().transpose()
    }

    fn required(&mut self, context: &'static str) -> XlsbResult<XlsbRecord> {
        self.next()?
            .ok_or_else(|| XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Consume records up to and including `end_type`, tolerating nested
    /// collections of the same record pair.
    fn skip_collection(
        &mut self,
        begin_type: u16,
        end_type: u16,
        context: &'static str,
    ) -> XlsbResult<()> {
        let mut depth = 1u32;
        while let Some(record) = self.next()? {
            if record.header.record_type == begin_type {
                depth += 1;
            } else if record.header.record_type == end_type {
                depth -= 1;
                if depth == 0 {
                    return Ok(());
                }
            }
        }
        Err(XlsbError::UnexpectedEndOfStream(context.to_string()))
    }

    /// Skip a record the parser does not handle: a balanced collection when
    /// the type is a known begin record, a single record otherwise.
    fn skip_unhandled(&mut self, record_type: u16, context: &'static str) -> XlsbResult<()> {
        if let Some(end_type) = paired_end(record_type) {
            self.skip_collection(record_type, end_type, context)?;
        }
        Ok(())
    }
}

/// Map a known begin record type to its matching end record type.
///
/// Returns `None` for standalone records and unknown types, which the parser
/// then skips as single records.
fn paired_end(record_type: u16) -> Option<u16> {
    Some(match record_type {
        rt::BEGIN_CS_VIEWS => rt::END_CS_VIEWS,
        rt::BEGIN_CS_VIEW => rt::END_CS_VIEW,
        rt::BEGIN_HEADER_FOOTER => rt::END_HEADER_FOOTER,
        rt::FRT_BEGIN => rt::FRT_END,
        rt::AC_BEGIN => rt::AC_END,
        _ => return None,
    })
}

/// Bounds-checked cursor over one record payload.
struct PayloadCursor<'a> {
    data: &'a [u8],
    offset: usize,
    context: &'static str,
}

impl<'a> PayloadCursor<'a> {
    fn new(data: &'a [u8], context: &'static str) -> Self {
        PayloadCursor {
            data,
            offset: 0,
            context,
        }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.offset
    }

    fn guard(&self, needed: usize) -> XlsbResult<()> {
        if self.remaining() < needed {
            return Err(XlsbError::InvalidLength {
                expected: self.offset + needed,
                found: self.data.len(),
            });
        }
        Ok(())
    }

    fn read_bytes(&mut self, len: usize) -> XlsbResult<&'a [u8]> {
        self.guard(len)?;
        let bytes = &self.data[self.offset..self.offset + len];
        self.offset += len;
        Ok(bytes)
    }

    fn read_u16(&mut self) -> XlsbResult<u16> {
        self.guard(2)?;
        let value = u16::from_le_bytes([self.data[self.offset], self.data[self.offset + 1]]);
        self.offset += 2;
        Ok(value)
    }

    fn read_i16(&mut self) -> XlsbResult<i16> {
        Ok(self.read_u16()? as i16)
    }

    fn read_u32(&mut self) -> XlsbResult<u32> {
        self.guard(4)?;
        let value = u32::from_le_bytes(
            self.data[self.offset..self.offset + 4]
                .try_into()
                .expect("slice length is guarded"),
        );
        self.offset += 4;
        Ok(value)
    }

    /// Read a `Boolean` encoded as a 32-bit integer (MS-XLSB 2.5.98.3).
    fn read_bool32(&mut self) -> XlsbResult<u32> {
        let value = self.read_u32()?;
        if value > 1 {
            return Err(malformed(self.context, "non-Boolean 32-bit flag"));
        }
        Ok(value)
    }

    /// Read an `XLWideString` (MS-XLSB 2.5.169).
    fn read_wide_string(&mut self) -> XlsbResult<String> {
        let (value, consumed) = wide_str_with_len(&self.data[self.offset..])?;
        self.offset += consumed;
        Ok(value)
    }

    /// Read an `XLNullableWideString` (MS-XLSB 2.5.167).
    fn read_nullable_wide_string(&mut self) -> XlsbResult<Option<String>> {
        self.guard(4)?;
        if u32::from_le_bytes(
            self.data[self.offset..self.offset + 4]
                .try_into()
                .expect("slice length is guarded"),
        ) == u32::MAX
        {
            self.offset += 4;
            return Ok(None);
        }
        self.read_wide_string().map(Some)
    }

    /// Read an `LPByteBuf` (MS-XLSB 2.5.91).
    fn read_blob(&mut self) -> XlsbResult<Vec<u8>> {
        let len = usize::try_from(self.read_u32()?)
            .map_err(|_| malformed(self.context, "byte blob length overflow"))?;
        self.guard(len)?;
        let blob = self.data[self.offset..self.offset + len].to_vec();
        self.offset += len;
        Ok(blob)
    }

    /// Reject payloads with unparsed trailing bytes.
    fn finish(&self) -> XlsbResult<()> {
        if self.remaining() != 0 {
            return Err(malformed(
                self.context,
                format!("{} trailing bytes", self.remaining()),
            ));
        }
        Ok(())
    }
}

fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}
