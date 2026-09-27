#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "limit fixtures favor explicit inputs and panic-driven assertions"
)]

//! Change 0766: the fresh XLS writer's remaining length fields, counts and
//! indices. Every limit is tested at N - 1, N and N + 1, with a surrogate pair
//! straddling it where the field counts UTF-16 code units; every accepted
//! input is read back by litchi's reader, and every refusal happens when the
//! input is registered, before any output, leaving the writer writing the
//! bytes of a writer that never saw the refused call.

use std::io::Cursor;

use litchi_core::sheet::{Cell as _, CellValue};
use litchi_xls::autofilter::FilterValue;
use litchi_xls::writer::{
    AutoFilterConditionWrite, CellStyle, ConditionalFormat, ConditionalFormatType, ExtendedFormat,
    ExternalCacheRowOptions, ExternalSheetOptions, ExternalWorkbookOptions, Font, PivotCacheValue,
    PivotDataItemConfig, PivotFieldConfig, PivotItemConfig, PivotTableConfig, Writer,
};
use litchi_xls::{
    CachedValue, Error, ExtProp, PivotCacheItem, StyleCategory, StyleExt, XfExt, XfProperties,
    XfProperty,
};

type Workbook = litchi_xls::Workbook<Cursor<Vec<u8>>>;

fn written(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read(bytes: Vec<u8>) -> Workbook {
    Workbook::new(Cursor::new(bytes)).unwrap()
}

/// The field, length and limit of a [`Error::StringTooLong`] refusal.
fn too_long<T: std::fmt::Debug>(result: Result<T, Error>) -> (&'static str, usize, usize) {
    match result {
        Err(Error::StringTooLong {
            field,
            utf16_units,
            limit,
        }) => (field, utf16_units, limit),
        other => panic!("expected StringTooLong, got {other:?}"),
    }
}

fn too_many<T: std::fmt::Debug>(result: Result<T, Error>) -> (&'static str, usize) {
    match result {
        Err(Error::TooMany { collection, limit }) => (collection, limit),
        other => panic!("expected TooMany, got {other:?}"),
    }
}

fn record_too_long<T: std::fmt::Debug>(result: Result<T, Error>) -> (&'static str, usize, usize) {
    match result {
        Err(Error::RecordTooLong {
            record,
            bytes,
            limit,
        }) => (record, bytes, limit),
        other => panic!("expected RecordTooLong, got {other:?}"),
    }
}

fn invalid_data<T: std::fmt::Debug>(result: Result<T, Error>) -> String {
    match result {
        Err(Error::InvalidData(message)) => message,
        other => panic!("expected InvalidData, got {other:?}"),
    }
}

/// Strings of exactly `units` UTF-16 code units: ASCII, Latin-1, CJK, and
/// one ending in a surrogate pair.
fn strings_of(units: usize) -> Vec<String> {
    vec![
        "b".repeat(units),
        "é".repeat(units),
        "漢".repeat(units),
        format!("{}😀", "a".repeat(units - 2)),
    ]
}

/// A string of `limit + 1` units whose surrogate pair straddles the limit:
/// its high surrogate is unit `limit` and its low surrogate unit `limit + 1`.
fn straddling(limit: usize) -> String {
    format!("{}😀", "a".repeat(limit - 1))
}

fn one_sheet_writer(name: &str) -> (Writer, usize) {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet(name).unwrap();
    writer.write_string(sheet, 0, 0, "kept").unwrap();
    (writer, sheet)
}

fn cell_value(workbook: &Workbook, sheet: usize, row: u32, column: u32) -> Option<CellValue> {
    workbook
        .xls_worksheet(sheet)
        .unwrap()
        .get_cell(row, column)
        .map(|cell| cell.value().clone())
}

// --- Hyperlinks ------------------------------------------------------------

#[test]
fn an_internal_hyperlink_to_an_astral_sheet_name_keeps_the_link_and_the_cells() {
    // Until change 0766 the location length counted chars, so this link made
    // the reader drop the link and every cell of its worksheet.
    let (mut writer, sheet) = one_sheet_writer("R😀");
    writer
        .set_hyperlink(sheet, 0, 1, "internal:'R😀'!A1")
        .unwrap();
    let workbook = read(written(&mut writer));
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(worksheet.hyperlinks().len(), 1);
    assert_eq!(worksheet.hyperlinks()[0].location(), Some("'R😀'!A1"));
    assert_eq!(
        cell_value(&workbook, 0, 0, 0),
        Some(CellValue::String("kept".to_string()))
    );
}

fn hyperlink_round_trip(url: &str) -> (Option<String>, Option<String>) {
    let (mut writer, sheet) = one_sheet_writer("Links");
    writer.set_hyperlink(sheet, 0, 1, url).unwrap();
    let workbook = read(written(&mut writer));
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(worksheet.hyperlinks().len(), 1, "{url:?}");
    let link = &worksheet.hyperlinks()[0];
    (
        link.location().map(str::to_string),
        link.address().map(str::to_string),
    )
}

#[test]
fn internal_hyperlink_targets_up_to_4093_units_fit_one_record() {
    for units in [4092, 4093] {
        for target in strings_of(units) {
            let (location, _) = hyperlink_round_trip(&format!("internal:{target}"));
            assert_eq!(location.as_deref(), Some(target.as_str()));
        }
    }
    for target in strings_of(4094).into_iter().chain([straddling(4093)]) {
        let (mut writer, sheet) = one_sheet_writer("Links");
        writer
            .set_hyperlink(sheet, 0, 1, "internal:Links!B2")
            .unwrap();
        let before = written(&mut writer);
        assert_eq!(
            too_long(writer.set_hyperlink(sheet, 0, 1, &format!("internal:{target}"))),
            ("internal hyperlink target", 4094, 4093)
        );
        // The refusal kept the cell's previous hyperlink and changed nothing.
        assert_eq!(written(&mut writer), before);
    }
    // 32,767 units wrapped the record length to 36 bytes before change 0766.
    let (mut writer, sheet) = one_sheet_writer("Links");
    assert_eq!(
        too_long(writer.set_hyperlink(sheet, 0, 1, &format!("internal:{}", "a".repeat(32_767)))),
        ("internal hyperlink target", 32_767, 4093)
    );
}

#[test]
fn url_hyperlinks_up_to_4085_units_fit_one_record() {
    for units in [4084, 4085] {
        for tail in strings_of(units - 8) {
            let url = format!("https://{tail}");
            let (_, address) = hyperlink_round_trip(&url);
            assert_eq!(address.as_deref(), Some(url.as_str()));
        }
    }
    for url in [
        format!("https://{}", "a".repeat(4086 - 8)),
        format!("https://{}", straddling(4085 - 8)),
    ] {
        let (mut writer, sheet) = one_sheet_writer("Links");
        let before = written(&mut writer);
        assert_eq!(
            too_long(writer.set_hyperlink(sheet, 0, 1, &url)),
            ("hyperlink URL", 4086, 4085)
        );
        assert_eq!(written(&mut writer), before);
    }
}

#[test]
fn hyperlink_targets_with_nul_are_refused_and_empty_targets_remove_the_link() {
    let (mut writer, sheet) = one_sheet_writer("Links");
    let before = written(&mut writer);
    for url in ["internal:'Links'!A1\0x", "https://example.com/\0"] {
        invalid_data(writer.set_hyperlink(sheet, 0, 1, url));
        assert_eq!(written(&mut writer), before);
    }
    writer
        .set_hyperlink(sheet, 0, 1, " https://example.com ")
        .unwrap();
    let workbook = read(written(&mut writer));
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().hyperlinks()[0].address(),
        Some("https://example.com")
    );
    writer.set_hyperlink(sheet, 0, 1, "  ").unwrap();
    assert_eq!(written(&mut writer), before);
    // A bare `internal:` names no location, so it is empty too.
    writer
        .set_hyperlink(sheet, 0, 1, "internal:Links!A1")
        .unwrap();
    writer.set_hyperlink(sheet, 0, 1, " internal: ").unwrap();
    assert_eq!(written(&mut writer), before);
}

#[test]
fn a_url_that_spells_the_moniker_serialization_guid_reads_back() {
    // The URL moniker's optional tail starts with this GUID; a URL whose own
    // last characters spell it must still read back as a URL.
    let guid = [
        0x5879u16, 0xF481, 0x1D3B, 0x487F, 0x2CAF, 0x5D82, 0x85C4, 0x6327,
    ];
    let mut units = "https://x/".encode_utf16().collect::<Vec<_>>();
    units.extend_from_slice(&guid);
    units.extend("abc".encode_utf16());
    let url = String::from_utf16(&units).unwrap();
    let (_, address) = hyperlink_round_trip(&url);
    assert_eq!(address.as_deref(), Some(url.as_str()));
}

// --- AutoFilter strings -----------------------------------------------------

fn filter_round_trip(condition: AutoFilterConditionWrite) -> (Vec<u8>, FilterValue) {
    let (mut writer, sheet) = one_sheet_writer("Filter");
    writer.set_auto_filter(sheet, 0, 9, 0, 1).unwrap();
    writer
        .add_filter_condition(sheet, 1, false, condition, AutoFilterConditionWrite::None)
        .unwrap();
    let bytes = written(&mut writer);
    let workbook = read(bytes.clone());
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let column = &worksheet.autofilter().unwrap().columns[0];
    assert_eq!(column.column_index, 1);
    (bytes, column.condition1.value.clone())
}

fn string_condition(value: &str) -> AutoFilterConditionWrite {
    AutoFilterConditionWrite::String {
        operator: 0x02,
        value: value.to_string(),
    }
}

#[test]
fn autofilter_strings_of_1_through_255_units_read_back_whole() {
    for value in ["x", "abc", "café", "x😀y", "漢字"]
        .into_iter()
        .map(str::to_string)
        .chain(strings_of(254))
        .chain(strings_of(255))
    {
        let (_, read_back) = filter_round_trip(string_condition(&value));
        assert_eq!(read_back, FilterValue::String(value));
    }
    for value in strings_of(256).into_iter().chain([straddling(255)]) {
        let (mut writer, sheet) = one_sheet_writer("Filter");
        writer.set_auto_filter(sheet, 0, 9, 0, 1).unwrap();
        let before = written(&mut writer);
        assert_eq!(
            too_long(writer.add_filter_condition(
                sheet,
                0,
                false,
                string_condition(&value),
                AutoFilterConditionWrite::None,
            )),
            ("AutoFilter string", 256, 255)
        );
        assert_eq!(written(&mut writer), before);
    }
    let (mut writer, sheet) = one_sheet_writer("Filter");
    writer.set_auto_filter(sheet, 0, 9, 0, 1).unwrap();
    invalid_data(writer.add_filter_condition(
        sheet,
        0,
        false,
        string_condition(""),
        AutoFilterConditionWrite::None,
    ));
}

/// The AFDOperStr layout ([MS-XLS] 2.5.8): four unused bytes, then `cch`
/// and `fCompare` (0 with a `?` or `*` wildcard, 1 without).
#[test]
fn autofilter_string_operands_carry_cch_and_fcompare_where_excel_reads_them() {
    for (value, compare) in [("café", 1), ("ca*", 0), ("c?t", 0)] {
        let (bytes, _) = filter_round_trip(string_condition(value));
        let doper = autofilter_doper(&bytes);
        assert_eq!(&doper[..2], &[0x06, 0x02]);
        assert_eq!(&doper[2..6], &[0, 0, 0, 0]);
        assert_eq!(usize::from(doper[6]), value.encode_utf16().count());
        assert_eq!(doper[7], compare, "{value}");
    }
}

fn autofilter_doper(bytes: &[u8]) -> Vec<u8> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut offset = 0;
    while offset + 4 <= stream.len() {
        let kind = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let len = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        if kind == 0x009E {
            return stream[offset + 8..offset + 18].to_vec();
        }
        offset += 4 + len;
    }
    panic!("no AutoFilter record");
}

#[test]
fn autofilter_operators_and_numbers_biff8_cannot_store_are_refused() {
    let (mut writer, sheet) = one_sheet_writer("Filter");
    writer.set_auto_filter(sheet, 0, 9, 0, 1).unwrap();
    let before = written(&mut writer);
    for condition in [
        AutoFilterConditionWrite::Number {
            operator: 0,
            value: 1.0,
        },
        AutoFilterConditionWrite::Number {
            operator: 7,
            value: 1.0,
        },
        AutoFilterConditionWrite::Number {
            operator: 2,
            value: f64::NAN,
        },
        AutoFilterConditionWrite::Number {
            operator: 2,
            value: f64::INFINITY,
        },
        AutoFilterConditionWrite::Number {
            operator: 2,
            value: -0.0,
        },
        AutoFilterConditionWrite::Number {
            operator: 2,
            value: f64::MIN_POSITIVE / 2.0,
        },
        AutoFilterConditionWrite::Bool {
            operator: 9,
            value: true,
        },
        AutoFilterConditionWrite::MatchAll { operator: 0 },
    ] {
        invalid_data(writer.add_filter_condition(
            sheet,
            0,
            false,
            condition,
            AutoFilterConditionWrite::None,
        ));
        assert_eq!(written(&mut writer), before);
    }
    for (operator, value) in [(0x01, -2.5), (0x06, 1e300), (0x04, f64::MIN_POSITIVE)] {
        let (_, read_back) =
            filter_round_trip(AutoFilterConditionWrite::Number { operator, value });
        assert_eq!(read_back, FilterValue::Number(value));
    }
}

// --- Fonts ------------------------------------------------------------------

fn font_named(name: &str) -> Font {
    Font {
        name: name.to_string(),
        ..Font::default()
    }
}

#[test]
fn font_names_up_to_31_units_are_written_whole_and_longer_ones_are_refused() {
    for units in [30, 31] {
        for name in strings_of(units) {
            let (mut writer, sheet) = one_sheet_writer("Fonts");
            let style = writer
                .add_cell_style(CellStyle {
                    font: font_named(&name),
                    ..CellStyle::default()
                })
                .unwrap();
            writer
                .write_string_with_format(sheet, 1, 0, "styled", style)
                .unwrap();
            let workbook = read(written(&mut writer));
            // The first added font has FontIndex 5: FontIndex skips 4.
            let font = workbook.fonts().last().unwrap();
            assert_eq!(font.index(), 5);
            assert_eq!(font.name(), name);
            let cell = workbook.xls_worksheet(0).unwrap().get_cell(1, 0).unwrap();
            let xf = &workbook.formatting().extended_formats()[usize::from(cell.xf_index())];
            assert_eq!(xf.font_index(), 5);
        }
    }
    for name in strings_of(32).into_iter().chain([straddling(31)]) {
        let (mut writer, _) = one_sheet_writer("Fonts");
        let before = written(&mut writer);
        assert_eq!(
            too_long(writer.add_cell_style(CellStyle {
                font: font_named(&name),
                number_format: Some("0.0\"x\"".to_string()),
                ..CellStyle::default()
            })),
            ("font name", 32, 31)
        );
        // Neither the font, the number format nor the XF was registered.
        assert_eq!(written(&mut writer), before);
    }
}

#[test]
fn fonts_the_reader_refuses_are_refused_when_registered() {
    for font in [
        font_named(""),
        font_named("A\0B"),
        Font {
            height: 19,
            ..Font::default()
        },
        Font {
            height: 8192,
            ..Font::default()
        },
        Font {
            weight: 99,
            ..Font::default()
        },
        Font {
            weight: 1001,
            ..Font::default()
        },
        Font {
            underline: 3,
            ..Font::default()
        },
        Font {
            color_index: 0x0042,
            ..Font::default()
        },
    ] {
        let (mut writer, _) = one_sheet_writer("Fonts");
        let before = written(&mut writer);
        invalid_data(writer.add_cell_style(CellStyle {
            font,
            ..CellStyle::default()
        }));
        assert_eq!(written(&mut writer), before);
    }
    for font in [
        Font {
            height: 0,
            weight: 0,
            ..Font::default()
        },
        Font {
            height: 20,
            weight: 100,
            underline: 0x22,
            color_index: 0x0008,
            ..Font::default()
        },
        Font {
            height: 8191,
            weight: 1000,
            underline: 0x21,
            color_index: 0x7FFF,
            italic: true,
            ..Font::default()
        },
    ] {
        let (mut writer, sheet) = one_sheet_writer("Fonts");
        let style = writer
            .add_cell_style(CellStyle {
                font: font.clone(),
                ..CellStyle::default()
            })
            .unwrap();
        writer
            .write_number_with_format(sheet, 1, 0, 1.0, style)
            .unwrap();
        let workbook = read(written(&mut writer));
        let read_back = workbook.fonts().last().unwrap();
        assert_eq!(read_back.height_twips(), font.height);
        assert_eq!(read_back.weight(), font.weight);
        assert_eq!(read_back.color_index(), font.color_index);
        assert_eq!(read_back.is_italic(), font.italic);
    }
}

/// A style with its own font: heights 20 through 8191 twips are all valid,
/// and none of the heights used below is the default font's 200.
fn style_with_height(height: u16) -> CellStyle {
    CellStyle {
        font: Font {
            height,
            ..Font::default()
        },
        ..CellStyle::default()
    }
}

#[test]
fn a_workbook_holds_1022_distinct_fonts_and_the_next_one_is_refused() {
    let (mut writer, sheet) = one_sheet_writer("Fonts");
    // Four default fonts and 1,018 added ones.
    let mut last = 0;
    for ordinal in 0..1018 {
        last = writer
            .add_cell_style(style_with_height(1000 + ordinal))
            .unwrap();
    }
    writer
        .write_number_with_format(sheet, 1, 0, 1.0, last)
        .unwrap();
    let before = written(&mut writer);
    assert_eq!(
        too_many(writer.add_cell_style(style_with_height(1000 + 1018))),
        ("fonts", 1022)
    );
    assert_eq!(written(&mut writer), before);
    // A style whose font the table already holds still fits.
    writer.add_cell_style(style_with_height(1000)).unwrap();
    let workbook = read(before);
    assert_eq!(workbook.fonts().len(), 1022);
    assert_eq!(workbook.fonts().last().unwrap().index(), 1022);
    assert_eq!(workbook.fonts().last().unwrap().height_twips(), 1000 + 1017);
}

/// Styles share an equal font instead of adding one each, so the 1022-font
/// bound limits distinct fonts, not styles.
#[test]
fn styles_share_an_equal_font() {
    let (mut writer, sheet) = one_sheet_writer("Fonts");
    let bold = Font {
        name: "Georgia".to_string(),
        weight: 700,
        ..Font::default()
    };
    for ordinal in 0..3000u32 {
        let style = writer
            .add_cell_style(CellStyle {
                font: bold.clone(),
                number_format: Some(format!("0.{}", "0".repeat(ordinal as usize % 20 + 1))),
                text_wrap: ordinal % 2 == 0,
                ..CellStyle::default()
            })
            .unwrap();
        writer
            .write_number_with_format(sheet, 1 + ordinal, 0, 1.0, style)
            .unwrap();
    }
    // The default font's style resolves to font 0.
    let plain = writer.add_cell_style(CellStyle::default()).unwrap();
    writer
        .write_number_with_format(sheet, 0, 1, 1.0, plain)
        .unwrap();
    let workbook = read(written(&mut writer));
    assert_eq!(workbook.fonts().len(), 5);
    assert_eq!(workbook.fonts()[4].name(), "Georgia");
    let worksheet = workbook.xls_worksheet(0).unwrap();
    let xfs = workbook.formatting().extended_formats();
    for row in [1, 1500, 3000] {
        let cell = worksheet.get_cell(row, 0).unwrap();
        assert_eq!(xfs[usize::from(cell.xf_index())].font_index(), 5);
    }
    let cell = worksheet.get_cell(0, 1).unwrap();
    assert_eq!(xfs[usize::from(cell.xf_index())].font_index(), 0);
}

// --- Number formats, cell formats and XF indices ---------------------------

#[test]
fn a_workbook_holds_210_custom_number_formats_and_the_next_is_refused() {
    let (mut writer, _) = one_sheet_writer("Formats");
    for ordinal in 0..210 {
        writer
            .register_number_format(&format!("0.0\"u{ordinal}\""))
            .unwrap();
    }
    let before = written(&mut writer);
    assert_eq!(
        too_many(writer.register_number_format("0.0\"one more\"")),
        ("custom number formats", 210)
    );
    assert_eq!(
        too_many(writer.add_cell_style(CellStyle {
            number_format: Some("0.0\"one more\"".to_string()),
            ..CellStyle::default()
        })),
        ("custom number formats", 210)
    );
    // Registered and built-in formats still resolve at capacity.
    assert_eq!(writer.register_number_format("0.0\"u7\"").unwrap(), 171);
    assert_eq!(writer.register_number_format("0.00").unwrap(), 2);
    assert_eq!(written(&mut writer), before);
    let workbook = read(before);
    let formats = workbook.formatting().number_formats();
    assert_eq!(formats.len(), 218);
    assert_eq!(formats.last().unwrap().id(), 164 + 209);
    assert_eq!(formats.last().unwrap().code(), "0.0\"u209\"");
}

#[test]
fn cell_formats_naming_no_font_or_number_format_are_refused() {
    let (mut writer, _) = one_sheet_writer("Formats");
    let before = written(&mut writer);
    for format in [
        ExtendedFormat {
            font_index: 4,
            ..ExtendedFormat::default()
        },
        ExtendedFormat {
            font_index: 5,
            ..ExtendedFormat::default()
        },
        ExtendedFormat {
            format_index: 164,
            ..ExtendedFormat::default()
        },
        ExtendedFormat {
            format_index: 100,
            ..ExtendedFormat::default()
        },
    ] {
        invalid_data(writer.add_cell_format(format));
        assert_eq!(written(&mut writer), before);
    }
    // Identifiers through 81 are built in, as the reader resolves them.
    for builtin in [49, 50, 56, 81] {
        let (mut builtin_writer, sheet) = one_sheet_writer("Formats");
        let format = builtin_writer
            .add_cell_format(ExtendedFormat {
                format_index: builtin,
                ..ExtendedFormat::default()
            })
            .unwrap();
        builtin_writer
            .write_number_with_format(sheet, 1, 0, 45_000.0, format)
            .unwrap();
        let workbook = read(written(&mut builtin_writer));
        let cell = workbook.xls_worksheet(0).unwrap().get_cell(1, 0).unwrap();
        assert_eq!(
            workbook.formatting().extended_formats()[usize::from(cell.xf_index())]
                .number_format_id(),
            builtin
        );
    }
    invalid_data(writer.add_cell_format(ExtendedFormat {
        format_index: 82,
        ..ExtendedFormat::default()
    }));
    let custom = writer.register_number_format("0.0\"k\"").unwrap();
    let style = writer
        .add_cell_style(CellStyle {
            font: font_named("Times"),
            ..CellStyle::default()
        })
        .unwrap();
    assert_eq!(style, 1);
    let format = writer
        .add_cell_format(ExtendedFormat {
            font_index: 5,
            format_index: custom,
            ..ExtendedFormat::default()
        })
        .unwrap();
    assert_eq!(format, 2);
}

#[test]
fn the_xf_index_space_holds_65512_cell_formats() {
    let (mut writer, sheet) = one_sheet_writer("Formats");
    let mut last = 0;
    for _ in 0..65_512 {
        last = writer.add_cell_format(ExtendedFormat::default()).unwrap();
    }
    assert_eq!(last, 65_512);
    writer
        .write_number_with_format(sheet, 1, 0, 1.0, last)
        .unwrap();
    let before = written(&mut writer);
    assert_eq!(
        too_many(writer.add_cell_format(ExtendedFormat::default())),
        ("cell formats", 65_512)
    );
    assert_eq!(
        too_many(writer.add_cell_style(CellStyle::default())),
        ("cell formats", 65_512)
    );
    assert_eq!(written(&mut writer), before);
    let workbook = read(before);
    // 21 fixed XFs and 65,512 user cell XFs; the three pivot XFs stay free.
    assert_eq!(workbook.formatting().extended_formats().len(), 65_533);
    let cell = workbook.xls_worksheet(0).unwrap().get_cell(1, 0).unwrap();
    assert_eq!(cell.xf_index(), 65_532);
}

#[test]
fn an_xfcrc_caps_the_xf_table_at_4050_records() {
    let (mut writer, sheet) = one_sheet_writer("Formats");
    writer
        .set_xf_extensions(vec![XfExt::try_new(15, vec![ExtProp::Indent(2)]).unwrap()])
        .unwrap();
    // 21 fixed XFs and 4,029 user cell XFs.
    let mut last = 0;
    for _ in 0..4029 {
        last = writer.add_cell_format(ExtendedFormat::default()).unwrap();
    }
    writer
        .write_number_with_format(sheet, 1, 0, 1.0, last)
        .unwrap();
    let before = written(&mut writer);
    let collection = "XF records of a workbook with XFExt records or a pivot table";
    assert_eq!(
        too_many(writer.add_cell_format(ExtendedFormat::default())),
        (collection, 4050)
    );
    assert_eq!(
        too_many(writer.add_cell_style(CellStyle::default())),
        (collection, 4050)
    );
    assert_eq!(
        too_many(writer.add_pivot_table(sheet, pivot_config("Sales", "Values"))),
        (collection, 4050)
    );
    assert_eq!(written(&mut writer), before);
    let workbook = read(before);
    assert_eq!(workbook.formatting().extended_formats().len(), 4050);

    // Without extensions the table can grow; extensions are then refused.
    writer.set_xf_extensions(Vec::new()).unwrap();
    writer.add_cell_format(ExtendedFormat::default()).unwrap();
    let before = written(&mut writer);
    assert_eq!(
        too_many(writer.set_xf_extensions(vec![XfExt::try_new(15, Vec::new()).unwrap()])),
        (collection, 4050)
    );
    assert_eq!(written(&mut writer), before);
}

// --- XFExt, StyleExt and CRN records -----------------------------------------

fn unknown_extension(data_len: usize) -> XfExt {
    XfExt::try_new(
        15,
        vec![ExtProp::Unknown {
            ext_type: 0x0100,
            data: vec![0xAB; data_len],
        }],
    )
    .unwrap()
}

#[test]
fn an_xfext_record_holds_8224_payload_bytes() {
    // A 20-byte fixed part, and a 4-byte header before each property's data.
    for data_len in [8199, 8200] {
        let (mut writer, _) = one_sheet_writer("Ext");
        writer
            .set_xf_extensions(vec![unknown_extension(data_len)])
            .unwrap();
        let workbook = read(written(&mut writer));
        assert_eq!(
            workbook.formatting().xf_extensions(),
            &[unknown_extension(data_len)]
        );
    }
    let (mut writer, _) = one_sheet_writer("Ext");
    let before = written(&mut writer);
    assert_eq!(
        record_too_long(writer.set_xf_extensions(vec![unknown_extension(8201)])),
        ("XFExt", 8225, 8224)
    );
    // 65,556 bytes wrapped to 20 before change 0766.
    assert_eq!(
        record_too_long(writer.set_xf_extensions(vec![
            XfExt::try_new(
                15,
                vec![
                    ExtProp::Unknown {
                        ext_type: 0x0100,
                        data: vec![0xAB; 252],
                    };
                    256
                ],
            )
            .unwrap()
        ])),
        ("XFExt", 65_556, 8224)
    );
    assert_eq!(written(&mut writer), before);
}

fn style_with_fonts(count: usize) -> StyleExt {
    StyleExt::try_new(
        false,
        StyleCategory::Custom,
        "Style".to_string(),
        XfProperties::try_new(vec![XfProperty::FontName("F".repeat(32)); count]).unwrap(),
    )
    .unwrap()
}

#[test]
fn a_style_ext_record_holds_one_record_of_properties() {
    // Find the most font-name properties one StyleExt record holds.
    let fits = |count| {
        let (mut writer, _) = one_sheet_writer("Styles");
        writer
            .set_style_extensions(vec![style_with_fonts(count)])
            .is_ok()
    };
    let most = (1..=2048).rev().find(|count| fits(*count)).unwrap();
    assert!(most > 100);
    let (mut writer, _) = one_sheet_writer("Styles");
    writer
        .set_style_extensions(vec![style_with_fonts(most)])
        .unwrap();
    let workbook = read(written(&mut writer));
    assert_eq!(
        workbook.style_extensions()[0]
            .properties()
            .properties()
            .len(),
        most
    );
    let before = written(&mut writer);
    let (record, bytes, limit) =
        record_too_long(writer.set_style_extensions(vec![style_with_fonts(most + 1)]));
    assert_eq!((record, limit), ("StyleExt", 8224));
    assert!(bytes > 8224);
    // 1,000 properties wrapped past 65,535 bytes before change 0766.
    record_too_long(writer.set_style_extensions(vec![style_with_fonts(1000)]));
    assert_eq!(written(&mut writer), before);
}

fn external_link(values: Vec<CachedValue>) -> ExternalWorkbookOptions {
    ExternalWorkbookOptions {
        encoded_virtual_path: "\u{1}book.xls".to_string(),
        sheets: vec![ExternalSheetOptions {
            name: "Data".to_string(),
            cache_rows: vec![ExternalCacheRowOptions {
                row: 3,
                first_column: 0,
                values,
            }],
        }],
    }
}

fn external_workbook_with_sheets(start: usize, count: usize) -> ExternalWorkbookOptions {
    ExternalWorkbookOptions {
        encoded_virtual_path: "\u{1}book.xls".to_string(),
        sheets: (start..start + count)
            .map(|index| ExternalSheetOptions {
                name: format!("S{index}"),
                cache_rows: Vec::new(),
            })
            .collect(),
    }
}

#[test]
fn external_sheet_reference_count_is_checked_at_registration() {
    let mut writer = Writer::new();
    writer.add_worksheet("Links").unwrap();

    // Five full SupBooks and one partial one produce N - 1 references.
    for (start, count) in [
        (0, 256),
        (256, 256),
        (512, 256),
        (768, 256),
        (1024, 256),
        (1280, 89),
    ] {
        writer
            .add_external_workbook_link(external_workbook_with_sheets(start, count))
            .unwrap();
    }
    let before_limit = written(&mut writer);

    // The BIFF8 ExternSheet record accepts exactly 1,370 references.
    writer
        .add_external_workbook_link(external_workbook_with_sheets(1369, 1))
        .unwrap();
    let at_limit = written(&mut writer);
    assert_ne!(at_limit, before_limit);

    // N + 1 is refused before the workbook state changes, and the accepted
    // N-byte workbook remains writable byte-for-byte.
    assert!(matches!(
        writer.add_external_workbook_link(external_workbook_with_sheets(1370, 1)),
        Err(Error::TooMany {
            collection: "ExternSheet references",
            limit: 1370,
        })
    ));
    assert_eq!(written(&mut writer), at_limit);
}

/// Fifteen 255-unit UTF-16 strings (514 bytes of `SerAr` each) and one of
/// `last` code units fill `4 + 7,710 + 4 + 2 * last` bytes of one `CRN`.
fn crn_values(last: &str) -> Vec<CachedValue> {
    let mut values = vec![CachedValue::Text("漢".repeat(255)); 15];
    values.push(CachedValue::Text(last.to_string()));
    values
}

#[test]
fn a_crn_record_holds_8224_payload_bytes() {
    for last in [
        "漢".repeat(252),
        "漢".repeat(253),
        format!("{}😀", "漢".repeat(251)),
    ] {
        let (mut writer, _) = one_sheet_writer("Links");
        writer
            .add_external_workbook_link(external_link(crn_values(&last)))
            .unwrap();
        let workbook = read(written(&mut writer));
        let book = workbook
            .external_links()
            .external_workbooks()
            .next()
            .unwrap();
        let row = &book.sheets()[0].cache_rows()[0];
        assert_eq!(row.values(), crn_values(&last).as_slice());
    }
    for last in ["漢".repeat(254), format!("{}😀", "漢".repeat(252))] {
        let (mut writer, _) = one_sheet_writer("Links");
        let before = written(&mut writer);
        assert_eq!(
            record_too_long(writer.add_external_workbook_link(external_link(crn_values(&last)))),
            ("CRN", 8226, 8224)
        );
        assert_eq!(written(&mut writer), before);
    }
    // 256 values: colLast is 255, which a debug build computed with an
    // overflowing subtraction before change 0766.
    let values = (0..256)
        .map(|value| CachedValue::Number(f64::from(value)))
        .collect::<Vec<_>>();
    let (mut writer, _) = one_sheet_writer("Links");
    writer
        .add_external_workbook_link(external_link(values.clone()))
        .unwrap();
    let workbook = read(written(&mut writer));
    let book = workbook
        .external_links()
        .external_workbooks()
        .next()
        .unwrap();
    assert_eq!(book.sheets()[0].cache_rows()[0].values(), values.as_slice());
}

// --- Formulas -----------------------------------------------------------------

#[test]
fn formulas_the_write_would_refuse_are_refused_when_they_are_set() {
    let (mut writer, sheet) = one_sheet_writer("Formulas");
    writer.write_formula(sheet, 1, 0, "1+2").unwrap();
    let before = written(&mut writer);
    for formula in ["SUM(", "", "=", "1+", "SUM(1,2"] {
        assert!(
            writer.write_formula(sheet, 1, 0, formula).is_err(),
            "{formula:?}"
        );
        // The cell keeps its formula and the writer still writes.
        assert_eq!(written(&mut writer), before);
    }
}

/// `n` terms of `1+1+...` encode to `4n - 1` bytes of tokens, and a Formula
/// record holds 8,202 bytes of them after its 22 fixed bytes.
fn sum_of_ones(terms: usize) -> String {
    vec!["1"; terms].join("+")
}

#[test]
fn a_formula_record_holds_8202_token_bytes() {
    for terms in [2049, 2050] {
        let (mut writer, sheet) = one_sheet_writer("Formulas");
        writer
            .write_formula(sheet, 1, 0, &sum_of_ones(terms))
            .unwrap();
        let workbook = read(written(&mut writer));
        let cell = workbook.xls_worksheet(0).unwrap().get_cell(1, 0).unwrap();
        assert_eq!(cell.formula_bytes().unwrap().len(), 4 * terms - 1);
    }
    let (mut writer, sheet) = one_sheet_writer("Formulas");
    let before = written(&mut writer);
    assert!(matches!(
        writer.write_formula(sheet, 1, 0, &sum_of_ones(2051)),
        Err(Error::InvalidFormula(_))
    ));
    assert_eq!(written(&mut writer), before);
}

#[test]
fn a_conditional_format_the_write_would_refuse_is_refused_when_added() {
    let (mut writer, sheet) = one_sheet_writer("Rules");
    let before = written(&mut writer);
    for formula in ["SUM(".to_string(), format!("\"{}\"", "a".repeat(256))] {
        assert!(
            writer
                .add_conditional_format(
                    sheet,
                    ConditionalFormat {
                        first_row: 0,
                        last_row: 3,
                        first_col: 0,
                        last_col: 0,
                        format_type: ConditionalFormatType::Formula { formula },
                        pattern: None,
                    },
                )
                .is_err()
        );
        assert_eq!(written(&mut writer), before);
    }
}

// --- Worksheet and defined names ------------------------------------------------

#[test]
fn worksheet_names_biff8_forbids_are_refused() {
    for name in [
        "a/b", "a\\b", "a?b", "a*b", "a[b", "a]b", "a:b", "a\0b", "a\u{3}b", "'lead", "trail'",
    ] {
        let mut writer = Writer::new();
        writer.add_worksheet("Kept").unwrap();
        let before = written(&mut writer);
        invalid_data(writer.add_worksheet(name));
        assert_eq!(written(&mut writer), before, "{name:?}");
    }
    for name in ["mid'quote", "a b", "漢字", "R😀", "a.b-c_d(1)"]
        .into_iter()
        .map(str::to_string)
        .chain(strings_of(30))
        .chain(strings_of(31))
    {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet(&name).unwrap();
        writer.write_number(sheet, 0, 0, 1.0).unwrap();
        let workbook = read(written(&mut writer));
        assert_eq!(workbook.sheets()[0].name(), name);
    }
    for name in strings_of(32).into_iter().chain([straddling(31)]) {
        let mut writer = Writer::new();
        assert_eq!(
            too_long(writer.add_worksheet(&name)),
            ("worksheet name", 32, 31)
        );
    }
}

#[test]
fn defined_names_with_nul_are_refused() {
    let (mut writer, _) = one_sheet_writer("Names");
    let before = written(&mut writer);
    invalid_data(writer.define_name("A\0B", "A1"));
    assert_eq!(written(&mut writer), before);
    writer.define_name("Größe", "A1:B2").unwrap();
    let workbook = read(written(&mut writer));
    assert!(
        workbook
            .defined_names()
            .iter()
            .any(|name| name.name == "Größe")
    );
}

// --- PivotTables --------------------------------------------------------------

fn pivot_config(name: &str, data_field_name: &str) -> PivotTableConfig {
    PivotTableConfig {
        name: name.to_string(),
        source_type: 1,
        source_sheet_name: "Source".to_string(),
        source_first_row: 0,
        source_last_row: 2,
        source_first_col: 0,
        source_last_col: 1,
        first_row: 8,
        last_row: 10,
        first_col: 0,
        last_col: 1,
        first_header_row: 8,
        first_data_row: 9,
        first_data_col: 1,
        data_field_name: data_field_name.to_string(),
        data_axis: 0,
        data_position: 0,
        fields: vec![
            PivotFieldConfig {
                axis: 1,
                subtotal_count: 0,
                subtotal_flags: 0,
                items: vec![
                    PivotItemConfig {
                        item_type: 0,
                        flags: 0,
                        cache_index: 0,
                        name: None,
                    },
                    PivotItemConfig {
                        item_type: 0,
                        flags: 0,
                        cache_index: 1,
                        name: None,
                    },
                ],
                name: None,
                cache_name: "Region".to_string(),
                cache_items: vec![PivotCacheItem::from("East"), PivotCacheItem::from("West")],
                is_numeric: false,
                grouping: None,
            },
            PivotFieldConfig {
                axis: 8,
                subtotal_count: 0,
                subtotal_flags: 0,
                items: Vec::new(),
                name: None,
                cache_name: "Amount".to_string(),
                cache_items: Vec::new(),
                is_numeric: true,
                grouping: None,
            },
        ],
        data_items: vec![PivotDataItemConfig {
            source_field_index: 1,
            function: 0,
            display_format: 0,
            base_field_index: 0,
            base_item_index: 0,
            num_format_index: 0,
            name: "Sum of Amount".to_string(),
        }],
        page_entries: Vec::new(),
        source_data: vec![
            vec![
                PivotCacheValue::StringIndex(0),
                PivotCacheValue::Number(10.0),
            ],
            vec![
                PivotCacheValue::StringIndex(1),
                PivotCacheValue::Number(20.0),
            ],
        ],
    }
}

fn pivot_writer() -> (Writer, usize) {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Source").unwrap();
    (writer, sheet)
}

#[test]
fn pivot_table_names_count_utf16_code_units_up_to_their_limits() {
    let named = |name: String, data: String, field: Option<String>, item: Option<String>| {
        let mut config = pivot_config(&name, &data);
        config.fields[0].name = field;
        config.fields[0].items[0].name = item;
        config.data_items[0].name = format!("Sum{}", "😀");
        config
    };
    for (name, data, field, item) in [
        ("", "V", None, None),
        ("S😀", "V😀", Some("F😀"), Some("I😀")),
    ]
    .map(|(a, b, c, d)| {
        (
            a.to_string(),
            b.to_string(),
            c.map(str::to_string),
            d.map(str::to_string),
        )
    })
    .into_iter()
    .chain([
        (
            straddling(254),
            straddling(253),
            Some(straddling(254)),
            Some(straddling(253)),
        ),
        (
            "t".repeat(255),
            "d".repeat(254),
            Some("f".repeat(255)),
            Some("i".repeat(254)),
        ),
    ]) {
        let (mut writer, sheet) = pivot_writer();
        let config = named(name.clone(), data.clone(), field.clone(), item.clone());
        writer.add_pivot_table(sheet, config).unwrap();
        let workbook = read(written(&mut writer));
        let table = &workbook.worksheet_pivot_tables(0).unwrap()[0];
        assert_eq!(table.view.name, name);
        assert_eq!(table.view.data_field_name, data);
        assert_eq!(table.fields[0].name, field);
        assert!(table.fields[0].items.iter().any(|entry| entry.name == item));
        assert_eq!(table.data_items[0].name, "Sum😀");
    }
    let refusals: [(PivotTableConfig, (&str, usize, usize)); 4] = [
        (
            named(straddling(255), "V".into(), None, None),
            ("PivotTable name", 256, 255),
        ),
        (
            named("T".into(), straddling(254), None, None),
            ("PivotTable data field name", 255, 254),
        ),
        (
            named("T".into(), "V".into(), Some(straddling(255)), None),
            ("PivotTable field name", 256, 255),
        ),
        (
            named("T".into(), "V".into(), None, Some(straddling(254))),
            ("PivotTable item name", 255, 254),
        ),
    ];
    for (config, expected) in refusals {
        let (mut writer, sheet) = pivot_writer();
        let before = written(&mut writer);
        assert_eq!(too_long(writer.add_pivot_table(sheet, config)), expected);
        // No output cell, pivot table or pivot XF was added.
        assert_eq!(written(&mut writer), before);
    }
    let (mut writer, sheet) = pivot_writer();
    invalid_data(writer.add_pivot_table(sheet, pivot_config("T", "")));
    let mut config = pivot_config("T", "V");
    config.source_sheet_name = "a/b".to_string();
    invalid_data(writer.add_pivot_table(sheet, config));
}

#[test]
fn pivot_cache_strings_fit_one_sxstring_record() {
    for (item, fits) in [
        ("a".repeat(8221), true),
        ("a".repeat(8222), false),
        ("é".repeat(4110), true),
        ("é".repeat(4111), false),
        (format!("{}😀", "é".repeat(4108)), true),
        (format!("{}😀", "é".repeat(4109)), false),
    ] {
        let (mut writer, sheet) = pivot_writer();
        let mut config = pivot_config("Sales", "Values");
        config.fields[0].cache_items[0] = PivotCacheItem::from(item.as_str());
        let before = written(&mut writer);
        let added = writer.add_pivot_table(sheet, config);
        assert_eq!(added.is_ok(), fits, "{} units", item.encode_utf16().count());
        if fits {
            let workbook = read(written(&mut writer));
            let table = &workbook.worksheet_pivot_tables(0).unwrap()[0];
            let cache = workbook.pivot_cache_for_table(table).unwrap();
            assert_eq!(
                table.cache_field(cache, 0).unwrap().items()[0],
                PivotCacheItem::from(item.as_str())
            );
        } else {
            too_long(added);
            assert_eq!(written(&mut writer), before);
        }
    }
}

#[test]
fn pivot_xfs_follow_user_cell_formats_that_reach_past_index_63() {
    for formats in [43, 44, 100] {
        let (mut writer, sheet) = pivot_writer();
        for _ in 0..formats {
            writer.add_cell_format(ExtendedFormat::default()).unwrap();
        }
        writer
            .add_pivot_table(sheet, pivot_config("Sales", "Values"))
            .unwrap();
        let workbook = read(written(&mut writer));
        let xfs = workbook.formatting().extended_formats();
        let pivot_start = (21 + formats).max(64);
        assert_eq!(xfs.len(), pivot_start + 3);
        // The grand-total value cell uses the third pivot XF.
        let cell = workbook.xls_worksheet(0).unwrap().get_cell(10, 1).unwrap();
        assert_eq!(usize::from(cell.xf_index()), pivot_start + 2);
    }
}

/// The PivotTable view editor re-encodes names with the same encoders, and
/// its own `SXVDEx`, `SxEx` and `QsiSXTag` strings count UTF-16 code units
/// too; until change 0766 a supplementary-plane character left each count
/// one short.
#[test]
fn the_pivot_view_editor_writes_astral_names_that_read_back() {
    let (mut writer, sheet) = pivot_writer();
    writer
        .add_pivot_table(sheet, pivot_config("Sales", "Values"))
        .unwrap();
    let mut editor = litchi_xls::PivotViewEditor::new(written(&mut writer)).unwrap();
    editor
        .update_by_name(0, "Sales", |table| {
            table.view.name = "S😀".to_string();
            table.view.data_field_name = "V😀".to_string();
            table.fields[0].name = Some("F😀".to_string());
            table.fields[0].extension.as_mut().unwrap().subtotal_name = Some("T😀".to_string());
            let extension = table.extension.as_mut().unwrap();
            extension.error_string = Some("E😀".to_string());
            extension.table_style = Some("PivotStyle😀".to_string());
            table.query_tag.as_mut().unwrap().table_name = "S😀".to_string();
        })
        .unwrap();
    let workbook = read(editor.finish().unwrap());
    let table = &workbook.worksheet_pivot_tables(0).unwrap()[0];
    assert_eq!(table.view.name, "S😀");
    assert_eq!(table.view.data_field_name, "V😀");
    assert_eq!(table.fields[0].name.as_deref(), Some("F😀"));
    assert_eq!(
        table.fields[0]
            .extension
            .as_ref()
            .unwrap()
            .subtotal_name
            .as_deref(),
        Some("T😀")
    );
    let extension = table.extension.as_ref().unwrap();
    assert_eq!(extension.error_string.as_deref(), Some("E😀"));
    assert_eq!(extension.table_style.as_deref(), Some("PivotStyle😀"));
    assert_eq!(table.query_tag.as_ref().unwrap().table_name, "S😀");

    // A name its record cannot count is refused, not wrapped.
    let (mut writer, sheet) = pivot_writer();
    writer
        .add_pivot_table(sheet, pivot_config("Sales", "Values"))
        .unwrap();
    let mut editor = litchi_xls::PivotViewEditor::new(written(&mut writer)).unwrap();
    editor
        .update_by_name(0, "Sales", |table| {
            table.view.name = straddling(255);
        })
        .unwrap();
    assert_eq!(too_long(editor.finish()), ("PivotTable name", 256, 255));
}

/// A formula's tokens are encoded once, when it is set, and the write reuses
/// them; replacing the cell drops them, so the write never emits the tokens
/// of a formula the cell no longer holds.
#[test]
fn a_replaced_formula_cell_writes_what_it_holds_now() {
    let expected = |formula: &str| {
        let tokens = litchi_xls::writer::FormulaTokenizer::new()
            .tokenize(formula)
            .unwrap();
        litchi_xls::writer::formula::encode_ptg_tokens(&tokens).unwrap()
    };
    let (mut writer, sheet) = one_sheet_writer("Formulas");
    writer.write_formula(sheet, 1, 0, "1+2").unwrap();
    writer.write_formula(sheet, 1, 1, "SUM(A1:B2)").unwrap();
    writer.write_formula(sheet, 1, 2, "LEN(\"x\")").unwrap();
    // A formula replaced by a string, by a number and by another formula.
    writer.write_string(sheet, 1, 0, "text").unwrap();
    writer.write_number(sheet, 1, 1, 7.0).unwrap();
    writer.write_formula(sheet, 1, 2, "ABS(-3)*2").unwrap();
    // A string replaced by a formula, and a formula with a leading `=`.
    writer.write_string(sheet, 2, 0, "gone").unwrap();
    writer.write_formula(sheet, 2, 0, "=MAX(1,2)").unwrap();
    let workbook = read(written(&mut writer));
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(
        worksheet.get_cell(1, 0).unwrap().value(),
        &CellValue::String("text".to_string())
    );
    assert!(worksheet.get_cell(1, 0).unwrap().formula_bytes().is_none());
    assert!(worksheet.get_cell(1, 1).unwrap().formula_bytes().is_none());
    assert_eq!(
        worksheet.get_cell(1, 2).unwrap().formula_bytes(),
        Some(expected("ABS(-3)*2").as_slice())
    );
    assert_eq!(
        worksheet.get_cell(2, 0).unwrap().formula_bytes(),
        Some(expected("MAX(1,2)").as_slice())
    );
}

/// User `XFExt` records may name style, default and cell-format XFs; the
/// writer's pivot-table XFs follow those and move as formats are added.
#[test]
fn xf_extensions_cannot_name_the_pivot_xfs() {
    let (mut writer, sheet) = pivot_writer();
    writer
        .add_pivot_table(sheet, pivot_config("Sales", "Values"))
        .unwrap();
    let before = written(&mut writer);
    for index in [21, 63, 64, 66] {
        invalid_data(writer.set_xf_extensions(vec![XfExt::try_new(index, Vec::new()).unwrap()]));
        assert_eq!(written(&mut writer), before);
    }
    writer
        .set_xf_extensions(vec![XfExt::try_new(20, vec![ExtProp::Indent(1)]).unwrap()])
        .unwrap();
    let workbook = read(written(&mut writer));
    let indices = workbook
        .formatting()
        .xf_extensions()
        .iter()
        .map(XfExt::xf_index)
        .collect::<Vec<_>>();
    assert_eq!(indices, [20, 64, 65, 66]);
}

/// An empty data-item name is written as absent (`cchName` 0xFFFF,
/// [MS-XLS] 2.4.278), and the reader now reads it so.
#[test]
fn an_empty_data_item_name_round_trips() {
    let (mut writer, sheet) = pivot_writer();
    let mut config = pivot_config("Sales", "Values");
    config.data_items[0].name = String::new();
    writer.add_pivot_table(sheet, config).unwrap();
    let workbook = read(written(&mut writer));
    let table = &workbook.worksheet_pivot_tables(0).unwrap()[0];
    assert_eq!(table.data_items[0].name, "");
}

fn grouping_config(base_items: usize) -> PivotTableConfig {
    let mut config = pivot_config("Grouped", "Values");
    config.fields = vec![
        PivotFieldConfig {
            axis: 0,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: Vec::new(),
            name: None,
            cache_name: "Base".to_string(),
            cache_items: (0..base_items)
                .map(|item| PivotCacheItem::from(format!("I{item}").as_str()))
                .collect(),
            is_numeric: false,
            grouping: None,
        },
        PivotFieldConfig {
            axis: 1,
            subtotal_count: 0,
            subtotal_flags: 0,
            items: Vec::new(),
            name: None,
            cache_name: "Group".to_string(),
            cache_items: Vec::new(),
            is_numeric: false,
            grouping: Some(litchi_xls::PivotCacheGrouping::Discrete(
                litchi_xls::PivotCacheDiscreteGrouping {
                    base_field_index: 0,
                    group_items: vec!["Low".into(), "High".into()],
                    item_to_group: (0..base_items)
                        .map(|item| u16::from(item % 2 == 1))
                        .collect(),
                },
            )),
        },
    ];
    config.data_items = Vec::new();
    config.source_data = Vec::new();
    config.source_last_row = 0;
    config.source_last_col = 0;
    config
}

/// The cache stream is built when the table is added, so a grouping whose
/// `SXGroupInfo` record (two bytes per base item) cannot fit is refused then;
/// no API removes a pivot table, so a write-time refusal would stop every
/// later write.
#[test]
fn a_pivot_cache_record_that_cannot_fit_is_refused_when_the_table_is_added() {
    for base_items in [4111, 4112] {
        let (mut writer, sheet) = pivot_writer();
        writer
            .add_pivot_table(sheet, grouping_config(base_items))
            .unwrap();
        let workbook = read(written(&mut writer));
        assert_eq!(workbook.worksheet_pivot_tables(0).unwrap().len(), 1);
    }
    let (mut writer, sheet) = pivot_writer();
    let before = written(&mut writer);
    assert_eq!(
        record_too_long(writer.add_pivot_table(sheet, grouping_config(4113))),
        ("SXGroupInfo", 8226, 8224)
    );
    assert_eq!(written(&mut writer), before);
}

/// Litchi's reader reads at most 4,096 fields per view and no overlapping
/// views on one worksheet; the writer refuses what it would not read back.
#[test]
fn pivot_tables_the_reader_would_refuse_are_refused_when_added() {
    let (mut writer, sheet) = pivot_writer();
    writer
        .add_pivot_table(sheet, pivot_config("Sales", "Values"))
        .unwrap();
    let before = written(&mut writer);
    let mut overlapping = pivot_config("Other", "Values");
    overlapping.first_col = 1;
    overlapping.last_col = 2;
    overlapping.first_data_col = 2;
    invalid_data(writer.add_pivot_table(sheet, overlapping));
    let mut wide = pivot_config("Wide", "Values");
    wide.first_row = 20;
    wide.last_row = 22;
    wide.first_header_row = 20;
    wide.first_data_row = 21;
    let extra = wide.fields[1].clone();
    wide.fields.resize(4097, extra);
    wide.source_data = Vec::new();
    assert_eq!(
        too_many(writer.add_pivot_table(sheet, wide)),
        ("PivotTable fields", 4096)
    );
    assert_eq!(written(&mut writer), before);
}
