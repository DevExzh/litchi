#![allow(
    clippy::unwrap_used,
    clippy::float_cmp,
    reason = "focused fixture assertions compare exact authored values"
)]

//! Readable cell formatting on both workbook owners.
//!
//! The eager owner reads it through `Style::format`, the source-backed owner
//! through the style key of a cell record. Both must report the same typed
//! values for the same cell, and those values must carry everything the
//! `OfficeCLI` cell readback needs (`schemas/help/xlsx/cell.json`).

use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

use litchi_xlsx::style::cell_format::{
    Color, Fill, Format, Line, Pattern, Reading, Script, Side, Underline,
};
use litchi_xlsx::{LocalStyle, Rect, SourceBackedWorkbook, WorksheetKind};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsx")
        .join(name)
}

/// Every typed value a format exposes, for equality across owners.
fn snapshot(format: &Format) -> String {
    format!(
        "{:?}|{:?}|{:?}|{:?}|{:?}|{:?}|{}|{}{}{}{}",
        format.number_format(),
        format.font(),
        format.fill(),
        format.border(),
        format.alignment(),
        format.protection(),
        format.quote_prefix(),
        format.is_base(),
        format.has_default_font(),
        format.has_default_fill(),
        format.has_default_border(),
    )
}

/// Render a format the way `OfficeCLI`'s cell readback does
/// (`ExcelHandler.Helpers.Node.cs`): nothing for the base format, no font,
/// fill or border keys for the workbook's default records, `#RRGGBB` or
/// scheme names with a tint suffix for colours, `11pt` for sizes. It shows
/// that each key is served from typed values without parsing again.
fn officecli(format: &Format) -> BTreeMap<&'static str, String> {
    let mut keys = BTreeMap::new();
    if format.is_base() {
        return keys;
    }
    if let Some(font) = format.font().filter(|_| !format.has_default_font()) {
        // `OfficeCLI` reports a present `<b val="0"/>` as `true`; the typed
        // value keeps it apart as `Some(false)`, so a node view can emit the
        // explicit `false` record 0010 §2.3 asks for instead.
        if font.bold().is_some() {
            keys.insert("font.bold", "true".to_owned());
        }
        if font.italic().is_some() {
            keys.insert("font.italic", "true".to_owned());
        }
        if font.strike().is_some() {
            keys.insert("strike", "true".to_owned());
        }
        match font.underline() {
            Some(Underline::None) | None => {},
            Some(underline) => {
                keys.insert("underline", underline.as_str().to_owned());
            },
        }
        if let Some(color) = font.color().and_then(color_text) {
            keys.insert("font.color", color);
        }
        match font.script() {
            Some(Script::Superscript) => {
                keys.insert("superscript", "true".to_owned());
            },
            Some(Script::Subscript) => {
                keys.insert("subscript", "true".to_owned());
            },
            Some(Script::Baseline) | None => {},
        }
        if let Some(size) = font.size() {
            keys.insert("font.size", format!("{size}pt"));
        }
        if let Some(name) = font.name() {
            keys.insert("font.name", name.to_owned());
        }
        if let Some(family) = font.family() {
            keys.insert("font.family", family.to_string());
        }
        if let Some(charset) = font.charset() {
            keys.insert("font.charset", charset.to_string());
        }
        if let Some(scheme) = font.scheme() {
            keys.insert("font.scheme", scheme.as_str().to_owned());
        }
    }
    if let Some(fill) = format.fill().filter(|_| !format.has_default_fill()) {
        match fill {
            Fill::Gradient(gradient) => {
                let colors = gradient
                    .stops()
                    .iter()
                    .filter_map(|stop| match stop.color() {
                        Color::Rgb { value, .. } => Some(hex(value)),
                        _ => None,
                    })
                    .collect::<Vec<_>>();
                if colors.len() >= 2 {
                    keys.insert(
                        "fill",
                        format!("gradient;{};{}", colors.join(";"), gradient.degree() as i64),
                    );
                }
            },
            Fill::Pattern(pattern) => {
                let solid = matches!(
                    pattern.pattern(),
                    None | Some(Pattern::Solid | Pattern::None)
                );
                if !solid {
                    keys.insert(
                        "fillPattern",
                        pattern.pattern().unwrap().as_str().to_owned(),
                    );
                    if let Some(color) = pattern.background().and_then(color_text) {
                        keys.insert("fillBg", color);
                    }
                }
                if let Some(color) = pattern.foreground().and_then(color_text) {
                    keys.insert("fill", color);
                }
            },
            _ => {},
        }
    }
    if let Some(border) = format.border().filter(|_| !format.has_default_border()) {
        let sides = [
            ("border.left", "border.left.color", &border.left),
            ("border.right", "border.right.color", &border.right),
            ("border.top", "border.top.color", &border.top),
            ("border.bottom", "border.bottom.color", &border.bottom),
        ];
        for (key, color_key, side) in sides {
            if let Some(Side { line, color }) = side {
                keys.insert(key, line.as_str().to_owned());
                if let Some(Color::Rgb { value, .. }) = color {
                    keys.insert(color_key, hex(*value));
                }
            }
        }
        if let Some(diagonal) = &border.diagonal {
            if let Some(side) = diagonal.side() {
                keys.insert("border.diagonal", side.line.as_str().to_owned());
            }
            if diagonal.dir().is_some_and(|dir| dir.is_up()) {
                keys.insert("border.diagonalUp", "true".to_owned());
            }
            if diagonal.dir().is_some_and(|dir| dir.is_down()) {
                keys.insert("border.diagonalDown", "true".to_owned());
            }
        }
    }
    if let Some(alignment) = format.alignment() {
        if alignment.wrap_text {
            keys.insert("alignment.wrapText", "true".to_owned());
        }
        if let Some(horizontal) = alignment.horizontal {
            keys.insert("alignment.horizontal", horizontal.as_str().to_owned());
        }
        if let Some(vertical) = alignment.vertical {
            keys.insert("alignment.vertical", vertical.as_str().to_owned());
        }
        if let Some(rotation) = alignment.text_rotation.filter(|value| value.get() != 0) {
            keys.insert("alignment.textRotation", rotation.get().to_string());
        }
        if let Some(indent) = alignment.indent.filter(|value| value.get() > 0) {
            keys.insert("alignment.indent", indent.get().to_string());
        }
        if alignment.shrink_to_fit {
            keys.insert("alignment.shrinkToFit", "true".to_owned());
        }
        match alignment.reading_order {
            Some(Reading::LeftToRight) => {
                keys.insert("alignment.readingOrder", "ltr".to_owned());
            },
            Some(Reading::RightToLeft) => {
                keys.insert("alignment.readingOrder", "rtl".to_owned());
            },
            Some(Reading::Context) | None => {},
        }
    }
    if let Some(protection) = format.protection() {
        if let Some(locked) = protection.locked() {
            keys.insert("protection.locked", locked.to_string());
        }
        if let Some(hidden) = protection.hidden() {
            keys.insert("protection.hidden", hidden.to_string());
        }
    }
    if format.quote_prefix() {
        keys.insert("quotePrefix", "true".to_owned());
    }
    let number_format = format.number_format();
    if litchi_xlsx::style::cell_format::raw::number_format_id(&number_format) > 0 {
        keys.insert(
            "numFmtId",
            litchi_xlsx::style::cell_format::raw::number_format_id(&number_format).to_string(),
        );
        keys.insert(
            "numberformat",
            number_format.code().map_or_else(
                || {
                    litchi_xlsx::style::cell_format::raw::number_format_id(&number_format)
                        .to_string()
                },
                str::to_owned,
            ),
        );
    }
    keys
}

fn hex(value: litchi_xlsx::style::cell_format::Rgb) -> String {
    if value.alpha() == 0xFF {
        format!(
            "#{:02X}{:02X}{:02X}",
            value.red(),
            value.green(),
            value.blue()
        )
    } else {
        format!(
            "#{:02X}{:02X}{:02X}{:02X}",
            value.red(),
            value.green(),
            value.blue(),
            value.alpha()
        )
    }
}

fn color_text(color: Color) -> Option<String> {
    match color {
        Color::Rgb { value, .. } => Some(hex(value)),
        Color::Theme { index, tint } => {
            let name = [
                "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
                "accent6", "hlink", "folHlink",
            ]
            .get(usize::try_from(index).ok()?)?;
            match tint
                .map(|tint| tint.get())
                .filter(|tint| tint.abs() >= 0.005)
            {
                Some(tint) => Some(format!("{name}+tint{}", (tint * 10_000.0).round() / 100.0)),
                None => Some((*name).to_owned()),
            }
        },
        _ => None,
    }
}

fn keys(pairs: &[(&'static str, &str)]) -> BTreeMap<&'static str, String> {
    pairs
        .iter()
        .map(|(key, value)| (*key, (*value).to_owned()))
        .collect()
}

#[test]
fn eager_style_format_reads_fonts_fills_alignment_and_number_formats() {
    let workbook = litchi_xlsx::Package::open(fixture("styles.xlsx")).unwrap();
    let sheet = workbook.workbook_ref().sheet(0_usize).unwrap().unwrap();
    let format = |at: &str| sheet.style(at).unwrap().unwrap().format().unwrap();

    let bold = format("A1");
    let font = bold.font().unwrap();
    assert_eq!(font.bold(), Some(true));
    assert_eq!(font.name(), Some("Arial"));
    assert_eq!(font.size(), Some(11.0));
    assert_eq!(font.color(), Some(Color::rgb(0, 0, 0)));
    assert_eq!(
        officecli(&bold),
        keys(&[
            ("font.bold", "true"),
            ("font.color", "#000000"),
            ("font.name", "Arial"),
            ("font.size", "11pt"),
        ])
    );
    assert_eq!(format("A2").font().unwrap().italic(), Some(true));
    assert_eq!(
        format("A3").font().unwrap().underline(),
        Some(Underline::Single)
    );
    assert_eq!(
        officecli(&format("A4")).get("alignment.horizontal"),
        Some(&"center".to_owned())
    );
    assert_eq!(
        officecli(&format("A5")).get("alignment.horizontal"),
        Some(&"right".to_owned())
    );

    let filled = format("A11");
    assert_eq!(
        officecli(&filled),
        keys(&[
            ("fill", "dk2+tint40"),
            ("font.name", "Arial"),
            ("font.size", "11pt"),
        ])
    );
    assert_eq!(
        filled.fill().and_then(Fill::pattern).unwrap().background(),
        Some(Color::indexed(64))
    );

    // The column style is a shared format as well.
    let column = sheet.column_style("A").unwrap();
    let Some(LocalStyle::Shared(style)) = column else {
        panic!("column A has a shared style: {column:?}");
    };
    assert_eq!(
        style.format().unwrap().font().unwrap().name(),
        Some("Arial")
    );

    let numbers = litchi_xlsx::Package::open(fixture("Formatting.xlsx")).unwrap();
    let styles = numbers.workbook_ref().styles().unwrap();
    let date = styles.get(1).unwrap().format().unwrap();
    assert_eq!(
        litchi_xlsx::style::cell_format::raw::number_format_id(&date.number_format()),
        14
    );
    assert_eq!(date.number_format().code(), Some("mm-dd-yy"));
    assert!(date.number_format().is_date_time());
    let pounds = styles.get(10).unwrap().format().unwrap();
    assert_eq!(pounds.number_format().code(), Some("\"£\"#,##0.00"));
    assert!(pounds.number_format().is_declared());
    assert_eq!(
        officecli(&pounds),
        keys(&[("numFmtId", "172"), ("numberformat", "\"£\"#,##0.00")])
    );
}

#[test]
fn eager_borders_and_colors_keep_their_typed_references() {
    let workbook = litchi_xlsx::Package::open(fixture("cell-borders.xlsx")).unwrap();
    let styles = workbook.workbook_ref().styles().unwrap();
    let double = styles.get(2).unwrap().format().unwrap();
    let border = double.border().unwrap();
    assert_eq!(
        border.top,
        Some(Side::new(Line::Double).with_color(Color::indexed(64)))
    );
    assert_eq!(officecli(&double), keys(&[("border.top", "double")]));

    let colors = litchi_xlsx::Package::open(fixture("TextColor.xlsx")).unwrap();
    let base = colors
        .workbook_ref()
        .styles()
        .unwrap()
        .base()
        .unwrap()
        .format()
        .unwrap();
    assert!(base.is_base());
    assert_eq!(base.font().unwrap().color(), Some(Color::theme(1)));
    assert!(officecli(&base).is_empty());
}

#[test]
fn formats_share_one_parsed_table_until_it_is_released() {
    let workbook = litchi_xlsx::Package::open(fixture("styles.xlsx")).unwrap();
    let styles = workbook.workbook_ref().styles().unwrap();
    let first = styles.get(1).unwrap().format().unwrap();
    let again = styles.get(1).unwrap().format().unwrap();
    assert!(first.same(&again), "a second read hits the cached table");

    // A clone and a descendant snapshot of the lineage share the cache.
    let clone = workbook.clone();
    assert!(
        clone
            .workbook_ref()
            .styles()
            .unwrap()
            .get(1)
            .unwrap()
            .format()
            .unwrap()
            .same(&first)
    );

    // An active format pins the table; once released, the next read parses
    // again and reports equal values.
    assert!(!workbook.workbook_ref().release_cached_formats());
    drop((first, again));
    assert!(workbook.workbook_ref().release_cached_formats());
    let reread = styles.get(1).unwrap().format().unwrap();
    assert_eq!(
        snapshot(&reread),
        snapshot(&styles.get(1).unwrap().format().unwrap())
    );
}

#[test]
fn source_records_carry_the_style_the_eager_owner_reports() {
    for name in [
        "styles.xlsx",
        "cell-borders.xlsx",
        "TextColor.xlsx",
        "Formatting.xlsx",
        "fontSize.xlsx",
        "ConditionalFormattingSamples.xlsx",
        "ExcelTables.xlsx",
    ] {
        compare_owners(&fixture(name), usize::MAX);
    }
}

#[test]
fn every_repository_workbook_reads_the_same_formats_on_both_owners() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/xlsx");
    let mut compared = 0usize;
    for entry in std::fs::read_dir(root).unwrap().flatten() {
        let path = entry.path();
        if path.extension().and_then(|value| value.to_str()) != Some("xlsx") {
            continue;
        }
        if compare_owners(&path, 4_096) {
            compared += 1;
        }
    }
    assert!(compared > 30, "only {compared} workbooks compared");
}

/// Compare every stored record of every worksheet, up to `limit` records per
/// sheet. Returns false for a workbook either owner refuses to open or whose
/// worksheets either owner refuses to read: those refusals belong to other
/// tests and must at least agree.
fn compare_owners(path: &Path, limit: usize) -> bool {
    let (Ok(eager), Ok(source)) = (
        litchi_xlsx::Package::open(path),
        SourceBackedWorkbook::open(path),
    ) else {
        return false;
    };
    for (eager_sheet, source_sheet) in eager.workbook_ref().sheets().zip(source.sheets().unwrap()) {
        if eager_sheet.kind() != WorksheetKind::Worksheet {
            continue;
        }
        let records = match source_sheet.cell_records(Rect::ALL) {
            Ok(records) => records,
            Err(_) => {
                assert!(eager_sheet.cells(Rect::ALL).is_err(), "{}", path.display());
                return false;
            },
        };
        let eager_cells = eager_sheet.cells(Rect::ALL).unwrap().count();
        assert_eq!(records.len(), eager_cells, "{}", path.display());
        for record in records.iter().take(limit) {
            let local = eager_sheet.local_style(record.address).unwrap().unwrap();
            assert_eq!(
                matches!(local, LocalStyle::Default),
                record.style.is_none(),
                "{} {}",
                path.display(),
                record.address
            );
            // A record is returned as stored, merged follower or not.
            let stored = eager_sheet
                .cells(Rect::single(record.address))
                .unwrap()
                .next()
                .map(|(_, cell)| cell);
            assert_eq!(Some(&record.cell), stored);
            let eager_format = eager_sheet
                .style(record.address)
                .unwrap()
                .map(|style| style.format().unwrap());
            let source_format = source_sheet.format(record.style.as_ref()).unwrap();
            assert_eq!(
                eager_format.as_ref().map(snapshot),
                source_format.as_ref().map(snapshot),
                "{} {}",
                path.display(),
                record.address
            );
            if let Some(single) = source_sheet.cell_record(record.address).unwrap() {
                assert_eq!(&single, record);
            } else {
                panic!("{} {} has no single record", path.display(), record.address);
            }
        }
    }
    true
}

#[test]
fn source_cell_records_follow_ranges_and_release_their_table() {
    let workbook = SourceBackedWorkbook::open(fixture("styles.xlsx")).unwrap();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let records = sheet.cell_records("A1:A3").unwrap();
    assert_eq!(
        records
            .iter()
            .map(|record| record.address.to_string())
            .collect::<Vec<_>>(),
        ["A1", "A2", "A3"]
    );
    assert!(records.iter().all(|record| record.style.is_some()));
    assert_ne!(records[0].style, records[1].style);
    assert_ne!(records[1].style, records[2].style);
    assert!(sheet.cell_record("C1").unwrap().is_none());
    assert!(sheet.cell_records("C1:D9").unwrap().is_empty());

    let base = workbook.format(None).unwrap().unwrap();
    assert!(base.is_base());
    let bold = sheet.format(records[0].style.as_ref()).unwrap().unwrap();
    assert_eq!(bold.font().unwrap().bold(), Some(true));
    assert!(bold.same(&workbook.format(records[0].style.as_ref()).unwrap().unwrap()));
    assert!(
        !workbook.release_cached_formats(),
        "active formats pin the table"
    );
    drop((base, bold));
    assert!(workbook.release_cached_formats());
    assert!(!workbook.release_cached_formats());

    // Converting a record drops only its style.
    let plain = litchi_xlsx::SourceCell::from(records[0].clone());
    assert_eq!(plain.address, records[0].address);
    assert_eq!(plain.cell, records[0].cell);
}
