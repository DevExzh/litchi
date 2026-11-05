//! The text, annotated, outline and stats views in OfficeCLI's shapes
//! (record 0010 §6, `ExcelHandler.View.cs`).

use super::{Mode, Options, SheetRange, both, kitchen, open_path, view};
use crate::Code;
use crate::format::xlsx::view_json;

const TEXT: &str = "=== Sheet: Data ===
[/sheet[1]/row[1]] A1=Region\tB1=Sales\tC1=Status\tE1=Merged
[/sheet[1]/row[2]] A2=West\tB2=1200\tC2=open
[/sheet[1]/row[3]] A3=East\tB3=800\tC3=closed\tD3=TRUE
[/sheet[1]/row[4]] A4=North\tB4=450.5\tC4=open\tD4=#DIV/0!
[/sheet[1]/row[5]] A5=South\tB5=2000\tC5=closed\tD5=South!
[/sheet[1]/row[6]] B6=4450.5\tC6=\tD6=2000
[/sheet[1]/row[7]] A7=West

=== Sheet: Notes ===
[/sheet[2]/row[2]] B2=Name\tC2=Score
[/sheet[2]/row[3]] B3=Ann\tC3=90
[/sheet[2]/row[4]] B4=Bob\tC4=75

=== Sheet: My Data ===
[/sheet[3]/row[1]] A1=42
[/sheet[3]/row[2]] B2=East

=== Sheet: Ärger ===
[/sheet[4]/row[1]] A1=ä";

const ANNOTATED: &str = "=== Sheet: Data ===
  A1: [Region] ← SharedString
  B1: [Sales] ← SharedString
  C1: [Status] ← SharedString
  E1: [Merged] ← SharedString
  A2: [West] ← SharedString
  B2: [1200] ← Number
  C2: [open] ← SharedString
  D2: [] ← Number ⚠ empty
  F2: [] ← Number ⚠ empty
  A3: [East] ← SharedString
  B3: [800] ← Number
  C3: [closed] ← SharedString
  D3: [TRUE] ← Boolean
  A4: [North] ← SharedString
  B4: [450.5] ← Number
  C4: [open] ← SharedString
  D4: [#DIV/0!] ← Error
  A5: [South] ← SharedString
  B5: [2000] ← Number
  C5: [closed] ← SharedString
  D5: [South!] ← =A5&\"!\"
  B6: [4450.5] ← =SUM(B2:B5)
  C6: [] ← =B6*2
  D6: [2000] ← =MAX(B2:B5)
  A7: [West] ← SharedString
=== Sheet: Notes ===
  B2: [Name] ← SharedString
  C2: [Score] ← SharedString
  B3: [Ann] ← SharedString
  C3: [90] ← Number
  B4: [Bob] ← SharedString
  C4: [75] ← Number
=== Sheet: My Data ===
  A1: [42] ← Number
  B2: [East] ← SharedString
=== Sheet: Ärger ===
  A1: [ä] ← SharedString";

const OUTLINE: &str = "├── \"Data\" (7 rows × 6 cols, 4 formula(s), 1 error cell(s), 1 table(s))
├── \"Notes\" (3 rows × 2 cols)
├── \"My Data\" (2 rows × 1 cols)
├── \"Ärger\" (1 rows × 1 cols)";

const STATS: &str = "Sheets: 4
Total Cells: 34
Empty Cells: 3
Formula Cells: 4
Error Cells: 1

Data Type Distribution:
  SharedString: 19
  Number: 12
  Boolean: 1
  Error: 1
  String: 1";

#[test]
fn modes_render_officecli_content() {
    for file in both(&kitchen()) {
        let options = Options::default();
        assert_eq!(view(&file, Mode::Text, &options), TEXT);
        assert_eq!(view(&file, Mode::Annotated, &options), ANNOTATED);
        assert_eq!(view(&file, Mode::Outline, &options), OUTLINE);
        assert_eq!(view(&file, Mode::Stats, &options), STATS);
    }
}

#[test]
fn text_options_window_the_rows() {
    let file = &both(&kitchen())[0];
    // `start` and `end` count row records per sheet.
    let options = Options {
        start: Some(2),
        end: Some(3),
        ..Options::default()
    };
    assert_eq!(
        view(file, Mode::Text, &options),
        "=== Sheet: Data ===
[/sheet[1]/row[2]] A2=West\tB2=1200\tC2=open
[/sheet[1]/row[3]] A3=East\tB3=800\tC3=closed\tD3=TRUE

=== Sheet: Notes ===
[/sheet[2]/row[3]] B3=Ann\tC3=90
[/sheet[2]/row[4]] B4=Bob\tC4=75

=== Sheet: My Data ===
[/sheet[3]/row[2]] B2=East

=== Sheet: Ärger ==="
    );
    // `max_lines` counts rows across sheets and ends with OfficeCLI's note.
    let options = Options {
        max_lines: Some(8),
        ..Options::default()
    };
    let text = view(file, Mode::Text, &options);
    assert!(text.starts_with(&TEXT[..TEXT.find("\n\n").expect("two sheets")]));
    assert!(
        text.ends_with(
            "=== Sheet: Notes ===
[/sheet[2]/row[2]] B2=Name\tC2=Score
... (showed 8 rows, 3 total in sheet, use --start/--end to view more)"
        ),
        "{text}"
    );
    // Columns and a range clip the cells; a range also picks the sheet,
    // without regard to case.
    let options = Options::default().with_columns("b,D").expect("letters");
    assert!(view(file, Mode::Text, &options).starts_with(
        "=== Sheet: Data ===
[/sheet[1]/row[1]] B1=Sales
[/sheet[1]/row[2]] B2=1200
[/sheet[1]/row[3]] B3=800\tD3=TRUE"
    ));
    let options = Options {
        range: Some(SheetRange::parse("NOTES!C3:B10").expect("a range")),
        ..Options::default()
    };
    assert_eq!(
        view(file, Mode::Text, &options),
        "=== Sheet: Notes ===
[/sheet[2]/row[3]] B3=Ann\tC3=90
[/sheet[2]/row[4]] B4=Bob\tC4=75"
    );
}

#[test]
fn annotated_options_window_the_rows() {
    let file = &both(&kitchen())[1];
    let options = Options {
        max_lines: Some(1),
        ..Options::default()
    };
    assert_eq!(
        view(file, Mode::Annotated, &options),
        "=== Sheet: Data ===
  A1: [Region] ← SharedString
  B1: [Sales] ← SharedString
  C1: [Status] ← SharedString
  E1: [Merged] ← SharedString
... (showed 1 rows, 7 total in sheet, use --start/--end to view more)"
    );
    let options = Options {
        start: Some(6),
        ..Options::default().with_columns("C").expect("a letter")
    };
    assert_eq!(
        view(file, Mode::Annotated, &options),
        "=== Sheet: Data ===
  C6: [] ← =B6*2
=== Sheet: Notes ===
=== Sheet: My Data ===
=== Sheet: Ärger ==="
    );
}

#[test]
fn views_serialize_as_officecli_json() {
    let file = &both(&kitchen())[0];
    let json = |mode: Mode, options: &Options| {
        serde_json::to_string(&view_json(file, mode, options).expect("renders"))
            .expect("serializes")
    };
    let options = Options {
        max_lines: Some(2),
        ..Options::default()
    };
    assert_eq!(
        json(Mode::Text, &options),
        r#"{"sheets":[{"name":"Data","rows":[{"row":1,"cells":{"A1":"Region","B1":"Sales","C1":"Status","E1":"Merged"}},{"row":2,"cells":{"A2":"West","B2":"1200","C2":"open"}}]}]}"#
    );
    assert_eq!(
        json(Mode::Outline, &Options::default()),
        r#"{"sheets":[{"name":"Data","rows":7,"cols":6,"formulas":4,"errorCells":1,"tables":1},{"name":"Notes","rows":3,"cols":2,"formulas":0,"errorCells":0,"tables":0},{"name":"My Data","rows":2,"cols":1,"formulas":0,"errorCells":0,"tables":0},{"name":"Ärger","rows":1,"cols":1,"formulas":0,"errorCells":0,"tables":0}]}"#
    );
    assert_eq!(
        json(Mode::Stats, &Options::default()),
        r#"{"sheets":4,"totalCells":34,"emptyCells":3,"formulaCells":4,"errorCells":1,"dataTypeDistribution":{"SharedString":19,"Number":12,"Boolean":1,"Error":1,"String":1}}"#
    );
    let annotated: serde_json::Value =
        serde_json::from_str(&json(Mode::Annotated, &Options::default())).expect("JSON");
    assert_eq!(annotated["view"], "annotated");
    assert_eq!(annotated["content"], ANNOTATED);
}

#[test]
fn unrendered_modes_are_refused() {
    let file = &both(&kitchen())[0];
    for mode in ["issues", "forms", "html", "svg", "screenshot", "pdf"] {
        let mode = Mode::parse(mode).expect("a mode");
        let mut out = String::new();
        let error = crate::format::xlsx::view_to(file, mode, &Options::default(), &mut out)
            .expect_err("not rendered");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
        assert!(out.is_empty());
        let error = view_json(file, mode, &Options::default()).expect_err("not rendered");
        assert_eq!(error.code(), Code::UnsupportedCapability, "{error}");
    }
}

/// A chart sheet has no rows: the outline names its kind, and the other
/// views skip it, as OfficeCLI's worksheet list does.
#[test]
fn chart_sheets_have_no_rows_to_view() {
    let file = open_path("poi/test-data/spreadsheet/chart_sheet.xlsx");
    let outline = view(&file, Mode::Outline, &Options::default());
    assert!(
        outline.contains("├── \"Chart1\" (chart sheet)"),
        "{outline}"
    );
    let text = view(&file, Mode::Text, &Options::default());
    assert!(!text.contains("Chart1"), "{text}");
    assert!(view(&file, Mode::Stats, &Options::default()).starts_with("Sheets: 3\n"));
}

/// A real fixture: `ExcelTables.xlsx` keeps row records without cells, which
/// the text view lists with an empty cell list.
#[test]
fn fixture_views() {
    let file = open_path("ooxml/xlsx/ExcelTables.xlsx");
    let text = view(&file, Mode::Text, &Options::default());
    let expected = [
        "=== Sheet: ExcelTable ===",
        "[/sheet[1]/row[1]] A1=Name\tB1=Age\tC1=AgeGroup\tG1=Name\tH1=Age\tI1=AgeGroup",
        "[/sheet[1]/row[2]] A2=Anton\tB2=44\tC2=4\tG2=Anton\tH2=44\tI2=4",
        "[/sheet[1]/row[3]] A3=Bertha\tB3=33\tC3=3\tG3=Bertha\tH3=33\tI3=3",
        "[/sheet[1]/row[4]] A4=Cäsar\tB4=21\tC4=2\tG4=Cäsar\tH4=21\tI4=2",
        "[/sheet[1]/row[5]] ",
        "[/sheet[1]/row[6]] ",
        "[/sheet[1]/row[7]] ",
        // The output's trailing white space is dropped.
        "[/sheet[1]/row[8]]",
    ];
    assert_eq!(text, expected.join("\n"));
    assert_eq!(
        view(&file, Mode::Outline, &Options::default()),
        "├── \"ExcelTable\" (8 rows × 9 cols, 6 formula(s), 1 table(s))"
    );
}
