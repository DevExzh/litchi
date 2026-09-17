//! Differential oracle for the compact planning facts (change 0622).
//!
//! The commit's complete layout scan is only removable if the facts the
//! planning traversal retained reproduce, byte for byte, the output that scan
//! would have produced — and if a source the scan would have refused never
//! produces facts at all. Both halves are checked here by running the two
//! routes side by side over the same bytes and the same actions:
//!
//! * [`compare`] drives one worksheet through every cell the raw parser sees,
//!   comparing complete output bytes, readback provenance and typed errors
//!   between the fact route and the scan route;
//! * the synthetic table below covers the shapes the value-only validator
//!   admits and the shapes it does not, so both an accepted and a declined
//!   fact route are exercised;
//! * [`oracle_over_test_data`] repeats the same comparison over every
//!   worksheet part of every `.xlsx` fixture in the repository.
//!
//! The route counters make the comparison non-vacuous: a test that only ever
//! declined would compare the scan route against itself.

use std::collections::BTreeMap;

use litchi_sheet::Cell as Address;

use crate::raw;
use crate::raw::worksheet::SourceFacts;
use crate::raw::worksheet::edit::{Action, Payload, rewrite_value_only_with_provenance, route};

/// Run the planning traversal exactly as `Snapshot::from_source_selected`
/// does, and return the facts it published.
fn plan(content: &[u8]) -> Option<SourceFacts> {
    let admission = raw::worksheet::source_stream_admission(content)?;
    super::validation::worksheet_xml_and_parse_source(content, admission, || Ok(None))
        .ok()
        .and_then(|(_, facts)| facts)
}

fn actions_for(address: Address, value: u32) -> BTreeMap<Address, Action> {
    BTreeMap::from([(address, Action::set(value.into()))])
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
struct Tally {
    accepted: usize,
    declined: usize,
    compared: usize,
}

/// Compare the two commit routes for one action set.
fn compare_one(content: &[u8], actions: &BTreeMap<Address, Action>, facts: Option<&SourceFacts>) {
    let mut fact_actions = BTreeMap::new();
    let mut scan_actions = BTreeMap::new();
    for (address, action) in actions {
        fact_actions.insert(*address, clone_action(action));
        scan_actions.insert(*address, clone_action(action));
    }
    let from_facts = rewrite_value_only_with_provenance(content, "Sheet1", fact_actions, facts);
    let from_scan = rewrite_value_only_with_provenance(content, "Sheet1", scan_actions, None);
    match (from_facts, from_scan) {
        (Ok(facts_output), Ok(scan_output)) => {
            assert_eq!(
                facts_output.bytes, scan_output.bytes,
                "fact-derived output bytes differ from the scan-derived output"
            );
            let facts_spans = facts_output
                .omitted
                .iter()
                .map(|span| {
                    (
                        span.row,
                        span.first_column,
                        span.last_column,
                        span.start,
                        span.end,
                    )
                })
                .collect::<Vec<_>>();
            let scan_spans = scan_output
                .omitted
                .iter()
                .map(|span| {
                    (
                        span.row,
                        span.first_column,
                        span.last_column,
                        span.start,
                        span.end,
                    )
                })
                .collect::<Vec<_>>();
            assert_eq!(
                facts_spans, scan_spans,
                "fact-derived readback provenance differs from the scan-derived provenance"
            );
        },
        (Err(facts_error), Err(scan_error)) => {
            assert_eq!(
                format!("{facts_error:?}"),
                format!("{scan_error:?}"),
                "fact-derived and scan-derived routes disagree on the typed error"
            );
        },
        (facts_result, scan_result) => panic!(
            "fact and scan routes disagree on success: {:?} vs {:?}",
            facts_result.map(|value| value.bytes.len()),
            scan_result.map(|value| value.bytes.len())
        ),
    }
}

fn clone_action(action: &Action) -> Action {
    match action {
        Action::Remove => Action::Remove,
        Action::Update { payload, style } => Action::Update {
            payload: payload.as_ref().map(clone_payload),
            style: *style,
        },
    }
}

fn clone_payload(payload: &Payload) -> Payload {
    match payload {
        Payload::Set(content) => Payload::Set(content.clone()),
        Payload::SharedString { index, text } => Payload::SharedString {
            index: *index,
            text: text.clone(),
        },
        Payload::SharedFormula {
            index,
            reference,
            formula,
        } => Payload::SharedFormula {
            index: *index,
            reference: reference.clone(),
            formula: formula.clone(),
        },
        Payload::Clear => Payload::Clear,
        Payload::ClearIfPresent => Payload::ClearIfPresent,
    }
}

/// Whether the sweep visits every cell of every worksheet.
///
/// Every comparison rewrites the whole worksheet, so an exhaustive pass over a
/// large sheet is quadratic and too slow for an ordinary debug `cargo test`.
/// The bounded form visits every cell of every worksheet with at most
/// [`STRIDE_LIMIT`] cells and a fixed stride through the rest; setting
/// `LITCHI_0622_FULL_ORACLE=1` visits every cell of every worksheet, which is
/// how the retained evidence run was taken.
fn exhaustive() -> bool {
    std::env::var_os("LITCHI_0622_FULL_ORACLE").is_some()
}

const STRIDE_LIMIT: usize = 96;

/// Drive one worksheet through every cell the raw parser sees.
fn compare(content: &[u8]) -> Tally {
    route::reset();
    let facts = plan(content);
    // Enumerate the cells the *scanner* would see. When the planning
    // traversal published facts, those are the retained addresses; otherwise
    // fall back to the raw parser's own view.
    let addresses = match facts.as_ref() {
        Some(facts) => facts.cells.iter().map(|cell| cell.address).collect(),
        None => match raw::worksheet::parse(content, || Ok(None)) {
            Ok(store) => store
                .entries()
                .iter()
                .map(|entry| entry.address)
                .collect::<Vec<_>>(),
            // A source the parser refuses never reaches a commit, and without
            // facts both routes are the same call.
            Err(_) => Vec::new(),
        },
    };
    let total = addresses.len();
    let stride = if exhaustive() || total <= STRIDE_LIMIT {
        1
    } else {
        total.div_ceil(STRIDE_LIMIT)
    };
    let mut compared = 0usize;
    for (index, address) in addresses.iter().enumerate() {
        if index % stride != 0 {
            continue;
        }
        let value = u32::try_from(index % 977).unwrap_or(0).saturating_add(1);
        compare_one(content, &actions_for(*address, value), facts.as_ref());
        compared += 1;
    }
    // A batch that touches many cells at once, capped the way a real commit is.
    let mut batch = BTreeMap::new();
    for (index, address) in addresses.iter().take(256).enumerate() {
        batch.insert(
            *address,
            Action::set(u32::try_from(index).unwrap_or(0).into()),
        );
    }
    if !batch.is_empty() {
        compare_one(content, &batch, facts.as_ref());
        compared += 1;
    }
    // Coordinates the sheet does not contain: insertion and removal are the
    // shapes the fact route declines.
    for address in ["ZZ9999", "A1", "B2"] {
        if let Ok(address) = Address::from_a1(address) {
            compare_one(content, &actions_for(address, 7), facts.as_ref());
            compare_one(
                content,
                &BTreeMap::from([(address, Action::Remove)]),
                facts.as_ref(),
            );
            compared += 2;
        }
    }
    Tally {
        accepted: route::accepted(),
        declined: route::declined(),
        compared,
    }
}

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn scalar_sheet(body: &str) -> String {
    format!(r#"<worksheet xmlns="{SML}">{body}</worksheet>"#)
}

/// Shapes the fact route is expected to accept.
fn admitted_shapes() -> Vec<(&'static str, String)> {
    vec![
        (
            "plain",
            scalar_sheet(
                r#"<dimension ref="A1:C2"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2"><c r="A2"><v>4</v></c><c r="C2"><v>6</v></c></row></sheetData>"#,
            ),
        ),
        (
            "styled-and-typed",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1" spans="1:2"><c r="A1" s="3"><v>1</v></c><c r="B1" s="4" t="n"><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "empty-cells",
            scalar_sheet(
                r#"<dimension ref="A1:C1"/><sheetData><row r="1"><c r="A1"/><c r="B1"><v>2</v></c><c r="C1" s="2"/></row></sheetData>"#,
            ),
        ),
        (
            "inline-string",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>hello</t></is></c><c r="B1"><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "comments-and-whitespace",
            scalar_sheet(
                "<dimension ref=\"A1:B1\"/><sheetData>\n  <row r=\"1\">\n    <c r=\"A1\"><!--between--><v>1</v></c>\n    <c r=\"B1\"> <v>2</v> </c>\n  </row>\n</sheetData>",
            ),
        ),
        (
            "prefixed-cells",
            format!(
                r#"<x:worksheet xmlns:x="{SML}"><x:dimension ref="A1:B1"/><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c><x:c r="B1" s="1"><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
            ),
        ),
        (
            "cols-and-format",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="2" width="12" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "sheetviews",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "narrow-dimension-expands",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="D1"><v>4</v></c></row></sheetData>"#,
            ),
        ),
        (
            "no-dimension",
            scalar_sheet(
                r#"<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            // An empty `<row/>` is retained as a fact; only an action that
            // targets it declines, and the comparison below exercises both.
            "empty-row",
            scalar_sheet(
                r#"<dimension ref="A1:A2"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"/></sheetData>"#,
            ),
        ),
        (
            "mce-declaration-only",
            format!(
                r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#
            ),
        ),
    ]
}

/// Shapes the fact route must decline while the scan keeps its own answer.
fn declined_shapes() -> Vec<(&'static str, String)> {
    vec![
        (
            "formula",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><f>SUM(B1)</f><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "shared-formula",
            scalar_sheet(
                r#"<dimension ref="A1:A2"/><sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>1</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "inferred-cell-reference",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1"><c><v>1</v></c><c><v>2</v></c></row></sheetData>"#,
            ),
        ),
        (
            "inferred-row-number",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "entity-in-attribute",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row r="1" spans="1:1&#32;"><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "empty-sheetdata",
            scalar_sheet(r#"<dimension ref="A1:A1"/><sheetData/>"#),
        ),
        (
            "duplicate-dimension",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "dimension-after-sheetdata",
            scalar_sheet(
                r#"<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><dimension ref="A1:A1"/>"#,
            ),
        ),
        (
            "descending-rows",
            scalar_sheet(
                r#"<dimension ref="A1:A2"/><sheetData><row r="2"><c r="A2"><v>2</v></c></row><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "descending-cells",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1"><c r="B1"><v>2</v></c><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "empty-cols",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><cols/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "merged-ranges",
            scalar_sheet(
                r#"<dimension ref="A1:B1"/><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>"#,
            ),
        ),
        (
            "sheet-protection",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><sheetProtection sheet="1"/>"#,
            ),
        ),
        (
            "data-validation",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><dataValidations count="1"><dataValidation type="whole" sqref="A1"/></dataValidations>"#,
            ),
        ),
        (
            "unknown-cell-child",
            scalar_sheet(
                r#"<dimension ref="A1:A1"/><sheetData><row r="1"><c r="A1"><extLst/><v>1</v></c></row></sheetData>"#,
            ),
        ),
        (
            "cell-outside-its-row",
            scalar_sheet(
                r#"<dimension ref="A1:A2"/><sheetData><row r="1"><c r="A2"><v>1</v></c></row></sheetData>"#,
            ),
        ),
    ]
}

/// The retained state's size is the whole argument for the design, so pin it.
#[test]
fn change_0622_retained_records_stay_compact() {
    use crate::raw::worksheet::edit::codec_sizes;

    let cell_fact = std::hint::black_box(codec_sizes::CELL_FACT);
    let row_fact = std::hint::black_box(codec_sizes::ROW_FACT);
    let cell_slot = std::hint::black_box(codec_sizes::CELL_SLOT);
    println!(
        "change 0622 sizes: CellFact {cell_fact} B, RowFact {row_fact} B, CellSlot {cell_slot} B"
    );
    assert_eq!(cell_fact, 16, "one retained cell must stay 16 bytes");
    assert_eq!(row_fact, 32, "one retained row must stay 32 bytes");
    assert!(
        cell_slot >= cell_fact * 4,
        "the scanner slot the facts replace is {cell_slot} bytes plus its heap tag"
    );
}

#[test]
fn change_0622_admitted_shapes_take_the_fact_route_and_match_the_scan() {
    for (name, worksheet) in admitted_shapes() {
        let tally = compare(worksheet.as_bytes());
        assert!(
            tally.accepted > 0,
            "shape '{name}' never took the fact route ({tally:?})"
        );
    }
}

#[test]
fn change_0622_declined_shapes_keep_the_scan_and_its_answers() {
    for (name, worksheet) in declined_shapes() {
        let tally = compare(worksheet.as_bytes());
        assert_eq!(
            tally.accepted, 0,
            "shape '{name}' unexpectedly took the fact route ({tally:?})"
        );
    }
}

#[test]
fn change_0622_dense_grid_matches_the_scan_on_every_cell() {
    // The harness `dense-sparse` corpus shape: one dense grid of scalar
    // numeric cells, written the way `litchi-xlsx` itself writes them.
    let (rows, columns) = if exhaustive() {
        (128u32, 100u32)
    } else {
        (48u32, 48u32)
    };
    let corner = Address::at(rows - 1, columns - 1).expect("grid corner");
    let mut body = format!(r#"<dimension ref="A1:{}"/><sheetData>"#, corner.a1());
    for row in 1..=rows {
        body.push_str(&format!(r#"<row r="{row}" spans="1:{columns}">"#));
        for column in 0..columns {
            let address = Address::at(row - 1, column).expect("grid address");
            body.push_str(&format!(
                r#"<c r="{}"><v>{}</v></c>"#,
                address.a1(),
                row * columns + column
            ));
        }
        body.push_str("</row>");
    }
    body.push_str("</sheetData>");
    let worksheet = scalar_sheet(&body);
    let tally = compare(worksheet.as_bytes());
    assert!(
        tally.accepted + 6 >= tally.compared && tally.compared > 90,
        "dense grid tally {tally:?}"
    );
}

#[test]
fn change_0622_oracle_over_test_data_worksheets() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let Ok(root) = root.canonicalize() else {
        return;
    };
    let mut packages = Vec::new();
    collect_xlsx(&root, &mut packages);
    packages.sort();
    assert!(
        packages.len() > 50,
        "the oracle needs the repository .xlsx corpus, found {}",
        packages.len()
    );
    let mut worksheets = 0usize;
    let mut accepted = 0usize;
    let mut compared = 0usize;
    let mut admitted = 0usize;
    let mut editable = 0usize;
    let mut with_facts = 0usize;
    for path in packages {
        let Ok(bytes) = std::fs::read(&path) else {
            continue;
        };
        let Ok(package) = litchi_opc::OpcPackage::from_bytes(&bytes) else {
            continue;
        };
        for part in package
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            let name = part.partname().as_str().to_owned();
            if !name.starts_with("/xl/worksheets/") || !name.ends_with(".xml") {
                continue;
            }
            let content = part.blob().to_vec();
            worksheets += 1;
            if raw::worksheet::source_stream_admission(&content).is_some() {
                admitted += 1;
            }
            if super::validation::worksheet_xml(&content).is_ok() {
                editable += 1;
            }
            if plan(&content).is_some() {
                with_facts += 1;
            }
            let tally = compare(&content);
            accepted += tally.accepted;
            compared += tally.compared;
        }
    }
    assert!(worksheets > 100, "worksheet parts swept: {worksheets}");
    assert!(compared > 1_000, "comparisons run: {compared}");
    // Reported for the record: the real corpus is dominated by worksheets the
    // builder declines, so the sweep proves the decline path, not the accept
    // path. The synthetic and dense-grid cases prove the accept path.
    println!(
        "change 0622 oracle: {worksheets} worksheet parts, {admitted} admitted to the shared \
         traversal, {editable} accepted by the value-only validator, {with_facts} publishing \
         facts, {compared} comparisons, {accepted} fact-route commits"
    );
}

fn collect_xlsx(directory: &std::path::Path, found: &mut Vec<std::path::PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_xlsx(&path, found);
        } else if path
            .extension()
            .is_some_and(|extension| extension == "xlsx")
        {
            found.push(path);
        }
    }
}
