//! Small downstream-style native ODS acceptance runner for the sheet metadata owner.
//!
//! Usage:
//!
//! ```text
//! native_roundtrip <candidate-worktree> <native-template.ods> <output-root>
//! ```
//!
//! The runner copies a real native Calc package, adds the four metadata owners
//! to its native `content.xml` shape, asks LibreOffice to open/save it twice
//! under a private profile, then reopens the final bytes through the public
//! `litchi-ods` facade and reports semantic survivors.

use std::{
    env,
    error::Error,
    fs,
    fmt,
    io::{self, Cursor, Read, Write},
    path::{Path, PathBuf},
    process::Command,
};

use litchi_ods::{
    Spreadsheet,
    sheet_metadata::{
        CellRange, CellSelector, Detective, Direction, HighlightedRange, LabelRange, Operation,
        OperationKind, Options, Orientation,
    },
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

type AnyResult<T> = Result<T, Box<dyn Error>>;

fn native_content(template: &Path) -> AnyResult<String> {
    let bytes = fs::read(template)?;
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut content = String::new();
    archive.by_name("content.xml")?.read_to_string(&mut content)?;
    // The source-publication seam intentionally accepts compact authored XML.
    // The native fixture has one formatting newline after the XML declaration;
    // remove that separator while retaining every native root child and byte
    // inside the document body.
    content = content.replace("?>\n<", "?><");

    let labels = concat!(
        "<table:label-ranges>",
        "<table:label-range table:label-cell-range-address=\"sheet1.A1:sheet1.A2\" ",
        "table:data-cell-range-address=\"sheet1.B1:sheet1.B2\" table:orientation=\"row\"/>",
        "</table:label-ranges>"
    );
    let table_marker = "<table:table table:name=\"sheet1\"";
    let table_start = content
        .find(table_marker)
        .ok_or("native template first table marker is missing")?;
    content.insert_str(table_start, labels);

    let cell_open = "<table:table-cell office:value-type=\"string\" calcext:value-type=\"string\">";
    let cell_metadata = concat!(
        "<table:cell-range-source table:name=\"NativeImport\" ",
        "table:last-column-spanned=\"2\" table:last-row-spanned=\"2\" ",
        "xlink:type=\"simple\" xlink:href=\"source.ods#sheet1.A1:sheet1.B2\"/>",
        "<table:detective>",
        "<table:highlighted-range table:cell-range-address=\"sheet1.B1:sheet1.B2\" ",
        "table:direction=\"from-same-table\" table:contains-error=\"false\"/>",
        "<table:operation table:name=\"trace-precedents\" table:index=\"0\"/>",
        "</table:detective>"
    );
    let cell_start = content
        .find(cell_open)
        .ok_or("native template first string cell marker is missing")?
        + cell_open.len();
    content.insert_str(cell_start, cell_metadata);

    let consolidation = concat!(
        "<table:consolidation table:function=\"sum\" ",
        "table:source-cell-range-addresses=\"sheet1.A1:sheet1.A2 sheet1.B1:sheet1.B2\" ",
        "table:target-cell-address=\"sheet1.C1\" table:use-labels=\"both\"/>",
    );
    let spreadsheet_end = "</office:spreadsheet>";
    let end = content
        .find(spreadsheet_end)
        .ok_or("native template spreadsheet end marker is missing")?;
    content.insert_str(end, consolidation);
    Ok(content)
}

fn package_with_content(template: &Path, content: &str) -> AnyResult<Vec<u8>> {
    let bytes = fs::read(template)?;
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut output = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut output);
    let options =
        SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);

    for index in 0..archive.len() {
        let mut entry = archive.by_index(index)?;
        let name = entry.name().to_owned();
        if entry.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }
        writer.start_file(name.as_str(), options)?;
        if name == "content.xml" {
            writer.write_all(content.as_bytes())?;
        } else {
            io::copy(&mut entry, &mut writer)?;
        }
    }
    writer.finish()?;
    Ok(output.into_inner())
}

fn run_libreoffice(profile: &Path, input: &Path, output_dir: &Path) -> AnyResult<()> {
    let profile_uri = format!("file://{}", profile.display());
    let status = Command::new("/usr/bin/libreoffice")
        .arg(format!("-env:UserInstallation={profile_uri}"))
        .args(["--headless", "--convert-to", "ods", "--outdir"])
        .arg(output_dir)
        .arg(input)
        .status()?;
    if !status.success() {
        return Err(format!("LibreOffice exited with {status}").into());
    }
    Ok(())
}

fn content_xml(bytes: &[u8]) -> AnyResult<String> {
    let mut archive = ZipArchive::new(Cursor::new(bytes))?;
    let mut content = String::new();
    archive.by_name("content.xml")?.read_to_string(&mut content)?;
    Ok(content)
}

fn print_snapshot(label: &str, bytes: &[u8]) -> AnyResult<()> {
    let content = content_xml(bytes)?;
    let spreadsheet = Spreadsheet::from_bytes(bytes.to_vec())?;
    let snapshot = spreadsheet.sheet_metadata()?;
    let consolidation = snapshot.consolidation()?;
    let labels = snapshot.label_ranges();
    let cell = snapshot.cell_metadata(CellSelector::by_name("sheet1", 0, 0))?;
    let source = cell.as_ref().and_then(|value| value.range_source());
    let detective = cell.as_ref().and_then(|value| value.detective());
    println!(
        "{label}:bytes={} content_bytes={} consolidation={} label_ranges={} cell_range_source={} detective={} raw_tokens={{consolidation:{},label_range:{},cell_range_source:{},detective:{}}}",
        bytes.len(),
        content.len(),
        consolidation.is_some(),
        labels.len(),
        source.is_some(),
        detective.is_some(),
        content.matches("table:consolidation").count(),
        content.matches("table:label-range").count(),
        content.matches("table:cell-range-source").count(),
        content.matches("table:detective").count(),
    );
    println!("{label}:typed.consolidation={consolidation:?}");
    println!(
        "{label}:typed.label_ranges.present={} values={:?}",
        labels.is_present(),
        labels.as_slice()
    );
    println!("{label}:typed.cell_metadata={cell:?}");
    println!("{label}:typed.cell_range_source={source:?}");
    if let Some(value) = source {
        println!(
            "{label}:typed.cell_range_source.scalars={{name:{:?},href:{:?},rows:{},columns:{},actuate_on_request:{},filter_name:{:?},filter_options:{:?},refresh_delay:{:?}}}",
            value.name(),
            value.href(),
            value.rows(),
            value.columns(),
            value.actuate_on_request(),
            value.filter_name(),
            value.filter_options(),
            value.refresh_delay(),
        );
    }
    println!("{label}:typed.detective={detective:?}");
    if let Some(value) = detective {
        for (index, range) in value.highlighted_ranges().iter().enumerate() {
            println!(
                "{label}:typed.detective.highlighted[{index}]={{address:{:?},direction:{:?},contains_error:{:?},marked_invalid:{:?}}}",
                range.cell_range_address(),
                range.direction(),
                range.contains_error(),
                range.marked_invalid(),
            );
        }
        for (index, operation) in value.operations().iter().enumerate() {
            println!(
                "{label}:typed.detective.operation[{index}]={{kind:{:?},index:{}}}",
                operation.kind, operation.index
            );
        }
    }
    Ok(())
}

fn report_check<T: fmt::Debug>(
    label: &str,
    field: &str,
    expected: &T,
    observed: &T,
    matches: bool,
) {
    println!(
        "{label}:expected.{field} expected={expected:?} observed={observed:?} status={}",
        if matches { "preserved" } else { "changed" }
    );
}

fn expected_detective() -> AnyResult<Detective> {
    let mut detective = Detective::new();
    detective.add_highlighted_range(HighlightedRange::valid(
        Some("sheet1.C1:sheet1.C2".to_owned()),
        Direction::FromSameTable,
        Some(true),
    )?);
    detective.add_operation(Operation::new(OperationKind::TraceErrors, 0));
    Ok(detective)
}

/// Compare the API-mutated semantic values with the expected state.  Native
/// round trips deliberately keep running when a field changes, so the output
/// records compatibility loss instead of turning a valid native-save result
/// into an opaque process failure.
fn compare_expected(label: &str, bytes: &[u8], expect_range_source: bool) -> AnyResult<()> {
    let spreadsheet = Spreadsheet::from_bytes(bytes.to_vec())?;
    let snapshot = spreadsheet.sheet_metadata()?;
    let observed_consolidation = snapshot.consolidation()?.cloned();
    let expected_consolidation = Some(Options {
        function: "average".to_owned(),
        source_cell_range_addresses: vec![
            "sheet1.A1:sheet1.A2".to_owned(),
            "sheet1.B1:sheet1.B2".to_owned(),
        ],
        target_cell_address: "sheet1.C1".to_owned(),
        use_labels: Some(litchi_ods::sheet_metadata::UseLabels::Both),
        link_to_source_data: None,
    });
    report_check(
        label,
        "consolidation",
        &expected_consolidation,
        &observed_consolidation,
        observed_consolidation == expected_consolidation,
    );
    let expected_function = Some("average".to_owned());
    let observed_function = observed_consolidation
        .as_ref()
        .map(|value| value.function.clone());
    report_check(
        label,
        "consolidation.function",
        &expected_function,
        &observed_function,
        observed_function == expected_function,
    );
    let expected_sources = expected_consolidation
        .as_ref()
        .map(|value| value.source_cell_range_addresses.clone());
    let observed_sources = observed_consolidation
        .as_ref()
        .map(|value| value.source_cell_range_addresses.clone());
    report_check(
        label,
        "consolidation.source_cell_range_addresses",
        &expected_sources,
        &observed_sources,
        observed_sources == expected_sources,
    );
    let expected_target = Some("sheet1.C1".to_owned());
    let observed_target = observed_consolidation
        .as_ref()
        .map(|value| value.target_cell_address.clone());
    report_check(
        label,
        "consolidation.target_cell_address",
        &expected_target,
        &observed_target,
        observed_target == expected_target,
    );
    let expected_use_labels = Some(litchi_ods::sheet_metadata::UseLabels::Both);
    let observed_use_labels = observed_consolidation
        .as_ref()
        .and_then(|value| value.use_labels);
    report_check(
        label,
        "consolidation.use_labels",
        &expected_use_labels,
        &observed_use_labels,
        observed_use_labels == expected_use_labels,
    );

    let expected_labels = vec![LabelRange::new(
        "sheet1.A1:sheet1.A2",
        "sheet1.B1:sheet1.B2",
        Orientation::Column,
    )?];
    let labels = snapshot.label_ranges();
    let observed_labels = labels.as_slice().to_vec();
    report_check(
        label,
        "label_ranges",
        &expected_labels,
        &observed_labels,
        labels.is_present() && observed_labels == expected_labels,
    );
    let expected_orientation = Some(Orientation::Column);
    let observed_orientation = observed_labels.first().map(|value| value.orientation);
    report_check(
        label,
        "label_ranges[0].orientation",
        &expected_orientation,
        &observed_orientation,
        observed_orientation == expected_orientation,
    );

    let expected_source = if expect_range_source {
        Some(CellRange::new(
            "NativeImportChanged",
            "source.ods#sheet1.A1:sheet1.B2",
            2,
            2,
        )?)
    } else {
        None
    };
    let cell = snapshot.cell_metadata(CellSelector::by_name("sheet1", 0, 0))?;
    let observed_source = cell
        .as_ref()
        .and_then(|value| value.range_source())
        .cloned();
    report_check(
        label,
        "cell_range_source",
        &expected_source,
        &observed_source,
        observed_source == expected_source,
    );
    let expected_source_name = expected_source.as_ref().map(|value| value.name().to_owned());
    let observed_source_name = observed_source
        .as_ref()
        .map(|value| value.name().to_owned());
    report_check(
        label,
        "cell_range_source.name",
        &expected_source_name,
        &observed_source_name,
        observed_source_name == expected_source_name,
    );
    let expected_source_href = expected_source.as_ref().map(|value| value.href().to_owned());
    let observed_source_href = observed_source
        .as_ref()
        .map(|value| value.href().to_owned());
    report_check(
        label,
        "cell_range_source.href",
        &expected_source_href,
        &observed_source_href,
        observed_source_href == expected_source_href,
    );
    let expected_source_dimensions = expected_source
        .as_ref()
        .map(|value| (value.rows(), value.columns()));
    let observed_source_dimensions = observed_source
        .as_ref()
        .map(|value| (value.rows(), value.columns()));
    report_check(
        label,
        "cell_range_source.dimensions",
        &expected_source_dimensions,
        &observed_source_dimensions,
        observed_source_dimensions == expected_source_dimensions,
    );

    let expected_detective = Some(expected_detective()?);
    let observed_detective = cell
        .as_ref()
        .and_then(|value| value.detective())
        .cloned();
    report_check(
        label,
        "detective",
        &expected_detective,
        &observed_detective,
        observed_detective == expected_detective,
    );
    let expected_highlights = expected_detective
        .as_ref()
        .map(|value| value.highlighted_ranges().to_vec());
    let observed_highlights = observed_detective
        .as_ref()
        .map(|value| value.highlighted_ranges().to_vec());
    report_check(
        label,
        "detective.highlighted_ranges",
        &expected_highlights,
        &observed_highlights,
        observed_highlights == expected_highlights,
    );
    let expected_operations = expected_detective
        .as_ref()
        .map(|value| value.operations().to_vec());
    let observed_operations = observed_detective
        .as_ref()
        .map(|value| value.operations().to_vec());
    report_check(
        label,
        "detective.operations",
        &expected_operations,
        &observed_operations,
        observed_operations == expected_operations,
    );
    let all_match = observed_consolidation == expected_consolidation
        && labels.is_present()
        && observed_labels == expected_labels
        && observed_source == expected_source
        && observed_detective == expected_detective;
    println!(
        "{label}:typed_expected_overall={}",
        if all_match { "preserved" } else { "changed" }
    );
    Ok(())
}

fn main() -> AnyResult<()> {
    let mut args = env::args_os().skip(1);
    let candidate = fs::canonicalize(PathBuf::from(
        args.next().ok_or("missing candidate worktree")?,
    ))?;
    let template = fs::canonicalize(PathBuf::from(
        args.next().ok_or("missing native template")?,
    ))?;
    let output_root = PathBuf::from(args.next().ok_or("missing output root")?);
    fs::create_dir_all(&output_root)?;
    let output_root = fs::canonicalize(output_root)?;
    let save1 = output_root.join("save1");
    let save2 = output_root.join("save2");
    let profile = output_root.join("libreoffice-profile");
    fs::create_dir_all(&save1)?;
    fs::create_dir_all(&save2)?;
    fs::create_dir_all(&profile)?;

    if !candidate.join("crates/litchi-ods").is_dir() {
        return Err("candidate worktree does not contain crates/litchi-ods".into());
    }
    let content = native_content(&template)?;
    let input = output_root.join("ods_metadata_native_input.ods");
    let input_bytes = package_with_content(&template, &content)?;
    fs::write(&input, &input_bytes)?;
    print_snapshot("injected_input", &input_bytes)?;

    let mut spreadsheet = Spreadsheet::from_bytes(input_bytes)?;
    spreadsheet.edit_sheet_metadata(|edit| {
        let mut consolidation = Options::new(
            "average",
            vec![
                "sheet1.A1:sheet1.A2".to_owned(),
                "sheet1.B1:sheet1.B2".to_owned(),
            ],
            "sheet1.C1",
        )?;
        consolidation.use_labels = Some(litchi_ods::sheet_metadata::UseLabels::Both);
        edit.set_consolidation(Some(consolidation))?;
        edit.replace_label_range(
            0,
            LabelRange::new(
                "sheet1.A1:sheet1.A2",
                "sheet1.B1:sheet1.B2",
                Orientation::Column,
            )?,
        )?;
        edit.set_cell_range_source(
            CellSelector::by_name("sheet1", 0, 0),
            CellRange::new(
                "NativeImportChanged",
                "source.ods#sheet1.A1:sheet1.B2",
                2,
                2,
            )?,
        )?;
        let mut detective = Detective::new();
        detective.add_highlighted_range(HighlightedRange::valid(
            Some("sheet1.C1:sheet1.C2".to_owned()),
            Direction::FromSameTable,
            Some(true),
        )?);
        detective.add_operation(Operation::new(OperationKind::TraceErrors, 0));
        edit.set_detective(CellSelector::by_name("sheet1", 0, 0), detective)
    })?;
    let api_changed_bytes = spreadsheet.into_bytes();
    print_snapshot("api_changed_input", &api_changed_bytes)?;
    compare_expected("api_changed_input", &api_changed_bytes, true)?;
    fs::write(&input, &api_changed_bytes)?;

    run_libreoffice(&profile, &input, &save1)?;
    let save1_path = save1.join("ods_metadata_native_input.ods");
    let save1_bytes = fs::read(&save1_path)?;
    print_snapshot("after_libreoffice_save1", &save1_bytes)?;
    compare_expected("after_libreoffice_save1", &save1_bytes, true)?;

    run_libreoffice(&profile, &save1_path, &save2)?;
    let save2_path = save2.join("ods_metadata_native_input.ods");
    let save2_bytes = fs::read(&save2_path)?;
    print_snapshot("after_libreoffice_save2_reopen", &save2_bytes)?;
    compare_expected("after_libreoffice_save2_reopen", &save2_bytes, true)?;

    println!("input_path={}", input.display());
    println!("save1_path={}", save1_path.display());
    println!("save2_path={}", save2_path.display());
    println!("profile_path={}", profile.display());
    Ok(())
}
