use std::env;
use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use xml_minifier::audit::{
    Limits, ReplacementError, ReplacementProof, verify_source, verify_source_replacement,
};

const PAIRS: &[(&str, &str, &str)] = &[
    (
        "medium-one-percent",
        "xl/workbook.xml",
        "medium/one-percent",
    ),
    (
        "medium-one-percent",
        "xl/worksheets/sheet1.xml",
        "medium/one-percent",
    ),
    (
        "medium-one-percent",
        "xl/worksheets/sheet2.xml",
        "medium/one-percent",
    ),
    (
        "medium-one-percent",
        "xl/worksheets/sheet3.xml",
        "medium/one-percent",
    ),
    (
        "medium-one-percent",
        "xl/worksheets/sheet4.xml",
        "medium/one-percent",
    ),
    ("medium-one-edit", "xl/workbook.xml", "medium/one-edit"),
    (
        "medium-one-edit",
        "xl/worksheets/sheet1.xml",
        "medium/one-edit",
    ),
    (
        "dense-sparse-one-percent",
        "xl/workbook.xml",
        "dense-sparse/one-percent",
    ),
    (
        "dense-sparse-one-percent",
        "xl/worksheets/sheet1.xml",
        "dense-sparse/one-percent",
    ),
    (
        "dense-sparse-one-percent",
        "xl/worksheets/sheet2.xml",
        "dense-sparse/one-percent",
    ),
    (
        "dense-sparse-one-percent",
        "xl/worksheets/sheet3.xml",
        "dense-sparse/one-percent",
    ),
    (
        "dense-sparse-one-percent",
        "xl/worksheets/sheet4.xml",
        "dense-sparse/one-percent",
    ),
];

fn usage() -> ! {
    eprintln!("usage: litchi-proof-diagnostic <0904-artifacts> <output.tsv>");
    std::process::exit(2);
}

fn path_pair(root: &Path, scenario: &str, member: &str, side: &str) -> PathBuf {
    root.join(scenario).join(side).join(member)
}

fn print_error(error: &ReplacementError) -> String {
    match error {
        ReplacementError::Original(error) => format!("original:{error:?}"),
        ReplacementError::Replacement(error) => format!("replacement:{error:?}"),
        _ => "unknown-replacement-error".to_string(),
    }
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut args = env::args_os().skip(1);
    let root = args.next().map(PathBuf::from).unwrap_or_else(|| usage());
    let output = args.next().map(PathBuf::from).unwrap_or_else(|| usage());
    if args.next().is_some() {
        usage();
    }
    let limits = Limits::default();
    let mut out = String::from(
        "case\tmember\toriginal_bytes\treplacement_bytes\toriginal_report\treplacement_report\tproof\n",
    );
    for &(case, member, scenario) in PAIRS {
        let original = fs::read(path_pair(&root, scenario, member, "original"))?;
        let replacement = fs::read(path_pair(&root, scenario, member, "replacement"))?;
        let original_report = verify_source(&original, limits);
        let replacement_report = verify_source(&replacement, limits);
        let original_summary = match original_report {
            Ok(report) => format!(
                "ok,attributes={},bytes={},events={},max_depth={},text_bytes={}",
                report.attributes(),
                report.bytes(),
                report.events(),
                report.max_depth(),
                report.text_bytes()
            ),
            Err(error) => format!("error,{error:?}"),
        };
        let replacement_summary = match replacement_report {
            Ok(report) => format!(
                "ok,attributes={},bytes={},events={},max_depth={},text_bytes={}",
                report.attributes(),
                report.bytes(),
                report.events(),
                report.max_depth(),
                report.text_bytes()
            ),
            Err(error) => format!("error,{error:?}"),
        };
        let proof = match verify_source_replacement(&original, &replacement, limits) {
            Ok(ReplacementProof::Identical) => "Identical".to_string(),
            Ok(ReplacementProof::Complete) => "Complete".to_string(),
            Ok(ReplacementProof::Window {
                original,
                replacement,
            }) => format!("Window,{original:?},{replacement:?}"),
            Err(error) => format!("Error,{}", print_error(&error)),
            Ok(_) => "UnknownProof".to_string(),
        };
        out.push_str(&format!(
            "{case}\t{member}\t{}\t{}\t{original_summary}\t{replacement_summary}\t{proof}\n",
            original.len(),
            replacement.len()
        ));
    }
    let mut file = fs::OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&output)
        .map_err(|error| {
            format!(
                "refusing existing or inaccessible output {}: {error}",
                output.display()
            )
        })?;
    file.write_all(out.as_bytes())?;
    Ok(())
}
