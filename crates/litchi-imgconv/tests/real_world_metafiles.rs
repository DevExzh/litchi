//! Conversion outcomes for metafiles carved out of real legacy documents.
//!
//! Every fixture under `test-data/metafiles/` is converted through the public
//! entry points a document reader uses. Each conversion must finish without
//! panicking and, when it succeeds, produce well-formed output. The outcome of
//! every operation is compared with `test-data/metafiles/expectations.tsv`, so
//! a change in what the library accepts, approximates or rejects is visible.
//!
//! After an intentional behaviour change, regenerate the expectations with
//! `LITCHI_UPDATE_METAFILE_EXPECTATIONS=1 cargo test -p litchi-imgconv --test
//! real_world_metafiles` and review the diff.

use std::{
    collections::BTreeMap,
    fmt::Write as _,
    fs,
    panic::{AssertUnwindSafe, catch_unwind},
    path::{Path, PathBuf},
};

use litchi_core::error::Error;
use litchi_imgconv::{ConvertedFormat, InputFormat, Options, OutputFormat, convert_metafile};
use litchi_odraw::image::{Blip, Limits as ArtLimits};
use quick_xml::{Reader, events::Event};

const EXPECTATIONS: &str = "expectations.tsv";
const RASTER_WIDTH: u32 = 96;

fn root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/metafiles")
}

#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum Kind {
    Blip,
    Emf,
    Pict,
    Wmf,
}

impl Kind {
    fn of(path: &Path) -> Option<Self> {
        match path.extension()?.to_str()? {
            "emf" => Some(Self::Emf),
            "wmf" => Some(Self::Wmf),
            "pict" => Some(Self::Pict),
            "blip" => Some(Self::Blip),
            _ => None,
        }
    }
}

fn fixtures() -> Vec<(String, Kind, Vec<u8>)> {
    let root = root();
    let mut found = Vec::new();
    for directory in ["blip", "emf", "pict", "wmf"] {
        let entries = fs::read_dir(root.join(directory)).unwrap_or_else(|error| {
            panic!("failed to read fixture directory {directory}: {error}")
        });
        for entry in entries {
            let path = entry
                .unwrap_or_else(|error| panic!("failed to enumerate {directory}: {error}"))
                .path();
            let Some(kind) = Kind::of(&path) else {
                continue;
            };
            let name = format!(
                "{directory}/{}",
                path.file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or_else(|| panic!("fixture name is not UTF-8: {}", path.display()))
            );
            let bytes = fs::read(&path)
                .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
            found.push((name, kind, bytes));
        }
    }
    found.sort_by(|left, right| left.0.cmp(&right.0));
    found
}

fn error_class(error: &Error) -> &'static str {
    match error {
        Error::Unsupported(_) => "unsupported",
        Error::ParseError(_) => "parse",
        Error::InvalidFormat(_) => "invalid",
        _ => "other",
    }
}

/// Runs one operation, returning its outcome class and a failure detail when
/// the operation produced invalid output.
fn outcome<F>(operation: F) -> Result<String, String>
where
    F: FnOnce() -> Result<String, String>,
{
    match catch_unwind(AssertUnwindSafe(operation)) {
        Ok(result) => result,
        Err(payload) => {
            let message = payload
                .downcast_ref::<String>()
                .cloned()
                .or_else(|| {
                    payload
                        .downcast_ref::<&str>()
                        .map(|text| (*text).to_string())
                })
                .unwrap_or_default();
            Err(format!("panicked: {message}"))
        },
    }
}

fn check_svg(svg: &str) -> Result<(), String> {
    // XML 1.0 forbids most control characters and U+FFFE/U+FFFF in a document;
    // the event reader below does not check for them.
    if let Some(character) = svg.chars().find(|character| {
        !matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
    }) {
        return Err(format!(
            "SVG contains the character U+{:04X}, which XML forbids",
            u32::from(character)
        ));
    }
    let mut reader = Reader::from_str(svg);
    let mut depth = 0usize;
    let mut roots = 0usize;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if depth == 0 {
                    roots += 1;
                    if start.name().as_ref() != b"svg" {
                        return Err("root element is not <svg>".into());
                    }
                }
                depth += 1;
            },
            Ok(Event::Empty(element)) => {
                if depth == 0 {
                    roots += 1;
                    if element.name().as_ref() != b"svg" {
                        return Err("root element is not <svg>".into());
                    }
                }
            },
            Ok(Event::End(_)) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| "unbalanced closing tag".to_string())?;
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(format!("SVG is not well-formed XML: {error}")),
        }
    }
    if depth != 0 || roots != 1 {
        return Err(format!(
            "SVG has {roots} roots and {depth} unclosed elements"
        ));
    }
    resvg::usvg::Tree::from_str(svg, &resvg::usvg::Options::default())
        .map_err(|error| format!("SVG is rejected by usvg: {error}"))?;
    Ok(())
}

/// Classifies a raster output as `ok` or `ok-blank` after checking it decodes.
fn check_raster(bytes: &[u8], width: Option<u32>) -> Result<String, String> {
    let image = image::load_from_memory(bytes)
        .map_err(|error| format!("raster output does not decode: {error}"))?;
    if let Some(width) = width
        && image.width() != width
    {
        return Err(format!(
            "raster output is {} pixels wide, requested {width}",
            image.width()
        ));
    }
    let blank = image.to_rgba8().pixels().all(|pixel| {
        let [red, green, blue, alpha] = pixel.0;
        alpha == 0 || (red >= 250 && green >= 250 && blue >= 250)
    });
    Ok(if blank { "ok-blank" } else { "ok" }.to_string())
}

fn lossy(report: &litchi_imgconv::ConversionReport) -> bool {
    report.diagnostics.iter().any(|diagnostic| {
        !(diagnostic.code.starts_with("noncanonical-")
            || diagnostic.code.starts_with("auto-")
            || diagnostic.code == "comment")
    })
}

fn metafile_outcomes(
    bytes: &[u8],
    input: InputFormat,
) -> Vec<(&'static str, Result<String, String>)> {
    let svg =
        outcome(
            || match convert_metafile(bytes, input, OutputFormat::Svg, Options::default()) {
                Ok(converted) => {
                    if converted.format != ConvertedFormat::Svg {
                        return Err("SVG request produced another format".into());
                    }
                    let text = std::str::from_utf8(&converted.bytes)
                        .map_err(|_| "SVG output is not UTF-8".to_string())?;
                    check_svg(text)?;
                    Ok(if lossy(&converted.report) {
                        "approx"
                    } else {
                        "ok"
                    }
                    .to_string())
                },
                Err(error) => Ok(error_class(&error).to_string()),
            },
        );
    let png = outcome(|| {
        let options = Options::default().width(RASTER_WIDTH);
        match convert_metafile(bytes, input, OutputFormat::Png, options) {
            Ok(converted) => check_raster(&converted.bytes, Some(RASTER_WIDTH)),
            Err(error) => Ok(error_class(&error).to_string()),
        }
    });
    let strict = outcome(|| {
        let result = match input {
            InputFormat::Emf => litchi_imgconv::emf::convert_emf_to_svg(bytes),
            InputFormat::Wmf => litchi_imgconv::wmf::convert_wmf_to_svg(bytes),
            InputFormat::Pict => litchi_imgconv::pict::convert_pict_to_svg(bytes),
        };
        match result {
            Ok(svg) => {
                check_svg(&svg)?;
                Ok("ok".to_string())
            },
            Err(error) => Ok(error_class(&error).to_string()),
        }
    });
    vec![("svg", svg), ("png", png), ("strict-svg", strict)]
}

fn blip_outcomes(bytes: &[u8]) -> Vec<(&'static str, Result<String, String>)> {
    let parsed = catch_unwind(AssertUnwindSafe(|| {
        Blip::parse(bytes, ArtLimits::default())
    }));
    let blip = match parsed {
        Ok(Ok(blip)) => blip,
        Ok(Err(_)) => return vec![("parse", Ok("rejected".to_string()))],
        Err(_) => return vec![("parse", Err("panicked".to_string()))],
    };
    let decode =
        outcome(
            || match litchi_imgconv::decode_data(&blip, &litchi_imgconv::Limits::default()) {
                Ok(_) => Ok("ok".to_string()),
                Err(error) => Ok(error_class(&error).to_string()),
            },
        );
    let png =
        outcome(
            || match litchi_imgconv::to_png(&blip, Options::default().width(RASTER_WIDTH)) {
                Ok(png) => check_raster(&png, Some(RASTER_WIDTH)),
                Err(error) => Ok(error_class(&error).to_string()),
            },
        );
    let svg = outcome(|| match litchi_imgconv::to_svg(&blip, Options::default()) {
        Ok(svg) => {
            check_svg(&svg)?;
            Ok("ok".to_string())
        },
        Err(error) => Ok(error_class(&error).to_string()),
    });
    vec![
        ("parse", Ok("ok".to_string())),
        ("decode", decode),
        ("png", png),
        ("svg", svg),
    ]
}

fn read_expectations() -> BTreeMap<String, String> {
    let text = fs::read_to_string(root().join(EXPECTATIONS)).unwrap_or_default();
    text.lines()
        .filter(|line| !line.starts_with('#') && !line.is_empty())
        .filter_map(|line| line.split_once('\t'))
        .map(|(name, outcomes)| (name.to_string(), outcomes.to_string()))
        .collect()
}

#[test]
fn real_world_metafiles_convert_as_expected() {
    let fixtures = fixtures();
    assert!(
        fixtures.len() >= 500,
        "the real-world metafile corpus unexpectedly shrank to {}",
        fixtures.len()
    );
    let expected = read_expectations();
    let mut actual = BTreeMap::new();
    let mut failures = Vec::new();
    for (name, kind, bytes) in &fixtures {
        let results = match kind {
            Kind::Emf => metafile_outcomes(bytes, InputFormat::Emf),
            Kind::Wmf => metafile_outcomes(bytes, InputFormat::Wmf),
            Kind::Pict => metafile_outcomes(bytes, InputFormat::Pict),
            Kind::Blip => blip_outcomes(bytes),
        };
        let mut summary = String::new();
        for (operation, result) in results {
            let class = match result {
                Ok(class) => class,
                Err(detail) => {
                    failures.push(format!("{name}: {operation}: {detail}"));
                    "invalid-output".to_string()
                },
            };
            if !summary.is_empty() {
                summary.push(' ');
            }
            let _ = write!(summary, "{operation}={class}");
        }
        actual.insert(name.clone(), summary);
    }

    if std::env::var_os("LITCHI_UPDATE_METAFILE_EXPECTATIONS").is_some() {
        let mut text = String::from(
            "# Expected conversion outcomes; see crates/litchi-imgconv/tests/real_world_metafiles.rs\n",
        );
        for (name, summary) in &actual {
            let _ = writeln!(text, "{name}\t{summary}");
        }
        fs::write(root().join(EXPECTATIONS), text)
            .unwrap_or_else(|error| panic!("failed to write expectations: {error}"));
    } else {
        for (name, summary) in &actual {
            match expected.get(name) {
                Some(want) if want == summary => {},
                Some(want) => failures.push(format!("{name}: expected `{want}`, got `{summary}`")),
                None => failures.push(format!("{name}: no expectation recorded (got `{summary}`)")),
            }
        }
        for name in expected.keys() {
            if !actual.contains_key(name) {
                failures.push(format!("{name}: expectation has no fixture"));
            }
        }
    }
    assert!(
        failures.is_empty(),
        "{} real-world metafile checks failed:\n{}",
        failures.len(),
        failures.join("\n")
    );
}

#[test]
fn minimized_fuzz_inputs_convert_or_fail_without_panicking() {
    let directory = root().join("fuzz");
    let mut checked = 0usize;
    let mut failures = Vec::new();
    let entries = fs::read_dir(&directory)
        .unwrap_or_else(|error| panic!("failed to read {}: {error}", directory.display()));
    for entry in entries {
        let path = entry
            .unwrap_or_else(|error| panic!("failed to enumerate fuzz inputs: {error}"))
            .path();
        let Some(kind) = Kind::of(&path) else {
            continue;
        };
        let bytes = fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        let results = match kind {
            Kind::Emf => metafile_outcomes(&bytes, InputFormat::Emf),
            Kind::Wmf => metafile_outcomes(&bytes, InputFormat::Wmf),
            Kind::Pict => metafile_outcomes(&bytes, InputFormat::Pict),
            Kind::Blip => blip_outcomes(&bytes),
        };
        for (operation, result) in results {
            if let Err(detail) = result {
                failures.push(format!("{}: {operation}: {detail}", path.display()));
            }
        }
        checked += 1;
    }
    assert!(checked >= 9, "only {checked} fuzz inputs were found");
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
