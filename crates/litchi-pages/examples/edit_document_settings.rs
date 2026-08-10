//! Edit Pages document and footnote formatter settings in one transaction.

#![allow(
    clippy::print_stdout,
    reason = "this command-line example reports its committed semantic change"
)]

use std::error::Error;
use std::ffi::{OsStr, OsString};
use std::io::{self, Write as _};
use std::path::{Path, PathBuf};

use litchi_pages::{
    Package,
    document_options::Options,
    document_settings::Settings,
    footnote::{Format, Gap, Kind, Numbering, Settings as FootnoteSettings},
};
use tempfile::NamedTempFile;

const USAGE: &str = "usage: edit_document_settings <input.pages> <output.pages> \\
                     <unset|true|false:body> <unset|true|false:headers> \\
                     <unset|true|false:footers> <unset|true|false:facing-pages> \\
                     <unset|true|false:hyphenation> <unset|true|false:ligatures> \\
                     <unset|footnotes|document-endnotes|section-endnotes> \\
                     <unset|numeric|roman|symbolic|japanese-numeric|japanese-ideographic|arabic-numeric> \\
                     <unset|continuous|page|section> <unset|gap-points> [--inverse PATH]";

fn main() -> Result<(), Box<dyn Error>> {
    let mut arguments = std::env::args_os().skip(1);
    let input = PathBuf::from(required_argument(&mut arguments, "missing input path")?);
    let output = PathBuf::from(required_argument(&mut arguments, "missing output path")?);
    let settings = parse_settings(&mut arguments)?;
    let inverse_output = parse_inverse_output(&mut arguments)?;

    if input == output {
        return Err(invalid_input("input and output paths must differ"));
    }
    if inverse_output
        .as_deref()
        .is_some_and(|path| path == input || path == output)
    {
        return Err(invalid_input(
            "inverse path must differ from input and output paths",
        ));
    }

    let package = Package::open(&input)?;
    let commit = package.edit_document_settings()?.set(settings).commit()?;

    let inverse = inverse_output
        .as_ref()
        .map(|_| {
            commit
                .package()
                .apply_document_settings(&commit.patch().inverse())
        })
        .transpose()?;
    if inverse
        .as_ref()
        .is_some_and(|restored| restored.package().source_bytes() != package.source_bytes())
    {
        return Err(invalid_input(
            "inverse patch did not restore the exact input package",
        ));
    }

    save_new(&output, commit.package().source_bytes())?;
    if let (Some(path), Some(restored)) = (inverse_output, inverse) {
        save_new(&path, restored.package().source_bytes())?;
    }

    println!(
        "document settings: changed={}, touched_components={}, deleted_previews={}",
        commit.diagnostics().changed(),
        commit.diagnostics().touched_components(),
        commit.diagnostics().deleted_previews(),
    );
    Ok(())
}

fn parse_settings(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Settings, Box<dyn Error>> {
    let options = Options::new(
        parse_optional_bool(required_text(arguments, "missing body option")?, "body")?,
        parse_optional_bool(
            required_text(arguments, "missing headers option")?,
            "headers",
        )?,
        parse_optional_bool(
            required_text(arguments, "missing footers option")?,
            "footers",
        )?,
        parse_optional_bool(
            required_text(arguments, "missing facing-pages option")?,
            "facing-pages",
        )?,
        parse_optional_bool(
            required_text(arguments, "missing hyphenation option")?,
            "hyphenation",
        )?,
        parse_optional_bool(
            required_text(arguments, "missing ligatures option")?,
            "ligatures",
        )?,
    );
    let footnotes = FootnoteSettings {
        kind: parse_kind(required_text(arguments, "missing footnote kind")?)?,
        format: parse_format(required_text(arguments, "missing footnote format")?)?,
        numbering: parse_numbering(required_text(arguments, "missing footnote numbering")?)?,
        gap: parse_gap(required_text(arguments, "missing footnote gap")?)?,
    };
    Settings::new(options, footnotes).map_err(|error| -> Box<dyn Error> { Box::new(error) })
}

fn parse_optional_bool(value: String, label: &str) -> Result<Option<bool>, Box<dyn Error>> {
    match value.as_str() {
        "unset" => Ok(None),
        "true" => Ok(Some(true)),
        "false" => Ok(Some(false)),
        _ => Err(invalid_input(format!(
            "{label} must be unset, true, or false"
        ))),
    }
}

fn parse_kind(value: String) -> Result<Option<Kind>, Box<dyn Error>> {
    match value.as_str() {
        "unset" => Ok(None),
        "footnotes" => Ok(Some(Kind::Footnotes)),
        "document-endnotes" => Ok(Some(Kind::DocumentEndnotes)),
        "section-endnotes" => Ok(Some(Kind::SectionEndnotes)),
        _ => Err(invalid_input(
            "footnote kind must be unset, footnotes, document-endnotes, or section-endnotes",
        )),
    }
}

fn parse_format(value: String) -> Result<Option<Format>, Box<dyn Error>> {
    match value.as_str() {
        "unset" => Ok(None),
        "numeric" => Ok(Some(Format::Numeric)),
        "roman" => Ok(Some(Format::Roman)),
        "symbolic" => Ok(Some(Format::Symbolic)),
        "japanese-numeric" => Ok(Some(Format::JapaneseNumeric)),
        "japanese-ideographic" => Ok(Some(Format::JapaneseIdeographic)),
        "arabic-numeric" => Ok(Some(Format::ArabicNumeric)),
        _ => Err(invalid_input(
            "footnote format must be unset, numeric, roman, symbolic, japanese-numeric, japanese-ideographic, or arabic-numeric",
        )),
    }
}

fn parse_numbering(value: String) -> Result<Option<Numbering>, Box<dyn Error>> {
    match value.as_str() {
        "unset" => Ok(None),
        "continuous" => Ok(Some(Numbering::Continuous)),
        "page" => Ok(Some(Numbering::RestartEachPage)),
        "section" => Ok(Some(Numbering::RestartEachSection)),
        _ => Err(invalid_input(
            "footnote numbering must be unset, continuous, page, or section",
        )),
    }
}

fn parse_gap(value: String) -> Result<Option<Gap>, Box<dyn Error>> {
    if value == "unset" {
        return Ok(None);
    }
    let points = value
        .parse()
        .map_err(|_| invalid_input("footnote gap must be unset or a non-negative integer"))?;
    Gap::new(points)
        .map(Some)
        .map_err(|error| -> Box<dyn Error> { Box::new(error) })
}

fn required_argument(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<OsString, Box<dyn Error>> {
    arguments.next().ok_or_else(|| invalid_input(message))
}

fn required_text(
    arguments: &mut impl Iterator<Item = OsString>,
    message: &'static str,
) -> Result<String, Box<dyn Error>> {
    required_argument(arguments, message)?
        .into_string()
        .map_err(|_| invalid_input("settings arguments must be valid UTF-8"))
}

fn parse_inverse_output(
    arguments: &mut impl Iterator<Item = OsString>,
) -> Result<Option<PathBuf>, Box<dyn Error>> {
    let Some(flag) = arguments.next() else {
        return Ok(None);
    };
    if flag != OsStr::new("--inverse") {
        return Err(invalid_input(
            "unexpected trailing argument; expected --inverse PATH",
        ));
    }
    let path = PathBuf::from(required_argument(arguments, "missing --inverse path")?);
    if arguments.next().is_some() {
        return Err(invalid_input("unexpected trailing arguments"));
    }
    Ok(Some(path))
}

/// Publishes through a sibling temporary file without overwriting an existing target.
fn save_new(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let parent = path
        .parent()
        .filter(|parent| !parent.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));
    let mut temporary = NamedTempFile::new_in(parent)?;
    temporary.write_all(bytes)?;
    temporary.as_file().sync_all()?;
    temporary
        .persist_noclobber(path)
        .map_err(|error| Box::new(error.error))?;
    Ok(())
}

fn invalid_input(message: impl Into<String>) -> Box<dyn Error> {
    Box::new(io::Error::new(
        io::ErrorKind::InvalidInput,
        format!("{}\n\n{USAGE}", message.into()),
    ))
}
