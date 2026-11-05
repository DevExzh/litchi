//! Format names on the command line (`--type`, `--format`, file extensions).
//!
//! `litchi_core::FileFormat` has no public name or extension table, so the
//! CLI keeps this one for argument parsing only (record 0256 follow-up F2).

use litchi_core::FileFormat;
use litchi_ops::{Code, Error, Result};

/// Names and extensions, ASCII case-insensitive, with OfficeCLI's optional
/// leading dot accepted by [`parse`].
const NAMES: &[(&str, FileFormat)] = &[
    ("doc", FileFormat::Doc),
    ("docx", FileFormat::Docx),
    ("ppt", FileFormat::Ppt),
    ("pptx", FileFormat::Pptx),
    ("xls", FileFormat::Xls),
    ("xlsx", FileFormat::Xlsx),
    ("xlsb", FileFormat::Xlsb),
    ("rtf", FileFormat::Rtf),
    ("pages", FileFormat::Pages),
    ("key", FileFormat::Keynote),
    ("keynote", FileFormat::Keynote),
    ("numbers", FileFormat::Numbers),
    ("odt", FileFormat::Odt),
    ("ods", FileFormat::Ods),
    ("odp", FileFormat::Odp),
    ("odg", FileFormat::Odg),
    ("odc", FileFormat::Odc),
    ("odf", FileFormat::Odf),
    ("odi", FileFormat::Odi),
    ("odm", FileFormat::Odm),
    ("oth", FileFormat::Oth),
    ("odb", FileFormat::Odb),
];

/// The format a name or extension (`docx`, `.docx`) denotes.
pub(crate) fn parse(name: &str) -> Result<FileFormat> {
    let bare = name.strip_prefix('.').unwrap_or(name);
    NAMES
        .iter()
        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(bare))
        .map(|&(_, format)| format)
        .ok_or_else(|| {
            Error::new(
                Code::InvalidArgument,
                format!("unknown document type `{name}`"),
            )
            .with_suggestion(
                "use one of doc, docx, ppt, pptx, xls, xlsx, xlsb, rtf, odt, ods, odp, odg, \
                     odc, odf, odi, odm, oth, odb, pages, key, numbers",
            )
        })
}

/// The format a file path's extension denotes, if any.
pub(crate) fn of_path(path: &std::path::Path) -> Option<FileFormat> {
    path.extension()
        .and_then(|extension| extension.to_str())
        .and_then(|extension| parse(extension).ok())
}

/// The canonical name of `format`, for messages and JSON.
pub(crate) fn name(format: FileFormat) -> &'static str {
    NAMES
        .iter()
        .find(|&&(_, candidate)| candidate == format)
        .map_or("unknown", |&(name, _)| name)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_round_trip() {
        for &(name, format) in NAMES {
            assert_eq!(parse(name).unwrap(), format);
            assert_eq!(parse(&format!(".{}", name.to_uppercase())).unwrap(), format);
            assert_eq!(parse(super::name(format)).unwrap(), format);
        }
        assert_eq!(parse("txt").unwrap_err().code(), Code::InvalidArgument);
        assert_eq!(
            of_path(std::path::Path::new("a/b.XLSX")),
            Some(FileFormat::Xlsx)
        );
        assert_eq!(of_path(std::path::Path::new("noext")), None);
    }
}
