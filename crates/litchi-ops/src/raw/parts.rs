//! OOXML part addressing (OfficeCLI's part aliases), `add_part` and
//! `raw_set`.

use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_opc::{OpcPackage, PackURI};

use super::package::{Container, part_not_found};
use super::{
    Action, AddedPart, CHART_CONTENT_TYPE, CHART_RELATIONSHIP, Edited, FOOTER_CONTENT_TYPE,
    FOOTER_RELATIONSHIP, HEADER_CONTENT_TYPE, HEADER_RELATIONSHIP, Kind, NewPart, is_xml,
    no_raw_write, with_format,
};
use crate::{Code, Error, File, Result};

/// The part a part name or OfficeCLI alias names.
pub(super) fn resolve(package: &OpcPackage, format: FileFormat, spec: &str) -> Result<PackURI> {
    if !spec.starts_with('/') {
        return Err(Error::new(Code::InvalidPath, "a part name starts with '/'")
            .with_suggestion("for example /word/document.xml or the alias /document"));
    }
    if spec == "/" {
        return Err(
            Error::new(Code::InvalidPath, "the package root is not a part")
                .with_suggestion("name a part, for example /document, /workbook or /presentation"),
        );
    }
    if let Ok(name) = PackURI::new(spec) {
        if let Some(part) = package.part_metadata(&name) {
            return Ok(part.partname().clone());
        }
    }
    let segments: Vec<&str> = spec[1..].split('/').collect();
    let found = match format {
        FileFormat::Docx => docx(package, &segments)?,
        FileFormat::Pptx => pptx(package, &segments)?,
        FileFormat::Xlsx => xlsx(package, &segments)?,
        _ => match segments.as_slice() {
            [name] if name.eq_ignore_ascii_case("workbook") => Some(main(package)?),
            _ => None,
        },
    };
    found.ok_or_else(part_not_found)
}

/// `name[N]` as `(name, Some(N))`, `name` as `(name, None)`.
fn segment(segment: &str) -> Option<(&str, Option<usize>)> {
    match segment.split_once('[') {
        None => Some((segment, None)),
        Some((name, rest)) => {
            let index: usize = rest.strip_suffix(']')?.trim().parse().ok()?;
            (index >= 1).then_some((name, Some(index)))
        },
    }
}

fn main(package: &OpcPackage) -> Result<PackURI> {
    Ok(package.main_document_part()?.partname().clone())
}

/// The internal relationships of `source` whose type ends with
/// `/<suffix>`, ordered by where the source XML first mentions their ids
/// (document order: slide ids, sheets, header references).
fn related(package: &OpcPackage, source: &PackURI, suffix: &str) -> Result<Vec<(String, PackURI)>> {
    let part = package.get_part(source)?;
    let xml = part.blob();
    let mut found: Vec<(usize, String, PackURI)> = Vec::new();
    for relationship in part.rels().iter() {
        if relationship.is_external()
            || !relationship
                .reltype()
                .rsplit('/')
                .next()
                .is_some_and(|kind| kind.eq_ignore_ascii_case(suffix))
        {
            continue;
        }
        let Ok(target) = relationship.target_partname() else {
            continue;
        };
        let id = relationship.r_id();
        let quoted = |quote: u8| {
            let mut needle = Vec::with_capacity(id.len() + 2);
            needle.push(quote);
            needle.extend_from_slice(id.as_bytes());
            needle.push(quote);
            memchr::memmem::find(xml, &needle)
        };
        let at = quoted(b'"')
            .into_iter()
            .chain(quoted(b'\''))
            .min()
            .unwrap_or(usize::MAX);
        found.push((at, id.to_owned(), target));
    }
    found.sort_by(|left, right| {
        left.0
            .cmp(&right.0)
            .then_with(|| natural(&left.1, &right.1))
    });
    Ok(found
        .into_iter()
        .map(|(_, id, target)| (id, target))
        .collect())
}

/// `rId2` before `rId10`.
fn natural(left: &str, right: &str) -> std::cmp::Ordering {
    let digits = |value: &str| {
        value
            .trim_start_matches(|c: char| !c.is_ascii_digit())
            .parse::<u64>()
            .unwrap_or(u64::MAX)
    };
    digits(left)
        .cmp(&digits(right))
        .then_with(|| left.cmp(right))
}

fn nth(parts: Vec<(String, PackURI)>, index: Option<usize>) -> Option<PackURI> {
    parts
        .into_iter()
        .nth(index.unwrap_or(1) - 1)
        .map(|(_, target)| target)
}

/// DOCX aliases: `/document`, the main part's single resources, and
/// `/header[N]`, `/footer[N]`, `/chart[N]`.
fn docx(package: &OpcPackage, segments: &[&str]) -> Result<Option<PackURI>> {
    let [only] = segments else {
        return Ok(None);
    };
    let Some((name, index)) = segment(only) else {
        return Ok(None);
    };
    let main = main(package)?;
    if name.eq_ignore_ascii_case("document") && index.is_none() {
        return Ok(Some(main));
    }
    const RESOURCES: [&str; 12] = [
        "styles",
        "numbering",
        "settings",
        "fontTable",
        "theme",
        "footnotes",
        "endnotes",
        "comments",
        "webSettings",
        "header",
        "footer",
        "chart",
    ];
    let Some(kind) = RESOURCES
        .iter()
        .find(|kind| name.eq_ignore_ascii_case(kind))
    else {
        return Ok(None);
    };
    Ok(nth(related(package, &main, kind)?, index))
}

/// PPTX aliases: `/presentation`, `/slide[N]`, `/slideMaster[N]`,
/// `/slideLayout[N]` (counted across masters), `/theme`, `/notesMaster`,
/// `/noteSlide[N]`, and `/slide[N]/<kind>[K]` for the slide's related parts.
fn pptx(package: &OpcPackage, segments: &[&str]) -> Result<Option<PackURI>> {
    let main = main(package)?;
    let parsed: Option<Vec<(&str, Option<usize>)>> = segments.iter().map(|s| segment(s)).collect();
    let Some(parsed) = parsed else {
        return Ok(None);
    };
    let slide = |index: Option<usize>| -> Result<Option<PackURI>> {
        Ok(nth(related(package, &main, "slide")?, index))
    };
    match parsed.as_slice() {
        [(name, None)] if name.eq_ignore_ascii_case("presentation") => Ok(Some(main)),
        [(name, index)] if name.eq_ignore_ascii_case("slide") => slide(*index),
        [(name, index)] if name.eq_ignore_ascii_case("slideMaster") => {
            Ok(nth(related(package, &main, "slideMaster")?, *index))
        },
        [(name, index)] if name.eq_ignore_ascii_case("slideLayout") => {
            let mut layouts = Vec::new();
            for (_, master) in related(package, &main, "slideMaster")? {
                layouts.extend(related(package, &master, "slideLayout")?);
            }
            Ok(nth(layouts, *index))
        },
        [(name, index)] if name.eq_ignore_ascii_case("theme") => {
            Ok(nth(related(package, &main, "theme")?, *index))
        },
        [(name, None)] if name.eq_ignore_ascii_case("notesMaster") => {
            Ok(nth(related(package, &main, "notesMaster")?, None))
        },
        [(name, index)]
            if name.eq_ignore_ascii_case("noteSlide")
                || name.eq_ignore_ascii_case("notesSlide") =>
        {
            let Some(slide) = slide(*index)? else {
                return Ok(None);
            };
            Ok(nth(related(package, &slide, "notesSlide")?, None))
        },
        [(name, index), (kind, sub)] if name.eq_ignore_ascii_case("slide") => {
            let Some(slide) = slide(*index)? else {
                return Ok(None);
            };
            Ok(nth(related(package, &slide, kind)?, *sub))
        },
        _ => Ok(None),
    }
}

/// XLSX aliases: `/workbook`, `/styles`, `/sharedStrings`, `/theme`, a sheet
/// by name or `/sheet[N]`, and `/<sheet>/drawing`, `/<sheet>/chart[K]`
/// (charts of the sheet's drawing), `/<sheet>/<kind>[K]` for the sheet's
/// other related parts.
fn xlsx(package: &OpcPackage, segments: &[&str]) -> Result<Option<PackURI>> {
    let main = main(package)?;
    let Some(first) = segments.first() else {
        return Ok(None);
    };
    let single =
        |kind: &str| -> Result<Option<PackURI>> { Ok(nth(related(package, &main, kind)?, None)) };
    if segments.len() == 1 {
        for (alias, kind) in [
            ("workbook", None),
            ("styles", Some("styles")),
            ("sharedStrings", Some("sharedStrings")),
            ("theme", Some("theme")),
        ] {
            if first.eq_ignore_ascii_case(alias) {
                return match kind {
                    None => Ok(Some(main)),
                    Some(kind) => single(kind),
                };
            }
        }
    }
    let Some(sheet) = sheet(package, &main, first)? else {
        return Ok(None);
    };
    match segments {
        [_] => Ok(Some(sheet)),
        [_, rest] => {
            let Some((kind, index)) = segment(rest) else {
                return Ok(None);
            };
            if kind.eq_ignore_ascii_case("chart") {
                let Some(drawing) = nth(related(package, &sheet, "drawing")?, None) else {
                    return Ok(None);
                };
                return Ok(nth(related(package, &drawing, "chart")?, index));
            }
            Ok(nth(related(package, &sheet, kind)?, index))
        },
        _ => Ok(None),
    }
}

/// The worksheet part `name` (a sheet name, ASCII and Unicode caseless, or
/// `sheet[N]`) names.
fn sheet(package: &OpcPackage, main: &PackURI, name: &str) -> Result<Option<PackURI>> {
    let sheets = sheets(package, main)?;
    if let Some((_, Some(index))) =
        segment(name).filter(|(kind, _)| kind.eq_ignore_ascii_case("sheet"))
    {
        return Ok(sheets.into_iter().nth(index - 1).map(|(_, target)| target));
    }
    let wanted = caseless::default_case_fold_str(name);
    Ok(sheets
        .into_iter()
        .find(|(sheet, _)| caseless::default_case_fold_str(sheet) == wanted)
        .map(|(_, target)| target))
}

/// The workbook's sheets in tab order: name and worksheet part.
fn sheets(package: &OpcPackage, main: &PackURI) -> Result<Vec<(String, PackURI)>> {
    use quick_xml::events::Event;
    let part = package.get_part(main)?;
    let mut reader = quick_xml::Reader::from_reader(part.blob());
    let mut out = Vec::new();
    loop {
        let event = reader.read_event().map_err(|_| {
            Error::new(
                Code::CorruptFile,
                "the workbook part is not well-formed XML",
            )
        })?;
        match event {
            Event::Start(start) | Event::Empty(start)
                if start.local_name().as_ref() == b"sheet" =>
            {
                let mut name = None;
                let mut id = None;
                let mut attributes = start.attributes();
                attributes.with_checks(false);
                for attribute in attributes.flatten() {
                    let key = attribute.key;
                    if key.as_ref() == b"name" {
                        name = attribute
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .ok()
                            .map(|value| value.into_owned());
                    } else if key.prefix().is_some() && key.local_name().as_ref() == b"id" {
                        id = attribute
                            .normalized_value(quick_xml::XmlVersion::Implicit1_0)
                            .ok()
                            .map(|value| value.into_owned());
                    }
                }
                if let (Some(name), Some(id)) = (name, id) {
                    if let Some(relationship) = part.rels().get(&id) {
                        if let Ok(target) = relationship.target_partname() {
                            out.push((name, target));
                        }
                    }
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(out)
}

fn writable(file: &File) -> Result<()> {
    match file {
        File::Signed(_) => Err(with_format(
            file,
            Error::new(
                Code::DocumentProtected,
                "a signed document cannot be edited",
            ),
        )
        .with_suggestion("consume the signature with `Signed::unsign` before editing")),
        File::Locked(locked) => Err(locked.read_error()),
        _ => Ok(()),
    }
}

/// The published OPC package of `file`, or a format-named refusal.
fn opc(file: &File, operation: &str) -> Result<(FileFormat, OpcPackage)> {
    writable(file)?;
    match Container::of(file) {
        Ok(Container::Opc { format, package }) => Ok((format, *package)),
        #[cfg(any(feature = "odt", feature = "ods", feature = "odp", odf_objects))]
        Ok(Container::Odf { .. }) => Err(no_raw_write(file, operation)),
        Err(error) if error.code() == Code::UnsupportedCapability => {
            Err(no_raw_write(file, operation))
        },
        Err(error) => Err(error),
    }
}

/// Publish `package` and open the result as the new handle.
fn reopen(package: &OpcPackage) -> Result<File> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    crate::open_bytes(bytes)
}

pub(super) fn add_part(file: &File, parent: &str, part: NewPart) -> Result<AddedPart> {
    let (format, mut package) = opc(file, "adding a part to")?;
    let owner = if parent == "/" {
        None
    } else {
        let mut owner = resolve(&package, format, parent)?;
        // OfficeCLI relates a sheet's chart from the sheet's drawing.
        if format == FileFormat::Xlsx && matches!(part.kind, Kind::Chart) {
            let content_type = package.get_part(&owner)?.content_type().to_owned();
            if content_type.ends_with("worksheet+xml") {
                owner = nth(related(&package, &owner, "drawing")?, None).ok_or_else(|| {
                    Error::new(Code::UnsupportedCapability, "the sheet has no drawing part")
                        .with_suggestion("add a chart or picture to the sheet first, then add-part relates the new chart from its drawing")
                })?;
            }
        }
        Some(owner)
    };
    let (content_type, relationship, bytes, template) = match part.kind {
        Kind::Chart => {
            let template = match format {
                FileFormat::Docx => "/word/charts/chart%d.xml",
                FileFormat::Xlsx => "/xl/charts/chart%d.xml",
                FileFormat::Pptx => "/ppt/charts/chart%d.xml",
                _ => return Err(unsupported_type(file)),
            };
            (
                CHART_CONTENT_TYPE.to_owned(),
                CHART_RELATIONSHIP.to_owned(),
                EMPTY_CHART.as_bytes().to_vec(),
                Some(template),
            )
        },
        Kind::Header | Kind::Footer if format != FileFormat::Docx => {
            return Err(unsupported_type(file));
        },
        Kind::Header => (
            HEADER_CONTENT_TYPE.to_owned(),
            HEADER_RELATIONSHIP.to_owned(),
            EMPTY_HEADER.as_bytes().to_vec(),
            Some("/word/header%d.xml"),
        ),
        Kind::Footer => (
            FOOTER_CONTENT_TYPE.to_owned(),
            FOOTER_RELATIONSHIP.to_owned(),
            EMPTY_FOOTER.as_bytes().to_vec(),
            Some("/word/footer%d.xml"),
        ),
        Kind::Custom {
            content_type,
            relationship_type,
            bytes,
        } => (content_type, relationship_type, bytes, None),
    };
    let name = match (part.name, template) {
        (Some(name), _) => PackURI::new(name).map_err(|_| {
            Error::new(
                Code::InvalidPath,
                "the part name is not a valid OPC part name",
            )
        })?,
        (None, Some(template)) => package.next_partname(template)?,
        (None, None) => {
            return Err(
                Error::new(Code::MissingArgument, "a custom part needs a name")
                    .with_suggestion("name it with `NewPart::named`"),
            );
        },
    };
    if package.part_metadata(&name).is_some() {
        return Err(Error::new(
            Code::FileExists,
            "the package already has a part of that name",
        ));
    }
    package.validate_new_part_name(&name)?;
    if content_type.trim().is_empty() || relationship.trim().is_empty() {
        return Err(Error::new(
            Code::MissingArgument,
            "a part needs a content type and a relationship type",
        ));
    }
    if is_xml(name.as_str(), &content_type) {
        package.try_add_owned_xml_part_bytes(name.clone(), content_type, Arc::new(bytes))?;
    } else {
        package.try_add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            content_type,
            bytes,
        )))?;
    }
    let relationship_id = match &owner {
        None => package.relate_to(&name.relative_ref("/"), &relationship),
        Some(owner) => {
            let target = name.relative_ref(owner.base_uri());
            package
                .get_part_mut(owner)?
                .relate_to(&target, &relationship)
        },
    };
    Ok(AddedPart {
        file: reopen(&package)?,
        part: name.as_str().to_owned(),
        relationship_id,
    })
}

fn unsupported_type(file: &File) -> Error {
    with_format(
        file,
        Error::new(
            Code::UnsupportedType,
            "this part type does not apply to this format",
        ),
    )
    .with_suggestion("chart applies to DOCX, XLSX and PPTX; header and footer to DOCX")
}

pub(super) fn raw_set(
    file: &File,
    part: &str,
    xpath: &str,
    action: Action,
    xml: Option<&str>,
) -> Result<Edited> {
    let (format, mut package) = opc(file, "raw XML edits of")?;
    let name = resolve(&package, format, part)?;
    let (original, content_type) = {
        let part = package.get_part(&name)?;
        (part.blob_arc(), part.content_type().to_owned())
    };
    if !is_xml(name.as_str(), &content_type) {
        return Err(with_format(
            file,
            Error::new(Code::UnsupportedCapability, "raw-set edits XML parts only"),
        ));
    }
    let (edited, matched) = super::xpath::edit(&original, xpath, action, xml)?;
    if edited.as_slice() == original.as_slice() {
        return Ok(Edited {
            file: file.clone(),
            matched,
        });
    }
    package.try_replace_owned_xml_part_bytes(&name, &original, Arc::new(edited))?;
    Ok(Edited {
        file: reopen(&package)?,
        matched,
    })
}

const EMPTY_CHART: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#;

const EMPTY_HEADER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p/></w:hdr>"#;

const EMPTY_FOOTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p/></w:ftr>"#;
