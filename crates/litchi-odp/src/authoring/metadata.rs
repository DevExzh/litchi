//! Retained-source `meta.xml` edits for the primary ODP transaction.
//!
//! The common ODF patcher owns five simple Dublin Core and ODF fields
//! (title, creator, subject, description, keywords); every other element of
//! the source `meta.xml` keeps its bytes. A request that changes a field the
//! patcher does not own is refused before any package bytes move.

use litchi_core::{Error, Metadata as CoreMetadata, Result};
use litchi_odf_common::core::{
    AuthoredXmlFragment, MetaXmlPatch, Structure, XmlSourcePart, XmlSplicePublication,
    patch_meta_xml,
};
use litchi_odf_common::{XmlStreamLimits, core::metadata::Metadata as OdfMetadata};

use crate::core::OwnedPackage;

const META: &str = "meta.xml";
const MAX_OUTPUT_BYTES: usize = 128 * 1024 * 1024;

/// The metadata `source` projects, with its parsed ODF form and source XML.
fn source_metadata(source: &OwnedPackage) -> Result<(Option<String>, OdfMetadata)> {
    if !source.has_file(META)? {
        return Ok((
            None,
            OdfMetadata::from_xml(
                &Structure::default_meta_xml(),
                XmlStreamLimits::owner_default_value(),
            )?,
        ));
    }
    let xml = String::from_utf8(source.get_file(META)?)
        .map_err(|error| Error::InvalidFormat(format!("ODP meta.xml is not UTF-8: {error}")))?;
    let metadata = OdfMetadata::from_xml(&xml, XmlStreamLimits::owner_default_value())?;
    Ok((Some(xml), metadata))
}

/// The common metadata projection of `source`.
pub(crate) fn read(source: &OwnedPackage) -> Result<CoreMetadata> {
    Ok(CoreMetadata::from(source_metadata(source)?.1))
}

/// Publish `metadata` into `source`'s `meta.xml`; `None` for an exact no-op.
///
/// Each changed field is published as its own checked splice of the
/// retained part, so the producer's formatting between fields never has to
/// pass the authored-XML audit.
pub(crate) fn apply(source: &OwnedPackage, metadata: &CoreMetadata) -> Result<Option<Vec<u8>>> {
    let (_, odf) = source_metadata(source)?;
    let current = CoreMetadata::from(odf);
    if !unsupported_equal(&current, metadata) {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odp.authoring.metadata",
            "ODP metadata edits support title, author, subject, description and keywords only",
        )));
    }
    if supported_equal(&current, metadata) {
        return Ok(None);
    }
    let mut package = source.clone();
    let mut staged = current;
    let steps: [fn(&mut CoreMetadata, &CoreMetadata); 5] = [
        |to, from| to.title.clone_from(&from.title),
        |to, from| to.author.clone_from(&from.author),
        |to, from| to.subject.clone_from(&from.subject),
        |to, from| to.description.clone_from(&from.description),
        |to, from| to.keywords.clone_from(&from.keywords),
    ];
    for step in steps {
        let mut next = staged.clone();
        step(&mut next, metadata);
        if next == staged {
            continue;
        }
        let bytes = publish_one(&package, &next)?;
        package = OwnedPackage::from_bytes(bytes)?;
        staged = next;
    }
    Ok(Some(package.as_bytes().to_vec()))
}

/// Publish one metadata state whose patch differs from `source` in one
/// field.
fn publish_one(source: &OwnedPackage, metadata: &CoreMetadata) -> Result<Vec<u8>> {
    let (xml, odf) = source_metadata(source)?;
    let base = xml.clone().unwrap_or_else(Structure::default_meta_xml);
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&odf, metadata);
    let patched =
        patch_meta_xml(&base, &patch, XmlStreamLimits::owner_default_value())?.unwrap_or(base);
    if patched.len() > MAX_OUTPUT_BYTES {
        return Err(Error::InvalidFormat(
            "patched ODP meta.xml exceeds the size limit".to_string(),
        ));
    }
    OdfMetadata::from_xml(&patched, XmlStreamLimits::owner_default_value())?;
    if let Some(xml) = &xml {
        let part = XmlSourcePart::load(source, META)?;
        if part.bytes() != xml.as_bytes() {
            return Err(Error::InvalidFormat(
                "ODP meta.xml changed while it was patched".to_string(),
            ));
        }
        // The patcher drops a byte-order mark the source part keeps.
        let patched = match (
            xml.strip_prefix('\u{feff}'),
            patched.starts_with('\u{feff}'),
        ) {
            (Some(_), false) => format!("\u{feff}{patched}"),
            _ => patched,
        };
        let (range, fragment) = element_splice(xml, &patched).ok_or_else(|| {
            Error::Refused(litchi_core::edit::Refusal::new(
                litchi_core::edit::RefusalKind::UnsupportedOperation,
                "odp.authoring.metadata",
                "the ODP meta.xml change does not align with one element",
            ))
        })?;
        let fragment = if fragment.is_empty() {
            AuthoredXmlFragment::deletion()
        } else {
            AuthoredXmlFragment::markup(fragment.into_bytes())?
        };
        let mut publication = XmlSplicePublication::new(part.clone());
        publication.replace(
            part.checked_range(range.clone(), &xml.as_bytes()[range])?,
            fragment,
        )?;
        return litchi_odf_common::core::rebuild_package_with_xml_splices(
            source,
            vec![publication],
            MAX_OUTPUT_BYTES,
        );
    }
    let content = String::from_utf8(source.get_file("content.xml")?)
        .map_err(|error| Error::InvalidFormat(format!("ODP content.xml is not UTF-8: {error}")))?;
    litchi_odf_common::package::rebuild_package(
        source,
        &content,
        vec![litchi_odf_common::package::Addition {
            path: META.to_string(),
            bytes: patched.into_bytes(),
            media_type: "text/xml".to_string(),
        }],
        Vec::new(),
        Vec::<String>::new(),
        Vec::<String>::new(),
    )
}

fn supported_equal(left: &CoreMetadata, right: &CoreMetadata) -> bool {
    left.title == right.title
        && left.subject == right.subject
        && left.author == right.author
        && left.keywords == right.keywords
        && left.description == right.description
}

fn unsupported_equal(left: &CoreMetadata, right: &CoreMetadata) -> bool {
    left.identifier == right.identifier
        && left.language == right.language
        && left.template == right.template
        && left.last_modified_by == right.last_modified_by
        && left.revision == right.revision
        && left.created == right.created
        && left.modified == right.modified
        && left.page_count == right.page_count
        && left.word_count == right.word_count
        && left.character_count == right.character_count
        && left.character_count_with_spaces == right.character_count_with_spaces
        && left.editing_time_minutes == right.editing_time_minutes
        && left.application == right.application
        && left.category == right.category
        && left.company == right.company
        && left.manager == right.manager
        && left.content_status == right.content_status
        && left.content_type == right.content_type
        && left.version == right.version
        && left.last_printed == right.last_printed
        && left.last_backup == right.last_backup
        && left.hyperlink_base == right.hyperlink_base
        && left.security == right.security
        && left.codepage == right.codepage
}

/// The one source range and balanced replacement markup that turn `old`
/// into `new`, aligned to element boundaries: the range starts at a `<` and
/// ends before one, and the replacement is empty or balanced markup.
fn element_splice(old: &str, new: &str) -> Option<(std::ops::Range<usize>, String)> {
    let (old_bytes, new_bytes) = (old.as_bytes(), new.as_bytes());
    let prefix = old_bytes
        .iter()
        .zip(new_bytes)
        .take_while(|(left, right)| left == right)
        .count();
    // Candidate starts: each `<` before the first difference, nearest first.
    let mut starts = old_bytes[..prefix]
        .iter()
        .enumerate()
        .rev()
        .filter(|(_, byte)| **byte == b'<')
        .map(|(index, _)| index);
    starts.find_map(|start| {
        let limit = old_bytes.len().min(new_bytes.len()) - start;
        let suffix = old_bytes
            .iter()
            .rev()
            .zip(new_bytes.iter().rev())
            .take(limit)
            .take_while(|(left, right)| left == right)
            .count();
        let (mut old_end, mut new_end) = (old_bytes.len() - suffix, new_bytes.len() - suffix);
        while old_end < old_bytes.len() {
            if old_bytes[old_end] == b'<'
                && let (Some(source), Some(candidate)) =
                    (old.get(start..old_end), new.get(start..new_end))
                && balanced(source)
                && balanced(candidate)
            {
                return Some((start..old_end, candidate.to_owned()));
            }
            old_end += 1;
            new_end += 1;
        }
        None
    })
}

/// Whether `fragment` is empty or balanced markup that starts with a tag.
fn balanced(fragment: &str) -> bool {
    if fragment.is_empty() {
        return true;
    }
    if !fragment.starts_with('<') || !fragment.ends_with('>') {
        return false;
    }
    let mut reader = quick_xml::Reader::from_str(fragment);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    loop {
        match reader.read_event() {
            Ok(quick_xml::events::Event::Start(_)) => depth += 1,
            Ok(quick_xml::events::Event::End(_)) => {
                let Some(next) = depth.checked_sub(1) else {
                    return false;
                };
                depth = next;
            },
            Ok(quick_xml::events::Event::Eof) => return depth == 0,
            Ok(_) => {},
            Err(_) => return false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::element_splice;

    fn apply(old: &str, new: &str) -> (String, String) {
        let (range, fragment) = element_splice(old, new).unwrap();
        assert!(fragment.is_empty() || fragment.starts_with('<'));
        assert_eq!(old.as_bytes().get(range.end), Some(&b'<'));
        let assembled = format!("{}{}{}", &old[..range.start], fragment, &old[range.end..]);
        (assembled, fragment)
    }

    #[test]
    fn splices_align_to_elements() {
        let old = "<m><a>1</a><b/></m>";
        for new in [
            "<m><t>T</t><a>1</a><b/></m>",
            "<m><a>2</a><b/></m>",
            "<m><b/></m>",
            "<m><a>1</a><b/><c>x y</c></m>",
        ] {
            assert_eq!(apply(old, new).0, new);
        }
        assert_eq!(apply(old, "<m><a>2</a><b/></m>").1, "<a>2</a>");
    }
}
