//! Retained-source `meta.xml` edits for the primary ODP transaction.
//!
//! The common ODF patcher owns five simple Dublin Core and ODF fields
//! (title, creator, subject, description, keywords); every other element of
//! the source `meta.xml` keeps its bytes. A request that changes a field the
//! patcher does not own is refused before any package bytes move.

use litchi_core::{Error, Metadata as CoreMetadata, Result};
use litchi_odf_common::core::metadata::Metadata as OdfMetadata;
use litchi_odf_common::core::{MetaXmlPatch, Structure, patch_meta_xml};

use crate::core::OwnedPackage;

const META: &str = "meta.xml";
const MAX_OUTPUT_BYTES: usize = 128 * 1024 * 1024;

/// The metadata `source` projects, with its parsed ODF form and source XML.
fn source_metadata(source: &OwnedPackage) -> Result<(Option<String>, OdfMetadata)> {
    if !source.has_file(META)? {
        return Ok((None, OdfMetadata::from_xml(&Structure::default_meta_xml())?));
    }
    let xml = String::from_utf8(source.get_file(META)?)
        .map_err(|error| Error::InvalidFormat(format!("ODP meta.xml is not UTF-8: {error}")))?;
    let metadata = OdfMetadata::from_xml(&xml)?;
    Ok((Some(xml), metadata))
}

/// The common metadata projection of `source`.
pub(crate) fn read(source: &OwnedPackage) -> Result<CoreMetadata> {
    Ok(CoreMetadata::from(source_metadata(source)?.1))
}

/// Publish `metadata` into `source`'s `meta.xml`; `None` for an exact no-op.
pub(crate) fn apply(source: &OwnedPackage, metadata: &CoreMetadata) -> Result<Option<Vec<u8>>> {
    let (xml, odf) = source_metadata(source)?;
    let current = CoreMetadata::from(odf.clone());
    if !unsupported_equal(&current, metadata) {
        return Err(Error::Unsupported(
            "ODP metadata edits support title, author, subject, description and keywords only"
                .to_string(),
        ));
    }
    if supported_equal(&current, metadata) {
        return Ok(None);
    }
    let base = xml.clone().unwrap_or_else(Structure::default_meta_xml);
    let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&odf, metadata);
    let patched = patch_meta_xml(&base, &patch)?.unwrap_or(base);
    if patched.len() > MAX_OUTPUT_BYTES {
        return Err(Error::InvalidFormat(
            "patched ODP meta.xml exceeds the size limit".to_string(),
        ));
    }
    OdfMetadata::from_xml(&patched)?;
    if xml.is_some()
        && let Ok(publication) =
            litchi_odf_common::package::xml_splice_publication(source, META, &patched)
    {
        return litchi_odf_common::core::rebuild_package_with_xml_splices(
            source,
            vec![publication],
            MAX_OUTPUT_BYTES,
        )
        .map(Some);
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
    .map(Some)
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
