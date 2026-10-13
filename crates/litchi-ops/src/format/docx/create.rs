//! DOCX's typed creator for the unified operations surface.

use litchi_core::Metadata;

use super::Docx;
use crate::create::{self, Conformance, CreateOptions, Flavor};
use crate::error::{Error, Result};
use crate::options::OpenOptions;

/// Build and adopt a DOCX package without a ZIP serialization round trip.
pub(crate) fn create(options: &CreateOptions) -> Result<Docx> {
    if options.conformance_value() == Conformance::Strict {
        return Err(create::unsupported_docx_option(
            "the DOCX creator does not author Strict conformance",
            "use Conformance::Transitional; Strict authoring is not integrated",
        ));
    }

    let mut package = match options.flavor_value() {
        Flavor::Document => litchi_docx::Package::new().map_err(Error::from)?,
        Flavor::Template => litchi_docx::Package::new_template().map_err(Error::from)?,
    };

    if let Some(metadata) = options.metadata_value() {
        apply_metadata(&mut package, metadata)?;
    }

    // A created package is already an owning leaf package.  Adopt it directly
    // so the unified handle does not decode, serialize, and reopen the ZIP.
    Docx::from_owned_package(package, &OpenOptions::new())
}

fn apply_metadata(package: &mut litchi_docx::Package, metadata: &Metadata) -> Result<()> {
    if has_unsupported_metadata(metadata) {
        return Err(create::invalid_created_metadata());
    }

    let mut edit = package.edit_metadata();
    edit.update(|props| {
        props.title = metadata.title.clone();
        props.subject = metadata.subject.clone();
        props.creator = metadata.author.clone();
        props.keywords = metadata.keywords.clone().map(Into::into);
        props.description = metadata.description.clone();
        props.identifier = metadata.identifier.clone();
        props.language = metadata.language.clone();
        props.last_modified_by = metadata.last_modified_by.clone();
        props.category = metadata.category.clone();
        props.content_status = metadata.content_status.clone();
        props.revision = metadata.revision.clone();
        props.version = metadata.version.clone();

        // `Props` uses the same typed W3CDTF and xsd:dateTime wrappers as the
        // unified metadata conversion.  The destination field supplies the
        // concrete type to `Into`, so this keeps conversion in the leaf codec.
        props.created = metadata.created.map(Into::into);
        props.modified = metadata.modified.map(Into::into);
        props.last_printed = metadata.last_printed_time.map(Into::into);
    });
    *package = edit.commit()?.into_package();
    Ok(())
}

fn has_unsupported_metadata(metadata: &Metadata) -> bool {
    // Deliberately exhaustive: a new unified field requires an explicit
    // admission decision instead of silently disappearing during creation.
    let Metadata {
        title: _,
        subject: _,
        author: _,
        keywords: _,
        description: _,
        identifier: _,
        language: _,
        template,
        last_modified_by: _,
        revision: _,
        created: _,
        created_local,
        modified: _,
        modified_local,
        page_count,
        word_count,
        character_count,
        character_count_with_spaces,
        editing_time_minutes,
        application,
        category: _,
        company,
        manager,
        content_status: _,
        content_type,
        version: _,
        last_printed_time: _,
        last_printed_local,
        last_backup_local,
        hyperlink_base,
        security,
        codepage,
    } = metadata;
    template.is_some()
        || created_local.is_some()
        || modified_local.is_some()
        || last_printed_local.is_some()
        || page_count.is_some()
        || word_count.is_some()
        || character_count.is_some()
        || character_count_with_spaces.is_some()
        || editing_time_minutes.is_some()
        || application.is_some()
        || company.is_some()
        || manager.is_some()
        || content_type.is_some()
        || last_backup_local.is_some()
        || hyperlink_base.is_some()
        || security.is_some()
        || codepage.is_some()
}
