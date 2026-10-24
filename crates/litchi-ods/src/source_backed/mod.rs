//! Edit families of the source-backed [`SourceBackedSpreadsheet`] (record
//! 0358a; the shape of DOCX record 0355a §1).
//!
//! ```text
//! SourceBackedSpreadsheet::edit_x(&self, ..) -> x::Edit       // shares the owner
//! x::Edit::{verbs..}(&mut self, ..)                           // staging only
//! x::Edit::commit(self) -> Result<x::Commit, x::Error>        // validates, writes nothing
//! x::Commit = litchi_core::edit::Commit<SourceBackedSpreadsheet, x::Patch, x::Diagnostics>
//! SourceBackedSpreadsheet::apply_x(&self, &x::Patch) -> Result<x::Commit, x::Error>
//! SourceBackedSpreadsheet::write_to(&self, impl Write) / save(..) // the only output
//! ```
//!
//! Families: [`cells`] (existing cell values, formulas and hyperlinks),
//! [`dde`] (inert DDE declarations) and [`sheet_metadata`] (consolidation,
//! label ranges, cell-range sources and detective metadata).
//!
//! Every family changes only `content.xml`. A commit validates the candidate
//! exactly as the former commit-then-publish route did and returns a new
//! spreadsheet over the same retained source holding the candidate; nothing
//! is written. Writing it streams the source with only `content.xml`
//! replaced and every other member copied raw, byte for byte the output of
//! the former `write_to` of the family commit. Reading or editing the
//! committed spreadsheet parses its `content.xml` once; memory stays bounded
//! by `content.xml`, never the archive.
//!
//! A [`Patch`](cells::Patch) holds the exact `content.xml` before and after
//! the edit. It replays onto a spreadsheet over the same retained source
//! whose `content.xml` is exactly the patch's before content, and refuses any
//! other with [`Error::PatchConflict`];
//! its inverse applied to the commit's spreadsheet restores the source
//! content, and writing that restores the exact source archive.

use std::sync::Arc;

use litchi_core::SourceVersion;
use litchi_odf_common::Error;
use litchi_odf_common::core::SourceBackedPackage;

use crate::SourceBackedSpreadsheet;

pub mod cells;
pub mod dde;
pub mod sheet_metadata;

/// The exact `content.xml` change of one source-backed edit.
#[derive(Clone)]
pub(crate) struct Change {
    before: Arc<str>,
    after: Arc<str>,
    package: Arc<SourceBackedPackage>,
    source_version: SourceVersion,
}

impl Change {
    /// The change from `base`'s `content.xml` to `after`.
    pub(crate) fn new(base: &SourceBackedSpreadsheet, after: Arc<str>) -> Self {
        Self {
            before: Arc::clone(base.written_content_xml()),
            after,
            package: Arc::clone(&base.package),
            source_version: base.source_version,
        }
    }

    pub(crate) fn is_noop(&self) -> bool {
        Arc::ptr_eq(&self.before, &self.after) || self.before == self.after
    }

    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            package: Arc::clone(&self.package),
            source_version: self.source_version,
        }
    }

    pub(crate) fn before(&self) -> &str {
        &self.before
    }

    pub(crate) fn after(&self) -> &str {
        &self.after
    }

    pub(crate) fn after_arc(&self) -> &Arc<str> {
        &self.after
    }

    /// The spreadsheet that publishes this change, committed over `owner`,
    /// or the patch conflict that refuses it.
    pub(crate) fn publish(
        &self,
        owner: &SourceBackedSpreadsheet,
        operation: &'static str,
    ) -> Result<SourceBackedSpreadsheet, Error> {
        let base = owner.effective()?;
        base.check_source()?;
        if !Arc::ptr_eq(&base.package, &self.package) || base.source_version != self.source_version
        {
            return Err(Error::patch_conflict(
                operation,
                "the patch was made on another source-backed spreadsheet",
            ));
        }
        let current = base.written_content_xml();
        if !Arc::ptr_eq(current, &self.before) && current.as_ref() != self.before.as_ref() {
            return Err(Error::patch_conflict(
                operation,
                "the spreadsheet does not hold the patch's before content.xml",
            ));
        }
        if self.is_noop() {
            return Ok(owner.share());
        }
        Ok(base.with_content(Arc::clone(&self.after)))
    }
}

impl std::fmt::Debug for Change {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Change")
            .field("source_version", &self.source_version)
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .finish()
    }
}

/// Declare a family's `Patch` over [`Change`].
macro_rules! content_patch {
    ($what:literal) => {
        #[doc = concat!("The exact, reversible `content.xml` change of ", $what, ".")]
        ///
        /// It replays onto a spreadsheet over the same retained source whose
        /// `content.xml` is exactly its before content (record 0358a).
        #[derive(Clone, Debug)]
        pub struct Patch {
            change: super::Change,
        }

        impl Patch {
            /// Whether the patch changes nothing.
            #[must_use]
            pub fn is_noop(&self) -> bool {
                self.change.is_noop()
            }

            /// The patch that restores the before content.
            #[must_use]
            pub fn inverse(&self) -> Self {
                Self {
                    change: self.change.inverse(),
                }
            }

            /// The exact `content.xml` the patch applies to.
            #[must_use]
            pub fn source_xml(&self) -> &str {
                self.change.before()
            }

            /// The exact `content.xml` the patch produces.
            #[must_use]
            pub fn target_xml(&self) -> &str {
                self.change.after()
            }
        }
    };
}

pub(crate) use content_patch;

#[cfg(test)]
mod test_support {
    use super::SourceBackedSpreadsheet;
    use litchi_core::OwnedSource;
    use std::sync::Arc;

    const MIME: &str = "application/vnd.oasis.opendocument.spreadsheet";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    pub(super) fn owner() -> SourceBackedSpreadsheet {
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes())))
            .expect("source-backed ODS fixture")
    }

    fn bytes() -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>value</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new(
            litchi_odf_common::core::StreamingArchiveLimits::default(),
        );
        writer.set_mimetype(MIME).expect("ODS MIME");
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("ODS content");
        writer.finish_to_bytes().expect("ODS package")
    }
}
