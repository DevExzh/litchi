use super::prelude::*;

impl Document {
    // ──────────────────────────────────────────────────────────────────
    // Bookmarks
    // ──────────────────────────────────────────────────────────────────

    /// Get all standard bookmarks in start-CP order.
    pub fn bookmarks(&self) -> Result<Vec<Bookmark>> {
        Ok(self.bookmarks_table.bookmarks().to_vec())
    }

    /// Legacy smart-tag metadata, when the document contains it.
    ///
    /// Recognition code and download URLs remain inert; this only exposes the
    /// validated ranges, property bags, types, and recognizer states.
    pub fn smart_tags(&self) -> Option<&DocumentSmartTags> {
        self.smart_tags.as_ref()
    }

    /// Revision-save identifiers assigned in the document (MS-DOC 2.9.203),
    /// when the document carries a `PLRSID` table.
    pub fn rsids(&self) -> Option<&DocumentRsids> {
        self.rsids.as_ref()
    }

    /// E-mail review threading data parallel to the revision-author table
    /// (MS-DOC 2.9.230), when the document carries an `RmdThreading`.
    ///
    /// The data is inert: message identifiers are exposed verbatim and no
    /// message is ever contacted, opened, or rendered.
    pub fn rmd_threading(&self) -> Option<&DocumentRmdThreading> {
        self.rmd_threading.as_ref()
    }

    /// Embedded TrueType font descriptions from the `SttbTtmbd` table
    /// (MS-DOC 2.9.296), when the document embeds fonts.
    ///
    /// The metadata is inert: font data stays in the `WordDocument` stream
    /// and is never loaded, installed, or executed.
    pub fn embedded_fonts(&self) -> Option<&DocumentEmbeddedFonts> {
        self.embedded_fonts.as_ref()
    }

    /// AutoSummary priority ranges for the main document (MS-DOC 2.8.4),
    /// when the document carries a `PlcfAsumy`.
    pub fn auto_summary(&self) -> Option<&DocumentAutoSummary> {
        self.auto_summary.as_ref()
    }

    /// Word 2003 range-level protection ("editable ranges") metadata, when
    /// the document carries it (MS-DOC 2.9.283 and 2.9.293).
    ///
    /// The metadata is inert: usernames are exposed verbatim, never
    /// authenticated, and no protection policy is enforced.
    pub fn protected_ranges(&self) -> Option<&Ranges> {
        self.protected_ranges.as_ref()
    }

    /// Format consistency-checker marks, when the document carries them
    /// (MS-DOC 2.9.282 and 2.9.64).
    ///
    /// The data is inert: it records which text regions the checker flagged
    /// and why; no formatting is analyzed or modified.
    pub fn format_consistency_marks(&self) -> Option<&DocumentFormatConsistencyMarks> {
        self.format_consistency_marks.as_ref()
    }

    /// Word 2003 structured document tag bookmarks, when the document
    /// carries them (MS-DOC 2.9.284 and 2.9.239).
    ///
    /// The data is inert: no XML schema is resolved and no placeholder is
    /// rendered.
    pub fn structured_tags(&self) -> Option<&DocumentStructuredTags> {
        self.structured_tags.as_ref()
    }

    /// The XML schema definition references of the document (`Hplxsdr`,
    /// MS-DOC 2.9.117), when it carries any.
    ///
    /// The data is inert: schema URIs and name tables are exposed verbatim;
    /// no schema is fetched, resolved, or applied.
    pub fn xml_schemas(&self) -> Option<&crate::parts::xml_schemas::Collection> {
        self.xml_schemas.as_ref()
    }

    /// The custom XML save transform path (`fcCustomXForm`, MS-DOC 2.5.9):
    /// the XML stylesheet Word applies when saving the document in XML
    /// format, when the document names one.
    ///
    /// The path is inert: it is exposed verbatim and never opened, resolved,
    /// or applied.
    pub fn custom_xml_transform_path(&self) -> Option<&str> {
        self.custom_xml_transform_path.as_deref()
    }

    /// The OLE controls recorded in the document (`RgxOcxInfo`, MS-DOC
    /// 2.9.229), when it contains any.
    ///
    /// The data is inert: no control is instantiated or activated and no
    /// control code is executed.
    pub fn ole_controls(&self) -> Option<&RgxOcxInfo> {
        self.ole_controls.as_ref()
    }

    /// The mail-merge data-source state of the document (`Pms` plus the ODSO
    /// property set), when the document carries any (MS-DOC 2.9.205, 2.9.162).
    ///
    /// The state is inert: data-source paths, connection strings, and SQL
    /// queries are stored verbatim, never opened, resolved, contacted, or
    /// executed, and no merge is performed.
    pub fn mail_merge(&self) -> Option<&DocumentMailMerge> {
        self.mail_merge.as_ref()
    }

    /// The master-document subdocument directory (`PlcfWKB`) and the
    /// referenced-file name table (`SttbFnm`), when the document carries
    /// either (MS-DOC 2.8.34, 2.9.288).
    ///
    /// The metadata is inert: file paths are exposed verbatim and are never
    /// opened, resolved, or followed, and no subdocument content is loaded.
    pub fn subdocuments(&self) -> Option<&Collection> {
        self.subdocuments.as_ref()
    }

    /// The Word 97 mail-merge state (`Pms`), when the document carries one.
    pub fn mail_merge_state(&self) -> Option<&crate::parts::mail_merge::Pms> {
        self.mail_merge.as_ref().and_then(DocumentMailMerge::state)
    }

    /// The Word 2002+ ODSO mail-merge properties, when the document carries
    /// mail-merge state. Never used to contact a data source.
    pub fn odso_properties(&self) -> Option<&[crate::parts::mail_merge::OdsoProperty]> {
        self.mail_merge
            .as_ref()
            .map(DocumentMailMerge::odso_properties)
    }

    /// Get author names used by tracked revisions and related annotations.
    pub fn revision_authors(&self) -> &[String] {
        self.revision_authors.authors()
    }

    /// Get section property revision marks in document order.
    pub fn section_revisions(&self) -> &[crate::revision::SectionRevisionMark] {
        self.sections.revisions()
    }

    /// Get sections in main-document character-position order.
    pub fn sections(&self) -> &[crate::section::Section] {
        self.sections.sections()
    }

    /// Find the section containing `cp` using half-open section ranges.
    pub fn section_at_cp(&self, cp: u32) -> Option<&crate::section::Section> {
        self.sections.section_at_cp(cp)
    }
}
