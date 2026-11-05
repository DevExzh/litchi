//! ODT projection construction and lossless byte access.
//!
//! Public ingress and output verbs live on [`crate::Package`]; the
//! constructors here are crate-internal so commit and authoring paths can
//! reparse an intermediate artifact without a public state verb.

use super::model::{Document, State};
#[cfg(test)]
use crate::core::PreparedPackage;
use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::elements::style::{StyleElements, StyleRegistry};
use litchi_core::{Error, Result};
use litchi_odf_common::core::private::neutralize_prologue_doctype_bytes;
use litchi_odf_common::leniency::ContentDefect;
use litchi_odf_common::{SemanticLimitKind, SemanticLimits, XmlStreamLimits};
use std::sync::Arc;

impl Document {
    pub(crate) fn transaction_xml_limits(&self) -> XmlStreamLimits {
        self.state.xml_limits
    }

    pub(crate) fn transaction_package(&self) -> &OwnedPackage {
        &self.state.package
    }

    pub(crate) fn transaction_content_xml(&self) -> &str {
        self.state.content.xml_content()
    }

    pub(crate) fn transaction_styles_xml(&self) -> Option<&str> {
        self.state.styles.as_ref().map(Styles::xml_content)
    }

    pub(crate) fn replace_transaction_bytes(&mut self, bytes: Vec<u8>) -> Result<()> {
        let xml_limits = self.state.xml_limits;
        *self = Self::from_bytes_with_xml_limits(bytes, xml_limits)?;
        Ok(())
    }

    /// Reparse a complete ODT artifact produced inside the crate.
    pub(crate) fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_xml_limits(bytes, XmlStreamLimits::owner_default_value())
    }

    pub(crate) fn from_bytes_with_xml_limits(
        bytes: Vec<u8>,
        xml_limits: XmlStreamLimits,
    ) -> Result<Self> {
        // Let the prepared detector perform the local MIME probe exactly
        // once.  A matching ODT result transfers its indexed archive; a
        // different ODF family still follows the historical package owner so
        // its MIME error is reported at the same boundary; rejected probes
        // recover the original allocation for the ordinary parser.
        match litchi_odf_common::detect::prepared_or_original(
            bytes,
            litchi_odf_common::SourceLimits::default(),
        ) {
            Ok(prepared) if prepared.format() == litchi_core::detection::FileFormat::Odt => {
                Self::from_owned_package_with_limits(
                    prepared.into_package(),
                    SemanticLimits::DEFAULT,
                    xml_limits,
                )
                .map_err(Error::from)
            },
            Ok(prepared) => Self::from_owned_package_with_limits(
                prepared.into_package(),
                SemanticLimits::DEFAULT,
                xml_limits,
            )
            .map_err(Error::from),
            Err(bytes) => Self::from_owned_package_with_limits(
                OwnedPackage::from_bytes(bytes)?,
                SemanticLimits::DEFAULT,
                xml_limits,
            )
            .map_err(Error::from),
        }
    }

    /// Adopt a prepared packaged ODF result produced by smart detection.
    #[cfg(test)]
    pub(crate) fn from_prepared_package(prepared: PreparedPackage) -> Result<Self> {
        if prepared.format() != litchi_core::detection::FileFormat::Odt {
            return Err(Error::UnexpectedFormat {
                detected: prepared.format(),
            });
        }
        Self::from_owned_package(prepared.into_package())
    }

    /// Return the identity of the archive index retained by this document.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        self.state.package.prepared_index_identity()
    }

    #[cfg(test)]
    pub(crate) fn from_owned_package(owned_package: OwnedPackage) -> Result<Self> {
        Self::from_owned_package_with_limits(
            owned_package,
            SemanticLimits::DEFAULT,
            XmlStreamLimits::owner_default_value(),
        )
        .map_err(Error::from)
    }

    /// Parse the projection under caller-selected semantic limits.
    pub(crate) fn from_owned_package_with_limits(
        owned_package: OwnedPackage,
        limits: SemanticLimits,
        xml_limits: XmlStreamLimits,
    ) -> std::result::Result<Self, litchi_odf_common::Error> {
        let mut tolerance = owned_package.tolerance_report().clone();
        let package = owned_package.package()?;

        // Verify this is a text document.
        validate_mimetype(package.mimetype())?;

        // Parse core components. Declared part sizes are checked against the
        // semantic limits before either part is inflated.
        if !limits.is_default() {
            check_part_size(
                &package,
                "content.xml",
                SemanticLimitKind::ContentBytes,
                limits,
            )?;
            check_part_size(
                &package,
                "styles.xml",
                SemanticLimitKind::StylesBytes,
                limits,
            )?;
        }
        let mut content_bytes = package.get_file("content.xml")?;
        if neutralize_prologue_doctype_bytes(&mut content_bytes) {
            tolerance.record(ContentDefect::DoctypeDeclaration, 0);
        }
        let content_xml = std::str::from_utf8(&content_bytes).map_err(|error| {
            Error::InvalidFormat(format!("ODT content.xml is not UTF-8: {error}"))
        })?;
        // One fused tokenization validates the content structure and collects
        // the automatic styles (the same `OpenParse` the source-backed facade
        // uses). Error precedence is identical to the historical sequential
        // passes: content validation (mid-stream and end-of-stream) returns
        // here, before `Content` adoption and before styles.xml/meta.xml are
        // fetched; a content-styles error is only recorded and surfaces via
        // `finish` after the styles.xml parse, and a `try_extend` conflict
        // still loses to it — the historical call order.
        let content_styles =
            super::open_parse::OpenParse::run_with_xml_limits(content_xml, xml_limits)?;
        limits.check(SemanticLimitKind::Units, content_styles.units())?;
        let content = Content::from_vec(content_bytes, limits)?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(Styles::from_vec(styles_bytes, limits)?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(Meta::from_vec(meta_bytes, xml_limits)?)
        } else {
            None
        };

        // Initialize style registry
        let mut style_registry = StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles {
            style_registry = StyleElements::parse_styles(styles_part.xml_content())?;
        }

        // Also parse styles from content.xml (automatic styles), collected by
        // the fused pass above; its recorded error surfaces at this point,
        // after the styles.xml parse.
        // Merge content styles into main registry (content styles take precedence)
        style_registry.try_extend(content_styles.finish()?)?;

        Ok(Self {
            state: Arc::new(State {
                package: owned_package,
                content,
                styles,
                meta,
                style_registry,
                tolerance,
                xml_limits,
                body: std::sync::OnceLock::new(),
            }),
        })
    }

    pub(crate) fn original_bytes(&self) -> &[u8] {
        original_bytes(self)
    }

    /// Return bytes only for inline or verified package-contained images.
    /// Linked images remain inert and are never fetched.
    pub fn image_bytes(&self, image: &crate::Image) -> Result<Option<Vec<u8>>> {
        match &image.source {
            litchi_odf_common::media::Source::Inline { bytes, .. } => Ok(Some(bytes.clone())),
            litchi_odf_common::media::Source::PackagePart { path, .. } => {
                self.state.package.get_file(path).map(Some)
            },
            _ => Ok(None),
        }
    }

    /// Get a file from the ODF package (useful for extracting images)
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the file within the package (e.g., "Pictures/image1.png")
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odt::Package;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let doc = Package::open("document.odt")?.document()?;
    /// let images = doc.image_paths()?;
    /// if let Some(first_image) = images.first() {
    ///     let image_bytes = doc.get_file(first_image)?;
    ///     std::fs::write("extracted_image.png", image_bytes)?;
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        get_file(self, path)
    }

    /// Read one package member, or `None` when the package has no member
    /// at `path` (the counterpart of
    /// [`crate::SourceBackedDocument::member_data`]).
    ///
    /// # Errors
    ///
    /// Returns an error when the member exists but cannot be read.
    pub fn member_data(&self, path: &str) -> Result<Option<Vec<u8>>> {
        if !self.state.package.has_file(path)? {
            return Ok(None);
        }
        get_file(self, path).map(Some)
    }

    // Ordinary paragraph edits should use `Document::edit()` and commit an
    // immutable transaction. `MutableDocument` remains the compatibility
    // surface for structural families that have not migrated yet.
}

pub(super) fn validate_mimetype(mimetype: &str) -> Result<()> {
    if mimetype.contains("opendocument.text") {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "Not an ODT file: MIME type is {mimetype}"
        )))
    }
}

pub(super) fn original_bytes(document: &Document) -> &[u8] {
    document.state.package.as_bytes()
}

pub(super) fn get_file(document: &Document, path: &str) -> Result<Vec<u8>> {
    document.state.package.get_file(path)
}

fn check_part_size(
    package: &litchi_odf_common::core::package::Package<'_>,
    path: &str,
    kind: SemanticLimitKind,
    limits: SemanticLimits,
) -> std::result::Result<(), litchi_odf_common::Error> {
    if let Some(size) = package.member_materialized_size(path)? {
        limits.check(kind, size)?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::Document;
    use super::{ContentDefect, neutralize_prologue_doctype_bytes};
    use crate::core::OwnedPackage;
    use crate::elements::style::StyleElements;
    use litchi_core::Result;
    use litchi_odf_common::XmlStreamLimits;
    use soapberry_zip::ZipArchiveWriter;

    const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    /// Pre-0222 oracle: the historical sequential owned-path open — full
    /// `validate_content_document_part` scan, then `styles.xml` parse, then
    /// the `content.xml` automatic-styles rescan, then `try_extend` —
    /// retained verbatim to cross-check the fused `OpenParse` open on error
    /// precedence, error messages, and the resulting document state.
    fn from_owned_package_sequential_oracle(owned_package: OwnedPackage) -> Result<Document> {
        let mut tolerance = owned_package.tolerance_report().clone();
        let package = owned_package.package()?;

        // Verify this is a text document.
        super::validate_mimetype(package.mimetype())?;

        // Parse core components
        let mut content_bytes = package.get_file("content.xml")?;
        if neutralize_prologue_doctype_bytes(&mut content_bytes) {
            tolerance.record(ContentDefect::DoctypeDeclaration, 0);
        }
        let content_xml = std::str::from_utf8(&content_bytes).map_err(|error| {
            litchi_core::Error::InvalidFormat(format!("ODT content.xml is not UTF-8: {error}"))
        })?;
        litchi_odf_common::core::validate_content_document_part(
            content_xml,
            "<office:text",
            "ODT",
        )?;
        let content = crate::core::Content::from_vec(
            content_bytes,
            litchi_odf_common::SemanticLimits::DEFAULT,
        )?;

        let styles = if package.has_file("styles.xml") {
            let styles_bytes = package.get_file("styles.xml")?;
            Some(crate::core::Styles::from_vec(
                styles_bytes,
                litchi_odf_common::SemanticLimits::DEFAULT,
            )?)
        } else {
            None
        };

        let meta = if package.has_file("meta.xml") {
            let meta_bytes = package.get_file("meta.xml")?;
            Some(crate::core::Meta::from_vec(
                meta_bytes,
                XmlStreamLimits::owner_default_value(),
            )?)
        } else {
            None
        };

        // Initialize style registry
        let mut style_registry = crate::elements::style::StyleRegistry::default();

        // Parse styles from styles.xml if available
        if let Some(ref styles_part) = styles {
            style_registry = StyleElements::parse_styles(styles_part.xml_content())?;
        }

        // Also parse styles from content.xml (automatic styles)
        // Merge content styles into main registry (content styles take precedence)
        style_registry.try_extend(StyleElements::parse_styles(content.xml_content())?)?;

        Ok(Document {
            state: std::sync::Arc::new(super::State {
                package: owned_package,
                content,
                styles,
                meta,
                style_registry,
                tolerance,
                xml_limits: XmlStreamLimits::default(),
                body: std::sync::OnceLock::new(),
            }),
        })
    }

    fn oracle_from_bytes(bytes: Vec<u8>) -> Result<Document> {
        from_owned_package_sequential_oracle(OwnedPackage::from_bytes(bytes)?)
    }

    /// Deterministic comparison projection: error string on failure, or the
    /// sorted style names with their families plus the full extracted text
    /// on success.
    fn projection(result: Result<Document>) -> std::result::Result<String, String> {
        result
            .and_then(|document| {
                let mut styles: Vec<String> = document
                    .state
                    .style_registry
                    .styles
                    .iter()
                    .map(|(name, style)| format!("{name}|{:?}", style.family()))
                    .collect();
                styles.sort_unstable();
                Ok(format!("{}\n{}", styles.join(","), document.text()?))
            })
            .map_err(|error| error.to_string())
    }

    fn assert_open_parity(label: &str, bytes: &[u8]) {
        let expected = projection(oracle_from_bytes(bytes.to_vec()));
        let actual = projection(Document::from_bytes(bytes.to_vec()));
        assert_eq!(
            expected, actual,
            "{label}: fused owned open and sequential oracle disagree"
        );
    }

    fn odt_package_with_mimetype(
        mimetype: &str,
        content: &[u8],
        styles: Option<&[u8]>,
        meta: Option<&[u8]>,
    ) -> Vec<u8> {
        // Stored members written directly: `PackageWriter` validates XML on
        // publication, which would reject the deliberately malformed
        // fixtures below.
        let mut writer = ZipArchiveWriter::new(Vec::new());
        let mut built = writer.write_stored_file("mimetype", mimetype.as_bytes());
        let mut extra_members: Vec<(&str, &[u8])> = vec![("content.xml", content)];
        if let Some(styles) = styles {
            extra_members.push(("styles.xml", styles));
        }
        if let Some(meta) = meta {
            extra_members.push(("meta.xml", meta));
        }
        let mut manifest = String::from(
            r#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">"#,
        );
        manifest.push_str(&format!(
            r#"<manifest:file-entry manifest:full-path="/" manifest:media-type="{mimetype}"/>"#
        ));
        for (path, _) in &extra_members {
            manifest.push_str(&format!(
                r#"<manifest:file-entry manifest:full-path="{path}" manifest:media-type="text/xml"/>"#
            ));
        }
        manifest.push_str("</manifest:manifest>");
        for (path, bytes) in extra_members {
            built = built.and_then(|()| writer.write_stored_file(path, bytes));
        }
        built = built
            .and_then(|()| writer.write_stored_file("META-INF/manifest.xml", manifest.as_bytes()));
        match built.and_then(|()| writer.finish()) {
            Ok(bytes) => bytes,
            Err(error) => panic!("synthetic ODT package must build: {error}"),
        }
    }

    fn odt_package(content: &str, styles: Option<&str>, meta: Option<&[u8]>) -> Vec<u8> {
        odt_package_with_mimetype(
            MIMETYPE,
            content.as_bytes(),
            styles.map(str::as_bytes),
            meta,
        )
    }

    fn content(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}" office:version="1.4">{body}</office:document-content>"#
        )
    }

    fn styles_document(styles: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="{OFFICE}" xmlns:style="{STYLE}" office:version="1.4">{styles}</office:document-styles>"#
        )
    }

    const VALID_BODY: &str =
        r#"<office:body><office:text><text:p>hello</text:p></office:text></office:body>"#;
    const STYLE_A: &str = r#"<style:style style:name="A" style:family="paragraph"></style:style>"#;
    // A duplicate style attribute: the styles parse fails with
    // "invalid ODT style attribute: …".
    const DUP_ATTR_STYLE: &str = r#"<style:style style:name="a" style:name="b"></style:style>"#;
    // A mismatched end tag: the styles parse fails with
    // "invalid ODT style XML: …".
    const BROKEN_STYLES_XML: &str = "<a></b>";

    #[test]
    fn fused_owned_open_matches_sequential_oracle_on_synthetic_packages() {
        let with_styles = styles_document(STYLE_A);
        // Each case: label, package bytes, expected outcome — `None` for a
        // successful open, `Some(fragment)` for an error whose message must
        // contain the fragment. The fragments make the cross-stage error
        // precedence observable, not just parity on identical failures.
        let cases: Vec<(&str, Vec<u8>, Option<&str>)> = vec![
            (
                "valid-full",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles><style:style style:name="B" style:family="paragraph"></style:style></office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    Some(br#"<?xml version="1.0"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#),
                ),
                None,
            ),
            (
                "valid-bare",
                odt_package(&content(VALID_BODY), None, None),
                None,
            ),
            // styles.xml and content.xml both define style "A": content
            // styles take precedence (try_extend overwrites), no error.
            (
                "same-style-name-in-both",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{STYLE_A}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    None,
                ),
                None,
            ),
            // A content-styles error loses to the styles.xml parse error.
            (
                "styles-error-beats-content-style-error",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(BROKEN_STYLES_XML),
                    None,
                ),
                Some("invalid ODT style XML:"),
            ),
            // A content-styles error alone surfaces after a clean styles.xml
            // parse.
            (
                "content-style-error-alone",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles>{VALID_BODY}"#
                    )),
                    Some(&with_styles),
                    None,
                ),
                Some("invalid ODT style attribute:"),
            ),
            // A content validation error beats any styles.xml failure: the
            // historical validation pass early-returned before styles.xml
            // was even fetched.
            (
                "validation-error-beats-styles-error",
                odt_package(
                    &content(r#"<office:body><office:spreadsheet/></office:body>"#),
                    Some(BROKEN_STYLES_XML),
                    None,
                ),
                Some("has the wrong office body"),
            ),
            // A style error recorded before a later validation failure still
            // loses to the validation error.
            (
                "early-style-error-loses-to-late-validation-error",
                odt_package(
                    &content(&format!(
                        r#"<office:automatic-styles>{DUP_ATTR_STYLE}</office:automatic-styles><office:body><office:spreadsheet/></office:body>"#
                    )),
                    None,
                    None,
                ),
                Some("has the wrong office body"),
            ),
            // Deep validation failure (mismatched end tag) maps to the
            // tokenizer message.
            (
                "mismatched-end-tag-deep",
                odt_package(
                    &content(
                        r#"<office:body><office:text><text:p></text:q></office:text></office:body>"#,
                    ),
                    Some(&with_styles),
                    None,
                ),
                Some("invalid ODT content.xml:"),
            ),
            (
                "wrong-root",
                odt_package(
                    &content(VALID_BODY).replace("document-content", "document"),
                    Some(&with_styles),
                    None,
                ),
                Some("has the wrong root"),
            ),
            (
                "missing-body",
                odt_package(&content(""), None, None),
                Some("has no complete expected body"),
            ),
            (
                "meta-error-after-clean-styles",
                odt_package(&content(VALID_BODY), Some(&with_styles), Some(b"\xff\xfe")),
                Some("Invalid UTF-8 in XML content"),
            ),
        ];
        for (label, bytes, expected) in &cases {
            assert_open_parity(label, bytes);
            let outcome = projection(Document::from_bytes(bytes.clone()));
            match (expected, &outcome) {
                (None, Ok(_)) => {},
                (Some(fragment), Err(error)) => assert!(
                    error.contains(fragment),
                    "{label}: error {error:?} does not contain {fragment:?}"
                ),
                _ => panic!("{label}: outcome {outcome:?} misses expectation {expected:?}"),
            }
        }

        // Mimetype and container-level failures agree as well.
        let wrong_mime = odt_package_with_mimetype(
            "application/vnd.oasis.opendocument.presentation",
            content(VALID_BODY).as_bytes(),
            None,
            None,
        );
        assert_open_parity("wrong-mimetype", &wrong_mime);
        let outcome = projection(Document::from_bytes(wrong_mime));
        match outcome {
            Err(error) => assert!(error.contains("Not an ODT file"), "unexpected: {error}"),
            Ok(_) => panic!("wrong mimetype accepted"),
        }
        assert_open_parity("not-a-zip", b"this is not a package");
        match projection(Document::from_bytes(b"this is not a package".to_vec())) {
            Err(_) => {},
            Ok(_) => panic!("non-ZIP bytes accepted"),
        }
    }

    #[test]
    fn fused_owned_open_matches_sequential_oracle_on_odt_corpus() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        collect_odt(&root.join("test-data"), &mut files);
        files.sort();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        for path in &files {
            let Ok(bytes) = std::fs::read(path) else {
                continue;
            };
            assert_open_parity(&path.display().to_string(), &bytes);
        }
    }

    fn collect_odt(directory: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odt(&path, files);
            } else if path.extension().is_some_and(|extension| extension == "odt") {
                files.push(path);
            }
        }
    }
}
