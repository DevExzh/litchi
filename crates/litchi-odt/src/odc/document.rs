//! Namespace-aware access to a standalone `OpenDocument` chart package.

use crate::{Family, GenericPackage as Package, constants, core::PackageWriter};
use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::calculation::{Settings, parse};
use litchi_odf_common::chart::authoring::{Definition, serialize_content};
use litchi_odf_common::chart::{Element, read};
use litchi_odf_common::{Error as OdfError, OpenOptions};
use std::io::Read;
use std::path::Path;

/// A validated standalone `OpenDocument` chart or chart template.
///
/// The document is immutable: [`Document::edit`] stages an
/// [`Edit`](super::Edit) whose commit publishes a new document.
pub struct Document {
    pub(crate) package: Package,
    pub(crate) chart: Element,
}

impl Document {
    /// Create a new packaged `.odc` document from a typed chart definition.
    pub fn create(definition: &Definition) -> Result<Self> {
        Self::create_with_mimetype(definition, constants::ODF_CHART)
    }

    /// Create a new packaged `.otc` chart template.
    pub fn create_template(definition: &Definition) -> Result<Self> {
        Self::create_with_mimetype(definition, constants::ODF_CHART_TEMPLATE)
    }

    fn create_with_mimetype(definition: &Definition, mimetype: &str) -> Result<Self> {
        let content = serialize_content(definition)?;
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        Self::from_vec(writer.finish_to_bytes()?)
    }

    /// Open a packaged chart from a path.
    ///
    /// # Errors
    ///
    /// Returns typed limit failures and package or family-validation errors.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::open(path)?)
    }

    /// Open a packaged chart from a path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::open_with_options(path, options)?)
    }

    /// Read a packaged chart through the input-limit guard.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader(reader: impl Read) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::from_reader(reader)?)
    }

    /// Read a packaged chart under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_reader_with_options(
        reader: impl Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::from_reader_with_options(reader, options)?)
    }

    /// Validate a packaged chart from bytes (the bytes are copied).
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::from_bytes(bytes)?)
    }

    /// Validate a packaged chart from bytes under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::open`].
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, OdfError> {
        Self::from_package(Package::from_bytes_with_options(bytes, options)?)
    }

    pub(crate) fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Ok(Self::from_package(Package::from_vec(bytes)?)?)
    }

    fn from_package(package: Package) -> std::result::Result<Self, OdfError> {
        if package.family() != Family::Chart {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument chart: MIME type is '{}'",
                package.mimetype()
            ))
            .into());
        }
        let chart = read(&package.content_xml()?, package.xml_limits_value())?;
        Ok(Self { package, chart })
    }

    /// Replace the typed chart content while preserving package entries.
    pub(crate) fn set_definition(&mut self, definition: &Definition) -> Result<()> {
        let content = serialize_content(definition)?;
        let parsed = read(&content, self.package.xml_limits_value())?;
        self.package.replace_content_xml(content)?;
        self.chart = parsed;
        Ok(())
    }

    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    pub fn chart(&self) -> &Element {
        &self.chart
    }

    pub fn plot_area(&self) -> Option<litchi_odf_common::chart::PlotArea<'_>> {
        self.chart.plot_area()
    }

    pub fn legend(&self) -> Option<litchi_odf_common::chart::Legend<'_>> {
        self.chart.legend()
    }

    /// Return inert calculation settings stored beside the chart.
    pub fn calculation_settings(&self) -> Result<Option<Settings>> {
        parse(
            &self.package.content_xml()?,
            self.package.xml_limits_value(),
        )
        .map_err(Error::from)
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        Ok(self.package.variable_declarations()?)
    }

    pub fn text(&self) -> String {
        self.chart.all_text()
    }

    pub fn metadata(&self) -> Result<Metadata> {
        Ok(self.package.metadata()?)
    }

    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        Ok(self.package.odf_metadata()?)
    }

    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    /// Atomically replace `path` with the chart package bytes.
    ///
    /// # Errors
    ///
    /// Returns the publication I/O failure; the destination is unchanged
    /// unless the error is a committed-durability failure.
    pub fn save(&self, path: impl AsRef<Path>) -> std::result::Result<(), OdfError> {
        self.package.save(path)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use litchi_odf_common::chart::Kind;
    use litchi_odf_common::namespace::CHARTNS;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn chart_xml() -> &'static str {
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:chart><c:chart c:class="c:bar"><c:title><x:p>Revenue &amp; margin</x:p></c:title><c:legend c:legend-position="end"/><c:plot-area t:cell-range-address="Data.A1:C4" c:data-source-has-labels="both"><c:axis c:dimension="x"><c:categories t:cell-range-address="Data.A2:A4"/><c:grid c:class="major"/></c:axis><c:series c:values-cell-range-address="Data.B2:B4"><c:domain t:cell-range-address="Data.A2:A4"/></c:series></c:plot-area></c:chart></o:chart></o:body></o:document-content>"#
    }

    #[test]
    fn opens_chart_package_with_common_retained_content() {
        let bytes = package(constants::ODF_CHART, chart_xml());
        let document = Document::from_bytes(&bytes).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.chart().kind(), Kind::Chart);
        assert_eq!(
            document.chart().attribute(Some(CHARTNS), "class"),
            Some("c:bar")
        );
        assert_eq!(document.plot_area().unwrap().axes().count(), 1);
        assert_eq!(document.plot_area().unwrap().series().count(), 1);
        assert_eq!(document.text(), "Revenue & margin");
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_chart_templates_and_rejects_other_families() {
        let bytes = package(constants::ODF_CHART_TEMPLATE, chart_xml());
        let document = Document::from_bytes(&bytes).unwrap();
        assert!(document.is_template());
        assert_eq!(document.into_bytes(), bytes);
        assert!(Document::from_bytes(&package(constants::ODF_DRAWING, chart_xml())).is_err());
    }
}
