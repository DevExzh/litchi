//! Namespace-aware access to a standalone OpenDocument chart package.

use crate::{Family, Package, constants, core::PackageWriter};
use litchi_core::{Error, Metadata, Result};
use litchi_odc::{Definition, serialize_content};
use litchi_odf_common::calculation::{Settings, parse};
use litchi_odf_common::chart::{Element, read};
use std::io::Read;
use std::path::Path;

/// A validated standalone OpenDocument chart or chart template.
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
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype)?;
        writer.add_file(constants::ODF_CONTENT, content.as_bytes())?;
        Self::from_bytes(writer.finish_to_bytes()?)
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = Package::from_bytes(bytes)?;
        if package.family() != Family::Chart {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument chart: MIME type is '{}'",
                package.mimetype()
            )));
        }
        let chart = read(&package.content_xml()?)?;
        Ok(Self { package, chart })
    }

    /// Replace the typed chart content while preserving package entries.
    pub fn set_definition(&mut self, definition: &Definition) -> Result<()> {
        let content = serialize_content(definition)?;
        let parsed = read(&content)?;
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
        parse(&self.package.content_xml()?)
    }

    /// Inspect ordered ODF variable declarations without evaluating fields or formulas.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        self.package.variable_declarations()
    }

    pub fn text(&self) -> String {
        self.chart.all_text()
    }

    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.package.odf_metadata()
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

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
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
        let mut writer = PackageWriter::new();
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
        let document = Document::from_bytes(bytes.clone()).unwrap();
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
        let document = Document::from_bytes(bytes.clone()).unwrap();
        assert!(document.is_template());
        assert_eq!(document.into_bytes(), bytes);
        assert!(Document::from_bytes(package(constants::ODF_DRAWING, chart_xml())).is_err());
    }
}
