//! Focused regression tests for the layered header/footer owner.

use litchi_core::Result;

use super::properties::{Color, Length, Region, StyleProperties};

#[test]
fn property_model_round_trips_through_the_canonical_owner() -> Result<()> {
    let properties = StyleProperties {
        height: Some(Length::new("1.25cm")?),
        background_color: Some(Color::Rgb(0x10, 0x20, 0x30)),
        dynamic_spacing: Some(true),
        ..Default::default()
    };
    let fragment = properties.to_region_fragment(Region::Header)?;
    let xml = format!(
        r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:automatic-styles><style:page-layout style:name="layout">{fragment}</style:page-layout></office:automatic-styles></office:document-styles>"#
    );
    let entry = super::parse_page_layout_header_footer_properties(&xml)?
        .pop()
        .ok_or_else(|| super::bad("header property readback is missing"))?;
    assert_eq!(entry.region, Region::Header);
    assert_eq!(entry.properties, properties);
    Ok(())
}

const SOURCE_LAYOUT: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles><style:page-layout style:name="pmA"><style:page-layout-properties style:print-orientation="portrait"/></style:page-layout><style:page-layout style:name="pmB"><style:page-layout-properties style:print-orientation="landscape"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="A" style:page-layout-name="pmA"><style:header><text:p>alpha</text:p></style:header></style:master-page><style:master-page style:name="B" style:page-layout-name="pmB"><style:footer><text:p>beta</text:p></style:footer></style:master-page></office:master-styles></office:document-styles>"#;

const DESTINATION_LAYOUT: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles><style:page-layout style:name="pmC"/></office:automatic-styles><office:master-styles><style:master-page style:name="C" style:page-layout-name="pmC"/></office:master-styles></office:document-styles>"#;

#[test]
fn master_transfer_carries_its_page_layout_dependency() -> Result<()> {
    let source =
        super::Snapshot::parse(SOURCE_LAYOUT, litchi_odf_common::XmlStreamLimits::default())?;
    let transfer = source.prepare_master_page_transfer("A")?;
    assert_eq!(transfer.page_layout_name(), "pmA");
    assert!(transfer.dependencies().is_empty());

    let destination = super::Snapshot::parse(
        DESTINATION_LAYOUT,
        litchi_odf_common::XmlStreamLimits::default(),
    )?;
    let mut edit = destination.edit();
    edit.insert_transfer(&transfer)?;
    let transferred = edit.commit()?;
    assert!(
        transferred
            .snapshot()
            .master_pages()
            .iter()
            .any(|master| master.name == "A")
    );
    assert!(
        transferred
            .snapshot()
            .page_layouts()
            .iter()
            .any(|layout| layout.name == "pmA")
    );
    Ok(())
}

const SECTION_LAYOUT: &str = r##"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:styles><style:style style:name="Sect" style:family="section"><style:section-properties fo:background-color="#112233" fo:margin-left="1cm" fo:margin-right="2cm" style:protect="true"><style:background-image xlink:href="Pictures/section.png" xlink:type="simple" xlink:actuate="onLoad"/><style:columns fo:column-count="2" fo:column-gap="0.5cm"/></style:section-properties></style:style><style:style style:name="Residual" style:family="section"/></office:styles><office:automatic-styles><style:page-layout style:name="pmA"/></office:automatic-styles><office:master-styles/></office:document-styles>"##;

const EMPTY_SECTION_DESTINATION: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#;

fn authored_section(column_count: u8, margin: &str) -> Result<super::SectionLayout> {
    use crate::section_properties::{
        SectionBackgroundColor, SectionLength, SectionProperties, SectionStyleProperties,
    };
    use crate::style::columns::{Columns, Length as ColumnLength};

    let style = SectionStyleProperties::new(
        "Sect",
        SectionProperties {
            background_color: Some(SectionBackgroundColor::Rgb(0x44, 0x55, 0x66)),
            margin_left: Some(SectionLength::new(margin)?),
            margin_right: Some(SectionLength::new("3cm")?),
            protect: Some(true),
            ..Default::default()
        },
    )?;
    let mut columns = Columns::new(column_count)?;
    columns.column_gap = Some(ColumnLength::new("0.75cm")?);
    super::SectionLayout::new(style, Some(columns))
}

#[test]
fn section_layout_is_typed_replaceable_and_transfer_safe() -> Result<()> {
    use crate::section_properties::SectionBackgroundColor;

    let source = super::Snapshot::parse(
        SECTION_LAYOUT,
        litchi_odf_common::XmlStreamLimits::default(),
    )?;
    assert_eq!(source.section_layouts().len(), 1);
    let section = source
        .section_layouts()
        .first()
        .ok_or_else(|| super::bad("section layout readback is missing"))?;
    assert_eq!(section.name(), "Sect");
    assert_eq!(
        section.properties().properties.background_color,
        Some(SectionBackgroundColor::Rgb(0x11, 0x22, 0x33))
    );
    assert_eq!(
        section
            .properties()
            .properties
            .margin_left
            .as_ref()
            .map(crate::section_properties::SectionLength::as_str),
        Some("1cm")
    );
    assert_eq!(
        section
            .columns()
            .ok_or_else(|| super::bad("section columns are missing"))?
            .column_count,
        2
    );

    let replacement = authored_section(3, "1.5cm")?;
    assert!(
        !replacement
            .xml()
            .chars()
            .any(|character| matches!(character, '\n' | '\r' | '\t'))
    );
    let mut edit = source.edit();
    edit.replace_section_layout(&replacement)?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.snapshot().section_layouts()[0].columns(),
        replacement.columns()
    );

    let transfer = source.prepare_section_layout_transfer("Sect")?;
    assert_eq!(transfer.dependencies().len(), 1);
    assert_eq!(transfer.dependencies()[0].href(), "Pictures/section.png");
    let destination = super::Snapshot::parse(
        EMPTY_SECTION_DESTINATION,
        litchi_odf_common::XmlStreamLimits::default(),
    )?;
    let mut unauthorized = destination.edit();
    assert!(unauthorized.insert_section_transfer(&transfer).is_err());
    let mut authorized = destination.edit();
    authorized.insert_section_transfer_with(&transfer, |dependency| {
        dependency.href() == "Pictures/section.png"
    })?;
    let transferred = authorized.commit()?;
    assert_eq!(transferred.snapshot().section_layouts().len(), 1);
    assert_eq!(
        transferred.snapshot().section_layouts()[0].xml(),
        transfer.xml()
    );
    Ok(())
}
