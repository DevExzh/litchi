//! Public ODS data-style vocabulary and native-corpus integration fixtures.
//!
//! The package-owner tests in this file deliberately stay on public APIs.  The
//! native fixtures are useful even before the source catalog/editor lands: a
//! green leaf projection must not be mistaken for proof that a packaged style
//! owner can edit the same bytes.

use std::{
    fmt::Write as _,
    io::{Cursor, Read, Write},
};

use litchi_core::{Error, Result};
use litchi_ods::{
    FlatSpreadsheet, Spreadsheet, data_style,
    document::{
        DataStyleAttributePatch, DataStyleFamily, DataStyleOwner, DataStyleSelector, Snapshot,
        StyleGraphExtension,
    },
};
use zip::ZipArchive;

mod support;

const FORMATS_ODS: &[u8] =
    include_bytes!("../../../test-data/libreoffice-core/sc/qa/unit/data/ods/formats.ods");
const YIELDDISC_FODS: &[u8] = include_bytes!(
    "../../../test-data/libreoffice-core/sc/qa/unit/data/functions/financial/fods/yielddisc.fods"
);

fn zip_member(source: &[u8], path: &str) -> Result<Vec<u8>> {
    let mut archive = ZipArchive::new(Cursor::new(source))
        .map_err(|error| Error::InvalidFormat(format!("native ODS ZIP: {error}")))?;
    let mut member = archive
        .by_name(path)
        .map_err(|error| Error::InvalidFormat(format!("native ODS member {path}: {error}")))?;
    let mut bytes = Vec::new();
    member
        .read_to_end(&mut bytes)
        .map_err(|error| Error::InvalidFormat(format!("native ODS member {path}: {error}")))?;
    Ok(bytes)
}

fn deflated_package(entries: &[(&str, &[u8], &str)]) -> Vec<u8> {
    let mut manifest = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>",
    );
    for (path, _, media_type) in entries {
        manifest.push_str("<manifest:file-entry manifest:full-path=\"");
        manifest.push_str(path);
        manifest.push_str("\" manifest:media-type=\"");
        manifest.push_str(media_type);
        manifest.push_str("\"/>");
    }
    manifest.push_str("</manifest:manifest>");

    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("mimetype", stored)
        .expect("deflated ODS mimetype");
    zip.write_all(support::MIMETYPE.as_bytes())
        .expect("deflated ODS mimetype bytes");
    zip.start_file("META-INF/manifest.xml", stored)
        .expect("deflated ODS manifest");
    zip.write_all(manifest.as_bytes())
        .expect("deflated ODS manifest bytes");
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (path, bytes, _) in entries {
        zip.start_file(*path, deflated)
            .expect("deflated ODS member");
        zip.write_all(bytes).expect("deflated ODS member bytes");
    }
    zip.finish().expect("deflated ODS package finish");
    output.into_inner()
}

fn formats_styles() -> Result<String> {
    String::from_utf8(zip_member(FORMATS_ODS, "styles.xml")?)
        .map_err(|error| Error::InvalidFormat(format!("native styles.xml UTF-8: {error}")))
}

fn formats_content() -> Result<String> {
    String::from_utf8(zip_member(FORMATS_ODS, "content.xml")?)
        .map_err(|error| Error::InvalidFormat(format!("native content.xml UTF-8: {error}")))
}

fn number_style(name: &str, format: data_style::Format) -> Result<data_style::Number> {
    data_style::Number::new(name).and_then(|mut style| {
        style.set_format(Some(format))?;
        Ok(style)
    })
}

const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const NUMBER_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn package_with_style_owners(content_styles: &str, common_styles: &str) -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles>{content_styles}</office:automatic-styles>
  <office:body><office:spreadsheet>
    <table:table table:name="Sheet1">
      <table:table-row><table:table-cell table:style-name="Cell"><text:p>1</text:p></table:table-cell></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:styles>{common_styles}</office:styles>
  <office:automatic-styles/>
  <office:master-styles/>
</office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn package_with_three_style_owners(
    content_styles: &str,
    common_styles: &str,
    styles_automatic: &str,
) -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles>{content_styles}</office:automatic-styles>
  <office:body><office:spreadsheet>
    <table:table table:name="Sheet1">
      <table:table-row><table:table-cell table:style-name="Cell"><text:p>1</text:p></table:table-cell></table:table-row>
    </table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:styles>{common_styles}</office:styles>
  <office:automatic-styles>{styles_automatic}</office:automatic-styles>
  <office:master-styles/>
</office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn three_owner_fixture() -> Vec<u8> {
    package_with_three_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="ThreeOwner"/>
<number:number-style style:name="ThreeOwner"><number:fraction number:min-numerator-digits="1"/></number:number-style>"#,
        r#"<number:number-style style:name="ThreeOwner"><number:scientific-number number:decimal-places="2"/></number:number-style>"#,
        r#"<number:number-style style:name="ThreeOwner"><number:fraction number:min-denominator-digits="3"/></number:number-style>"#,
    )
}

fn lexical_metadata_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="Lexical"/>
<number:number-style style:name="Lexical" number:title="A&#x20;B" style:display-name="A&#32;B"><number:fraction/></number:number-style>"#,
        "",
    )
}

fn inherited_prefix_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles><n:number-style s:name="Inherited" s:display-name="Display" s:volatile="false"><n:fraction/></n:number-style></office:automatic-styles>
  <office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn explicit_xml_namespace_fixture(escaped: bool) -> Vec<u8> {
    let xml_namespace = if escaped {
        "http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace"
    } else {
        "http://www.w3.org/XML/1998/namespace"
    };
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" xmlns:xml="{xml_namespace}" office:version="1.3">
  <office:automatic-styles><number:number-style style:name="ExplicitXml"><number:fraction/></number:number-style></office:automatic-styles>
  <office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn date_style_with_number_style(style: &str) -> Vec<u8> {
    let content_styles = format!(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="DateStyle"/>
<number:date-style style:name="DateStyle"><number:year number:style="{style}"/><number:month number:style="long"/><number:day number:style="short"/></number:date-style>"#
    );
    package_with_style_owners(&content_styles, "")
}

fn whitespace_number_text_fixture() -> Vec<u8> {
    let content_styles = concat!(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="WhitespaceText"/>
<number:number-style style:name="WhitespaceText"><number:text>"#,
        "\t\n\r",
        r#"</number:text><number:fraction/></number:number-style>"#,
    );
    package_with_style_owners(content_styles, "")
}

fn token_whitespace_language_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="TokenLanguage"/>
<number:number-style style:name="TokenLanguage" number:language="  en&#x9;&#xA; "><number:fraction/></number:number-style>"#,
        "",
    )
}

fn content_document_styles_root_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn styles_document_content_root_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3"><office:automatic-styles><style:style style:name="Cell" style:family="table-cell" style:data-style-name="Content"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Cell"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:automatic-styles><number:number-style style:name="WrongRoot"><number:fraction/></number:number-style></office:automatic-styles><office:body/></office:document-content>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn malformed_styles_query_fixture(fragment: &str) -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3"><office:automatic-styles><style:style style:name="Cell" style:family="table-cell" style:data-style-name="Content"/></office:automatic-styles><office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Cell"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:styles><number:number-style style:name="Common"><number:fraction/></number:number-style></office:styles>{fragment}<office:automatic-styles><number:number-style style:name="Automatic"><number:fraction/></number:number-style></office:automatic-styles><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn literal_whitespace_metadata_fixture() -> Vec<u8> {
    let content_styles = concat!(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="LiteralWhitespace"/>
<number:number-style style:name="LiteralWhitespace" number:title="A"#,
        "\t",
        "B",
        "\r",
        "C",
        "\n",
        r#"D" style:display-name="E"#,
        "\t",
        "F",
        "\r",
        "G",
        "\n",
        r#"H"><number:fraction/></number:number-style>"#,
    );
    package_with_style_owners(content_styles, "")
}

fn crlf_string_metadata_fixture() -> Vec<u8> {
    let content_styles = concat!(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="CrlfStrings"/>
<number:number-style style:name="CrlfStrings" number:title="a"#,
        "\r\n",
        r#"b" style:display-name="a"#,
        "\r\n",
        r#"b"><number:fraction/></number:number-style>"#,
    );
    package_with_style_owners(content_styles, "")
}

fn schema_string_charref_metadata_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="SchemaStrings"/>
<number:number-style style:name="SchemaStrings" number:title="a  b" style:display-name="a&#x9;b"><number:fraction/></number:number-style>"#,
        "",
    )
}

fn plain_string_metadata_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="ControlStrings"/>
<number:number-style style:name="ControlStrings" number:title="plain title" style:display-name="plain display"><number:fraction/></number:number-style>"#,
        "",
    )
}

fn self_closing_automatic_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:o="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <o:automatic-styles   xml:lang = "fr"  />
  <office:body><office:spreadsheet>
    <table:table table:name="Sheet1"><table:table-row><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row></table:table>
  </office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:o="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn owner_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="AutoFraction"/>
<number:number-style style:name="AutoFraction"><number:fraction number:min-numerator-digits="2" number:min-denominator-digits="2"/></number:number-style>
<number:text-style style:name="OpaqueText" style:display-name="before"><!--opaque comment--><number:text-content>opaque body</number:text-content></number:text-style>"#,
        r#"<number:number-style style:name="SharedName"><number:scientific-number number:decimal-places="3"/></number:number-style>
<number:number-style style:name="OtherCommon"><number:number number:decimal-places="2"/></number:number-style>"#,
    )
}

fn multi_number_owner_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="FirstFraction"/>
<number:number-style style:name="FirstFraction"><number:fraction number:min-numerator-digits="2"/></number:number-style>
<number:number-style style:name="SecondFraction"><number:fraction number:min-denominator-digits="2"/></number:number-style>
<number:text-style style:name="OpaqueText"><!--opaque comment--><number:text-content>opaque body</number:text-content></number:text-style>"#,
        r#"<number:number-style style:name="CommonFraction"><number:fraction/></number:number-style>"#,
    )
}

fn multi_lexical_graph_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles>
    <n:number-style s:name='FirstFraction' n:title='A&#x20;B' s:display-name='A&#32;B'><n:fraction n:min-numerator-digits='2'/></n:number-style>
    <n:number-style s:name="SecondFraction"><n:fraction n:min-denominator-digits="2"/></n:number-style>
  </office:automatic-styles>
  <office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn compressed_late_graph_fixture() -> Vec<u8> {
    let padding = " \n".repeat(32_768);
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles>{padding}<number:number-style style:name="CompressedTarget"><number:fraction number:min-numerator-digits="1"/></number:number-style></office:automatic-styles>
  <office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Cell"><text:p>1</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:style="{STYLE_NS}" xmlns:number="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    deflated_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn package_content(source: &[u8]) -> Result<String> {
    Ok(Spreadsheet::from_bytes(source.to_vec())?
        .content_xml()
        .to_owned())
}

fn two_fraction_graph(first_digits: i64, second_digits: i64) -> Result<StyleGraphExtension> {
    let first = data_style::NumberBuilder::fraction("FirstFraction")?
        .min_numerator_digits(first_digits)?
        .build()?;
    let second = data_style::NumberBuilder::fraction("SecondFraction")?
        .min_denominator_digits(second_digits)?
        .build()?;
    Ok(StyleGraphExtension {
        number_styles: vec![first, second],
        data_styles: Vec::new(),
    })
}

fn same_name_owner_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="SharedName"/>
<number:number-style style:name="SharedName"><number:fraction/></number:number-style>
<number:date-style style:name="SharedName"/>"#,
        r#"<number:number-style style:name="SharedName"><number:scientific-number/></number:number-style>"#,
    )
}

fn aliased_owner_fixture() -> Vec<u8> {
    let content = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" xmlns:table="{TABLE_NS}" xmlns:text="{TEXT_NS}" office:version="1.3">
  <office:automatic-styles><s:style s:name="Cell" s:family="table-cell" s:data-style-name="Aliased"/>
    <n:number-style s:name="Aliased"><n:scientific-number n:decimal-places="2"/></n:number-style>
  </office:automatic-styles>
  <office:body><office:spreadsheet><table:table table:name="Sheet1"><table:table-row><table:table-cell table:style-name="Cell"><text:p>x</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body>
</office:document-content>"#
    );
    let styles = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="{OFFICE_NS}" xmlns:s="{STYLE_NS}" xmlns:n="{NUMBER_NS}" office:version="1.3"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>"#
    );
    support::raw_package(&[
        ("content.xml", content.as_bytes(), "text/xml"),
        ("styles.xml", styles.as_bytes(), "text/xml"),
    ])
}

fn legacy_owner_fixture() -> Vec<u8> {
    package_with_style_owners(
        r#"<number:date-style style:name="DateRead"/>
<number:time-style style:name="TimeRead"><number:hours/><number:text>:</number:text><number:minutes/><number:text>:</number:text><number:seconds number:decimal-places="3"/></number:time-style>
<number:currency-style style:name="CurrencyRead"><number:currency-symbol>$</number:currency-symbol><number:number number:decimal-places="2"/></number:currency-style>
<number:percentage-style style:name="PercentageRead"><number:number number:decimal-places="1"/><number:text>%</number:text></number:percentage-style>
<number:boolean-style style:name="BooleanRead"><number:boolean/></number:boolean-style>"#,
        "",
    )
}

fn graph_with_number(style: data_style::Number) -> StyleGraphExtension {
    StyleGraphExtension {
        number_styles: vec![style],
        data_styles: Vec::new(),
    }
}

fn auto_selector(name: &str, family: DataStyleFamily) -> DataStyleSelector<'_> {
    DataStyleSelector::automatic(name, family)
}

fn typed_number(
    snapshot: &Snapshot,
    selector: DataStyleSelector<'_>,
) -> Result<data_style::Number> {
    match snapshot.data_style(selector)? {
        data_style::Entry::Number(value) => Ok(value.value.clone()),
        other => Err(Error::InvalidFormat(format!(
            "expected typed number style, got {other:?}"
        ))),
    }
}

#[test]
fn native_packaged_fixture_contains_the_missing_number_particles() -> Result<()> {
    let spreadsheet = Spreadsheet::from_bytes(FORMATS_ODS.to_vec())?;
    let styles = spreadsheet
        .styles_xml()
        .ok_or_else(|| Error::InvalidFormat("native formats.ods has no styles.xml".to_owned()))?;
    assert!(styles.contains("number:scientific-number"));
    assert!(styles.contains("number:fraction"));
    assert!(styles.contains("loext:max-numerator-digits"));

    let content = spreadsheet.content_xml();
    assert!(content.contains("number:scientific-number"));
    assert!(content.contains("office:automatic-styles"));
    Ok(())
}

#[test]
fn native_flat_fixture_is_read_as_flat_and_retains_extension_lexemes() -> Result<()> {
    let flat = FlatSpreadsheet::from_bytes(YIELDDISC_FODS.to_vec())?;
    assert_eq!(flat.as_bytes(), YIELDDISC_FODS);
    let source = std::str::from_utf8(flat.as_bytes())
        .map_err(|error| Error::InvalidFormat(format!("yielddisc UTF-8: {error}")))?;
    assert!(source.contains("number:scientific-number"));
    assert!(source.contains("number:embedded-text"));
    assert!(source.contains("loext:min-decimal-places"));
    assert!(source.contains("loext:exponent-interval"));
    assert!(source.contains("loext:forced-exponent-sign"));
    Ok(())
}

#[test]
fn fraction_builder_covers_all_core_fraction_attributes() -> Result<()> {
    let fraction = data_style::Fraction {
        min_numerator_digits: Some(2),
        min_denominator_digits: Some(3),
        denominator_value: Some(16),
        // ODF ignores max-denominator-value when denominator-value is present.
        max_denominator_value: None,
        min_integer_digits: Some(1),
        grouping: Some(true),
    };
    let style = number_style("FractionAll", data_style::Format::Fraction(fraction))?;
    let xml = style.to_xml()?;
    for attribute in [
        "number:min-numerator-digits=\"2\"",
        "number:min-denominator-digits=\"3\"",
        "number:denominator-value=\"16\"",
        "number:min-integer-digits=\"1\"",
        "number:grouping=\"true\"",
    ] {
        assert!(xml.contains(attribute), "missing {attribute} in {xml}");
    }
    assert!(xml.contains("number:fraction"));
    Ok(())
}

#[test]
fn scientific_builder_covers_all_core_scientific_attributes_and_defaults() -> Result<()> {
    let scientific = data_style::Scientific {
        decimal_places: Some(4),
        min_decimal_places: Some(2),
        min_integer_digits: Some(1),
        grouping: Some(true),
        min_exponent_digits: Some(3),
        exponent_interval: Some(2),
        forced_exponent_sign: Some(false),
    };
    let style = number_style("ScientificAll", data_style::Format::Scientific(scientific))?;
    let xml = style.to_xml()?;
    for attribute in [
        "number:decimal-places=\"4\"",
        "number:min-decimal-places=\"2\"",
        "number:min-integer-digits=\"1\"",
        "number:grouping=\"true\"",
        "number:min-exponent-digits=\"3\"",
        "number:exponent-interval=\"2\"",
        "number:forced-exponent-sign=\"false\"",
    ] {
        assert!(xml.contains(attribute), "missing {attribute} in {xml}");
    }

    let defaults = data_style::Scientific::new();
    assert_eq!(defaults.effective_exponent_interval(), 1);
    assert!(defaults.effective_forced_exponent_sign());
    assert_eq!(
        defaults.resolve_decimal_places(None),
        data_style::Resolution::Unresolved
    );
    assert_eq!(
        defaults.resolve_decimal_places(Some(6)),
        data_style::Resolution::Inherited(6)
    );
    Ok(())
}

#[test]
fn decimal_embedded_text_is_shared_by_currency_and_percentage() -> Result<()> {
    let embedded = data_style::EmbeddedText::new(5, " ")?;
    let currency = data_style::DataBuilder::currency("CurrencyEmbedded", "€")?
        .embedded_text(embedded.clone())?
        .decimal_places(2)?
        .build()?;
    let percentage = data_style::DataBuilder::percentage("PercentageEmbedded")?
        .embedded_text(embedded)?
        .decimal_places(2)?
        .build()?;
    let currency_xml = currency.to_xml()?;
    let percentage_xml = percentage.to_xml()?;
    assert!(currency_xml.contains("number:currency-symbol"));
    assert!(currency_xml.contains("number:embedded-text"));
    assert!(percentage_xml.contains("number:embedded-text"));
    assert!(percentage_xml.contains("number:text"));
    Ok(())
}

#[test]
fn transliteration_metadata_accepts_unicode_digit_one_and_applies_defaults() -> Result<()> {
    let transliteration = data_style::Transliteration::new(
        Some("١".to_owned()),
        Some("ar".to_owned()),
        Some("EG".to_owned()),
        Some(data_style::TransliterationStyle::Long),
    )?;
    assert_eq!(transliteration.effective_format(), "١");
    assert_eq!(transliteration.effective_language(), Some("ar"));
    assert_eq!(transliteration.effective_country(), Some("EG"));
    assert_eq!(
        transliteration.effective_style(),
        data_style::TransliterationStyle::Long
    );

    let omitted = data_style::Transliteration::default();
    assert_eq!(omitted.effective_format(), "1");
    assert_eq!(omitted.effective_language(), None);
    assert_eq!(
        omitted.effective_style(),
        data_style::TransliterationStyle::Short
    );
    Ok(())
}

#[test]
fn common_attributes_preserve_omitted_fields_and_emit_transliteration() -> Result<()> {
    let attributes = data_style::Attributes::try_from_borrowed(
        Some("Display"),
        Some("en"),
        Some("US"),
        Some("Latn"),
        Some("en-US"),
        Some("Title"),
        Some(true),
        (
            Some("1"),
            Some("en"),
            Some("US"),
            Some(data_style::TransliterationStyle::Medium),
        ),
    )?;
    assert_eq!(attributes.display_name.as_deref(), Some("Display"));
    assert_eq!(attributes.language.as_deref(), Some("en"));
    assert_eq!(attributes.country.as_deref(), Some("US"));
    assert_eq!(attributes.script.as_deref(), Some("Latn"));
    assert_eq!(attributes.rfc_language_tag.as_deref(), Some("en-US"));
    assert_eq!(attributes.title.as_deref(), Some("Title"));
    assert_eq!(attributes.volatile, Some(true));
    assert_eq!(
        attributes.transliteration.style,
        Some(data_style::TransliterationStyle::Medium)
    );

    let number = data_style::NumberBuilder::scientific("MetadataNumber")?
        .attributes(attributes)
        .build()?;
    let xml = number.to_xml()?;
    assert!(xml.contains("style:display-name=\"Display\""));
    assert!(xml.contains("number:transliteration-format=\"1\""));
    assert!(xml.contains("number:transliteration-style=\"medium\""));
    Ok(())
}

#[test]
fn decimal_resolution_retains_explicit_inherited_and_unresolved_states() -> Result<()> {
    let explicit = data_style::Decimal {
        decimal_places: Some(2),
        ..data_style::Decimal::new()
    };
    assert_eq!(
        explicit.resolve_decimal_places(Some(8)),
        data_style::Resolution::Explicit(2)
    );
    let inherited = data_style::Decimal::new();
    assert_eq!(
        inherited.resolve_decimal_places(Some(8)),
        data_style::Resolution::Inherited(8)
    );
    assert_eq!(
        inherited.resolve_decimal_places(None),
        data_style::Resolution::Unresolved
    );
    Ok(())
}

#[test]
fn fraction_and_scientific_invalid_cross_field_values_are_refused() -> Result<()> {
    let decimal = data_style::Decimal {
        decimal_places: Some(1),
        min_decimal_places: Some(2),
        ..data_style::Decimal::new()
    };
    assert!(decimal.validate().is_err());

    let scientific = data_style::Scientific {
        exponent_interval: Some(0),
        ..data_style::Scientific::new()
    };
    assert!(scientific.validate().is_err());

    let fraction = data_style::Fraction {
        max_denominator_value: Some(0),
        ..data_style::Fraction::new()
    };
    assert!(fraction.validate().is_err());
    Ok(())
}

#[test]
fn particle_validation_rejects_invalid_positions_and_affix_order() -> Result<()> {
    assert!(data_style::EmbeddedText::new(0, "x").is_err());
    assert!(data_style::Affix::try_from_borrowed(None, None, Some("x")).is_err());
    assert!(data_style::Affix::try_from_borrowed(Some("a"), None, Some("b")).is_err());

    let leading = data_style::Affix::fill(".")?;
    let trailing = data_style::Affix::fill(",")?;
    let style = data_style::Number {
        name: "TwoFill".to_owned(),
        attributes: data_style::Attributes::new(),
        leading: Some(leading),
        format: Some(data_style::Format::Decimal(data_style::Decimal::new())),
        trailing: Some(trailing),
    };
    assert!(style.validate().is_err());
    Ok(())
}

#[test]
fn legacy_date_time_currency_percentage_and_boolean_values_remain_typed() -> Result<()> {
    let date = data_style::DataBuilder::date("DateLegacy")?.build()?;
    let time = data_style::DataBuilder::time("TimeLegacy")?
        .decimal_places(3)?
        .build()?;
    let currency = data_style::DataBuilder::currency("CurrencyLegacy", "$")?
        .decimal_places(2)?
        .build()?;
    let percentage = data_style::DataBuilder::percentage("PercentageLegacy")?
        .decimal_places(1)?
        .build()?;
    let boolean = data_style::DataBuilder::boolean("BooleanLegacy")?.build()?;

    assert!(date.to_xml()?.contains("number:date-style"));
    assert!(time.to_xml()?.contains("number:time-style"));
    assert!(currency.to_xml()?.contains("number:currency-style"));
    assert!(percentage.to_xml()?.contains("number:percentage-style"));
    assert!(boolean.to_xml()?.contains("number:boolean-style"));
    Ok(())
}

#[test]
fn package_catalog_preserves_legacy_data_style_families() -> Result<()> {
    let snapshot = Snapshot::from_bytes(legacy_owner_fixture())?;
    for (name, family) in [
        ("DateRead", DataStyleFamily::Date),
        ("TimeRead", DataStyleFamily::Time),
        ("CurrencyRead", DataStyleFamily::Currency),
        ("PercentageRead", DataStyleFamily::Percentage),
        ("BooleanRead", DataStyleFamily::Boolean),
    ] {
        let entry = snapshot.data_style(auto_selector(name, family))?;
        assert_eq!(entry.family(), family);
        assert_eq!(entry.owner(), DataStyleOwner::ContentAutomatic);
    }
    Ok(())
}

#[test]
fn graph_builder_rejects_duplicate_same_family_names_and_emits_all_nodes() -> Result<()> {
    let first = data_style::NumberBuilder::fraction("FractionGraph")?.build()?;
    let second = data_style::NumberBuilder::scientific("ScientificGraph")?.build()?;
    let mut builder = data_style::Graph::builder();
    builder.number_style(first)?;
    builder.number_style(second)?;
    let graph = builder.build()?;
    let xml = graph.to_xml()?;
    assert!(xml.contains("style:name=\"FractionGraph\""));
    assert!(xml.contains("style:name=\"ScientificGraph\""));

    let duplicate = data_style::NumberBuilder::decimal("Duplicate")?.build()?;
    let mut duplicate_builder = data_style::Graph::builder();
    duplicate_builder.number_style(duplicate.clone())?;
    assert!(duplicate_builder.number_style(duplicate).is_err());
    Ok(())
}

#[test]
fn duplicate_source_names_are_rejected_before_catalog_publication() -> Result<()> {
    let source = package_with_style_owners(
        r#"<number:number-style style:name="Duplicate"><number:fraction/></number:number-style>
<number:number-style style:name="Duplicate"><number:scientific-number/></number:number-style>"#,
        "",
    );
    let snapshot = Snapshot::from_bytes(source)?;
    assert!(
        snapshot
            .data_styles(DataStyleOwner::ContentAutomatic)
            .is_err()
    );
    Ok(())
}

#[test]
fn metadata_patch_is_atomic_and_preserves_unmentioned_fields() -> Result<()> {
    let mut attributes = data_style::Attributes::builder()
        .display_name("Original")?
        .language("en")?
        .volatile(true)
        .build()?;
    let original = attributes.clone();
    let patch = data_style::Patch::default()
        .set_display_name("Changed")?
        .set_transliteration_format("١")?;
    patch.apply(&mut attributes)?;
    assert_eq!(attributes.display_name.as_deref(), Some("Changed"));
    assert_eq!(attributes.language, original.language);
    assert_eq!(attributes.volatile, original.volatile);

    let before_failed = attributes.clone();
    let invalid_patch = data_style::Patch {
        transliteration_format: data_style::Op::Set("12".to_owned()),
        ..data_style::Patch::default()
    };
    assert!(invalid_patch.apply(&mut attributes).is_err());
    assert_eq!(attributes, before_failed);
    Ok(())
}

#[test]
fn native_fixture_helpers_expose_exact_member_bytes_for_future_owner_tests() -> Result<()> {
    let styles = formats_styles()?;
    let content = formats_content()?;
    assert_eq!(styles.len(), 23_029);
    assert_eq!(content.len(), 25_605);
    assert!(styles.contains("number:fraction"));
    assert!(content.contains("number:scientific-number"));
    Ok(())
}

#[test]
fn package_catalog_is_source_qualified_and_common_styles_are_read_only() -> Result<()> {
    let source = owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let automatic = snapshot.data_styles(DataStyleOwner::ContentAutomatic)?;
    assert!(automatic.iter().any(|entry| entry.name() == "AutoFraction"));
    assert!(automatic.iter().any(|entry| entry.name() == "OpaqueText"));
    assert!(matches!(
        snapshot.data_style(auto_selector("OpaqueText", DataStyleFamily::Text))?,
        data_style::Entry::Text(_)
    ));

    let common = snapshot.data_styles(DataStyleOwner::CommonStyles)?;
    assert!(common.iter().any(|entry| entry.name() == "SharedName"));
    assert_eq!(
        snapshot
            .data_style(auto_selector("AutoFraction", DataStyleFamily::Number))?
            .name(),
        "AutoFraction"
    );
    assert_eq!(
        snapshot
            .data_style(DataStyleSelector::common(
                "SharedName",
                DataStyleFamily::Number
            ))?
            .name(),
        "SharedName"
    );

    let mut edit = snapshot.edit();
    let patch = DataStyleAttributePatch::default().set_title("must-refuse")?;
    assert!(
        edit.patch_data_style(
            DataStyleSelector::common("SharedName", DataStyleFamily::Number),
            &patch,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn styles_automatic_owner_is_distinct_and_read_only() -> Result<()> {
    let source = three_owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;

    let content = snapshot.data_style(DataStyleSelector::automatic(
        "ThreeOwner",
        DataStyleFamily::Number,
    ))?;
    assert_eq!(content.owner(), DataStyleOwner::ContentAutomatic);
    assert_eq!(content.family(), DataStyleFamily::Number);

    let common = snapshot.data_style(DataStyleSelector::common(
        "ThreeOwner",
        DataStyleFamily::Number,
    ))?;
    assert_eq!(common.owner(), DataStyleOwner::CommonStyles);

    let styles_automatic = snapshot.data_style(DataStyleSelector::styles_automatic(
        "ThreeOwner",
        DataStyleFamily::Number,
    ))?;
    assert_eq!(styles_automatic.owner(), DataStyleOwner::StylesAutomatic);
    assert_eq!(styles_automatic.family(), DataStyleFamily::Number);
    assert!(
        snapshot
            .data_styles(DataStyleOwner::StylesAutomatic)?
            .iter()
            .any(|entry| entry.name() == "ThreeOwner")
    );

    let patch = DataStyleAttributePatch::default().set_title("read-only")?;
    let mut edit = snapshot.edit();
    assert!(
        edit.patch_data_style(
            DataStyleSelector::styles_automatic("ThreeOwner", DataStyleFamily::Number),
            &patch,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn data_style_patch_distinguishes_empty_set_clear_and_all_keep_noop() -> Result<()> {
    let source = lexical_metadata_fixture();
    let selector = auto_selector("Lexical", DataStyleFamily::Number);
    let snapshot = Snapshot::from_bytes(source.clone())?;

    let mut set_attributes = data_style::Attributes::builder().title("old")?.build()?;
    data_style::Patch::default()
        .set_title("")?
        .apply(&mut set_attributes)?;
    assert_eq!(set_attributes.title.as_deref(), Some(""));

    let mut clear_attributes = data_style::Attributes::builder().title("old")?.build()?;
    data_style::Patch::default()
        .clear_title()
        .apply(&mut clear_attributes)?;
    assert_eq!(clear_attributes.title, None);

    let mut set_edit = snapshot.edit();
    set_edit.patch_data_style(selector, &data_style::Patch::default().set_title("")?)?;
    let set_xml = package_content(set_edit.as_bytes())?;
    assert!(set_xml.contains("number:title=\"\""), "{set_xml}");

    let mut clear_edit = snapshot.edit();
    clear_edit.patch_data_style(selector, &data_style::Patch::default().clear_title())?;
    let clear_xml = package_content(clear_edit.as_bytes())?;
    assert!(!clear_xml.contains("number:title="));

    let mut noop_edit = snapshot.edit();
    noop_edit.patch_data_style(selector, &data_style::Patch::default())?;
    assert_eq!(noop_edit.as_bytes(), source.as_slice());
    let noop_commit = noop_edit.commit()?;
    assert!(!noop_commit.changed());
    assert_eq!(noop_commit.snapshot().as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn source_metadata_patch_preserves_decoded_equal_attribute_spelling() -> Result<()> {
    let source = lexical_metadata_fixture();
    let mut edit = Snapshot::from_bytes(source.clone())?.edit();
    let patch = data_style::Patch::default().set_volatile(true);
    edit.patch_data_style(auto_selector("Lexical", DataStyleFamily::Number), &patch)?;
    let content = package_content(edit.as_bytes())?;
    assert!(content.contains(r#"number:title="A&#x20;B""#));
    assert!(content.contains(r#"style:display-name="A&#32;B""#));
    assert!(content.contains(r#"style:volatile="true""#));
    Ok(())
}

#[test]
fn inherited_prefix_metadata_patch_preserves_style_alias_without_redundant_bindings() -> Result<()>
{
    let source = inherited_prefix_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let entry = snapshot.data_style(auto_selector("Inherited", DataStyleFamily::Number))?;
    match entry {
        data_style::Entry::Number(value) => {
            assert_eq!(
                value.value.attributes.display_name.as_deref(),
                Some("Display")
            );
            assert_eq!(value.value.attributes.volatile, Some(false));
        },
        other => panic!("unexpected inherited-prefix entry: {other:?}"),
    }

    let mut edit = snapshot.edit();
    let patch = data_style::Patch::default().set_title("Inherited title")?;
    edit.patch_data_style(auto_selector("Inherited", DataStyleFamily::Number), &patch)?;
    let content = package_content(edit.as_bytes())?;
    let start = content.find("<n:number-style").ok_or_else(|| {
        Error::InvalidFormat("inherited number-style opening is missing".to_owned())
    })?;
    let end = content[start..]
        .find('>')
        .map(|offset| start + offset)
        .ok_or_else(|| {
            Error::InvalidFormat("inherited number-style opening is unterminated".to_owned())
        })?;
    let opening = &content[start..=end];
    assert!(opening.contains(r#"n:title="Inherited title""#));
    assert!(!opening.contains("xmlns:s="));
    assert!(!opening.contains("xmlns:n="));
    Ok(())
}

#[test]
fn explicit_xml_namespace_bindings_literal_and_escaped_are_read_and_retained() -> Result<()> {
    for escaped in [false, true] {
        let source = explicit_xml_namespace_fixture(escaped);
        let snapshot = Snapshot::from_bytes(source.clone())?;
        assert_eq!(
            snapshot
                .data_style(auto_selector("ExplicitXml", DataStyleFamily::Number))?
                .name(),
            "ExplicitXml"
        );
        let content = package_content(&source)?;
        if escaped {
            assert!(content.contains(
                r#"xmlns:xml="http:&#x2F;&#x2F;www.w3.org&#x2F;XML&#x2F;1998&#x2F;namespace""#
            ));
        } else {
            assert!(content.contains(r#"xmlns:xml="http://www.w3.org/XML/1998/namespace""#));
        }
        let mut edit = snapshot.edit();
        edit.patch_data_style(
            auto_selector("ExplicitXml", DataStyleFamily::Number),
            &data_style::Patch::default(),
        )?;
        assert_eq!(edit.as_bytes(), source.as_slice());
    }
    Ok(())
}

#[test]
fn number_style_short_and_long_are_valid_but_bogus_is_rejected() -> Result<()> {
    for style in ["short", "long"] {
        let snapshot = Snapshot::from_bytes(date_style_with_number_style(style))?;
        assert_eq!(
            snapshot
                .data_style(auto_selector("DateStyle", DataStyleFamily::Date))?
                .family(),
            DataStyleFamily::Date
        );
    }
    // Data-style schema projection is lazy; the XML package itself is well formed.
    let malformed = Snapshot::from_bytes(date_style_with_number_style("bogus"))?;
    assert!(
        malformed
            .data_style(auto_selector("DateStyle", DataStyleFamily::Date))
            .is_err()
    );
    Ok(())
}

#[test]
fn number_text_accepts_xml_tab_linefeed_and_carriage_return() -> Result<()> {
    let snapshot = Snapshot::from_bytes(whitespace_number_text_fixture())?;
    assert!(matches!(
        snapshot.data_style(auto_selector("WhitespaceText", DataStyleFamily::Number))?,
        data_style::Entry::Number(_)
    ));
    Ok(())
}

#[test]
fn token_whitespace_language_collapses_for_semantic_noop() -> Result<()> {
    let source = token_whitespace_language_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let entry = snapshot.data_style(auto_selector("TokenLanguage", DataStyleFamily::Number))?;
    match entry {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.language.as_deref(), Some("en"));
        },
        other => panic!("unexpected token-language entry: {other:?}"),
    }
    let mut edit = snapshot.edit();
    let patch = data_style::Patch::default().set_language("en")?;
    edit.patch_data_style(
        auto_selector("TokenLanguage", DataStyleFamily::Number),
        &patch,
    )?;
    assert_eq!(edit.as_bytes(), source.as_slice());
    assert!(!edit.commit()?.changed());
    Ok(())
}

#[test]
fn data_style_owner_queries_reject_opposite_document_roots() -> Result<()> {
    assert!(Snapshot::from_bytes(content_document_styles_root_fixture()).is_err());

    let source = styles_document_content_root_fixture();
    let snapshot = Snapshot::from_bytes(source)?;
    assert!(
        snapshot
            .data_style(DataStyleSelector::common(
                "WrongRoot",
                DataStyleFamily::Number,
            ))
            .is_err()
    );
    assert!(
        snapshot
            .data_style(DataStyleSelector::styles_automatic(
                "WrongRoot",
                DataStyleFamily::Number,
            ))
            .is_err()
    );
    Ok(())
}

#[test]
fn literal_attribute_whitespace_projects_to_spaces_and_set_equal_is_exact_noop() -> Result<()> {
    let source = literal_whitespace_metadata_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let entry = snapshot.data_style(auto_selector("LiteralWhitespace", DataStyleFamily::Number))?;
    match entry {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.title.as_deref(), Some("A B C D"));
            assert_eq!(
                value.value.attributes.display_name.as_deref(),
                Some("E F G H")
            );
        },
        other => panic!("unexpected literal-whitespace entry: {other:?}"),
    }

    let patch = data_style::Patch::default()
        .set_title("A B C D")?
        .set_display_name("E F G H")?;
    let mut edit = snapshot.edit();
    edit.patch_data_style(
        auto_selector("LiteralWhitespace", DataStyleFamily::Number),
        &patch,
    )?;
    assert_eq!(edit.as_bytes(), source.as_slice());
    assert!(!edit.commit()?.changed());
    Ok(())
}

#[test]
fn schema_string_metadata_preserves_repeated_spaces_and_charref_tabs() -> Result<()> {
    let source = schema_string_charref_metadata_fixture();
    let selector = auto_selector("SchemaStrings", DataStyleFamily::Number);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let entry = snapshot.data_style(selector)?;
    match entry {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.title.as_deref(), Some("a  b"));
            assert_eq!(value.value.attributes.display_name.as_deref(), Some("a\tb"));
        },
        other => panic!("unexpected schema-string entry: {other:?}"),
    }

    // The same decoded values are an exact source no-op, including the
    // character-reference spelling that carries the tab through XML 1.0
    // attribute normalization.
    let same_values = data_style::Patch::default()
        .set_title("a  b")?
        .set_display_name("a\tb")?;
    let mut noop = snapshot.edit();
    noop.patch_data_style(selector, &same_values)?;
    assert_eq!(noop.as_bytes(), source.as_slice());
    assert!(!noop.commit()?.changed());

    // Collapsing either value is a real semantic edit.  It must not be
    // treated as an xsd:token-style whitespace-equivalent update.
    let changed_values = data_style::Patch::default()
        .set_title("a b")?
        .set_display_name("a b")?;
    let mut changed = snapshot.edit();
    changed.patch_data_style(selector, &changed_values)?;
    assert_ne!(changed.as_bytes(), source.as_slice());
    let reopened = Snapshot::from_bytes(changed.as_bytes().to_vec())?;
    match reopened.data_style(selector)? {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.title.as_deref(), Some("a b"));
            assert_eq!(value.value.attributes.display_name.as_deref(), Some("a b"));
        },
        other => panic!("unexpected changed schema-string entry: {other:?}"),
    }
    Ok(())
}

#[test]
fn literal_crlf_metadata_projects_to_one_space_and_equal_set_is_exact_noop() -> Result<()> {
    let source = crlf_string_metadata_fixture();
    let selector = auto_selector("CrlfStrings", DataStyleFamily::Number);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    match snapshot.data_style(selector)? {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.title.as_deref(), Some("a b"));
            assert_eq!(value.value.attributes.display_name.as_deref(), Some("a b"));
        },
        other => panic!("unexpected CRLF schema-string entry: {other:?}"),
    }

    let patch = data_style::Patch::default()
        .set_title("a b")?
        .set_display_name("a b")?;
    let mut edit = snapshot.edit();
    edit.patch_data_style(selector, &patch)?;
    assert_eq!(edit.as_bytes(), source.as_slice());
    assert!(!edit.commit()?.changed());
    Ok(())
}

#[test]
fn setting_schema_string_controls_uses_numeric_refs_and_inverse_preserves_values() -> Result<()> {
    let source = plain_string_metadata_fixture();
    let selector = auto_selector("ControlStrings", DataStyleFamily::Number);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let controls = "A\tB\nC\rD";
    let patch = data_style::Patch::default()
        .set_title(controls)?
        .set_display_name(controls)?;
    let mut edit = snapshot.edit();
    edit.patch_data_style(selector, &patch)?;
    let commit = edit.commit()?;
    let changed = commit.snapshot();
    let content = package_content(changed.as_bytes())?;
    assert!(
        content.contains("number:title=\"A&#x9;B&#xA;C&#xD;D\""),
        "{content}"
    );
    assert!(
        content.contains("style:display-name=\"A&#x9;B&#xA;C&#xD;D\""),
        "{content}"
    );
    match changed.data_style(selector)? {
        data_style::Entry::Number(value) => {
            assert_eq!(value.value.attributes.title.as_deref(), Some(controls));
            assert_eq!(
                value.value.attributes.display_name.as_deref(),
                Some(controls)
            );
        },
        other => panic!("unexpected control-string entry: {other:?}"),
    }

    let restored = commit.patch().inverse().apply(changed)?;
    assert_eq!(restored.snapshot().as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn malformed_styles_xml_outside_selected_owner_fails_on_style_query() -> Result<()> {
    let mut failures = Vec::new();
    for fragment in [
        "<!-- malformed -- comment -->",
        "<?xml version=\"1.0\"?>",
        "&illegalGeneralRef;",
        "<office:body>&#x1;</office:body>",
        "<office:body>&#1;</office:body>",
    ] {
        let snapshot = Snapshot::from_bytes(malformed_styles_query_fixture(fragment))?;
        let common_rejected = snapshot
            .data_style(DataStyleSelector::common("Common", DataStyleFamily::Number))
            .is_err();
        let styles_automatic_rejected = snapshot
            .data_style(DataStyleSelector::styles_automatic(
                "Automatic",
                DataStyleFamily::Number,
            ))
            .is_err();
        if !common_rejected || !styles_automatic_rejected {
            failures.push(format!(
                "{fragment:?}: common_rejected={common_rejected}, styles_automatic_rejected={styles_automatic_rejected}"
            ));
        }
    }
    assert!(failures.is_empty(), "{failures:?}");
    Ok(())
}

#[test]
fn effective_cell_data_style_resolves_the_content_automatic_owner() -> Result<()> {
    let snapshot = Snapshot::from_bytes(owner_fixture())?;
    let selected = snapshot
        .effective_cell_data_style("Cell")?
        .ok_or_else(|| Error::InvalidFormat("Cell has no data-style reference".to_owned()))?;
    assert_eq!(selected.name(), "AutoFraction");
    assert_eq!(selected.family(), DataStyleFamily::Number);
    assert_eq!(selected.owner(), DataStyleOwner::ContentAutomatic);
    Ok(())
}

#[test]
fn same_name_owners_and_families_require_explicit_selectors() -> Result<()> {
    let snapshot = Snapshot::from_bytes(same_name_owner_fixture())?;
    let automatic_number = snapshot.data_style(DataStyleSelector::automatic(
        "SharedName",
        DataStyleFamily::Number,
    ))?;
    assert_eq!(automatic_number.owner(), DataStyleOwner::ContentAutomatic);
    assert_eq!(automatic_number.family(), DataStyleFamily::Number);
    let common_number = snapshot.data_style(DataStyleSelector::common(
        "SharedName",
        DataStyleFamily::Number,
    ))?;
    assert_eq!(common_number.owner(), DataStyleOwner::CommonStyles);
    assert_eq!(common_number.family(), DataStyleFamily::Number);
    assert!(snapshot.effective_cell_data_style("Cell").is_err());
    Ok(())
}

#[test]
fn namespace_aliases_survive_source_qualified_metadata_patch() -> Result<()> {
    let source = aliased_owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    assert_eq!(
        snapshot
            .data_style(auto_selector("Aliased", DataStyleFamily::Number))?
            .name(),
        "Aliased"
    );
    let mut edit = snapshot.edit();
    let patch = DataStyleAttributePatch::default().set_title("alias-title")?;
    edit.patch_data_style(auto_selector("Aliased", DataStyleFamily::Number), &patch)?;
    let after = Spreadsheet::from_bytes(edit.as_bytes().to_vec())?
        .content_xml()
        .to_owned();
    assert!(
        after
            .as_bytes()
            .windows(b"xmlns:s".len())
            .any(|window| window == b"xmlns:s")
    );
    assert!(
        after
            .as_bytes()
            .windows(b"n:number-style".len())
            .any(|window| window == b"n:number-style")
    );
    let commit = edit.commit()?;
    assert_eq!(
        commit
            .snapshot()
            .data_style(auto_selector("Aliased", DataStyleFamily::Number))?
            .name(),
        "Aliased"
    );
    Ok(())
}

#[test]
fn extended_graph_add_replace_and_patch_reopen_through_the_public_transaction() -> Result<()> {
    let snapshot = Snapshot::from_bytes(owner_fixture())?;
    let added = data_style::NumberBuilder::scientific("AddedScientific")?
        .decimal_places(4)?
        .min_exponent_digits(2)?
        .exponent_interval(3)?
        .forced_exponent_sign(false)?
        .build()?;
    let replacement = data_style::NumberBuilder::fraction("AutoFraction")?
        .min_numerator_digits(3)?
        .min_denominator_digits(4)?
        .max_denominator_value(999)?
        .build()?;

    let mut edit = snapshot.edit();
    edit.put_extended_style_graph(&graph_with_number(added))?;
    edit.replace_extended_style_graph(
        auto_selector("AutoFraction", DataStyleFamily::Number),
        &graph_with_number(replacement),
    )?;
    let metadata_patch = DataStyleAttributePatch::default().set_title("edited")?;
    edit.patch_data_style(
        auto_selector("OpaqueText", DataStyleFamily::Text),
        &metadata_patch,
    )?;
    let pending_text = package_content(edit.as_bytes())?;
    assert!(pending_text.contains("AddedScientific"));
    let commit = edit.commit()?;
    assert!(commit.changed());
    let reopened = commit.snapshot();
    assert_eq!(
        reopened
            .data_style(auto_selector("AddedScientific", DataStyleFamily::Number))?
            .name(),
        "AddedScientific"
    );
    assert_eq!(
        reopened
            .data_style(auto_selector("AutoFraction", DataStyleFamily::Number))?
            .name(),
        "AutoFraction"
    );
    let content = package_content(reopened.as_bytes())?;
    assert!(
        content
            .as_bytes()
            .windows(b"opaque comment".len())
            .any(|window| { window == b"opaque comment" })
    );
    Ok(())
}

#[test]
fn identical_typed_graph_replacement_is_an_exact_source_noop() -> Result<()> {
    let source = lexical_metadata_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let selector = auto_selector("Lexical", DataStyleFamily::Number);
    let graph = StyleGraphExtension {
        number_styles: vec![typed_number(&snapshot, selector)?],
        data_styles: Vec::new(),
    };

    let mut edit = snapshot.edit();
    edit.replace_extended_style_graph(selector, &graph)?;
    assert_eq!(edit.as_bytes(), source.as_slice());
    let commit = edit.commit()?;
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn mixed_graph_replacement_preserves_unchanged_lexical_member_bytes() -> Result<()> {
    let source = multi_lexical_graph_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let first_selector = auto_selector("FirstFraction", DataStyleFamily::Number);
    let second_selector = auto_selector("SecondFraction", DataStyleFamily::Number);
    let first = typed_number(&snapshot, first_selector)?;
    let replacement = data_style::NumberBuilder::fraction("SecondFraction")?
        .min_denominator_digits(7)?
        .build()?;
    let graph = StyleGraphExtension {
        number_styles: vec![first, replacement],
        data_styles: Vec::new(),
    };

    let mut edit = snapshot.edit();
    edit.replace_extended_style_graph(second_selector, &graph)?;
    let content = String::from_utf8(zip_member(edit.as_bytes(), "content.xml")?)
        .map_err(|error| Error::InvalidFormat(format!("edited content.xml UTF-8: {error}")))?;
    assert!(
        content.contains(
            "<n:number-style s:name='FirstFraction' n:title='A&#x20;B' s:display-name='A&#32;B'><n:fraction n:min-numerator-digits='2'/></n:number-style>"
        ),
        "unchanged lexical member was rewritten: {content}"
    );
    assert!(content.contains(
        "<number:number-style xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" style:name=\"SecondFraction\"><number:fraction number:min-denominator-digits=\"7\"/></number:number-style>"
    ));
    let reopened = Snapshot::from_bytes(edit.as_bytes().to_vec())?;
    assert!(reopened.data_style(second_selector).is_ok());
    Ok(())
}

#[test]
fn compressed_late_content_replacement_uses_content_xml_ranges() -> Result<()> {
    let source = compressed_late_graph_fixture();
    let content = zip_member(&source, "content.xml")?;
    let marker = b"style:name=\"CompressedTarget\"";
    let target_offset = content
        .windows(marker.len())
        .position(|window| window == marker)
        .ok_or_else(|| Error::InvalidFormat("compressed target marker is missing".to_owned()))?;
    assert!(content.len() > source.len());
    assert!(target_offset + marker.len() > source.len());

    let selector = auto_selector("CompressedTarget", DataStyleFamily::Number);
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let replacement = data_style::NumberBuilder::fraction("CompressedTarget")?
        .min_numerator_digits(2)?
        .build()?;
    let mut edit = snapshot.edit();
    edit.replace_extended_style_graph(selector, &graph_with_number(replacement))?;
    let commit = edit.commit()?;
    assert!(commit.changed());

    let changed_content = zip_member(commit.snapshot().as_bytes(), "content.xml")?;
    assert!(
        changed_content
            .windows(b"min-numerator-digits=\"2\"".len())
            .any(|window| window == b"min-numerator-digits=\"2\"")
    );
    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn extended_graph_replacement_applies_every_member_atomically() -> Result<()> {
    let snapshot = Snapshot::from_bytes(multi_number_owner_fixture())?;
    let mut edit = snapshot.edit();
    let graph = two_fraction_graph(5, 7)?;
    edit.replace_extended_style_graph(
        auto_selector("FirstFraction", DataStyleFamily::Number),
        &graph,
    )?;

    let pending = package_content(edit.as_bytes())?;
    assert!(pending.contains(
        r#"<number:number-style xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="FirstFraction"><number:fraction number:min-numerator-digits="5"/></number:number-style>"#
    ));
    assert!(pending.contains(
        r#"<number:number-style xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="SecondFraction"><number:fraction number:min-denominator-digits="7"/></number:number-style>"#
    ));

    let commit = edit.commit()?;
    let reopened = commit.snapshot();
    let first = reopened.data_style(auto_selector("FirstFraction", DataStyleFamily::Number))?;
    let second = reopened.data_style(auto_selector("SecondFraction", DataStyleFamily::Number))?;
    match first {
        data_style::Entry::Number(value) => match value.value.format {
            Some(data_style::Format::Fraction(fraction)) => {
                assert_eq!(fraction.min_numerator_digits, Some(5));
            },
            other => panic!("unexpected FirstFraction body: {other:?}"),
        },
        other => panic!("unexpected FirstFraction entry: {other:?}"),
    }
    match second {
        data_style::Entry::Number(value) => match value.value.format {
            Some(data_style::Format::Fraction(fraction)) => {
                assert_eq!(fraction.min_denominator_digits, Some(7));
            },
            other => panic!("unexpected SecondFraction body: {other:?}"),
        },
        other => panic!("unexpected SecondFraction entry: {other:?}"),
    }
    Ok(())
}

#[test]
fn extended_graph_replacement_refuses_missing_member_without_partial_change() -> Result<()> {
    let source = multi_number_owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let first = data_style::NumberBuilder::fraction("FirstFraction")?
        .min_numerator_digits(5)?
        .build()?;
    let missing = data_style::NumberBuilder::fraction("MissingFraction")?
        .min_denominator_digits(7)?
        .build()?;
    let graph = StyleGraphExtension {
        number_styles: vec![first, missing],
        data_styles: Vec::new(),
    };
    let before = edit.as_bytes().to_vec();
    assert!(
        edit.replace_extended_style_graph(
            auto_selector("FirstFraction", DataStyleFamily::Number),
            &graph,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before.as_slice());
    assert_eq!(package_content(edit.as_bytes())?, package_content(&source)?);
    Ok(())
}

#[test]
fn extended_graph_replacement_refuses_body_kind_mismatch_without_partial_change() -> Result<()> {
    let source = multi_number_owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let mismatch = data_style::NumberBuilder::scientific("FirstFraction")?
        .decimal_places(3)?
        .build()?;
    let second = data_style::NumberBuilder::fraction("SecondFraction")?
        .min_denominator_digits(7)?
        .build()?;
    let graph = StyleGraphExtension {
        number_styles: vec![mismatch, second],
        data_styles: Vec::new(),
    };
    let before = edit.as_bytes().to_vec();
    assert!(
        edit.replace_extended_style_graph(
            auto_selector("FirstFraction", DataStyleFamily::Number),
            &graph,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before.as_slice());
    assert_eq!(package_content(edit.as_bytes())?, package_content(&source)?);
    Ok(())
}

#[test]
fn extended_graph_replacement_requires_selector_to_be_a_graph_member() -> Result<()> {
    let source = multi_number_owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let graph = two_fraction_graph(5, 7)?;
    let before = edit.as_bytes().to_vec();
    assert!(
        edit.replace_extended_style_graph(
            auto_selector("NotInReplacementGraph", DataStyleFamily::Number),
            &graph,
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), before.as_slice());
    assert_eq!(package_content(edit.as_bytes())?, package_content(&source)?);
    Ok(())
}

#[test]
fn extended_graph_empty_put_is_an_exact_noop() -> Result<()> {
    let source = owner_fixture();
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    edit.put_extended_style_graph(&StyleGraphExtension::default())?;
    assert_eq!(edit.as_bytes(), source.as_slice());
    let commit = edit.commit()?;
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn self_closing_automatic_styles_insertion_preserves_wrapper_and_exact_caps() -> Result<()> {
    let source = self_closing_automatic_fixture();
    let mut title = String::new();
    for index in 0..8_192 {
        write!(title, "{index:04x}")
            .map_err(|_| Error::InvalidFormat("test title formatting failed".to_owned()))?;
    }
    let attributes = data_style::Attributes::builder().title(&title)?.build()?;
    let graph = graph_with_number(
        data_style::NumberBuilder::scientific("SelfClosingScientific")?
            .decimal_places(4)?
            .attributes(attributes)
            .build()?,
    );

    let mut expected_edit = Snapshot::from_bytes(source.clone())?.edit();
    expected_edit.put_extended_style_graph(&graph)?;
    let expected = expected_edit.as_bytes().to_vec();
    assert!(expected.len() > source.len());
    let expected_content = package_content(&expected)?;
    assert!(expected_content.contains("<o:automatic-styles"));
    assert!(expected_content.contains(r#"xml:lang = "fr""#));
    assert!(expected_content.contains("SelfClosingScientific"));

    let defaults = litchi_ods::document::Limits::default();
    let output_cap = source.len().max(expected_content.len());
    let exact_limits = litchi_ods::document::Limits::new(
        output_cap,
        defaults.max_resources(),
        defaults.max_resource_bytes(),
        defaults.patch(),
        defaults.composition(),
        defaults.history(),
    );
    let mut exact_edit = Snapshot::from_bytes_with(source.clone(), exact_limits)?.edit();
    exact_edit.put_extended_style_graph(&graph)?;
    assert_eq!(exact_edit.as_bytes(), expected.as_slice());

    let under_limits = litchi_ods::document::Limits::new(
        output_cap.saturating_sub(1),
        defaults.max_resources(),
        defaults.max_resource_bytes(),
        defaults.patch(),
        defaults.composition(),
        defaults.history(),
    );
    let mut under_edit = Snapshot::from_bytes_with(source.clone(), under_limits)?.edit();
    assert!(under_edit.put_extended_style_graph(&graph).is_err());
    assert_eq!(under_edit.as_bytes(), source.as_slice());
    Ok(())
}

#[test]
fn unsupported_source_body_refuses_replacement_without_losing_candidate_state() -> Result<()> {
    let source = package_with_style_owners(
        r#"<style:style style:name="Cell" style:family="table-cell" style:data-style-name="Unsupported"/>
<number:number-style style:name="Unsupported"><number:scientific-number number:decimal-places="2"/><number:map style:condition="true" style:apply-style-name="x"/></number:number-style>"#,
        "",
    );
    let snapshot = Snapshot::from_bytes(source.clone())?;
    let mut edit = snapshot.edit();
    let pending = edit.as_bytes().to_vec();
    let replacement = data_style::NumberBuilder::decimal("Unsupported")?
        .decimal_places(2)?
        .build()?;
    assert!(
        edit.replace_extended_style_graph(
            auto_selector("Unsupported", DataStyleFamily::Number),
            &graph_with_number(replacement),
        )
        .is_err()
    );
    assert_eq!(edit.as_bytes(), pending.as_slice());
    Ok(())
}

#[test]
fn public_patch_inverse_and_stale_source_checks_are_exact() -> Result<()> {
    let snapshot = Snapshot::from_bytes(owner_fixture())?;
    let mut edit = snapshot.edit();
    let patch = DataStyleAttributePatch::default().set_title("round-trip")?;
    edit.patch_data_style(auto_selector("OpaqueText", DataStyleFamily::Text), &patch)?;
    let commit = edit.commit()?;
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), snapshot.as_bytes());

    let mut unrelated = snapshot.edit();
    unrelated.set_cell_style("Sheet1", 0, 0, "Cell")?;
    let changed = unrelated.commit()?.into_snapshot();
    assert!(commit.patch().apply(&changed).is_err());
    Ok(())
}

#[test]
fn exact_package_limit_refuses_graph_growth_before_staging() -> Result<()> {
    let source = owner_fixture();
    let defaults = litchi_ods::document::Limits::default();
    let limits = litchi_ods::document::Limits::new(
        source.len(),
        defaults.max_resources(),
        defaults.max_resource_bytes(),
        defaults.patch(),
        defaults.composition(),
        defaults.history(),
    );
    let snapshot = Snapshot::from_bytes_with(source.clone(), limits)?;
    let mut edit = snapshot.edit();
    let mut title = String::new();
    for index in 0..16_384 {
        write!(title, "{index:04x}")
            .map_err(|_| Error::InvalidFormat("test title formatting failed".to_owned()))?;
    }
    let attributes = data_style::Attributes::builder().title(&title)?.build()?;
    let style = data_style::NumberBuilder::scientific("TooLargeForPackage")?
        .decimal_places(6)?
        .attributes(attributes)
        .build()?;
    assert!(
        edit.put_extended_style_graph(&graph_with_number(style))
            .is_err()
    );
    assert_eq!(edit.as_bytes(), source.as_slice());
    assert!(!edit.commit()?.changed());
    Ok(())
}
