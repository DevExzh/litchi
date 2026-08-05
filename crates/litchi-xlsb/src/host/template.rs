//! Deterministic producer templates for newly authored XLSB packages.

/// Build the extended-properties part for the authored sheet inventory.
pub(crate) fn app(sheet_count: usize, sheet_names_xml: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>The Litchi Rust Library</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Sheet</vt:lpstr></vt:variant><vt:variant><vt:i4>{sheet_count}</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="{sheet_count}" baseType="lpstr">{sheet_names_xml}</vt:vector></TitlesOfParts><Company/><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>"#
    )
}

/// Return deterministic core properties for a newly authored package.
pub(crate) fn core() -> &'static str {
    include_str!("resources/generated/docProps/core.xml")
}

/// Return the default Office theme used by newly authored workbooks.
pub(crate) fn theme() -> &'static str {
    include_str!("resources/generated/theme/theme1.xml")
}
