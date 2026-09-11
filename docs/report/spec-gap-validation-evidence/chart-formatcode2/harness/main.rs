use std::{error::Error, fs, path::PathBuf};

use litchi_drawingml::chart::extension::formatcode2::{self, Element};

fn main() -> Result<(), Box<dyn Error>> {
    let output = PathBuf::from(std::env::args().nth(1).ok_or("missing output directory")?);
    fs::create_dir_all(&output)?;
    let source = br#"<f:formatcode2 xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart"><![CDATA[General]]><!--keep--></f:formatcode2>"#;
    let mut parsed = formatcode2::read(source)?;
    assert_eq!(formatcode2::write(&parsed)?, source);
    fs::write(output.join("source.xml"), formatcode2::write(&parsed)?)?;
    parsed.set_value("[$-fr-FR]#,##0.00")?;
    let edited = formatcode2::write(&parsed)?;
    assert_eq!(formatcode2::read(&edited)?.value(), parsed.value());
    assert!(
        edited
            .windows(b"<!--keep-->".len())
            .any(|v| v == b"<!--keep-->")
    );
    fs::write(output.join("edited.xml"), edited)?;
    let control = Element::new("literal _x0041_; control \u{1}; CR\rLF\n")?;
    let authored = formatcode2::write(&control)?;
    assert_eq!(formatcode2::read(&authored)?.value(), control.value());
    fs::write(output.join("authored.xml"), authored)?;

    let tag = br#"<c:numFmt xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:f="http://schemas.microsoft.com/office/drawing/2015/06/chart" formatCode="General" f:formatcode2='General' sourceLinked="0"/>"#;
    let mut attribute = formatcode2::read_attribute(tag)?;
    assert_eq!(formatcode2::write_attribute(&attribute)?, tag);
    attribute.set_value("[$-en-US]0.00 & 'quoted'\t\r\n")?;
    let edited = formatcode2::write_attribute(&attribute)?;
    assert_eq!(
        formatcode2::read_attribute(&edited)?.value(),
        attribute.value()
    );
    fs::write(output.join("attribute-start-tag.xml"), edited)?;
    let inherited = br#"<c:numFmt f:formatcode2='General' x:opaque='keep'/>"#;
    let bindings = [
        (
            "c",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ),
        ("f", formatcode2::NAMESPACE),
        ("x", "urn:opaque"),
    ];
    let mut projected = formatcode2::read_attribute_with_bindings(inherited, &bindings)?;
    assert_eq!(formatcode2::write_attribute(&projected)?, inherited);
    projected.set_value("0.00")?;
    let edited = formatcode2::write_attribute(&projected)?;
    assert_eq!(
        formatcode2::read_attribute_with_bindings(&edited, &bindings)?.value(),
        "0.00"
    );
    fs::write(output.join("attribute-inherited-start-tag.xml"), edited)?;
    println!(
        "element no-op, opaque comment, edit, escaped semantic value, attribute edit and inherited context: passed"
    );
    Ok(())
}
