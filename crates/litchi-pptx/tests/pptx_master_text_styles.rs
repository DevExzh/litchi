#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::collections::BTreeMap;

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::Package;
use quick_xml::Reader;
use quick_xml::events::Event;

const MASTER_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/master-text-styles/master.xml");

#[test]
fn master_text_style_inventories_are_exposed_by_the_owner_part() {
    let package = package_with_master_xml();
    let presentation = package.presentation().unwrap();
    let master = presentation.slide_masters().unwrap().remove(0);
    let styles = parse_text_styles(master.part().part().blob());

    // The standalone master facade currently exposes graph, shape, and theme
    // semantics; txStyles remains a borrowed wire inventory at this boundary.
    assert!(!styles["titleStyle"].0);
    assert_eq!(styles["titleStyle"].1, [1, 2]);
    assert!(styles["bodyStyle"].0);
    assert_eq!(styles["bodyStyle"].1, [1, 9]);
    assert!(!styles["otherStyle"].0);
    assert!(styles["otherStyle"].1.is_empty());
}

fn package_with_master_xml() -> Package {
    let mut package = Package::new().unwrap();
    package.to_bytes().unwrap();
    let package_bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&package_bytes).unwrap();
    let part_name = PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
    opc.get_part_mut(&part_name)
        .unwrap()
        .set_blob(MASTER_XML.to_vec());
    Package::from_opc_package(opc).unwrap()
}

fn parse_text_styles(xml: &[u8]) -> BTreeMap<String, (bool, Vec<u8>)> {
    let mut reader = Reader::from_reader(xml);
    let mut styles = BTreeMap::new();
    let mut current = None;
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) => {
                let name = element.local_name();
                if matches!(name.as_ref(), b"titleStyle" | b"bodyStyle" | b"otherStyle") {
                    let key = String::from_utf8(name.as_ref().to_vec()).unwrap();
                    styles.entry(key.clone()).or_insert((false, Vec::new()));
                    current = Some(key);
                } else if name.as_ref() == b"defPPr" {
                    if let Some(style) = current.as_ref() {
                        styles.get_mut(style).unwrap().0 = true;
                    }
                } else if let Some(level) = level_number(name.as_ref())
                    && let Some(style) = current.as_ref()
                {
                    styles.get_mut(style).unwrap().1.push(level);
                }
            },
            Event::Empty(element) => {
                let name = element.local_name();
                if matches!(name.as_ref(), b"titleStyle" | b"bodyStyle" | b"otherStyle") {
                    let key = String::from_utf8(name.as_ref().to_vec()).unwrap();
                    styles.entry(key).or_insert((false, Vec::new()));
                } else if name.as_ref() == b"defPPr" {
                    if let Some(style) = current.as_ref() {
                        styles.get_mut(style).unwrap().0 = true;
                    }
                } else if let Some(level) = level_number(name.as_ref())
                    && let Some(style) = current.as_ref()
                {
                    styles.get_mut(style).unwrap().1.push(level);
                }
            },
            Event::End(element)
                if matches!(
                    element.local_name().as_ref(),
                    b"titleStyle" | b"bodyStyle" | b"otherStyle"
                ) =>
            {
                current = None;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    styles
}

fn level_number(name: &[u8]) -> Option<u8> {
    let number = name.strip_prefix(b"lvl")?.strip_suffix(b"pPr")?;
    std::str::from_utf8(number).ok()?.parse().ok()
}
