//! Typed handout-master metadata with a transactional PresentationML round-trip.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::presentation_properties::metadata::handout::{Layout, Master};
use litchi_pptx::{Error, Package};
use std::str::FromStr;

const HANDOUT_PART: &str = "/ppt/handoutMasters/handoutMaster1.xml";
const HANDOUT_TARGET: &str = "handoutMasters/handoutMaster1.xml";
const HANDOUT_RELATIONSHIP_ID: &str = "rIdHandout";

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Handout Master Example ===\n");

    let handout = Master::new()
        .with_layout(Layout::ThreeSlides)
        .with_header("My Presentation Title")
        .with_footer("Confidential - Do Not Distribute")
        .with_slide_numbers()
        .with_date_time();

    println!("Handout Configuration:");
    println!("  Layout: {:?}", handout.layout);
    println!("  Show header: {}", handout.header_footer.show_header);
    println!("  Header text: {:?}", handout.header_footer.header_text);
    println!("  Show footer: {}", handout.header_footer.show_footer);
    println!("  Footer text: {:?}", handout.header_footer.footer_text);
    println!(
        "  Show slide numbers: {}",
        handout.header_footer.show_slide_number
    );
    println!("  Show date/time: {}", handout.header_footer.show_date_time);
    println!("  Auto date: {}", handout.header_footer.auto_date);

    let xml = handout.to_xml();
    println!("\nGenerated XML length: {} bytes", xml.len());
    assert!(xml.contains("<p:handoutMaster"));
    assert!(Master::parse_xml(&xml).is_ok());

    println!("\n--- Handout Layouts ---");
    let layouts = [
        Layout::OneSlide,
        Layout::TwoSlides,
        Layout::ThreeSlides,
        Layout::FourSlides,
        Layout::SixSlides,
        Layout::NineSlides,
        Layout::Outline,
    ];

    for layout in layouts {
        let str_repr = layout.as_str();
        let parsed = Layout::from_str(str_repr).expect("known handout layout");
        println!("  {:?} -> '{}' -> {:?}", layout, str_repr, parsed);
        assert_eq!(layout, parsed);
    }

    let colored_handout = Master::new()
        .with_layout(Layout::SixSlides)
        .with_background_color("E6E6FA")
        .with_footer("Page Footer");
    assert_eq!(colored_handout.background_color.as_deref(), Some("E6E6FA"));

    println!("\nColored Handout:");
    println!("  Layout: {:?}", colored_handout.layout);
    println!("  Background: {:?}", colored_handout.background_color);
    assert!(colored_handout.to_xml().contains("<p:handoutMaster"));

    publish_and_reopen(&xml, &handout)?;

    println!("\n✅ Handout master example completed successfully!");
    Ok(())
}

fn publish_and_reopen(xml: &str, expected: &Master) -> Result<(), Box<dyn std::error::Error>> {
    let mut package = Package::new()?;
    package
        .presentation_mut()?
        .add_slide()?
        .set_title("Handout Demo");

    // Flush the typed slide writer, then make one atomic OPC edit for the
    // handout part, relationship, and presentation-level handout ID list.
    let bytes = package.to_bytes()?;
    let mut package = Package::from_bytes(&bytes)?;
    package.edit_opc(|opc| store_handout(opc, xml))?;

    let bytes = package.to_bytes()?;
    let reopened = Package::from_bytes(&bytes)?;
    let parsed = reopened.with_opc(|opc| load_handout(opc))?;

    assert_eq!(parsed.layout, Layout::default());
    assert_eq!(
        parsed.header_footer.show_header,
        expected.header_footer.show_header
    );
    assert_eq!(
        parsed.header_footer.header_text,
        expected.header_footer.header_text
    );
    assert_eq!(
        parsed.header_footer.show_footer,
        expected.header_footer.show_footer
    );
    assert_eq!(
        parsed.header_footer.footer_text,
        expected.header_footer.footer_text
    );
    assert_eq!(
        parsed.header_footer.show_slide_number,
        expected.header_footer.show_slide_number
    );
    assert_eq!(
        parsed.header_footer.show_date_time,
        expected.header_footer.show_date_time
    );
    assert_eq!(
        parsed.header_footer.auto_date,
        expected.header_footer.auto_date
    );
    println!("Published and reopened {HANDOUT_PART} transactionally.");
    Ok(())
}

fn store_handout(opc: &mut OpcPackage, xml: &str) -> litchi_pptx::Result<()> {
    let presentation_name = PackURI::new("/ppt/presentation.xml").map_err(Error::Invalid)?;
    let handout_name = PackURI::new(HANDOUT_PART).map_err(Error::Invalid)?;
    let original = String::from_utf8(opc.get_part(&presentation_name)?.blob().to_vec())
        .map_err(|error| Error::Xml(error.to_string()))?;
    let marker = "</p:sldMasterIdLst>";
    let handout_ids =
        r#"<p:handoutMasterIdLst><p:handoutMasterId r:id="rIdHandout"/></p:handoutMasterIdLst>"#;
    let presentation_xml = if original.contains("<p:handoutMasterIdLst") {
        original
    } else {
        original
            .replace(marker, &format!("{marker}{handout_ids}"))
            .into()
    };
    if !presentation_xml.contains("<p:handoutMasterId r:id=\"rIdHandout\"") {
        return Err(Error::Invalid(
            "presentation has no handout-master ID list insertion point".to_string(),
        ));
    }

    opc.get_part_mut(&presentation_name)?
        .set_blob(presentation_xml.into_bytes());
    opc.get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            rt::HANDOUT_MASTER.to_string(),
            HANDOUT_TARGET.to_string(),
            HANDOUT_RELATIONSHIP_ID.to_string(),
            false,
        );
    opc.try_add_part(Box::new(BlobPart::new(
        handout_name,
        ct::PML_HANDOUT_MASTER.to_string(),
        xml.as_bytes().to_vec(),
    )))?;
    Ok(())
}

fn load_handout(opc: &OpcPackage) -> litchi_pptx::Result<Master> {
    let presentation = opc.main_document_part()?;
    let relationship = presentation
        .rels()
        .iter()
        .find(|relationship| relationship.r_id() == HANDOUT_RELATIONSHIP_ID)
        .ok_or_else(|| Error::Relationship("handout-master relationship is missing".into()))?;
    if relationship.reltype() != rt::HANDOUT_MASTER {
        return Err(Error::Relationship(
            "relationship is not a handout-master relationship".into(),
        ));
    }

    let handout_name = PackURI::new(HANDOUT_PART).map_err(Error::Invalid)?;
    let part = opc.get_part(&handout_name)?;
    if part.content_type() != ct::PML_HANDOUT_MASTER {
        return Err(Error::ContentType {
            expected: ct::PML_HANDOUT_MASTER.to_string(),
            actual: part.content_type().to_string(),
        });
    }
    let xml = std::str::from_utf8(part.blob()).map_err(|error| Error::Xml(error.to_string()))?;
    Master::parse_xml(xml)
}
