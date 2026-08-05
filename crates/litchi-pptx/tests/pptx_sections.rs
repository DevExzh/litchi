use litchi_opc::constants::content_type as ct;
use litchi_opc::part::BlobPart;
use litchi_opc::PackURI;
use litchi_pptx::parts::PresentationPart;
use litchi_pptx::presentation_properties::metadata::sections::List;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/sections/presentation.xml");

#[test]
fn presentation_sections_include_slide_membership() {
    // Sections are owned by the presentation-properties metadata layer. The
    // direct codec keeps the same stable IDs while PresentationPart owns the
    // ordered slide references; no umbrella Presentation method is needed.
    let part = BlobPart::new(
        PackURI::new("/ppt/presentation.xml").unwrap(),
        ct::PML_PRESENTATION_MAIN.to_owned(),
        PRESENTATION_XML.to_vec(),
    );
    let presentation = PresentationPart::from_part(&part).unwrap();
    assert_eq!(
        presentation
            .slide_references()
            .unwrap()
            .into_iter()
            .map(|reference| reference.id())
            .collect::<Vec<_>>(),
        [256, 257, 258]
    );

    let sections = List::from_xml(PRESENTATION_XML).unwrap();
    assert_eq!(sections.sections().len(), 2);
    assert_eq!(sections.sections()[0].name.as_deref(), Some("Opening"));
    assert_eq!(
        sections.sections()[0].id.as_deref(),
        Some("{11111111-1111-1111-1111-111111111111}")
    );
    assert_eq!(sections.sections()[0].slide_ids, [256, 258]);
    assert_eq!(sections.sections()[1].name.as_deref(), Some("Recap"));
    assert_eq!(sections.sections()[1].slide_ids, [257]);

    let positions = [256_u32, 257, 258]
        .into_iter()
        .enumerate()
        .map(|(index, id)| (id, index))
        .collect::<std::collections::HashMap<_, _>>();
    let membership = sections
        .sections()
        .iter()
        .map(|section| {
            (
                section.name.clone().unwrap(),
                section
                    .slide_ids
                    .iter()
                    .map(|id| positions[id])
                    .collect::<Vec<_>>(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(membership, [("Opening".to_owned(), vec![0, 2]), ("Recap".to_owned(), vec![1])]);
}
