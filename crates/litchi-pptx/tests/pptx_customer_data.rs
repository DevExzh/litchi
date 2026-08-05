use quick_xml::Reader;
use quick_xml::events::Event;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/customer-data/presentation.xml");

#[test]
fn presentation_customer_data_relationships_are_exposed_by_the_part_owner() {
    // The standalone PPTX facade intentionally has no public customer-data
    // projection yet. Keep the regression at the PresentationML part owner so
    // the relationship-bearing vocabulary remains covered without restoring
    // an umbrella compatibility API.
    let (custom_data, tags) = relationship_ids(PRESENTATION_XML);

    assert_eq!(
        custom_data,
        ["rIdCustomerDataOne", "rIdCustomerDataTwo"]
    );
    assert_eq!(tags.as_deref(), Some("rIdCustomerDataTags"));
}

fn relationship_ids(xml: &[u8]) -> (Vec<String>, Option<String>) {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut custom_data = Vec::new();
    let mut tags = None;

    loop {
        match reader.read_event().unwrap() {
            Event::Empty(element) | Event::Start(element) => {
                let id = element
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.as_ref() == b"r:id")
                    .map(|attribute| String::from_utf8(attribute.value.into_owned()).unwrap());
                match element.local_name().as_ref() {
                    b"custData" => {
                        if let Some(id) = id {
                            custom_data.push(id);
                        }
                    },
                    b"tags" => tags = id,
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }

    (custom_data, tags)
}
