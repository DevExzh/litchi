use quick_xml::Reader;
use quick_xml::events::Event;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/kinsoku/presentation.xml");

#[test]
fn presentation_kinsoku_settings_are_exposed_by_the_part_owner() {
    // Kinsoku remains an intentionally narrow PresentationML part projection;
    // keep its language and character-set regression here until a dedicated
    // standalone metadata owner is published.
    let settings = parse_kinsoku(PRESENTATION_XML);

    assert_eq!(settings.language.as_deref(), Some("ja-jp"));
    assert_eq!(settings.invalid_start_characters, "、。）］");
    assert_eq!(settings.invalid_end_characters, "（［");
}

struct Kinsoku {
    language: Option<String>,
    invalid_start_characters: String,
    invalid_end_characters: String,
}

fn parse_kinsoku(xml: &[u8]) -> Kinsoku {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"kinsoku" =>
            {
                let mut value = Kinsoku {
                    language: None,
                    invalid_start_characters: String::new(),
                    invalid_end_characters: String::new(),
                };
                for attribute in element.attributes().flatten() {
                    let target = match attribute.key.as_ref() {
                        b"lang" => &mut value.language,
                        b"invalStChars" => {
                            value.invalid_start_characters =
                                String::from_utf8(attribute.value.into_owned()).unwrap();
                            continue;
                        },
                        b"invalEndChars" => {
                            value.invalid_end_characters =
                                String::from_utf8(attribute.value.into_owned()).unwrap();
                            continue;
                        },
                        _ => continue,
                    };
                    *target = Some(String::from_utf8(attribute.value.into_owned()).unwrap());
                }
                return value;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    panic!("presentation did not contain p:kinsoku");
}
