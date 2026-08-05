use quick_xml::Reader;
use quick_xml::events::Event;

const PRESENTATION_XML: &[u8] =
    include_bytes!("../../../test-data/ooxml/pptx/default-text-style/presentation.xml");

#[test]
fn presentation_default_text_style_is_exposed_by_the_part_owner() {
    // Default text style is a PresentationML part concern and is not yet
    // published as a standalone PPTX projection. Cover its semantic shape
    // directly instead of reintroducing the retired umbrella facade.
    let style = parse_default_text_style(PRESENTATION_XML);

    assert!(style.has_default_paragraph_properties);
    assert_eq!(style.levels, [2, 5]);
    assert!(style.levels.contains(&5));
    assert!(!style.levels.contains(&1));
}

struct DefaultTextStyle {
    has_default_paragraph_properties: bool,
    levels: Vec<u8>,
}

fn parse_default_text_style(xml: &[u8]) -> DefaultTextStyle {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut style = DefaultTextStyle {
        has_default_paragraph_properties: false,
        levels: Vec::new(),
    };
    let mut in_style = false;

    loop {
        match reader.read_event().unwrap() {
            Event::Start(element) => {
                let local = element.local_name();
                if local.as_ref() == b"defaultTextStyle" {
                    in_style = true;
                } else if in_style && local.as_ref() == b"defPPr" {
                    style.has_default_paragraph_properties = true;
                } else if in_style
                    && let Some(level) = level_number(local.as_ref())
                {
                    style.levels.push(level);
                }
            },
            Event::Empty(element) => {
                let local = element.local_name();
                if in_style && local.as_ref() == b"defPPr" {
                    style.has_default_paragraph_properties = true;
                } else if in_style
                    && let Some(level) = level_number(local.as_ref())
                {
                    style.levels.push(level);
                }
            },
            Event::End(element) if element.local_name().as_ref() == b"defaultTextStyle" => {
                in_style = false;
            },
            Event::Eof => break,
            _ => {},
        }
    }

    style
}

fn level_number(name: &[u8]) -> Option<u8> {
    name.strip_prefix(b"lvl")?
        .strip_suffix(b"pPr")?
        .iter()
        .copied()
        .try_fold(0u8, |value, digit| {
            value
                .checked_mul(10)?
                .checked_add(digit.checked_sub(b'0')?)
        })
}
