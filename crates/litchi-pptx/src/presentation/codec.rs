//! Bounded XML scanning for presentation-level semantic values.

use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::hyperlinks::Hyperlink;
use crate::{Error, Result};

pub(super) fn parse_inline_hyperlinks(xml: &[u8]) -> Result<Vec<Hyperlink>> {
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(true);
    let mut hyperlinks = Vec::new();
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader.read_resolved_event()?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if litchi_ooxml_common::xml::is_drawingml_name(
                    &namespace,
                    element.name(),
                    b"hlinkClick",
                ) =>
            {
                let action = litchi_ooxml_common::xml::unqualified_attribute_value(
                    &element, b"action", decoder,
                )?;
                let tooltip = litchi_ooxml_common::xml::unqualified_attribute_value(
                    &element, b"tooltip", decoder,
                )?;
                if let Some(action) = action {
                    if action.is_empty() {
                        return Err(Error::Invalid(
                            "inline hyperlink action cannot be empty".into(),
                        ));
                    }
                    hyperlinks.push(Hyperlink::from_xml(&action, tooltip)?);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(hyperlinks)
}
