//! Package-wide collision avoidance for authored nonvisual drawing identifiers.

use std::collections::HashSet;

use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

pub(super) struct DrawingIds {
    used: HashSet<u32>,
    next: u32,
    maximum: usize,
}

impl DrawingIds {
    pub(super) fn new(maximum: usize) -> Self {
        Self {
            used: HashSet::new(),
            next: 1,
            maximum,
        }
    }

    /// Inspect the complete physical story, including inactive MCE alternatives.
    pub(super) fn observe(&mut self, source: &[u8]) -> Result<()> {
        let mut reader = NsReader::from_reader(source);
        loop {
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    let numeric = match namespace {
                        ResolveResult::Bound(Namespace(uri)) => {
                            (element.local_name().as_ref() == b"docPr" && matches!(uri,
                                b"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" |
                                b"http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing")) ||
                            (element.local_name().as_ref() == b"cNvPr" && matches!(uri,
                                b"http://schemas.openxmlformats.org/drawingml/2006/main" |
                                b"http://schemas.microsoft.com/office/word/2010/wordml"))
                        },
                        _ => false,
                    };
                    for attribute in element.attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        if !matches!(attribute.key.as_ref(), b"id" | b"xml:id") {
                            continue;
                        }
                        let value = attribute
                            .decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                element.decoder(),
                            )
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        let authored = value
                            .strip_prefix("litchiInk")
                            .and_then(|value| value.parse::<u32>().ok());
                        let id = if numeric && attribute.key.as_ref() == b"id" {
                            value.parse::<u32>().ok().or(authored)
                        } else {
                            authored
                        };
                        if let Some(id) = id {
                            self.reserve(id)?;
                        }
                    }
                },
                Event::Eof => return Ok(()),
                _ => {},
            }
        }
    }

    fn reserve(&mut self, id: u32) -> Result<()> {
        if self.used.contains(&id) {
            return Ok(());
        }
        super::bound(
            "drawing identifiers",
            self.used.len().saturating_add(1),
            self.maximum,
        )?;
        self.used
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "DOCX Ink drawing identifiers",
                source,
            })?;
        self.used.insert(id);
        Ok(())
    }

    pub(super) fn allocate(&mut self) -> Result<u32> {
        while self.used.contains(&self.next) {
            self.next = self.next.checked_add(1).ok_or_else(|| {
                Error::Invalid("DOCX Ink drawing identifier space exhausted".into())
            })?;
        }
        let id = self.next;
        self.reserve(id)?;
        Ok(id)
    }
}

#[cfg(test)]
mod tests {
    use super::DrawingIds;

    #[test]
    fn observes_inactive_and_aliased_drawing_ids_and_authored_vml_ids() {
        let mut ids = DrawingIds::new(10);
        ids.observe(br#"<root xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"><mc:Fallback><x:docPr id="1"/></mc:Fallback><v:shape id="litchiInk2"/><foreign id="4294967295"/></root>"#).unwrap();
        assert_eq!(ids.allocate().unwrap(), 3);
        assert_eq!(ids.allocate().unwrap(), 4);
        ids.observe(br#"<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="litchiInk5"/>"#)
            .unwrap();
        assert_eq!(ids.allocate().unwrap(), 6);
    }

    #[test]
    fn duplicate_ids_do_not_consume_extra_budget_and_limit_refusal_is_stable() {
        let mut ids = DrawingIds::new(1);
        ids.observe(br#"<root><shape id="litchiInk1"/><shape id="litchiInk1"/></root>"#)
            .unwrap();
        assert!(ids.allocate().is_err());
        assert!(ids.allocate().is_err());
    }
}
