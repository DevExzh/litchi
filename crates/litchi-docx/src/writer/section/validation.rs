use crate::error::{Error, Result};
use quick_xml::events::Event;

use super::model::{
    MAX_PAGE_BORDER_ART_SIZE, MAX_PAGE_BORDER_LINE_SIZE, MAX_PAGE_BORDER_SPACE, SectionProperties,
    Style,
};

pub(super) fn validate_header_footer_xml(xml: &str, header: bool) -> Result<()> {
    use quick_xml::reader::NsReader;
    let mut reader = NsReader::from_str(xml);
    let mut depth = 0usize;
    let mut root = false;
    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                if depth == 0 {
                    let expected = if header {
                        b"hdr".as_slice()
                    } else {
                        b"ftr".as_slice()
                    };
                    if root
                        || !crate::namespace::is_wordprocessing_namespace(&namespace)
                        || element.local_name().as_ref() != expected
                    {
                        return Err(Error::InvalidFormat(
                            "section header/footer XML has an invalid root".to_string(),
                        ));
                    }
                    root = true;
                }
                depth += 1;
            },
            Event::Empty(element) if depth == 0 => {
                let expected = if header {
                    b"hdr".as_slice()
                } else {
                    b"ftr".as_slice()
                };
                if root
                    || !crate::namespace::is_wordprocessing_namespace(&namespace)
                    || element.local_name().as_ref() != expected
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer XML has an invalid root".to_string(),
                    ));
                }
                root = true;
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid header/footer XML nesting".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root || depth != 0 {
        return Err(Error::InvalidFormat(
            "unterminated section header/footer XML".to_string(),
        ));
    }
    Ok(())
}

impl SectionProperties {
    pub(crate) fn validate(&self) -> Result<()> {
        if self.page_width == 0 || self.page_height == 0 {
            return Err(Error::InvalidFormat(
                "section page dimensions must be nonzero".to_string(),
            ));
        }
        for references in [&self.headers, &self.footers] {
            let mut kinds = std::collections::HashSet::new();
            for reference in references {
                if !kinds.insert(reference.kind) {
                    return Err(Error::InvalidFormat(
                        "section has duplicate header/footer reference type".to_string(),
                    ));
                }
                if reference.relationship_id.as_deref() == Some("") {
                    return Err(Error::InvalidFormat(
                        "section header/footer relationship ID is empty".to_string(),
                    ));
                }
                if reference.relationship_id.is_some() && reference.part.is_some() {
                    return Err(Error::InvalidFormat(
                        "section header/footer cannot be both existing and owned".to_string(),
                    ));
                }
                if let Some(part) = &reference.part
                    && (part.key.is_empty() || part.xml.is_empty())
                {
                    return Err(Error::InvalidFormat(
                        "section header/footer part key and XML must be non-empty".to_string(),
                    ));
                }
                if let Some(part) = &reference.part {
                    validate_header_footer_xml(&part.xml, std::ptr::eq(references, &self.headers))?;
                }
            }
        }
        if let Some(columns) = &self.columns {
            if columns.count == 0 || columns.count > 45 {
                return Err(Error::InvalidFormat(
                    "section column count must be in 1..=45".to_string(),
                ));
            }
            if !columns.equal_width && usize::from(columns.count) != columns.columns.len() {
                return Err(Error::InvalidFormat(
                    "unequal section columns require one width per column".to_string(),
                ));
            }
        }
        if let Some(borders) = &self.page_borders {
            for border in [&borders.top, &borders.left, &borders.bottom, &borders.right]
                .into_iter()
                .flatten()
            {
                if let Some(size) = border.size {
                    let max = match border.style {
                        Style::Art(_) => MAX_PAGE_BORDER_ART_SIZE,
                        _ => MAX_PAGE_BORDER_LINE_SIZE,
                    };
                    if size > max {
                        return Err(Error::InvalidFormat(format!(
                            "page border size {size} exceeds the {max} limit"
                        )));
                    }
                }
                if let Some(space) = border.space
                    && space > MAX_PAGE_BORDER_SPACE
                {
                    return Err(Error::InvalidFormat(format!(
                        "page border space {space} exceeds the {MAX_PAGE_BORDER_SPACE} limit"
                    )));
                }
            }
        }
        if self.printer_settings_relationship_id.as_deref() == Some("") {
            return Err(Error::InvalidFormat(
                "section printer-settings relationship ID is empty".to_string(),
            ));
        }
        Ok(())
    }
}
