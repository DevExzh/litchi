use crate::alt::{Chunk, Conformance, scan};
use crate::error::{Error, Result};
use crate::parts::document_part::active_block_ranges;

use super::super::paragraph::MutableParagraph;
use super::super::section::SectionProperties;
use super::super::table::MutableTable;
use super::super::toc::TableOfContents;

/// The document body containing all content elements.
#[derive(Debug)]
pub(crate) struct DocumentBody {
    /// Content elements (paragraphs, tables, etc.) in document order
    pub(crate) elements: Vec<BodyElement>,
}

/// Keep a pending TOC insertion point anchored after an inserted element.
pub(super) fn shift_toc_index_on_insert(
    toc_config: &mut Option<(usize, TableOfContents)>,
    position: usize,
) {
    if let Some((index, _)) = toc_config
        && position <= *index
    {
        *index += 1;
    }
}

/// Keep a pending TOC insertion point anchored after a removed element.
pub(super) fn shift_toc_index_on_remove(
    toc_config: &mut Option<(usize, TableOfContents)>,
    position: usize,
) {
    if let Some((index, _)) = toc_config
        && position < *index
    {
        *index -= 1;
    }
}

pub(super) struct ParsedDocumentBody {
    pub(super) body: DocumentBody,
    pub(super) prefix: String,
    pub(super) suffix: String,
}

#[derive(Clone, Copy)]
enum PreservedBodyKind {
    Paragraph,
    Table,
    SectionProperties,
    Alt,
    Other,
}

struct PreservedAltRange {
    start: usize,
    end: usize,
    chunk: Chunk,
}

impl DocumentBody {
    pub(super) fn new() -> Self {
        Self {
            elements: Vec::new(),
        }
    }

    pub(super) fn from_xml(xml: &str) -> Result<ParsedDocumentBody> {
        use crate::namespace::is_wordprocessing_namespace;
        use quick_xml::events::Event;
        use quick_xml::reader::NsReader;

        enum ScanEvent {
            StartBody,
            StartChild(PreservedBodyKind),
            NestedStart,
            EmptyChild(PreservedBodyKind),
            EndCaptured,
            EndBody,
            StartOther,
            EndOther,
            Eof,
            Other,
        }

        let bytes = xml.as_bytes();
        let mut chunks = scan(bytes)?;
        let mut active_alts = Vec::new();
        for (target, start, length) in active_block_ranges(bytes)? {
            if target != 2 {
                continue;
            }
            let start_usize = usize::try_from(start).map_err(|_| {
                Error::InvalidFormat("altChunk offset does not fit usize".to_string())
            })?;
            let end = start_usize
                .checked_add(usize::try_from(length).map_err(|_| {
                    Error::InvalidFormat("altChunk length does not fit usize".to_string())
                })?)
                .ok_or_else(|| Error::InvalidFormat("altChunk range overflowed".to_string()))?;
            let chunk = chunks.remove(&start).ok_or_else(|| {
                Error::InvalidFormat(
                    "active altChunk range lacks parsed anchor metadata".to_string(),
                )
            })?;
            active_alts.push(PreservedAltRange {
                start: start_usize,
                end,
                chunk,
            });
        }
        let mut reader = NsReader::from_reader(bytes);
        let mut body = Self::new();
        let mut depth = 0usize;
        let mut body_depth = None;
        let mut prefix_end = None;
        let mut suffix_start = None;
        let mut last_content_end = 0usize;
        let mut capture: Option<(PreservedBodyKind, usize, usize)> = None;

        loop {
            let event_start = usize::try_from(reader.buffer_position()).map_err(|_| {
                Error::InvalidFormat("Word document offset does not fit usize".to_string())
            })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                match event {
                    Event::Start(_) if capture.is_some() => ScanEvent::NestedStart,
                    Event::Start(element)
                        if is_wordprocessing_namespace(&namespace)
                            && element.local_name().as_ref() == b"body" =>
                    {
                        ScanEvent::StartBody
                    },
                    Event::Start(element) if body_depth == Some(depth) => ScanEvent::StartChild(
                        preserved_body_kind(&namespace, element.local_name().as_ref()),
                    ),
                    Event::Start(_) => ScanEvent::StartOther,
                    Event::Empty(element) if capture.is_none() && body_depth == Some(depth) => {
                        ScanEvent::EmptyChild(preserved_body_kind(
                            &namespace,
                            element.local_name().as_ref(),
                        ))
                    },
                    Event::End(_) if capture.is_some() => ScanEvent::EndCaptured,
                    Event::End(element)
                        if is_wordprocessing_namespace(&namespace)
                            && element.local_name().as_ref() == b"body"
                            && body_depth == Some(depth) =>
                    {
                        ScanEvent::EndBody
                    },
                    Event::End(_) => ScanEvent::EndOther,
                    Event::Eof => ScanEvent::Eof,
                    _ => ScanEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                Error::InvalidFormat("Word document offset does not fit usize".to_string())
            })?;

            match event {
                ScanEvent::StartBody => {
                    if body_depth.is_some() || prefix_end.is_some() {
                        return Err(Error::InvalidFormat(
                            "document contains multiple Word body elements".to_string(),
                        ));
                    }
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                    body_depth = Some(depth);
                    prefix_end = Some(event_end);
                    last_content_end = event_end;
                },
                ScanEvent::StartChild(kind) => {
                    capture = Some((kind, event_start, 1));
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::NestedStart => {
                    let Some((_, _, capture_depth)) = capture.as_mut() else {
                        return Err(Error::InvalidFormat(
                            "missing preserved body element".to_string(),
                        ));
                    };
                    *capture_depth = capture_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::EmptyChild(kind) => {
                    push_preserved_body_range(
                        &mut body,
                        xml,
                        &mut last_content_end,
                        kind,
                        event_start,
                        event_end,
                        &active_alts,
                    )?;
                },
                ScanEvent::EndCaptured => {
                    let Some((_, _, capture_depth)) = capture.as_mut() else {
                        return Err(Error::InvalidFormat(
                            "missing preserved body element".to_string(),
                        ));
                    };
                    *capture_depth = capture_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                    if *capture_depth == 0 {
                        let Some((kind, start, _)) = capture.take() else {
                            return Err(Error::InvalidFormat(
                                "missing preserved body element range".to_string(),
                            ));
                        };
                        push_preserved_body_range(
                            &mut body,
                            xml,
                            &mut last_content_end,
                            kind,
                            start,
                            event_end,
                            &active_alts,
                        )?;
                    }
                },
                ScanEvent::EndBody => {
                    if event_start > last_content_end {
                        push_raw_body_xml(
                            &mut body,
                            PreservedBodyKind::Other,
                            xml,
                            last_content_end,
                            event_start,
                            &active_alts,
                        )?;
                    }
                    suffix_start = Some(event_start);
                    body_depth = None;
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                },
                ScanEvent::StartOther => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".to_string())
                    })?;
                },
                ScanEvent::EndOther => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word XML nesting".to_string())
                    })?;
                },
                ScanEvent::Eof if depth != 0 || capture.is_some() || body_depth.is_some() => {
                    return Err(Error::InvalidFormat(
                        "unterminated Word document XML".to_string(),
                    ));
                },
                ScanEvent::Eof => break,
                ScanEvent::Other => {},
            }
        }

        let prefix_end = prefix_end
            .ok_or_else(|| Error::InvalidFormat("Word document has no body element".to_string()))?;
        let suffix_start = suffix_start
            .ok_or_else(|| Error::InvalidFormat("Word document body is not closed".to_string()))?;
        Ok(ParsedDocumentBody {
            body,
            prefix: ensure_writer_namespace_declarations(xml.get(..prefix_end).ok_or_else(
                || Error::InvalidFormat("invalid Word document prefix range".to_string()),
            )?)?,
            suffix: xml
                .get(suffix_start..)
                .ok_or_else(|| {
                    Error::InvalidFormat("invalid Word document suffix range".to_string())
                })?
                .to_string(),
        })
    }
    pub(super) fn add_paragraph(&mut self) -> &mut MutableParagraph {
        let index = self.content_insertion_index();
        self.elements
            .insert(index, BodyElement::Paragraph(MutableParagraph::new()));
        match self.elements.get_mut(index) {
            Some(BodyElement::Paragraph(p)) => p,
            _ => unreachable!(),
        }
    }

    pub(super) fn add_table(&mut self, rows: usize, cols: usize) -> &mut MutableTable {
        let index = self.content_insertion_index();
        self.elements
            .insert(index, BodyElement::Table(MutableTable::new(rows, cols)));
        match self.elements.get_mut(index) {
            Some(BodyElement::Table(t)) => t,
            _ => unreachable!(),
        }
    }

    pub(super) fn content_insertion_index(&self) -> usize {
        self.elements
            .iter()
            .position(|element| matches!(element, BodyElement::PreservedSectionProperties(_)))
            .unwrap_or(self.elements.len())
    }

    /// Element positions of all paragraphs, typed and preserved, in body order.
    pub(super) fn paragraph_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(
                    element,
                    BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
                )
                .then_some(index)
            })
            .collect()
    }

    /// Element positions of all tables, typed and preserved, in body order.
    pub(super) fn table_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(
                    element,
                    BodyElement::Table(_) | BodyElement::PreservedTable(_)
                )
                .then_some(index)
            })
            .collect()
    }

    /// Insert an empty paragraph before the paragraph at paragraph-relative
    /// `index`; returns the element position and the new paragraph.
    pub(super) fn insert_paragraph(
        &mut self,
        index: usize,
    ) -> Result<(usize, &mut MutableParagraph)> {
        let positions = self.paragraph_positions();
        if index > positions.len() {
            return Err(Error::InvalidFormat(format!(
                "paragraph insertion index {index} is out of range"
            )));
        }
        let position = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        self.elements
            .insert(position, BodyElement::Paragraph(MutableParagraph::new()));
        match self.elements.get_mut(position) {
            Some(BodyElement::Paragraph(paragraph)) => Ok((position, paragraph)),
            _ => unreachable!(),
        }
    }

    /// Insert an empty table before the table at table-relative `index`;
    /// returns the element position and the new table.
    pub(super) fn insert_table(
        &mut self,
        index: usize,
        rows: usize,
        cols: usize,
    ) -> Result<(usize, &mut MutableTable)> {
        let positions = self.table_positions();
        if index > positions.len() {
            return Err(Error::InvalidFormat(format!(
                "table insertion index {index} is out of range"
            )));
        }
        let position = positions
            .get(index)
            .copied()
            .unwrap_or_else(|| self.content_insertion_index());
        self.elements
            .insert(position, BodyElement::Table(MutableTable::new(rows, cols)));
        match self.elements.get_mut(position) {
            Some(BodyElement::Table(table)) => Ok((position, table)),
            _ => unreachable!(),
        }
    }

    /// Remove the paragraph at paragraph-relative `index`; returns the
    /// vacated element position.
    pub(super) fn remove_paragraph(&mut self, index: usize) -> Result<usize> {
        let position = self
            .paragraph_positions()
            .get(index)
            .copied()
            .ok_or_else(|| {
                Error::InvalidFormat(format!("paragraph index {index} is out of range"))
            })?;
        self.elements.remove(position);
        Ok(position)
    }

    /// Remove the table at table-relative `index`; returns the vacated
    /// element position.
    pub(super) fn remove_table(&mut self, index: usize) -> Result<usize> {
        let position =
            self.table_positions().get(index).copied().ok_or_else(|| {
                Error::InvalidFormat(format!("table index {index} is out of range"))
            })?;
        self.elements.remove(position);
        Ok(position)
    }

    pub(super) fn alt_positions(&self) -> Vec<usize> {
        self.elements
            .iter()
            .enumerate()
            .filter_map(|(index, element)| {
                matches!(element, BodyElement::PreservedAlt(_, _)).then_some(index)
            })
            .collect()
    }

    pub(super) fn alts(&self) -> Vec<Chunk> {
        self.elements
            .iter()
            .filter_map(|element| match element {
                BodyElement::PreservedAlt(_, chunk) => Some(chunk.clone()),
                _ => None,
            })
            .collect()
    }

    pub(super) fn insert_alt(
        &mut self,
        index: usize,
        chunk: Chunk,
        namespace: Conformance,
    ) -> Result<usize> {
        let positions = self.alt_positions();
        if index > positions.len() {
            return Err(Error::InvalidFormat(format!(
                "altChunk index {index} is out of range"
            )));
        }
        let position = match positions.get(index).copied() {
            Some(position) => position,
            None => self.content_insertion_index(),
        };
        let xml = chunk.xml(namespace);
        self.elements
            .insert(position, BodyElement::PreservedAlt(xml, chunk));
        Ok(position)
    }

    pub(super) fn replace_alt(
        &mut self,
        index: usize,
        chunk: Chunk,
        namespace: Conformance,
    ) -> Result<Chunk> {
        let position = self.alt_positions().get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!("altChunk index {index} is out of range"))
        })?;
        let xml = chunk.xml(namespace);
        let slot = self.elements.get_mut(position).ok_or_else(|| {
            Error::InvalidFormat("alternative-format position became invalid".into())
        })?;
        match std::mem::replace(slot, BodyElement::PreservedAlt(xml, chunk)) {
            BodyElement::PreservedAlt(_, old) => Ok(old),
            other => {
                *slot = other;
                Err(Error::InvalidFormat(
                    "alternative-format position does not contain an anchor".into(),
                ))
            },
        }
    }

    pub(super) fn remove_alt(&mut self, index: usize) -> Result<(usize, Chunk)> {
        let position = self.alt_positions().get(index).copied().ok_or_else(|| {
            Error::InvalidFormat(format!("altChunk index {index} is out of range"))
        })?;
        if position >= self.elements.len() {
            return Err(Error::InvalidFormat(
                "alternative-format position became invalid".into(),
            ));
        }
        match self.elements.remove(position) {
            BodyElement::PreservedAlt(_, chunk) => Ok((position, chunk)),
            other => {
                self.elements.insert(position, other);
                Err(Error::InvalidFormat(
                    "alternative-format position does not contain an anchor".into(),
                ))
            },
        }
    }

    pub(super) fn move_alt(&mut self, from: usize, to: usize) -> Result<()> {
        let positions = self.alt_positions();
        let count = positions.len();
        if from >= count || to >= count {
            return Err(Error::InvalidFormat(format!(
                "altChunk move {from} -> {to} is out of range"
            )));
        }
        if from == to {
            return Ok(());
        }
        let source = positions.get(from).copied().ok_or_else(|| {
            Error::InvalidFormat("alternative-format source position is missing".into())
        })?;
        if source >= self.elements.len() {
            return Err(Error::InvalidFormat(
                "alternative-format source position became invalid".into(),
            ));
        }
        let element = self.elements.remove(source);
        if !matches!(&element, BodyElement::PreservedAlt(_, _)) {
            self.elements.insert(source, element);
            return Err(Error::InvalidFormat(
                "alternative-format source does not contain an anchor".into(),
            ));
        }
        let remaining = self.alt_positions();
        let destination = match remaining.get(to).copied() {
            Some(position) => position,
            None => self.content_insertion_index(),
        };
        self.elements.insert(destination, element);
        Ok(())
    }

    pub(super) fn paragraph_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
                )
            })
            .count()
    }

    pub(super) fn table_count(&self) -> usize {
        self.elements
            .iter()
            .filter(|element| {
                matches!(
                    element,
                    BodyElement::Table(_) | BodyElement::PreservedTable(_)
                )
            })
            .count()
    }

    pub(super) fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        let mut count = 0;
        for elem in &mut self.elements {
            match elem {
                BodyElement::Paragraph(paragraph) => {
                    if count == index {
                        return Some(paragraph);
                    }
                    count += 1;
                },
                BodyElement::PreservedParagraph(_) => {
                    if count == index {
                        return None;
                    }
                    count += 1;
                },
                _ => {},
            }
        }
        None
    }

    pub(super) fn table(&mut self, index: usize) -> Option<&mut MutableTable> {
        let mut count = 0;
        for elem in &mut self.elements {
            match elem {
                BodyElement::Table(table) => {
                    if count == index {
                        return Some(table);
                    }
                    count += 1;
                },
                BodyElement::PreservedTable(_) => {
                    if count == index {
                        return None;
                    }
                    count += 1;
                },
                _ => {},
            }
        }
        None
    }

    pub(super) fn write_contents(&self, xml: &mut String, preserve_section: bool) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(p) => p.to_xml(xml)?,
                BodyElement::Table(t) => t.to_xml(xml)?,
                BodyElement::PreservedParagraph(raw)
                | BodyElement::PreservedTable(raw)
                | BodyElement::PreservedOther(raw)
                | BodyElement::PreservedAlt(raw, _) => xml.push_str(raw),
                BodyElement::PreservedSectionProperties(raw) if preserve_section => {
                    xml.push_str(raw);
                },
                BodyElement::PreservedSectionProperties(_) => {},
            }
        }
        Ok(())
    }

    pub(super) fn has_preserved_section(&self) -> bool {
        self.elements
            .iter()
            .any(|element| matches!(element, BodyElement::PreservedSectionProperties(_)))
    }

    pub(super) fn final_section_properties(&self) -> Result<Option<SectionProperties>> {
        self.elements
            .iter()
            .find_map(|element| match element {
                BodyElement::PreservedSectionProperties(raw) => {
                    Some(SectionProperties::from_xml(raw))
                },
                _ => None,
            })
            .transpose()
    }

    pub(super) fn validate_section_placement(&self) -> Result<()> {
        let mut final_section = None;
        for (index, element) in self.elements.iter().enumerate() {
            match element {
                BodyElement::PreservedSectionProperties(raw) => {
                    if final_section.replace(index).is_some() {
                        return Err(Error::InvalidFormat(
                            "document body contains multiple final section properties".to_string(),
                        ));
                    }
                    SectionProperties::from_xml(raw)?;
                },
                BodyElement::PreservedParagraph(raw) => {
                    paragraph_section_range(raw)?;
                },
                _ => {},
            }
        }
        if let Some(index) = final_section
            && self.elements[index + 1..].iter().any(|element| {
                !matches!(element, BodyElement::PreservedOther(raw) if raw.trim().is_empty())
            })
        {
            return Err(Error::InvalidFormat(
                "body-final section properties are not the final body child".to_string(),
            ));
        }
        Ok(())
    }

    pub(super) fn section_break_count(&self) -> Result<usize> {
        let mut count = 0usize;
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) if paragraph.properties.section.is_some() => {
                    count += 1;
                },
                BodyElement::PreservedParagraph(raw) if paragraph_section_range(raw)?.is_some() => {
                    count += 1;
                },
                _ => {},
            }
        }
        Ok(count)
    }

    pub(super) fn insert_section_break(
        &mut self,
        paragraph_index: usize,
        properties: SectionProperties,
    ) -> Result<()> {
        let element = self.paragraph_element_mut(paragraph_index).ok_or_else(|| {
            Error::InvalidFormat(format!("paragraph index {paragraph_index} is out of range"))
        })?;
        match element {
            BodyElement::Paragraph(paragraph) => {
                if paragraph.properties.section.is_some() {
                    return Err(Error::InvalidFormat(
                        "paragraph already ends a section".to_string(),
                    ));
                }
                paragraph.set_section_break(properties)
            },
            BodyElement::PreservedParagraph(raw) => {
                if paragraph_section_range(raw)?.is_some() {
                    return Err(Error::InvalidFormat(
                        "paragraph already ends a section".to_string(),
                    ));
                }
                let mut section_xml = String::new();
                properties.write_xml(&mut section_xml, None)?;
                *raw = insert_paragraph_property(raw, &section_xml)?;
                Ok(())
            },
            _ => unreachable!(),
        }
    }

    pub(super) fn section_break(&self, index: usize) -> Result<SectionProperties> {
        let mut current = 0usize;
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        if current == index {
                            return Ok(section.clone());
                        }
                        current += 1;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            return SectionProperties::from_xml(&raw[start..end]);
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(Error::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    pub(super) fn update_section_break(
        &mut self,
        index: usize,
        update: impl FnOnce(&mut SectionProperties),
    ) -> Result<()> {
        let mut current = 0usize;
        let mut update = Some(update);
        for element in &mut self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = paragraph.properties.section.as_mut() {
                        if current == index {
                            update.take().expect("called once")(section);
                            return section.validate();
                        }
                        current += 1;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            let mut section = SectionProperties::from_xml(&raw[start..end])?;
                            update.take().expect("called once")(&mut section);
                            section.validate()?;
                            let mut replacement = String::new();
                            section.write_xml(&mut replacement, None)?;
                            raw.replace_range(start..end, &replacement);
                            return Ok(());
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(Error::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    pub(super) fn remove_section_break(&mut self, index: usize) -> Result<SectionProperties> {
        let mut current = 0usize;
        for element in &mut self.elements {
            match element {
                BodyElement::Paragraph(paragraph) if paragraph.properties.section.is_some() => {
                    if current == index {
                        return paragraph.remove_section_break().ok_or_else(|| {
                            Error::InvalidFormat("section break disappeared".to_string())
                        });
                    }
                    current += 1;
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        if current == index {
                            let section = SectionProperties::from_xml(&raw[start..end])?;
                            raw.replace_range(start..end, "");
                            return Ok(section);
                        }
                        current += 1;
                    }
                },
                _ => {},
            }
        }
        Err(Error::InvalidFormat(format!(
            "section break index {index} is out of range"
        )))
    }

    pub(super) fn paragraph_element_mut(&mut self, index: usize) -> Option<&mut BodyElement> {
        let mut current = 0usize;
        for element in &mut self.elements {
            if matches!(
                element,
                BodyElement::Paragraph(_) | BodyElement::PreservedParagraph(_)
            ) {
                if current == index {
                    return Some(element);
                }
                current += 1;
            }
        }
        None
    }

    pub(super) fn collect_section_parts(
        &self,
        parts: &mut Vec<(bool, super::super::section::SectionHeaderFooterPart)>,
    ) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        collect_section_parts(section, parts)?;
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        collect_section_parts(
                            &SectionProperties::from_xml(&raw[start..end])?,
                            parts,
                        )?;
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }

    pub(super) fn collect_explicit_section_relationships(
        &self,
        relationships: &mut Vec<(String, bool)>,
    ) -> Result<()> {
        for element in &self.elements {
            match element {
                BodyElement::Paragraph(paragraph) => {
                    if let Some(section) = &paragraph.properties.section {
                        collect_explicit_section_relationships(section, relationships);
                    }
                },
                BodyElement::PreservedParagraph(raw) => {
                    if let Some((start, end)) = paragraph_section_range(raw)? {
                        collect_explicit_section_relationships(
                            &SectionProperties::from_xml(&raw[start..end])?,
                            relationships,
                        );
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }

    /// Generate XML with actual relationship IDs from the mapper.
    pub(super) fn write_contents_with_rels(
        &self,
        xml: &mut String,
        rel_mapper: &crate::writer::relmap::RelationshipMapper,
        preserve_section: bool,
    ) -> Result<()> {
        // Global counters for hyperlinks and images across all paragraphs
        let mut hyperlink_counter = 0;
        let mut image_counter = 0;

        for element in &self.elements {
            match element {
                BodyElement::Paragraph(p) => {
                    p.to_xml_with_rels(
                        xml,
                        rel_mapper,
                        &mut hyperlink_counter,
                        &mut image_counter,
                    )?;
                },
                BodyElement::Table(t) => t.to_xml(xml)?, // Tables don't need rel mapping for now
                BodyElement::PreservedParagraph(raw)
                | BodyElement::PreservedTable(raw)
                | BodyElement::PreservedOther(raw)
                | BodyElement::PreservedAlt(raw, _) => xml.push_str(raw),
                BodyElement::PreservedSectionProperties(raw) if preserve_section => {
                    xml.push_str(raw);
                },
                BodyElement::PreservedSectionProperties(_) => {},
            }
        }
        Ok(())
    }
}

pub(super) fn collect_section_parts(
    section: &SectionProperties,
    parts: &mut Vec<(bool, super::super::section::SectionHeaderFooterPart)>,
) -> Result<()> {
    section.validate()?;
    for reference in &section.headers {
        if let Some(part) = &reference.part {
            parts.push((true, part.clone()));
        }
    }
    for reference in &section.footers {
        if let Some(part) = &reference.part {
            parts.push((false, part.clone()));
        }
    }
    Ok(())
}

pub(super) fn collect_explicit_section_relationships(
    section: &SectionProperties,
    relationships: &mut Vec<(String, bool)>,
) {
    for reference in &section.headers {
        if let Some(id) = &reference.relationship_id {
            relationships.push((id.clone(), true));
        }
    }
    for reference in &section.footers {
        if let Some(id) = &reference.relationship_id {
            relationships.push((id.clone(), false));
        }
    }
}

fn word_ranges(xml: &str, target: &[u8]) -> Result<Vec<(usize, usize)>> {
    let mut ranges = Vec::new();
    crate::namespace::scan_word_element_ranges(xml.as_bytes(), &[target], |_, start, length| {
        let start = usize::try_from(start)
            .map_err(|_| Error::InvalidFormat("Word range overflow".to_string()))?;
        let length = usize::try_from(length)
            .map_err(|_| Error::InvalidFormat("Word range overflow".to_string()))?;
        ranges.push((start, start + length));
        Ok(())
    })?;
    Ok(ranges)
}

fn paragraph_section_range(xml: &str) -> Result<Option<(usize, usize)>> {
    let sections = word_ranges(xml, b"sectPr")?;
    if sections.len() > 1 {
        return Err(Error::InvalidFormat(
            "paragraph contains multiple section properties".to_string(),
        ));
    }
    let Some(section) = sections.first().copied() else {
        return Ok(None);
    };
    let properties = word_ranges(xml, b"pPr")?;
    if properties.len() != 1 || section.0 < properties[0].0 || section.1 > properties[0].1 {
        return Err(Error::InvalidFormat(
            "paragraph section properties must be inside one pPr".to_string(),
        ));
    }
    let close = xml[..properties[0].1]
        .rfind("</")
        .unwrap_or(properties[0].1);
    if !xml[section.1..close].trim().is_empty() {
        return Err(Error::InvalidFormat(
            "paragraph section properties must be the final pPr child".to_string(),
        ));
    }
    SectionProperties::from_xml(&xml[section.0..section.1])?;
    Ok(Some(section))
}

fn insert_paragraph_property(xml: &str, property: &str) -> Result<String> {
    let properties = word_ranges(xml, b"pPr")?;
    if properties.len() > 1 {
        return Err(Error::InvalidFormat(
            "paragraph contains multiple pPr elements".to_string(),
        ));
    }
    if let Some((start, end)) = properties.first().copied() {
        if xml[start..end].trim_end().ends_with("/>") {
            let empty_end = xml[..end].rfind("/>").ok_or_else(|| {
                Error::InvalidFormat("invalid empty paragraph properties".to_string())
            })?;
            let name_end = xml[start + 1..]
                .find(|ch: char| ch.is_whitespace() || ch == '/' || ch == '>')
                .map(|offset| start + 1 + offset)
                .ok_or_else(|| Error::InvalidFormat("invalid pPr name".to_string()))?;
            let name = &xml[start + 1..name_end];
            return Ok(format!(
                "{}>{property}</{name}>{}",
                &xml[..empty_end],
                &xml[end..]
            ));
        }
        let close = xml[..end].rfind("</").ok_or_else(|| {
            Error::InvalidFormat("paragraph properties are not closed".to_string())
        })?;
        return Ok(format!("{}{property}{}", &xml[..close], &xml[close..]));
    }

    let open_end = xml
        .find('>')
        .ok_or_else(|| Error::InvalidFormat("paragraph opening element is missing".to_string()))?;
    if xml[..=open_end].trim_end().ends_with("/>") {
        let empty_end = xml[..=open_end].rfind("/>").expect("checked");
        let name_end = xml[1..]
            .find(|ch: char| ch.is_whitespace() || ch == '/' || ch == '>')
            .map(|offset| 1 + offset)
            .ok_or_else(|| Error::InvalidFormat("invalid paragraph name".to_string()))?;
        let name = &xml[1..name_end];
        return Ok(format!(
            "{}><w:pPr>{property}</w:pPr></{name}>{}",
            &xml[..empty_end],
            &xml[open_end + 1..]
        ));
    }
    Ok(format!(
        "{}<w:pPr>{property}</w:pPr>{}",
        &xml[..=open_end],
        &xml[open_end + 1..]
    ))
}

fn preserved_body_kind(
    namespace: &quick_xml::name::ResolveResult<'_>,
    local_name: &[u8],
) -> PreservedBodyKind {
    if crate::namespace::is_wordprocessing_namespace(namespace) {
        return match local_name {
            b"p" => PreservedBodyKind::Paragraph,
            b"tbl" => PreservedBodyKind::Table,
            b"sectPr" => PreservedBodyKind::SectionProperties,
            b"altChunk" => PreservedBodyKind::Alt,
            _ => PreservedBodyKind::Other,
        };
    }
    PreservedBodyKind::Other
}

fn ensure_writer_namespace_declarations(prefix: &str) -> Result<String> {
    const REQUIRED: [(&str, &str); 4] = [
        (
            "w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ),
        (
            "r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        (
            "wp",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ),
        ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    ];

    let declarations = REQUIRED
        .iter()
        .filter(|(namespace_prefix, _)| !has_namespace_declaration(prefix, namespace_prefix))
        .map(|(namespace_prefix, namespace)| format!(r#" xmlns:{namespace_prefix}="{namespace}""#))
        .collect::<String>();
    if declarations.is_empty() {
        return Ok(prefix.to_string());
    }
    let insertion = prefix
        .rfind('>')
        .ok_or_else(|| Error::InvalidFormat("Word body opening tag is incomplete".to_string()))?;
    let mut augmented = String::with_capacity(prefix.len() + declarations.len());
    augmented.push_str(&prefix[..insertion]);
    augmented.push_str(&declarations);
    augmented.push_str(&prefix[insertion..]);
    Ok(augmented)
}

fn has_namespace_declaration(xml: &str, namespace_prefix: &str) -> bool {
    let needle = format!("xmlns:{namespace_prefix}");
    xml.match_indices(&needle).any(|(start, _)| {
        let before_is_boundary = start == 0
            || xml.as_bytes()[start - 1].is_ascii_whitespace()
            || xml.as_bytes()[start - 1] == b'<';
        let mut after = start + needle.len();
        while xml
            .as_bytes()
            .get(after)
            .is_some_and(u8::is_ascii_whitespace)
        {
            after += 1;
        }
        before_is_boundary && xml.as_bytes().get(after) == Some(&b'=')
    })
}

fn push_preserved_body_range(
    body: &mut DocumentBody,
    xml: &str,
    last_content_end: &mut usize,
    kind: PreservedBodyKind,
    start: usize,
    end: usize,
    active_alts: &[PreservedAltRange],
) -> Result<()> {
    if start > *last_content_end {
        push_raw_body_xml(
            body,
            PreservedBodyKind::Other,
            xml,
            *last_content_end,
            start,
            active_alts,
        )?;
    }
    push_raw_body_xml(body, kind, xml, start, end, active_alts)?;
    *last_content_end = end;
    Ok(())
}

fn push_raw_body_xml(
    body: &mut DocumentBody,
    kind: PreservedBodyKind,
    xml: &str,
    start: usize,
    end: usize,
    active_alts: &[PreservedAltRange],
) -> Result<()> {
    let source = xml
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid Word body element range".to_string()))?;
    match kind {
        PreservedBodyKind::Paragraph => body
            .elements
            .push(BodyElement::PreservedParagraph(source.to_string())),
        PreservedBodyKind::Table => body
            .elements
            .push(BodyElement::PreservedTable(source.to_string())),
        PreservedBodyKind::SectionProperties => body
            .elements
            .push(BodyElement::PreservedSectionProperties(source.to_string())),
        PreservedBodyKind::Alt => {
            let anchor = active_alts
                .iter()
                .find(|anchor| anchor.start == start && anchor.end == end)
                .ok_or_else(|| {
                    Error::InvalidFormat("direct altChunk body child is not MCE-active".to_string())
                })?;
            body.elements.push(BodyElement::PreservedAlt(
                source.to_string(),
                anchor.chunk.clone(),
            ));
        },
        PreservedBodyKind::Other => {
            let mut cursor = start;
            for anchor in active_alts
                .iter()
                .filter(|anchor| anchor.start >= start && anchor.end <= end)
            {
                let prefix = xml.get(cursor..anchor.start).ok_or_else(|| {
                    Error::InvalidFormat("active altChunk prefix range is invalid".to_string())
                })?;
                if !prefix.is_empty() {
                    body.elements
                        .push(BodyElement::PreservedOther(prefix.to_string()));
                }
                let raw = xml.get(anchor.start..anchor.end).ok_or_else(|| {
                    Error::InvalidFormat("active altChunk range is invalid".to_string())
                })?;
                body.elements.push(BodyElement::PreservedAlt(
                    raw.to_string(),
                    anchor.chunk.clone(),
                ));
                cursor = anchor.end;
            }
            let suffix = xml.get(cursor..end).ok_or_else(|| {
                Error::InvalidFormat("active altChunk suffix range is invalid".to_string())
            })?;
            if !suffix.is_empty() {
                body.elements
                    .push(BodyElement::PreservedOther(suffix.to_string()));
            }
        },
    }
    Ok(())
}

/// A body element (paragraph, table, or exact preserved XML).
#[derive(Debug)]
#[allow(clippy::large_enum_variant)] // writer-internal type; variants are moved, not compared
pub(crate) enum BodyElement {
    Paragraph(MutableParagraph),
    Table(MutableTable),
    PreservedParagraph(String),
    PreservedTable(String),
    PreservedSectionProperties(String),
    PreservedAlt(String, Chunk),
    PreservedOther(String),
}
