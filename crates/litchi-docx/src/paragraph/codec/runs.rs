#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Zero-copy paragraph child traversal.

use crate::error::{Error, Result};
use crate::smart_tag::SmartTag;
use litchi_core::XmlSlice;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::sync::Arc;

use super::super::model::{Paragraph, Run};
use super::xml::is_fragment_word_name;

impl Paragraph {
    /// Get an iterator over the runs in this paragraph.
    ///
    /// Each run represents a `<w:r>` element and may have different formatting.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming boundary detection and shared XML slices.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn runs(&self) -> Result<SmallVec<[Run; 8]>> {
        enum RunEvent {
            Start,
            NestedStart,
            Empty,
            End,
            Eof,
            Other,
        }

        let xml_bytes = self.xml_bytes();
        let (source_arc, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut runs = SmallVec::new();
        let mut run_start = None;
        let mut run_depth = 0usize;
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;

        loop {
            let event_start =
                usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                    Error::InvalidFormat("Word paragraph offset does not fit usize".to_string())
                })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                match event {
                    Event::Start(ref element)
                        if fragment_prefix.is_none()
                            && element.local_name().as_ref() == b"p"
                            && !matches!(namespace, ResolveResult::Bound(_)) =>
                    {
                        fragment_prefix = Some(
                            element
                                .name()
                                .prefix()
                                .map(|prefix| prefix.into_inner().to_vec()),
                        );
                        RunEvent::Other
                    },
                    Event::Start(_) if run_start.is_some() => RunEvent::NestedStart,
                    Event::Start(element)
                        if is_fragment_word_name(
                            &namespace,
                            element.name(),
                            b"r",
                            &fragment_prefix,
                        ) =>
                    {
                        RunEvent::Start
                    },
                    Event::Empty(element)
                        if run_start.is_none()
                            && is_fragment_word_name(
                                &namespace,
                                element.name(),
                                b"r",
                                &fragment_prefix,
                            ) =>
                    {
                        RunEvent::Empty
                    },
                    Event::End(_) if run_start.is_some() => RunEvent::End,
                    Event::Eof => RunEvent::Eof,
                    Event::Start(_)
                    | Event::End(_)
                    | Event::Empty(_)
                    | Event::Text(_)
                    | Event::CData(_)
                    | Event::Comment(_)
                    | Event::Decl(_)
                    | Event::PI(_)
                    | Event::DocType(_)
                    | Event::GeneralRef(_) => RunEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                Error::InvalidFormat("Word paragraph offset does not fit usize".to_string())
            })?;

            match event {
                RunEvent::Start => {
                    run_start = Some(event_start);
                    run_depth = 1;
                },
                RunEvent::NestedStart => {
                    run_depth = run_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word run nesting is too deep".to_string())
                    })?;
                },
                RunEvent::Empty => {
                    Self::push_run_slice(
                        &mut runs,
                        &source_arc,
                        base_offset,
                        event_start,
                        event_end,
                    )?;
                },
                RunEvent::End => {
                    run_depth = run_depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("invalid Word run nesting".to_string())
                    })?;
                    if run_depth == 0 {
                        let Some(start) = run_start.take() else {
                            return Err(Error::InvalidFormat(
                                "missing Word run start offset".to_string(),
                            ));
                        };
                        Self::push_run_slice(
                            &mut runs,
                            &source_arc,
                            base_offset,
                            start,
                            event_end,
                        )?;
                    }
                },
                RunEvent::Eof if run_start.is_some() => {
                    return Err(Error::InvalidFormat("unterminated Word run".to_string()));
                },
                RunEvent::Eof => break,
                RunEvent::Other => {},
            }
        }

        Ok(runs)
    }

    /// Return all run-level smart tags in document order, including nested tags.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn smart_tags(&self) -> Result<Vec<SmartTag>> {
        enum SmartTagEvent {
            Start(bool),
            Empty(bool),
            End(bool),
            Eof,
            Other,
        }

        let xml_bytes = self.xml_bytes();
        let (source_arc, base_offset) = self.xml_data.get_or_create_arc();
        let mut reader = NsReader::from_reader(xml_bytes);
        let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
        let mut depth = 0usize;
        let mut open_tags = Vec::new();
        let mut ranges = Vec::new();

        loop {
            let event_start =
                usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                    Error::InvalidFormat("Word smart-tag offset does not fit usize".into())
                })?;
            let event = {
                let (namespace, event) = reader
                    .read_resolved_event()
                    .map_err(|error| Error::Xml(error.to_string()))?;

                if fragment_prefix.is_none()
                    && let Event::Start(element) = &event
                    && element.local_name().as_ref() == b"p"
                    && !matches!(namespace, ResolveResult::Bound(_))
                {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }

                match event {
                    Event::Start(element) => SmartTagEvent::Start(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::Empty(element) => SmartTagEvent::Empty(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::End(element) => SmartTagEvent::End(is_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"smartTag",
                        &fragment_prefix,
                    )),
                    Event::Eof => SmartTagEvent::Eof,
                    Event::Text(_)
                    | Event::CData(_)
                    | Event::Comment(_)
                    | Event::Decl(_)
                    | Event::PI(_)
                    | Event::DocType(_)
                    | Event::GeneralRef(_) => SmartTagEvent::Other,
                }
            };
            let event_end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
                Error::InvalidFormat("Word smart-tag offset does not fit usize".into())
            })?;

            match event {
                SmartTagEvent::Start(is_smart_tag) => {
                    depth = depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("Word XML nesting is too deep".into())
                    })?;
                    if is_smart_tag {
                        open_tags.push((event_start, depth));
                    }
                },
                SmartTagEvent::Empty(true) => {
                    ranges.push((event_start, event_end));
                },
                SmartTagEvent::End(is_smart_tag) => {
                    if is_smart_tag {
                        let Some((start, tag_depth)) = open_tags.pop() else {
                            return Err(Error::InvalidFormat(
                                "Word smart tag has no opening element".into(),
                            ));
                        };
                        if tag_depth != depth {
                            return Err(Error::InvalidFormat(
                                "invalid nested Word smart tag".into(),
                            ));
                        }
                        ranges.push((start, event_end));
                    }
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| Error::InvalidFormat("invalid Word XML nesting".into()))?;
                },
                SmartTagEvent::Eof if !open_tags.is_empty() || depth != 0 => {
                    return Err(Error::InvalidFormat(
                        "unterminated Word smart-tag XML".into(),
                    ));
                },
                SmartTagEvent::Eof => break,
                SmartTagEvent::Empty(_) | SmartTagEvent::Other => {},
            }
        }

        ranges.sort_unstable_by_key(|&(start, _)| start);
        ranges
            .into_iter()
            .map(|(start, end)| {
                let start = u32::try_from(start).map_err(|_source_error| {
                    Error::InvalidFormat("Word smart-tag offset exceeds u32".into())
                })?;
                let length = u32::try_from(end.checked_sub(start as usize).ok_or_else(|| {
                    Error::InvalidFormat("invalid Word smart-tag byte range".into())
                })?)
                .map_err(|_source_error| {
                    Error::InvalidFormat("Word smart-tag length exceeds u32".into())
                })?;
                let absolute_start = base_offset.checked_add(start).ok_or_else(|| {
                    Error::InvalidFormat("Word smart-tag absolute offset exceeds u32".into())
                })?;
                SmartTag::parse(XmlSlice::new(
                    Arc::clone(&source_arc),
                    absolute_start,
                    length,
                ))
            })
            .collect()
    }

    fn push_run_slice(
        runs: &mut SmallVec<[Run; 8]>,
        source: &Arc<Vec<u8>>,
        base_offset: u32,
        start: usize,
        end: usize,
    ) -> Result<()> {
        let start = u32::try_from(start).map_err(|_source_error| {
            Error::InvalidFormat("Word run offset exceeds u32".to_string())
        })?;
        let length = u32::try_from(
            end.checked_sub(start as usize)
                .ok_or_else(|| Error::InvalidFormat("invalid Word run byte range".to_string()))?,
        )
        .map_err(|_source_error| Error::InvalidFormat("Word run length exceeds u32".to_string()))?;
        let absolute_start = base_offset.checked_add(start).ok_or_else(|| {
            Error::InvalidFormat("Word run absolute offset exceeds u32".to_string())
        })?;
        runs.push(Run::from_slice(XmlSlice::new(
            Arc::clone(source),
            absolute_start,
            length,
        )));
        Ok(())
    }
}
