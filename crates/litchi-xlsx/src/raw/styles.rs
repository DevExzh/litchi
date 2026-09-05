//! Bounded parser for the shared `SpreadsheetML` cell-format table.

use std::io::BufRead;

use litchi_ooxml_common::{
    mce::{
        Capabilities, Name, SemanticElement, SemanticEnd, SemanticEvent, StreamError, StreamLimits,
        process_markup_compatibility_stream_with_observers,
    },
    xml::unqualified_attribute_value,
};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, allocation, invalid};
use crate::raw::namespace::is_spreadsheetml_name;

// [MS-OE376] 2.1.728 limits the cellXfs collection to 65,430 records.
const MAX_CELL_FORMATS: u32 = 65_430;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Context {
    Styles,
    CellFormats,
    Format,
    Other,
}

/// Validated facts needed by the semantic shared-style facade.
#[derive(Debug)]
pub(crate) struct Catalog {
    cell_formats: u32,
}

impl Catalog {
    pub(crate) const fn len(&self) -> u32 {
        self.cell_formats
    }
}

pub(crate) fn parse(content: &[u8]) -> Result<Catalog> {
    let processed = litchi_ooxml_common::mce::process_ooxml(content)?;
    let mut reader = NsReader::from_reader(processed.as_ref());
    let mut stack = Vec::new();
    let mut closed_root = false;
    let mut seen_cell_formats = false;
    let mut declared_count = None;
    let mut actual_count = 0u32;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet")
                {
                    return Err(invalid(
                        "styles XML must have one SpreadsheetML styleSheet root",
                    ));
                }
                stack.push(Context::Styles);
            },
            Event::Empty(element) if stack.is_empty() => {
                if closed_root || !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet")
                {
                    return Err(invalid(
                        "styles XML must have one SpreadsheetML styleSheet root",
                    ));
                }
                closed_root = true;
            },
            Event::Start(element) => {
                let parent = current(&stack)?;
                let context = start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut seen_cell_formats,
                    &mut declared_count,
                    &mut actual_count,
                )?;
                stack.push(context);
            },
            Event::Empty(element) => {
                let parent = current(&stack)?;
                start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &mut seen_cell_formats,
                    &mut declared_count,
                    &mut actual_count,
                )?;
            },
            Event::End(element) => {
                let ended = stack
                    .pop()
                    .ok_or_else(|| invalid("styles XML has a closing element outside its root"))?;
                if ended == Context::Styles {
                    if !is_spreadsheetml_name(&namespace, element.name(), b"styleSheet") {
                        return Err(invalid("styles XML has an invalid root closing element"));
                    }
                    closed_root = true;
                }
            },
            Event::Eof if !closed_root || !stack.is_empty() => {
                return Err(invalid(
                    "styles XML has a missing or unterminated SpreadsheetML styleSheet root",
                ));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }

    if !seen_cell_formats {
        return Err(invalid("styles XML is missing the cellXfs collection"));
    }
    if actual_count == 0 {
        return Err(invalid("styles XML cellXfs collection must not be empty"));
    }
    if let Some(declared) = declared_count
        && declared != actual_count
    {
        return Err(invalid(format!(
            "styles XML declares {declared} cell formats but contains {actual_count}"
        )));
    }
    Ok(Catalog {
        cell_formats: actual_count,
    })
}

/// Count direct semantic `cellXfs/xf` records from a source-backed XML stream.
///
/// Markup compatibility is selected by the shared streaming processor and the
/// observer consumes only its semantic events. No projected XML, style record,
/// or catalog is materialized. The observer still validates the one
/// `styleSheet` root, every semantic frame closure, and the same `cellXfs`
/// invariants as [`parse`].
#[expect(
    clippy::result_large_err,
    reason = "The stream error retains typed MCE, input, and callback diagnostics; boxing it would change the established API."
)]
pub(crate) fn stream_count(
    input: &mut dyn BufRead,
    capabilities: &Capabilities,
    limits: &StreamLimits,
) -> std::result::Result<u32, StreamError<crate::Error, crate::Error>> {
    let mut state = StreamCounter::new(limits.processing.max_depth);
    let result = process_markup_compatibility_stream_with_observers(
        input,
        capabilities,
        limits,
        |_| Ok::<(), crate::Error>(()),
        |event| state.event(event),
    );
    match result {
        Ok(_) => state.finish().map_err(|error| StreamError::Callback {
            raw_error: None,
            active_error: Some(error),
        }),
        Err(error) => Err(error),
    }
}

#[derive(Debug)]
struct StreamFrame {
    context: Context,
    namespace: String,
    local_name: String,
}

#[derive(Debug)]
struct StreamCounter {
    max_depth: usize,
    stack: Vec<StreamFrame>,
    closed_root: bool,
    seen_cell_formats: bool,
    declared_count: Option<u32>,
    actual_count: u32,
}

impl StreamCounter {
    fn new(max_depth: usize) -> Self {
        Self {
            max_depth,
            stack: Vec::new(),
            closed_root: false,
            seen_cell_formats: false,
            declared_count: None,
            actual_count: 0,
        }
    }

    fn event(&mut self, event: SemanticEvent<'_>) -> Result<()> {
        match event {
            SemanticEvent::Start(element) => self.start(element, false),
            SemanticEvent::Empty(element) => self.start(element, true),
            SemanticEvent::End(element) => self.end(element),
            SemanticEvent::Text(_)
            | SemanticEvent::CData(_)
            | SemanticEvent::Comment(_)
            | SemanticEvent::Decl(_)
            | SemanticEvent::GeneralRef(_) => Ok(()),
            _ => Ok(()),
        }
    }

    fn start(&mut self, element: SemanticElement<'_>, empty: bool) -> Result<()> {
        if self.stack.is_empty() {
            if self.closed_root
                || !is_spreadsheetml_name_semantic(&element.expanded_name, b"styleSheet")
            {
                return Err(invalid(
                    "styles XML must have one SpreadsheetML styleSheet root",
                ));
            }
            if empty {
                self.closed_root = true;
            } else {
                self.push(Context::Styles, &element.expanded_name)?;
            }
            return Ok(());
        }

        let parent = self
            .stack
            .last()
            .map(|frame| frame.context)
            .ok_or_else(|| invalid("styles XML content appears outside its root"))?;
        let context = stream_start(
            parent,
            &element,
            &mut self.seen_cell_formats,
            &mut self.declared_count,
            &mut self.actual_count,
        )?;
        if !empty {
            self.push(context, &element.expanded_name)?;
        }
        Ok(())
    }

    fn end(&mut self, element: SemanticEnd<'_>) -> Result<()> {
        let frame = self
            .stack
            .pop()
            .ok_or_else(|| invalid("styles XML has a closing element outside its root"))?;
        if frame.namespace != element.expanded_name.namespace
            || frame.local_name != element.expanded_name.local_name
        {
            return Err(invalid("styles XML has an unexpected semantic end"));
        }
        if frame.context == Context::Styles {
            self.closed_root = true;
        }
        Ok(())
    }

    fn push(&mut self, context: Context, name: &Name) -> Result<()> {
        if self.stack.len() >= self.max_depth {
            return Err(invalid(format!(
                "styles XML exceeds {} levels",
                self.max_depth
            )));
        }
        self.stack
            .try_reserve(1)
            .map_err(|source| allocation("styles XML element stack", source))?;

        let mut namespace = String::new();
        namespace
            .try_reserve(name.namespace.len())
            .map_err(|source| allocation("styles XML element namespace", source))?;
        namespace.push_str(&name.namespace);

        let mut local_name = String::new();
        local_name
            .try_reserve(name.local_name.len())
            .map_err(|source| allocation("styles XML element local name", source))?;
        local_name.push_str(&name.local_name);

        self.stack.push(StreamFrame {
            context,
            namespace,
            local_name,
        });
        Ok(())
    }

    fn finish(self) -> Result<u32> {
        if !self.closed_root || !self.stack.is_empty() {
            return Err(invalid(
                "styles XML has a missing or unterminated SpreadsheetML styleSheet root",
            ));
        }
        if !self.seen_cell_formats {
            return Err(invalid("styles XML is missing the cellXfs collection"));
        }
        if self.actual_count == 0 {
            return Err(invalid("styles XML cellXfs collection must not be empty"));
        }
        if let Some(declared) = self.declared_count
            && declared != self.actual_count
        {
            return Err(invalid(format!(
                "styles XML declares {declared} cell formats but contains {}",
                self.actual_count
            )));
        }
        Ok(self.actual_count)
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments correspond directly to the cell-format wire attributes"
)]
fn stream_start(
    parent: Context,
    element: &SemanticElement<'_>,
    seen_cell_formats: &mut bool,
    declared_count: &mut Option<u32>,
    actual_count: &mut u32,
) -> Result<Context> {
    if parent == Context::Styles
        && is_spreadsheetml_name_semantic(&element.expanded_name, b"cellXfs")
    {
        if *seen_cell_formats {
            return Err(invalid("styles XML has duplicate cellXfs collections"));
        }
        *seen_cell_formats = true;
        *declared_count = semantic_count(element)?;
        return Ok(Context::CellFormats);
    }
    if parent == Context::CellFormats
        && is_spreadsheetml_name_semantic(&element.expanded_name, b"xf")
    {
        *actual_count = actual_count
            .checked_add(1)
            .filter(|count| *count <= MAX_CELL_FORMATS)
            .ok_or_else(|| {
                invalid(format!(
                    "styles XML contains more than {MAX_CELL_FORMATS} cell formats"
                ))
            })?;
        return Ok(Context::Format);
    }
    Ok(Context::Other)
}

fn semantic_count(element: &SemanticElement<'_>) -> Result<Option<u32>> {
    let mut value = None;
    for attribute in &element.attributes {
        if attribute.expanded_name.namespace.is_empty()
            && attribute.expanded_name.local_name == "count"
        {
            if value.is_some() {
                return Err(invalid("duplicate XML attribute 'count'"));
            }
            let decoded = attribute.decoded_value.as_ref();
            let count = decoded
                .parse::<u32>()
                .map_err(|_source| invalid(format!("invalid cellXfs count '{decoded}'")))?;
            if count > MAX_CELL_FORMATS {
                return Err(invalid(format!("cellXfs count exceeds {MAX_CELL_FORMATS}")));
            }
            value = Some(count);
        }
    }
    Ok(value)
}

fn is_spreadsheetml_name_semantic(name: &Name, local_name: &[u8]) -> bool {
    name.local_name.as_bytes() == local_name
        && (name.namespace.as_bytes() == crate::raw::namespace::SPREADSHEETML_NAMESPACE
            || name.namespace.as_bytes() == crate::raw::namespace::STRICT_SPREADSHEETML_NAMESPACE)
}

#[allow(
    clippy::too_many_arguments,
    reason = "arguments correspond directly to the cell-format wire attributes"
)]
fn start(
    parent: Context,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    seen_cell_formats: &mut bool,
    declared_count: &mut Option<u32>,
    actual_count: &mut u32,
) -> Result<Context> {
    if parent == Context::Styles && is_spreadsheetml_name(namespace, element.name(), b"cellXfs") {
        if *seen_cell_formats {
            return Err(invalid("styles XML has duplicate cellXfs collections"));
        }
        *seen_cell_formats = true;
        *declared_count = count(element, decoder)?;
        return Ok(Context::CellFormats);
    }
    if parent == Context::CellFormats && is_spreadsheetml_name(namespace, element.name(), b"xf") {
        *actual_count = actual_count
            .checked_add(1)
            .filter(|count| *count <= MAX_CELL_FORMATS)
            .ok_or_else(|| {
                invalid(format!(
                    "styles XML contains more than {MAX_CELL_FORMATS} cell formats"
                ))
            })?;
        return Ok(Context::Format);
    }
    Ok(Context::Other)
}

fn count(element: &BytesStart<'_>, decoder: Decoder) -> Result<Option<u32>> {
    let Some(value) = unqualified_attribute_value(element, b"count", decoder)? else {
        return Ok(None);
    };
    let value = value
        .parse::<u32>()
        .map_err(|_source| invalid(format!("invalid cellXfs count '{value}'")))?;
    if value > MAX_CELL_FORMATS {
        return Err(invalid(format!("cellXfs count exceeds {MAX_CELL_FORMATS}")));
    }
    Ok(Some(value))
}

fn current(stack: &[Context]) -> Result<Context> {
    stack
        .last()
        .copied()
        .ok_or_else(|| invalid("styles XML content appears outside its root"))
}

#[cfg(test)]
mod tests {
    use litchi_ooxml_common::mce::{Capabilities, StreamError, StreamLimits};

    use super::*;

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    type Streaming0364Result =
        std::result::Result<u32, Box<StreamError<crate::Error, crate::Error>>>;

    fn streaming_0364_count(xml: &str) -> Streaming0364Result {
        let mut input = std::io::Cursor::new(xml.as_bytes());
        stream_count(
            &mut input,
            &Capabilities::default(),
            &StreamLimits::default(),
        )
        .map_err(Box::new)
    }

    fn streaming_0364_count_with(
        xml: &str,
        capabilities: &Capabilities,
        limits: &StreamLimits,
    ) -> Streaming0364Result {
        let mut input = std::io::Cursor::new(xml.as_bytes());
        stream_count(&mut input, capabilities, limits).map_err(Box::new)
    }

    #[test]
    fn counts_only_direct_cell_formats_after_mce_processing() {
        let xml = format!(
            r#"<styleSheet xmlns="{S}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:f="urn:future" mc:Ignorable="f"><cellXfs count="2"><xf/><xf><alignment/></xf></cellXfs><f:cellXfs><f:xf/></f:cellXfs></styleSheet>"#
        );
        assert_eq!(parse(xml.as_bytes()).expect("styles").len(), 2);
    }

    #[test]
    fn rejects_missing_duplicate_empty_and_mismatched_tables() {
        let cases = [
            format!(r#"<styleSheet xmlns="{S}"/>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs/><cellXfs><xf/></cellXfs></styleSheet>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs/></styleSheet>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs count="2"><xf/></cellXfs></styleSheet>"#),
        ];
        for xml in cases {
            assert!(parse(xml.as_bytes()).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn streaming_0364_counts_direct_cell_xfs() {
        let xml = format!(
            r#"<styleSheet xmlns="{S}"><cellXfs count="3"><xf/><xf><alignment/></xf><xf/></cellXfs></styleSheet>"#
        );
        assert_eq!(streaming_0364_count(&xml).expect("styles"), 3);
    }

    #[test]
    fn streaming_0364_ignores_nested_and_unrelated_xfs() {
        let xml = format!(
            r#"<styleSheet xmlns="{S}"><unrelated><xf/></unrelated><cellXfs count="2"><xf><nested><xf/></nested></xf><unrelated><xf/></unrelated><xf/></cellXfs></styleSheet>"#
        );
        assert_eq!(streaming_0364_count(&xml).expect("styles"), 2);
    }

    #[test]
    fn streaming_0364_enforces_declared_count() {
        let exact = format!(
            r#"<styleSheet xmlns="{S}"><cellXfs count="2"><xf/><xf/></cellXfs></styleSheet>"#
        );
        let mismatch = format!(
            r#"<styleSheet xmlns="{S}"><cellXfs count="1"><xf/><xf/></cellXfs></styleSheet>"#
        );
        assert_eq!(streaming_0364_count(&exact).expect("styles"), 2);
        assert!(streaming_0364_count(&mismatch).is_err());
    }

    #[test]
    fn streaming_0364_requires_nonempty_cell_xfs() {
        let missing = format!(r#"<styleSheet xmlns="{S}"/>"#);
        let empty = format!(r#"<styleSheet xmlns="{S}"><cellXfs/></styleSheet>"#);
        assert!(streaming_0364_count(&missing).is_err());
        assert!(streaming_0364_count(&empty).is_err());
    }

    #[test]
    fn streaming_0364_selects_mce_choice_or_fallback() {
        let xml = format!(
            r#"<styleSheet xmlns="{S}" xmlns:mc="{MC}" xmlns:x="urn:future"><mc:AlternateContent><mc:Choice Requires="x"><cellXfs count="1"><xf/></cellXfs></mc:Choice><mc:Fallback><cellXfs count="2"><xf/><xf/></cellXfs></mc:Fallback></mc:AlternateContent></styleSheet>"#
        );
        assert_eq!(streaming_0364_count(&xml).expect("fallback styles"), 2);

        let mut capabilities = Capabilities::default();
        capabilities.understand_namespace("urn:future");
        assert_eq!(
            streaming_0364_count_with(&xml, &capabilities, &StreamLimits::default())
                .expect("choice styles"),
            1
        );
    }

    #[test]
    fn streaming_0364_rejects_malformed_tail_and_endings() {
        let cases = [
            format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet><tail/>"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet>tail"#),
            format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></stylesSheet>"#),
        ];
        for xml in cases {
            assert!(streaming_0364_count(&xml).is_err(), "accepted {xml}");
        }
    }

    #[test]
    fn streaming_0364_honors_exact_event_limit() {
        let xml = format!(r#"<styleSheet xmlns="{S}"><cellXfs><xf/></cellXfs></styleSheet>"#);
        let exact = StreamLimits {
            max_events: 5,
            ..StreamLimits::default()
        };
        assert_eq!(
            streaming_0364_count_with(&xml, &Capabilities::default(), &exact).expect("five events"),
            1
        );

        let under = StreamLimits {
            max_events: 4,
            ..exact
        };
        assert!(streaming_0364_count_with(&xml, &Capabilities::default(), &under).is_err());
    }
}
