//! Bounded XML codecs for the modern comment and author parts.

#[path = "codec/comments/mod.rs"]
mod comments;

mod authors {
    use super::super::model::{Author, Authors, NamespaceDeclaration};
    use super::super::{MAX_AUTHORS, MAX_BYTES, MAX_DEPTH, MAX_NODES, MAX_STRING_BYTES, P188};
    use crate::{Error, Result};
    use litchi_core::xml::ReaderOrigin;
    use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
    use litchi_ooxml_common::{custom_xml::valid_guid, mce::process_ooxml};
    use quick_xml::encoding::Decoder;
    use quick_xml::events::{BytesStart, Event};
    use quick_xml::name::ResolveResult;
    use quick_xml::reader::NsReader;
    use std::collections::{HashMap, HashSet};

    #[derive(Debug)]
    enum FrameKind {
        Root,
        Author { index: usize, seen_extension: bool },
        RawExtension { index: usize, start: usize },
        Opaque,
    }

    #[derive(Debug)]
    struct Frame {
        kind: FrameKind,
        namespace: String,
        local: String,
    }

    fn parse_author_list(xml: &[u8]) -> Result<Authors> {
        if xml.len() > MAX_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::PartBytes,
                xml.len(),
                MAX_BYTES,
            ));
        }
        let selected = process_ooxml(xml)?;
        if selected.len() > MAX_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                selected.len(),
                MAX_BYTES,
            ));
        }
        parse_processed_author_list(selected.as_ref())
    }

    /// Where the author-list reader is relative to the `p188:authorLst` root.
    ///
    /// ```text
    /// Prolog --p188:authorLst start tag--> Open
    /// Prolog --empty p188:authorLst tag--> Closed
    /// Open   --p188:authorLst end tag----> Closed
    /// any    --declaration, comment, whitespace--> unchanged
    /// ```
    ///
    /// The last line holds in every phase: a declaration is accepted anywhere.
    /// Refused: any other top-level element, before or after the root; text, CDATA
    /// or a character reference outside the root unless it is whitespace; a
    /// processing instruction or DOCTYPE anywhere; and the end of input before
    /// `Closed`.
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum RootPhase {
        /// No root start tag has been read.
        Prolog,
        /// Inside the root.
        Open,
        /// The root has ended.
        Closed,
    }

    fn parse_processed_author_list(bytes: &[u8]) -> Result<Authors> {
        if bytes.len() > MAX_BYTES {
            return Err(limit(
                crate::error::CodecLimitKind::XmlBytes,
                bytes.len(),
                MAX_BYTES,
            ));
        }
        let mut reader = NsReader::from_reader(bytes);
        let origin = ReaderOrigin::of(bytes);
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack: Vec<Frame> = Vec::new();
        let mut phase = RootPhase::Prolog;
        let mut root_prefix = String::new();
        let mut namespace_declarations = Vec::new();
        let mut authors = Vec::new();
        let mut nodes = 0usize;

        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("XML source position exceeds usize"))?;
            let decoder = reader.decoder();
            let (resolved, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(xml_error)?;
            let namespace = resolve_namespace(resolved)?;
            let empty = matches!(&event, Event::Empty(_));
            match event {
                Event::Start(element) | Event::Empty(element) => {
                    nodes = nodes.checked_add(1).ok_or_else(|| {
                        limit(
                            crate::error::CodecLimitKind::XmlNodes,
                            usize::MAX,
                            MAX_NODES,
                        )
                    })?;
                    if nodes > MAX_NODES {
                        return Err(limit(
                            crate::error::CodecLimitKind::XmlNodes,
                            nodes,
                            MAX_NODES,
                        ));
                    }
                    if stack.len() + 1 > MAX_DEPTH {
                        return Err(limit(
                            crate::error::CodecLimitKind::XmlDepth,
                            stack.len() + 1,
                            MAX_DEPTH,
                        ));
                    }
                    let local = decode_name(element.local_name().as_ref())?;
                    let kind = if stack.is_empty() {
                        if phase != RootPhase::Prolog || namespace != P188 || local != "authorLst" {
                            return Err(invalid(
                                "modern Comment Author root must be p188:authorLst",
                            ));
                        }
                        root_prefix = element_prefix(&element)?;
                        namespace_declarations =
                            namespace_declarations_from(&element, decoder, Some(&root_prefix))?;
                        no_non_namespace_attributes(&element)?;
                        phase = RootPhase::Open;
                        FrameKind::Root
                    } else {
                        child_frame(
                            &mut authors,
                            stack.last_mut().ok_or_else(|| {
                                invalid("modern Comment Author child has no parent")
                            })?,
                            &element,
                            decoder,
                            &namespace,
                            &local,
                            start,
                        )?
                    };
                    let frame = Frame {
                        kind,
                        namespace,
                        local,
                    };
                    if empty {
                        attach_extension(
                            &frame.kind,
                            bytes,
                            origin
                                .offset(reader.buffer_position())
                                .ok_or_else(|| invalid("XML source position exceeds usize"))?,
                            &mut authors,
                        )?;
                        if matches!(frame.kind, FrameKind::Root) {
                            phase = RootPhase::Closed;
                        }
                    } else {
                        stack.push(frame);
                    }
                },
                Event::End(element) => {
                    let frame = stack.pop().ok_or_else(|| {
                        invalid("unexpected modern Comment Author closing element")
                    })?;
                    let local = decode_name(element.local_name().as_ref())?;
                    if frame.namespace != namespace || frame.local != local {
                        return Err(invalid("mismatched modern Comment Author closing element"));
                    }
                    attach_extension(
                        &frame.kind,
                        bytes,
                        origin
                            .offset(reader.buffer_position())
                            .ok_or_else(|| invalid("XML source position exceeds usize"))?,
                        &mut authors,
                    )?;
                    if matches!(frame.kind, FrameKind::Root) {
                        phase = RootPhase::Closed;
                    }
                },
                Event::Text(text) => {
                    if !inside_extension(&stack) {
                        let decoded = text.decode().map_err(xml_error)?;
                        let value = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                        if !value.trim().is_empty() {
                            return Err(invalid(
                                "unexpected text in modern Comment Author metadata",
                            ));
                        }
                    }
                },
                Event::CData(text) => {
                    if !inside_extension(&stack)
                        && !text.decode().map_err(xml_error)?.trim().is_empty()
                    {
                        return Err(invalid(
                            "unexpected CDATA in modern Comment Author metadata",
                        ));
                    }
                },
                Event::GeneralRef(reference) => {
                    let value = litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(xml_error)?;
                    if !inside_extension(&stack) && !value.trim().is_empty() {
                        return Err(invalid(
                            "unexpected character reference in modern Comment Author metadata",
                        ));
                    }
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(invalid("DTD and processing instructions are rejected"));
                },
                // Accepted in every phase, including after the root.
                Event::Decl(_) | Event::Comment(_) => {},
                Event::Eof => break,
            }
            buffer.clear();
        }
        match phase {
            RootPhase::Closed if stack.is_empty() => {},
            RootPhase::Prolog | RootPhase::Open | RootPhase::Closed => {
                return Err(invalid("unterminated modern Comment Author part"));
            },
        }
        let value = Authors {
            root_prefix,
            namespace_declarations,
            authors,
        };
        validate_author_model(&value)?;
        Ok(value)
    }

    fn child_frame(
        authors: &mut Vec<Author>,
        parent: &mut Frame,
        element: &BytesStart<'_>,
        decoder: Decoder,
        namespace: &str,
        local: &str,
        start: usize,
    ) -> Result<FrameKind> {
        match &mut parent.kind {
            FrameKind::Root => {
                if namespace != P188 || local != "author" {
                    return Err(invalid("authorLst permits only p188:author children"));
                }
                if authors.len() >= MAX_AUTHORS {
                    return Err(limit(
                        crate::error::CodecLimitKind::Authors,
                        authors.len().checked_add(1).unwrap_or(usize::MAX),
                        MAX_AUTHORS,
                    ));
                }
                let declarations = namespace_declarations_from(element, decoder, None)?;
                let attributes = known_attributes(
                    element,
                    decoder,
                    &["id", "name", "initials", "userId", "providerId"],
                )?;
                authors.push(parse_author(attributes, declarations)?);
                Ok(FrameKind::Author {
                    index: authors.len() - 1,
                    seen_extension: false,
                })
            },
            FrameKind::Author {
                index,
                seen_extension,
            } => {
                if namespace != P188 || local != "extLst" || *seen_extension {
                    return Err(invalid(
                        "modern Comment Author permits at most one p188:extLst child",
                    ));
                }
                *seen_extension = true;
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::RawExtension {
                    index: *index,
                    start,
                })
            },
            FrameKind::RawExtension { .. } | FrameKind::Opaque => {
                validate_any_attributes(element, decoder)?;
                Ok(FrameKind::Opaque)
            },
        }
    }

    fn parse_author(
        attributes: HashMap<String, String>,
        namespace_declarations: Vec<NamespaceDeclaration>,
    ) -> Result<Author> {
        let id = required(&attributes, "id")?.to_owned();
        validate_guid(&id)?;
        Ok(Author {
            id,
            name: required(&attributes, "name")?.to_owned(),
            initials: attributes.get("initials").cloned(),
            user_id: required(&attributes, "userId")?.to_owned(),
            provider_id: required(&attributes, "providerId")?.to_owned(),
            namespace_declarations,
            extension_xml: None,
        })
    }

    fn attach_extension(
        kind: &FrameKind,
        bytes: &[u8],
        end: usize,
        authors: &mut [Author],
    ) -> Result<()> {
        let FrameKind::RawExtension { index, start } = kind else {
            return Ok(());
        };
        if *start > end || end > bytes.len() {
            return Err(invalid("invalid modern Comment Author extension bounds"));
        }
        authors[*index].extension_xml = Some(bytes[*start..end].to_vec());
        Ok(())
    }

    fn inside_extension(stack: &[Frame]) -> bool {
        stack
            .iter()
            .any(|frame| matches!(frame.kind, FrameKind::RawExtension { .. }))
    }

    fn validate_author_model(value: &Authors) -> Result<()> {
        validate_prefix(&value.root_prefix)?;
        validate_namespaces(&value.namespace_declarations, Some(&value.root_prefix))?;
        if value.authors.len() > MAX_AUTHORS {
            return Err(limit(
                crate::error::CodecLimitKind::Authors,
                value.authors.len(),
                MAX_AUTHORS,
            ));
        }
        let mut ids = HashSet::new();
        for author in &value.authors {
            validate_guid(&author.id)?;
            if !ids.insert(author.id.as_str()) {
                return Err(invalid("duplicate modern Comment Author ID"));
            }
            bounded(&author.name)?;
            if let Some(initials) = &author.initials {
                bounded(initials)?;
            }
            bounded(&author.user_id)?;
            bounded(&author.provider_id)?;
            validate_namespaces(&author.namespace_declarations, None)?;
            if author
                .extension_xml
                .as_ref()
                .is_some_and(|xml| xml.len() > MAX_BYTES)
            {
                return Err(limit(
                    crate::error::CodecLimitKind::ExtensionBytes,
                    author.extension_xml.as_ref().map_or(0, Vec::len),
                    MAX_BYTES,
                ));
            }
        }
        Ok(())
    }

    fn write_author(out: &mut Vec<u8>, prefix: &str, author: &Author) {
        open_tag(out, prefix, "author");
        write_attr(out, "id", &author.id);
        write_attr(out, "name", &author.name);
        if let Some(initials) = &author.initials {
            write_attr(out, "initials", initials);
        }
        write_attr(out, "userId", &author.user_id);
        write_attr(out, "providerId", &author.provider_id);
        write_namespaces(out, &author.namespace_declarations);
        if let Some(extension) = &author.extension_xml {
            out.push(b'>');
            out.extend_from_slice(extension);
            close_tag(out, prefix, "author");
        } else {
            out.extend_from_slice(b"/>");
        }
    }

    fn known_attributes(
        element: &BytesStart<'_>,
        decoder: Decoder,
        allowed: &[&str],
    ) -> Result<HashMap<String, String>> {
        let mut values = HashMap::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let key = decode_name(attribute.key.as_ref())?;
            if is_namespace_attribute(&key) {
                continue;
            }
            if key.contains(':') || !allowed.contains(&key.as_str()) {
                return Err(invalid(format!("unexpected author attribute '{key}'")));
            }
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            bounded(&value)?;
            if values.insert(key.clone(), value).is_some() {
                return Err(invalid(format!("duplicate author attribute '{key}'")));
            }
        }
        Ok(values)
    }

    fn validate_any_attributes(element: &BytesStart<'_>, decoder: Decoder) -> Result<()> {
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?;
            bounded(&value)?;
        }
        Ok(())
    }

    fn no_non_namespace_attributes(element: &BytesStart<'_>) -> Result<()> {
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            if !is_namespace_attribute(&decode_name(attribute.key.as_ref())?) {
                return Err(invalid(
                    "unexpected attribute on modern Comment Author container",
                ));
            }
        }
        Ok(())
    }

    fn namespace_declarations_from(
        element: &BytesStart<'_>,
        decoder: Decoder,
        exclude_prefix: Option<&str>,
    ) -> Result<Vec<NamespaceDeclaration>> {
        let mut result = Vec::new();
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(xml_error)?;
            let key = decode_name(attribute.key.as_ref())?;
            let prefix = if key == "xmlns" {
                Some(String::new())
            } else {
                key.strip_prefix("xmlns:").map(str::to_owned)
            };
            let Some(prefix) = prefix else {
                continue;
            };
            if exclude_prefix == Some(prefix.as_str()) {
                continue;
            }
            let uri = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                .map_err(xml_error)?
                .into_owned();
            result.push(NamespaceDeclaration { prefix, uri });
        }
        validate_namespaces(&result, None)?;
        Ok(result)
    }

    fn validate_namespaces(value: &[NamespaceDeclaration], excluded: Option<&str>) -> Result<()> {
        let mut seen = HashSet::new();
        for declaration in value {
            validate_prefix(&declaration.prefix)?;
            bounded(&declaration.uri)?;
            if declaration.prefix == "xml" || declaration.prefix == "xmlns" {
                return Err(invalid(
                    "reserved XML namespace prefix cannot be redeclared",
                ));
            }
            if excluded == Some(declaration.prefix.as_str()) {
                return Err(invalid(
                    "modern Comment Author namespace prefix is declared twice",
                ));
            }
            if !seen.insert(&declaration.prefix) {
                return Err(invalid("duplicate namespace declaration"));
            }
        }
        Ok(())
    }

    fn validate_prefix(value: &str) -> Result<()> {
        if value.is_empty()
            || (value.bytes().enumerate().all(|(index, byte)| {
                byte.is_ascii_alphanumeric()
                    || byte == b'_'
                    || byte == b'-'
                    || (byte == b'.' && index != 0)
            }) && !value.as_bytes()[0].is_ascii_digit()
                && !value.starts_with('-'))
        {
            Ok(())
        } else {
            Err(invalid(format!("invalid XML namespace prefix '{value}'")))
        }
    }

    fn validate_guid(value: &str) -> Result<()> {
        if valid_guid(value) {
            Ok(())
        } else {
            Err(invalid(format!(
                "invalid modern Comment Author GUID '{value}'"
            )))
        }
    }

    fn required<'a>(attributes: &'a HashMap<String, String>, name: &str) -> Result<&'a str> {
        attributes.get(name).map(String::as_str).ok_or_else(|| {
            invalid(format!(
                "modern Comment Author is missing required '{name}'"
            ))
        })
    }

    fn bounded(value: &str) -> Result<()> {
        if value.len() <= MAX_STRING_BYTES {
            Ok(())
        } else {
            Err(limit(
                crate::error::CodecLimitKind::StringBytes,
                value.len(),
                MAX_STRING_BYTES,
            ))
        }
    }

    fn resolve_namespace(value: ResolveResult<'_>) -> Result<String> {
        match value {
            ResolveResult::Bound(value) => Ok(std::str::from_utf8(value.as_ref())
                .map_err(xml_error)?
                .to_owned()),
            ResolveResult::Unbound => Ok(String::new()),
            ResolveResult::Unknown(prefix) => Err(invalid(format!(
                "unbound XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            ))),
        }
    }

    fn element_prefix(element: &BytesStart<'_>) -> Result<String> {
        let name = decode_name(element.name().as_ref())?;
        Ok(name
            .rsplit_once(':')
            .map_or(String::new(), |(prefix, _)| prefix.to_owned()))
    }

    fn decode_name(value: &[u8]) -> Result<String> {
        Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
    }

    fn is_namespace_attribute(value: &str) -> bool {
        value == "xmlns" || value.starts_with("xmlns:")
    }

    fn open_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
        out.push(b'<');
        qname(out, prefix, local);
    }

    fn close_tag(out: &mut Vec<u8>, prefix: &str, local: &str) {
        out.extend_from_slice(b"</");
        qname(out, prefix, local);
        out.push(b'>');
    }

    fn qname(out: &mut Vec<u8>, prefix: &str, local: &str) {
        if !prefix.is_empty() {
            out.extend_from_slice(prefix.as_bytes());
            out.push(b':');
        }
        out.extend_from_slice(local.as_bytes());
    }

    fn write_namespace_binding(out: &mut Vec<u8>, prefix: &str, uri: &str) {
        out.extend_from_slice(b" xmlns");
        if !prefix.is_empty() {
            out.push(b':');
            out.extend_from_slice(prefix.as_bytes());
        }
        out.extend_from_slice(b"=\"");
        escape(out, uri);
        out.push(b'"');
    }

    fn write_namespaces(out: &mut Vec<u8>, declarations: &[NamespaceDeclaration]) {
        for declaration in declarations {
            write_namespace_binding(out, &declaration.prefix, &declaration.uri);
        }
    }

    fn write_attr(out: &mut Vec<u8>, name: &str, value: &str) {
        out.push(b' ');
        out.extend_from_slice(name.as_bytes());
        out.extend_from_slice(b"=\"");
        escape(out, value);
        out.push(b'"');
    }

    fn escape(out: &mut Vec<u8>, value: &str) {
        for character in value.chars() {
            match character {
                '&' => out.extend_from_slice(b"&amp;"),
                '<' => out.extend_from_slice(b"&lt;"),
                '>' => out.extend_from_slice(b"&gt;"),
                '"' => out.extend_from_slice(b"&quot;"),
                '\t' => out.extend_from_slice(b"&#x9;"),
                '\n' => out.extend_from_slice(b"&#xA;"),
                '\r' => out.extend_from_slice(b"&#xD;"),
                _ => {
                    let mut bytes = [0; 4];
                    out.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
                },
            }
        }
    }

    fn xml_error(error: impl std::fmt::Display) -> Error {
        Error::Xml(error.to_string())
    }

    fn invalid(message: impl Into<String>) -> Error {
        Error::Invalid(message.into())
    }

    fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
        Error::Limit {
            kind: crate::error::LimitKind::Codec(kind),
            observed,
            maximum,
        }
    }

    impl Authors {
        /// Decode XML already admitted by the source physical/MCE boundary.
        pub(crate) fn from_processed(xml: &[u8]) -> Result<Self> {
            parse_processed_author_list(xml)
        }

        /// # Errors
        ///
        /// Returns an error if the input cannot be read or is malformed.
        pub fn parse(xml: &[u8]) -> Result<Self> {
            parse_author_list(xml)
        }

        /// # Errors
        ///
        /// Returns an error if the output cannot be encoded or written.
        pub fn to_xml(&self) -> Result<Vec<u8>> {
            validate_author_model(self)?;
            let mut out = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec();
            open_tag(&mut out, &self.root_prefix, "authorLst");
            write_namespace_binding(&mut out, &self.root_prefix, P188);
            write_namespaces(&mut out, &self.namespace_declarations);
            if self.authors.is_empty() {
                out.extend_from_slice(b"/>");
            } else {
                out.push(b'>');
                for author in &self.authors {
                    write_author(&mut out, &self.root_prefix, author);
                }
                close_tag(&mut out, &self.root_prefix, "authorLst");
            }
            if out.len() > MAX_BYTES {
                return Err(limit(
                    crate::error::CodecLimitKind::OutputBytes,
                    out.len(),
                    MAX_BYTES,
                ));
            }
            parse_author_list(&out)?;
            Ok(out)
        }
    }
}

#[cfg(test)]
mod envelope_tests {
    //! What the author-list reader accepts before, around and after `p188:authorLst`, and the exact error reported otherwise. Each expected
    //! line is the row label followed by the result.

    const DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;

    /// One reader's root element in the forms the rows are built from.
    struct Root<'a> {
        /// The start tag.
        open: &'a str,
        /// Content that makes the element valid for its reader.
        body: &'a str,
        /// The end tag.
        close: &'a str,
        /// The element as one empty-element tag.
        empty: &'a str,
    }

    fn rows(root: &Root<'_>) -> Vec<(&'static str, String)> {
        let Root {
            open,
            body,
            close,
            empty,
        } = *root;
        let document = format!("{open}{body}{close}");
        let d = document.as_str();
        vec![
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECLARATION.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", format!("{open}{close}")),
            ("root then whitespace", format!("{d}\n ")),
            ("root then comment", format!("{d}<!--c-->")),
            ("root then PI", format!("{d}<?pi x?>")),
            ("root then text", format!("{d}text")),
            ("root then second root", format!("{d}{d}")),
            ("root then foreign start", format!("{d}<foreign></foreign>")),
            ("root then foreign empty", format!("{d}<foreign/>")),
            ("declaration after root", format!("{d}{DECLARATION}")),
            ("two declarations", format!("{DECLARATION}{DECLARATION}{d}")),
            (
                "comment before declaration",
                format!("<!--c-->{DECLARATION}{d}"),
            ),
            ("text before root", format!("text{d}")),
            ("unterminated root", format!("{open}{body}")),
            ("stray end tag", format!("{d}</foreign>")),
            ("DOCTYPE", format!("<!DOCTYPE x>{d}")),
            ("top-level CDATA", format!("<![CDATA[x]]>{d}")),
            ("declaration then root", format!("{DECLARATION}{d}")),
            (
                "whitespace before declaration",
                format!(" {DECLARATION}{d}"),
            ),
            ("PI before declaration", format!("<?pi x?>{DECLARATION}{d}")),
            ("comment then root", format!("<!--c-->{d}")),
            ("PI then root", format!("<?pi x?>{d}")),
            ("root then text then second root", format!("{d}text{d}")),
            ("empty-element root then root", format!("{empty}{d}")),
            (
                "declaration after empty-element root",
                format!("{empty}{DECLARATION}"),
            ),
        ]
    }

    fn assert_matrix(root: &Root<'_>, expected: &[&str], read: impl Fn(&[u8]) -> String) {
        let actual = rows(root)
            .into_iter()
            .map(|(label, xml)| format!("{label}: {}", read(xml.as_bytes())))
            .collect::<Vec<_>>();
        assert_eq!(actual, expected, "actual rows:\n{actual:#?}");
    }

    fn outcome<T, E: std::fmt::Debug>(
        result: Result<T, E>,
        summary: impl Fn(T) -> String,
    ) -> String {
        match result {
            Ok(value) => format!("Ok({})", summary(value)),
            Err(error) => format!("{error:?}"),
        }
    }

    /// The reader of an author part already past markup compatibility.
    #[test]
    fn author_list_envelope() {
        let root = Root {
            open: r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">"#,
            body: r#"<p188:author id="{00000000-0000-0000-0000-000000000001}" name="Ada" initials="A" userId="ada" providerId="None"/>"#,
            close: r"</p188:authorLst>",
            empty: r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"/>"#,
        };
        assert_matrix(
            &root,
            &[
                "empty input: Invalid(\"unterminated modern Comment Author part\")",
                "whitespace only: Invalid(\"unterminated modern Comment Author part\")",
                "declaration only: Invalid(\"unterminated modern Comment Author part\")",
                "empty-element root: Ok(0)",
                "start and end root: Ok(0)",
                "root then whitespace: Ok(1)",
                "root then comment: Ok(1)",
                "root then PI: Invalid(\"DTD and processing instructions are rejected\")",
                "root then text: Invalid(\"unexpected text in modern Comment Author metadata\")",
                "root then second root: Invalid(\"modern Comment Author root must be p188:authorLst\")",
                "root then foreign start: Invalid(\"modern Comment Author root must be p188:authorLst\")",
                "root then foreign empty: Invalid(\"modern Comment Author root must be p188:authorLst\")",
                "declaration after root: Ok(1)",
                "two declarations: Ok(1)",
                "comment before declaration: Ok(1)",
                "text before root: Invalid(\"unexpected text in modern Comment Author metadata\")",
                "unterminated root: Invalid(\"unterminated modern Comment Author part\")",
                "stray end tag: Xml(\"ill-formed document: close tag `</foreign>` does not match any open tag\")",
                "DOCTYPE: Invalid(\"DTD and processing instructions are rejected\")",
                "top-level CDATA: Invalid(\"unexpected CDATA in modern Comment Author metadata\")",
                "declaration then root: Ok(1)",
                "whitespace before declaration: Ok(1)",
                "PI before declaration: Invalid(\"DTD and processing instructions are rejected\")",
                "comment then root: Ok(1)",
                "PI then root: Invalid(\"DTD and processing instructions are rejected\")",
                "root then text then second root: Invalid(\"unexpected text in modern Comment Author metadata\")",
                "empty-element root then root: Invalid(\"modern Comment Author root must be p188:authorLst\")",
                "declaration after empty-element root: Ok(0)",
            ],
            |xml| {
                outcome(super::super::model::Authors::from_processed(xml), |value| {
                    value.authors.len().to_string()
                })
            },
        );
    }
}
