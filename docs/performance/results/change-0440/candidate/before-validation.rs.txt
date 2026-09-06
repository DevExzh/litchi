//! Namespace, attribute, and structural validation for ODP XML.

use super::super::{
    Attribute, Attributes, BytesStart, DR3D_NAMESPACE, DRAW_NAMESPACE, DrawingAttribute,
    DrawingAttributeNamespace, DrawingShapeKind, Element, Error, HashSet, LocalName, Namespace,
    NsReader, Parser, RawAttribute, ResolveResult, Result, SVG_NAMESPACE, ShapeBuilder,
    TABLE_NAMESPACE, XLINK_NAMESPACE, XmlNamespace, XmlVersion,
};

impl Parser {
    pub(super) fn validate_shape_parent(
        parent: &ShapeBuilder,
        child: DrawingShapeKind,
    ) -> Result<()> {
        match parent.drawing_kind {
            Some(DrawingShapeKind::Group) => {
                if child.is_three_dimensional() && child != DrawingShapeKind::ThreeDimensionalScene
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing objects require a dr3d:scene parent".to_string(),
                    ));
                }
            },
            Some(DrawingShapeKind::ThreeDimensionalScene) => {
                if !child.is_three_dimensional() {
                    return Err(Error::InvalidFormat(
                        "dr3d:scene can only contain 3D lights and objects".to_string(),
                    ));
                }
                if child == DrawingShapeKind::ThreeDimensionalLight
                    && parent.children.iter().any(|existing| {
                        existing.drawing_kind() != Some(DrawingShapeKind::ThreeDimensionalLight)
                    })
                {
                    return Err(Error::InvalidFormat(
                        "dr3d:light elements must precede 3D objects".to_string(),
                    ));
                }
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "nested drawing shapes require a draw:g or dr3d:scene parent".to_string(),
                ));
            },
        }
        Ok(())
    }

    pub(super) fn validate_three_dimensional_child_element(
        parent: Option<&ShapeBuilder>,
        child: Element,
    ) -> Result<()> {
        let Some(parent_kind) = parent.and_then(|builder| builder.drawing_kind) else {
            return Ok(());
        };
        if !parent_kind.is_three_dimensional() {
            return Ok(());
        }
        if parent_kind != DrawingShapeKind::ThreeDimensionalScene {
            return Err(Error::InvalidFormat(
                "3D light and object elements cannot contain child elements".to_string(),
            ));
        }
        match child {
            Element::Shape(shape) if Self::drawing_kind(shape).is_three_dimensional() => Ok(()),
            // `svg:title`, `svg:desc`, `draw:glue-point`, and foreign
            // extension elements are intentionally handled as opaque content.
            Element::Other => Ok(()),
            Element::Page
            | Element::Notes
            | Element::SheetShapes
            | Element::SpreadsheetRoot
            | Element::Shape(_)
            | Element::Image
            | Element::Table
            | Element::Object
            | Element::Plugin
            | Element::PluginParameter
            | Element::DrawingHyperlink
            | Element::EnhancedGeometry
            | Element::EnhancedEquation
            | Element::EnhancedHandle
            | Element::EventListeners
            | Element::EventListener
            | Element::ScriptEventListener
            | Element::Sound
            | Element::TextParagraph
            | Element::TextSpace
            | Element::TextTab
            | Element::TextLineBreak
            | Element::Animation(_)
            | Element::UnknownAnimation
            | Element::LegacyAnimation(_) => Err(Error::InvalidFormat(
                "dr3d:scene can only contain 3D content".to_string(),
            )),
        }
    }

    pub(super) fn validate_required_three_dimensional_attributes(
        kind: DrawingShapeKind,
        attributes: &[DrawingAttribute],
    ) -> Result<()> {
        let has = |namespace, local_name| {
            attributes.iter().any(|attribute| {
                attribute.namespace() == namespace && attribute.local_name() == local_name
            })
        };
        if kind == DrawingShapeKind::ThreeDimensionalLight
            && !has(DrawingAttributeNamespace::Dr3d, "direction")
        {
            return Err(Error::InvalidFormat(
                "dr3d:light requires dr3d:direction".to_string(),
            ));
        }
        if matches!(
            kind,
            DrawingShapeKind::ThreeDimensionalExtrude | DrawingShapeKind::ThreeDimensionalRotate
        ) {
            for local_name in ["viewBox", "d"] {
                if !has(DrawingAttributeNamespace::Svg, local_name) {
                    return Err(Error::InvalidFormat(format!(
                        "{} requires svg:{local_name}",
                        kind.element_name()
                    )));
                }
            }
        }
        Ok(())
    }

    pub(super) fn require_simple_xlink(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        description: &str,
    ) -> Result<()> {
        ElementAttrs::new(element).require_simple_xlink(reader, description)
    }

    pub(super) fn parse_optional_bool(
        value: Option<String>,
        attribute: &str,
    ) -> Result<Option<bool>> {
        value
            .map(|text| match text.as_str() {
                "true" | "1" => Ok(true),
                "false" | "0" => Ok(false),
                _ => Err(Error::InvalidFormat(format!(
                    "invalid {attribute} value '{text}'"
                ))),
            })
            .transpose()
    }
    pub(super) fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
        matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
    }

    /// Rewinds the running element depth for a subtree consumed in place.
    ///
    /// Nested parsers such as [`Self::parse_enhanced_geometry`] read through
    /// their own element's `Event::End`, so the main event loop never observes
    /// it and would otherwise keep counting that element as open. Spreadsheet
    /// `table:shapes` container tracking compares depths exactly, so the
    /// increment taken on `Event::Start` has to be given back here.
    pub(super) const fn rewind_consumed_subtree(element_depth: usize) -> usize {
        element_depth.saturating_sub(1)
    }

    pub(super) fn animation_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<Attribute>> {
        if element.attributes().count() > 256 {
            return Err(Error::InvalidFormat(
                "ODP animation node exceeds 256 attributes".to_string(),
            ));
        }
        let mut attributes = Vec::with_capacity(element.attributes().count());
        let mut expanded_names = HashSet::new();
        for attribute_result in element.attributes() {
            let attribute = attribute_result
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let qualified_name = attribute.key.as_ref();
            if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
                continue;
            }
            let (resolved_namespace, local_name) =
                reader.resolver().resolve_attribute(attribute.key);
            let namespace_uri = match resolved_namespace {
                ResolveResult::Unbound => None,
                ResolveResult::Bound(XmlNamespace(uri)) => {
                    Some(std::str::from_utf8(uri).map_err(|_err| {
                        Error::InvalidFormat("non-UTF-8 animation namespace URI".to_string())
                    })?)
                },
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::InvalidFormat(format!(
                        "unknown animation attribute namespace prefix '{}'",
                        String::from_utf8_lossy(&prefix)
                    )));
                },
            };
            let local_name_text = std::str::from_utf8(local_name.as_ref())
                .map_err(|_err| {
                    Error::InvalidFormat("non-UTF-8 animation attribute name".to_string())
                })?
                .to_string();
            let namespace = Namespace::from_uri(namespace_uri);
            if !expanded_names.insert((namespace.clone(), local_name_text.clone())) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate animation attribute '{local_name_text}'"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                })?
                .into_owned();
            if value.len() > 1_048_576 {
                return Err(Error::InvalidFormat(
                    "ODP animation attribute exceeds 1 MiB".to_string(),
                ));
            }
            attributes.push(Attribute::from_parsed(namespace, local_name_text, value)?);
        }
        Ok(attributes)
    }

    pub(super) fn exact_geometry_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute_result in element.attributes() {
            let attribute = attribute_result.map_err(|error| {
                Error::InvalidFormat(format!("invalid enhanced-geometry attribute: {error}"))
            })?;
            let (resolved_namespace, local_name) =
                reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Self::is_namespace(&resolved_namespace, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Self::is_namespace(&resolved_namespace, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Self::is_namespace(&resolved_namespace, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid enhanced-geometry attribute value: {error}"
                    ))
                })?
                .into_owned();
            attributes.push(DrawingAttribute::new(
                namespace,
                String::from_utf8(local_name.as_ref().to_vec()).map_err(|_err| {
                    Error::InvalidFormat("non-UTF-8 enhanced-geometry attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    /// Helper to extract attribute values
    pub(super) fn get_attr(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace_uri: &[u8],
        local_name: &[u8],
    ) -> Result<Option<String>> {
        ElementAttrs::new(element).get(reader, namespace_uri, local_name)
    }
}

/// Lazy incremental raw-attribute cache for one element.
///
/// A fresh [`BytesStart::attributes`] iterator restarts at the first attribute
/// on every call, so a function reading `k` attributes of one element pays
/// O(n·k) scan steps. `ElementAttrs` shares a single iterator across lookups:
/// every yielded attribute is appended to `parsed`, and each lookup replays
/// the cached prefix in document order before advancing the shared iterator.
/// Each attribute's key is namespace-resolved exactly once, when the scan
/// first reaches it, and the owned outcome is cached next to the raw
/// attribute; replays compare the cached resolution and local-name bytes
/// instead of calling the resolver again. This is exact because no event may
/// be read between [`Self::new`] and the last lookup — every call site reads
/// one element's attributes consecutively — so the resolver state at append
/// time equals the state at lookup time, and resolution itself cannot fail
/// (unknown prefixes simply never match). Value decoding still happens per
/// match at lookup time, so results are identical to repeated
/// [`Parser::get_attr`] calls: first match in document order wins, duplicate
/// detection stays with the underlying iterator, and a malformed attribute
/// surfaces at exactly the lookup whose scan first reaches it. The formatted
/// error is retained and replayed for later lookups that find no match in the
/// cached prefix, mirroring a fresh re-scan that reaches the same position.
///
/// After the lookups of one element, [`Self::drawing_attributes`] harvests the
/// non-modeled drawing attributes from the cached prefix plus the rest of the
/// shared iterator, replacing the historical second fresh scan of
/// `element.attributes()`.
pub(super) struct ElementAttrs<'a> {
    attributes: Attributes<'a>,
    parsed: Vec<ResolvedAttribute<'a>>,
    malformed: Option<String>,
}

/// One scanned attribute together with its namespace resolution snapshot.
struct ResolvedAttribute<'a> {
    attribute: RawAttribute<'a>,
    namespace: ResolvedAttributeNamespace,
    local_name: LocalName<'a>,
}

/// Owned snapshot of one attribute key's [`ResolveResult`].
///
/// [`ResolveResult`] borrows the reader's namespace stack, so it cannot be
/// stored next to the attribute; this owned copy preserves exactly the
/// information [`Parser::is_namespace`] inspects.
enum ResolvedAttributeNamespace {
    /// Bound namespace URI bytes.
    Bound(Vec<u8>),
    /// No prefix: attributes never take the default namespace.
    Unbound,
    /// Prefix not declared in scope.
    Unknown,
}

impl ResolvedAttributeNamespace {
    /// Matches exactly like [`Parser::is_namespace`] on the live outcome.
    fn matches(&self, namespace_uri: &[u8]) -> bool {
        matches!(self, Self::Bound(uri) if uri.as_slice() == namespace_uri)
    }
}

impl<'a> ElementAttrs<'a> {
    pub(super) fn new(element: &'a BytesStart<'_>) -> Self {
        Self {
            attributes: element.attributes(),
            parsed: Vec::new(),
            malformed: None,
        }
    }

    /// Looks up `namespace_uri:local_name`, continuing the shared scan.
    pub(super) fn get(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace_uri: &[u8],
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for cached in &self.parsed {
            if let Some(value) = Self::lookup(reader, cached, namespace_uri, local_name)? {
                return Ok(Some(value));
            }
        }
        if let Some(message) = &self.malformed {
            return Err(Error::InvalidFormat(message.clone()));
        }
        for attribute_result in self.attributes.by_ref() {
            let attribute = match attribute_result {
                Ok(attribute) => attribute,
                Err(error) => {
                    let message = format!("invalid XML attribute: {error}");
                    self.malformed = Some(message.clone());
                    return Err(Error::InvalidFormat(message));
                },
            };
            let cached = Self::resolve_at_scan(reader, attribute);
            if let Some(value) = Self::lookup(reader, &cached, namespace_uri, local_name)? {
                self.parsed.push(cached);
                return Ok(Some(value));
            }
            self.parsed.push(cached);
        }
        Ok(None)
    }

    /// Looks up a required attribute, failing with `qualified_name` when absent.
    pub(super) fn required(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: &[u8],
        local_name: &[u8],
        qualified_name: &str,
    ) -> Result<String> {
        self.get(reader, namespace, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!("element is missing required {qualified_name}"))
        })
    }

    /// Requires `xlink:type="simple"` on the element.
    pub(super) fn require_simple_xlink(
        &mut self,
        reader: &NsReader<&[u8]>,
        description: &str,
    ) -> Result<()> {
        let link_type = self.required(reader, XLINK_NAMESPACE, b"type", "xlink:type")?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "{description} xlink:type must be 'simple', found '{link_type}'"
            )));
        }
        Ok(())
    }

    /// Resolves one freshly scanned attribute key exactly once with the
    /// current resolver, snapshotting the outcome next to the raw attribute.
    fn resolve_at_scan(
        reader: &NsReader<&[u8]>,
        attribute: RawAttribute<'a>,
    ) -> ResolvedAttribute<'a> {
        let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(XmlNamespace(uri)) => {
                ResolvedAttributeNamespace::Bound(uri.to_vec())
            },
            ResolveResult::Unbound => ResolvedAttributeNamespace::Unbound,
            ResolveResult::Unknown(_) => ResolvedAttributeNamespace::Unknown,
        };
        ResolvedAttribute {
            attribute,
            namespace,
            local_name,
        }
    }

    /// Matches one cached attribute and decodes its value on a hit, exactly
    /// like [`Parser::get_attr`]: the snapshotted resolution stands in for a
    /// fresh resolver call, and only the matching attribute is decoded.
    fn lookup(
        reader: &NsReader<&[u8]>,
        cached: &ResolvedAttribute<'_>,
        namespace_uri: &[u8],
        local_name: &[u8],
    ) -> Result<Option<String>> {
        if cached.namespace.matches(namespace_uri) && cached.local_name.as_ref() == local_name {
            return cached
                .attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                });
        }
        Ok(None)
    }

    /// Harvests the non-modeled drawing attributes in document order,
    /// replacing the historical second fresh scan of `element.attributes()`.
    ///
    /// The cached prefix (already scanned by earlier lookups) is replayed
    /// first, then the shared iterator runs to completion; both halves emit
    /// in document order, so the result is exactly the fresh scan's order.
    /// Exactness against the removed fresh re-scan:
    ///
    /// - Resolution: the cached per-attribute snapshots equal a fresh
    ///   resolver call (no event is read between the first lookup and this
    ///   harvest, so the resolver state is unchanged).
    /// - Malformed attributes: a lookup whose scan reached one already
    ///   returned `"invalid XML attribute: …"` and the caller (`?`) never
    ///   reaches this harvest — `self.malformed` is therefore always `None`
    ///   here. A malformed attribute first reached by this harvest's
    ///   continuation maps to `"invalid ODP shape attribute: …"`, identical
    ///   to the fresh scan reaching the same position (quick-xml attribute
    ///   errors carry byte offsets into the element, not iterator-relative
    ///   positions, so continued and fresh iteration report identically).
    /// - Duplicates: detection stays in the single underlying iterator. A
    ///   duplicate reached during lookups already errored there; one first
    ///   reached here errors with the same message the fresh scan produced.
    /// - Decoding: modeled and foreign-namespace attributes are skipped
    ///   without decoding, and every harvested attribute is decoded at the
    ///   same position as in the fresh scan, so decode-error positions and
    ///   messages are unchanged.
    pub(super) fn drawing_attributes(
        &mut self,
        reader: &NsReader<&[u8]>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for cached in &self.parsed {
            if let Some(attribute) = Self::harvest_drawing_attribute(reader, cached)? {
                attributes.push(attribute);
            }
        }
        for attribute_result in self.attributes.by_ref() {
            let attribute = attribute_result.map_err(|error| {
                Error::InvalidFormat(format!("invalid ODP shape attribute: {error}"))
            })?;
            let cached = Self::resolve_at_scan(reader, attribute);
            if let Some(attribute) = Self::harvest_drawing_attribute(reader, &cached)? {
                attributes.push(attribute);
            }
            self.parsed.push(cached);
        }
        Ok(attributes)
    }

    /// Classifies one scanned attribute exactly like the historical fresh
    /// scan: DRAW/SVG/DR3D/TABLE namespaces are kept, names modeled by
    /// dedicated `ShapeBuilder` fields are skipped, and the value of each
    /// harvested attribute is decoded with the same decoder settings.
    fn harvest_drawing_attribute(
        reader: &NsReader<&[u8]>,
        cached: &ResolvedAttribute<'_>,
    ) -> Result<Option<DrawingAttribute>> {
        let namespace = if cached.namespace.matches(DRAW_NAMESPACE) {
            DrawingAttributeNamespace::Drawing
        } else if cached.namespace.matches(SVG_NAMESPACE) {
            DrawingAttributeNamespace::Svg
        } else if cached.namespace.matches(DR3D_NAMESPACE) {
            DrawingAttributeNamespace::Dr3d
        } else if cached.namespace.matches(TABLE_NAMESPACE) {
            DrawingAttributeNamespace::Table
        } else {
            return Ok(None);
        };
        let local_name = cached.local_name.as_ref();
        let modeled = match namespace {
            DrawingAttributeNamespace::Drawing => matches!(
                local_name,
                b"name" | b"style-name" | b"layer" | b"z-index" | b"transform"
            ),
            DrawingAttributeNamespace::Svg => matches!(
                local_name,
                b"x" | b"y" | b"width" | b"height" | b"x1" | b"y1" | b"x2" | b"y2"
            ),
            DrawingAttributeNamespace::Dr3d | DrawingAttributeNamespace::Table => false,
        };
        if modeled {
            return Ok(None);
        }
        let value = cached
            .attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODP shape attribute value: {error}"))
            })?
            .into_owned();
        Ok(Some(DrawingAttribute::new(
            namespace,
            String::from_utf8(local_name.to_vec()).map_err(|_err| {
                Error::InvalidFormat("non-UTF-8 ODP shape attribute name".to_string())
            })?,
            value,
        )?))
    }
}

#[cfg(test)]
mod tests {
    use super::super::super::{Event, PRESENTATION_NAMESPACE};
    use super::*;

    /// Reads the first event of `xml`, which must be the element under test.
    /// Every namespace prefix the lookups rely on is declared on that element.
    fn first_element<'b>(reader: &mut NsReader<&[u8]>, buf: &'b mut Vec<u8>) -> BytesStart<'b> {
        match reader.read_resolved_event_into(buf) {
            Ok((_, Event::Start(element))) | Ok((_, Event::Empty(element))) => element,
            Ok(_) => panic!("test document must start with an element"),
            Err(error) => panic!("test document must parse: {error}"),
        }
    }

    fn invalid_format_message(result: Result<Option<String>>) -> String {
        match result {
            Err(Error::InvalidFormat(message)) => message,
            Err(error) => panic!("expected invalid-format error, got {error}"),
            Ok(value) => panic!("expected invalid-format error, got {value:?}"),
        }
    }

    fn draw_uri() -> &'static str {
        std::str::from_utf8(DRAW_NAMESPACE).expect("draw namespace is UTF-8")
    }

    #[test]
    fn malformed_attribute_after_match_is_not_reached() {
        let xml = format!("<e xmlns:d=\"{}\" d:name=\"ok\" broken/>", draw_uri());
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        // The match precedes the malformed attribute, so the lookup succeeds.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("ok".to_string())
        );
        assert_eq!(
            Parser::get_attr(&reader, &element, DRAW_NAMESPACE, b"name").unwrap(),
            Some("ok".to_string())
        );
        // A lookup whose scan reaches the malformed attribute fails with the
        // same message a fresh scan produces.
        let cached = invalid_format_message(attributes.get(&reader, DRAW_NAMESPACE, b"missing"));
        let one_shot = invalid_format_message(Parser::get_attr(
            &reader,
            &element,
            DRAW_NAMESPACE,
            b"missing",
        ));
        assert_eq!(cached, one_shot);
    }

    #[test]
    fn malformed_attribute_errors_at_first_reaching_lookup_and_replays() {
        let xml = format!(
            "<e xmlns:d=\"{}\" d:first=\"1\" broken d:name=\"x\"/>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        // The match precedes the malformed attribute: Ok in both paths.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"first").unwrap(),
            Some("1".to_string())
        );
        // The match sits past the malformed attribute: the scan errors first.
        let message = invalid_format_message(attributes.get(&reader, DRAW_NAMESPACE, b"name"));
        assert_eq!(
            message,
            invalid_format_message(Parser::get_attr(&reader, &element, DRAW_NAMESPACE, b"name"))
        );
        // Later unmatched lookups replay the identical message.
        assert_eq!(
            message,
            invalid_format_message(attributes.get(&reader, DRAW_NAMESPACE, b"other"))
        );
        assert_eq!(
            message,
            invalid_format_message(Parser::get_attr(
                &reader,
                &element,
                DRAW_NAMESPACE,
                b"other"
            ))
        );
        // A lookup matching the cached prefix still succeeds afterwards.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"first").unwrap(),
            Some("1".to_string())
        );
    }

    #[test]
    fn duplicate_attribute_is_detected_like_a_fresh_scan() {
        let xml = format!("<e xmlns:d=\"{}\" d:name=\"x\" d:name=\"y\"/>", draw_uri());
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        // The first occurrence matches before the duplicate is reached.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("x".to_string())
        );
        // A lookup scanning past the duplicate errors like a fresh scan.
        let message = invalid_format_message(attributes.get(&reader, DRAW_NAMESPACE, b"missing"));
        assert_eq!(
            message,
            invalid_format_message(Parser::get_attr(
                &reader,
                &element,
                DRAW_NAMESPACE,
                b"missing"
            ))
        );
        // The cached first occurrence still matches after the stored error.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("x".to_string())
        );
    }

    #[test]
    fn entity_and_whitespace_normalization_matches_one_shot_lookup() {
        let xml = format!(
            "<e xmlns:d=\"{}\" d:name=\"a&amp;b&#x41;c\" d:ws=\"x\ty\"/>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        for local_name in [b"name".as_slice(), b"ws".as_slice()] {
            let cached = attributes.get(&reader, DRAW_NAMESPACE, local_name).unwrap();
            let one_shot = Parser::get_attr(&reader, &element, DRAW_NAMESPACE, local_name).unwrap();
            assert_eq!(cached, one_shot);
        }
        // Repeated lookups of cached attributes keep returning identical values.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("a&bAc".to_string())
        );
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"ws").unwrap(),
            Some("x y".to_string())
        );
    }

    #[test]
    fn unknown_prefix_attributes_are_skipped_without_error() {
        let xml = format!("<e xmlns:d=\"{}\" u:skip=\"v\" d:name=\"x\"/>", draw_uri());
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("x".to_string())
        );
        // The unknown-prefix attribute is skipped, never an error; once the
        // iterator is exhausted, absent lookups keep returning `None`.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"missing").unwrap(),
            None
        );
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"other").unwrap(),
            None
        );
        assert_eq!(
            Parser::get_attr(&reader, &element, DRAW_NAMESPACE, b"missing").unwrap(),
            None
        );
    }

    #[test]
    fn style_name_fallback_prefers_draw_namespace() {
        let presentation_uri =
            std::str::from_utf8(PRESENTATION_NAMESPACE).expect("presentation namespace is UTF-8");
        // Both present, presentation first in document order: draw still wins.
        let xml = format!(
            "<e xmlns:d=\"{}\" xmlns:p=\"{presentation_uri}\" p:style-name=\"pres\" d:style-name=\"draw\"/>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        let style_name = attributes
            .get(&reader, DRAW_NAMESPACE, b"style-name")
            .unwrap()
            .or(attributes
                .get(&reader, PRESENTATION_NAMESPACE, b"style-name")
                .unwrap());
        assert_eq!(style_name, Some("draw".to_string()));
        let one_shot = Parser::get_attr(&reader, &element, DRAW_NAMESPACE, b"style-name")
            .unwrap()
            .or(
                Parser::get_attr(&reader, &element, PRESENTATION_NAMESPACE, b"style-name").unwrap(),
            );
        assert_eq!(style_name, one_shot);
        // The presentation attribute remains readable afterwards.
        assert_eq!(
            attributes
                .get(&reader, PRESENTATION_NAMESPACE, b"style-name")
                .unwrap(),
            Some("pres".to_string())
        );

        // Only presentation present: the fallback supplies the value.
        let xml = format!(
            "<e xmlns:d=\"{}\" xmlns:p=\"{presentation_uri}\" p:style-name=\"pres\"/>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        let style_name = attributes
            .get(&reader, DRAW_NAMESPACE, b"style-name")
            .unwrap()
            .or(attributes
                .get(&reader, PRESENTATION_NAMESPACE, b"style-name")
                .unwrap());
        assert_eq!(style_name, Some("pres".to_string()));
    }

    #[test]
    fn cached_resolution_matches_fresh_resolution() {
        let presentation_uri =
            std::str::from_utf8(PRESENTATION_NAMESPACE).expect("presentation namespace is UTF-8");
        let xlink_uri = std::str::from_utf8(XLINK_NAMESPACE).expect("xlink namespace is UTF-8");
        let xml = format!(
            "<e xmlns:d=\"{}\" xmlns:p=\"{presentation_uri}\" xmlns:xlink=\"{xlink_uri}\" \
             plain=\"1\" d:name=\"x\" p:name=\"y\" xlink:href=\"z\" d:extra=\"w\"/>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        // Interleave advancing scans, replays, and misses across namespaces;
        // every result must equal a fresh one-shot scan.
        let targets: [(&[u8], &[u8]); 8] = [
            (DRAW_NAMESPACE, b"name"),
            (PRESENTATION_NAMESPACE, b"name"),
            (XLINK_NAMESPACE, b"href"),
            (DRAW_NAMESPACE, b"extra"),
            (PRESENTATION_NAMESPACE, b"extra"),
            (DRAW_NAMESPACE, b"name"),
            (XLINK_NAMESPACE, b"type"),
            (DRAW_NAMESPACE, b"extra"),
        ];
        for (namespace_uri, local_name) in targets {
            let cached = attributes.get(&reader, namespace_uri, local_name).unwrap();
            let one_shot = Parser::get_attr(&reader, &element, namespace_uri, local_name).unwrap();
            assert_eq!(
                cached,
                one_shot,
                "mismatch for {}",
                String::from_utf8_lossy(local_name)
            );
        }
    }

    #[test]
    fn shadowed_prefix_resolves_with_element_scope_bindings() {
        const SHADOW_URI: &[u8] = b"urn:example:shadow";
        // The child re-binds prefix `d`; its attributes must resolve with the
        // bindings in scope at the child, not the root's binding.
        let xml = format!(
            "<root xmlns:d=\"{}\"><child xmlns:d=\"urn:example:shadow\" d:name=\"inner\"/></root>",
            draw_uri()
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        match reader.read_resolved_event_into(&mut buf) {
            Ok((_, Event::Start(_))) => {},
            other => panic!("expected root start, got {other:?}"),
        }
        buf.clear();
        let child = match reader.read_resolved_event_into(&mut buf) {
            Ok((_, Event::Start(element))) | Ok((_, Event::Empty(element))) => element,
            other => panic!("expected child element, got {other:?}"),
        };
        let mut attributes = ElementAttrs::new(&child);
        // Advancing scan: the shadow binding matches, the root binding does not.
        assert_eq!(
            attributes.get(&reader, SHADOW_URI, b"name").unwrap(),
            Some("inner".to_string())
        );
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            None
        );
        // Replay takes the cached-resolution path: identical outcomes.
        assert_eq!(
            attributes.get(&reader, SHADOW_URI, b"name").unwrap(),
            Some("inner".to_string())
        );
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            None
        );
        // And both equal fresh one-shot scans at the same reader position.
        assert_eq!(
            Parser::get_attr(&reader, &child, SHADOW_URI, b"name").unwrap(),
            Some("inner".to_string())
        );
        assert_eq!(
            Parser::get_attr(&reader, &child, DRAW_NAMESPACE, b"name").unwrap(),
            None
        );
    }

    fn presentation_uri() -> &'static str {
        std::str::from_utf8(PRESENTATION_NAMESPACE).expect("presentation namespace is UTF-8")
    }

    fn svg_uri() -> &'static str {
        std::str::from_utf8(SVG_NAMESPACE).expect("svg namespace is UTF-8")
    }

    fn dr3d_uri() -> &'static str {
        std::str::from_utf8(DR3D_NAMESPACE).expect("dr3d namespace is UTF-8")
    }

    fn table_uri() -> &'static str {
        std::str::from_utf8(TABLE_NAMESPACE).expect("table namespace is UTF-8")
    }

    /// Wraps `attributes` in a shape-like element declaring every prefix the
    /// harvest tests rely on.
    fn shape_element_xml(attributes: &str) -> String {
        format!(
            "<e xmlns:d=\"{}\" xmlns:p=\"{}\" xmlns:svg=\"{}\" xmlns:dr3d=\"{}\" \
             xmlns:table=\"{}\" xmlns:text=\"urn:example:text\" {attributes}/>",
            draw_uri(),
            presentation_uri(),
            svg_uri(),
            dr3d_uri(),
            table_uri(),
        )
    }

    /// Runs the exact lookup sequence `Parser::shape_builder` performs for a
    /// non-line shape, leaving the shared scan mid-element.
    fn shape_builder_lookups(attributes: &mut ElementAttrs<'_>, reader: &NsReader<&[u8]>) {
        attributes
            .get(reader, PRESENTATION_NAMESPACE, b"class")
            .unwrap();
        attributes.get(reader, DRAW_NAMESPACE, b"name").unwrap();
        for local_name in [b"x".as_slice(), b"y", b"width", b"height"] {
            attributes.get(reader, SVG_NAMESPACE, local_name).unwrap();
        }
        attributes
            .get(reader, DRAW_NAMESPACE, b"style-name")
            .unwrap();
        attributes
            .get(reader, PRESENTATION_NAMESPACE, b"style-name")
            .unwrap();
        attributes.get(reader, DRAW_NAMESPACE, b"layer").unwrap();
        attributes.get(reader, DRAW_NAMESPACE, b"z-index").unwrap();
        attributes
            .get(reader, DRAW_NAMESPACE, b"transform")
            .unwrap();
        attributes
            .get(reader, PRESENTATION_NAMESPACE, b"placeholder")
            .unwrap();
        attributes
            .get(reader, PRESENTATION_NAMESPACE, b"user-transformed")
            .unwrap();
    }

    fn harvest_error_message(result: Result<Vec<DrawingAttribute>>) -> String {
        match result {
            Err(Error::InvalidFormat(message)) => message,
            Err(error) => panic!("expected invalid-format error, got {error}"),
            Ok(attributes) => panic!("expected invalid-format error, got {attributes:?}"),
        }
    }

    /// Byte-identical copy of the pre-0215 fresh-scan
    /// `Parser::drawing_attributes`, kept as the parity oracle for
    /// `ElementAttrs::drawing_attributes`.
    fn oracle_drawing_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute_result in element.attributes() {
            let attribute = attribute_result.map_err(|error| {
                Error::InvalidFormat(format!("invalid ODP shape attribute: {error}"))
            })?;
            let (namespace_uri, local_name_ref) =
                reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Parser::is_namespace(&namespace_uri, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Parser::is_namespace(&namespace_uri, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Parser::is_namespace(&namespace_uri, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else if Parser::is_namespace(&namespace_uri, TABLE_NAMESPACE) {
                DrawingAttributeNamespace::Table
            } else {
                continue;
            };
            let local_name = local_name_ref.as_ref();
            let modeled = match namespace {
                DrawingAttributeNamespace::Drawing => matches!(
                    local_name,
                    b"name" | b"style-name" | b"layer" | b"z-index" | b"transform"
                ),
                DrawingAttributeNamespace::Svg => matches!(
                    local_name,
                    b"x" | b"y" | b"width" | b"height" | b"x1" | b"y1" | b"x2" | b"y2"
                ),
                DrawingAttributeNamespace::Dr3d | DrawingAttributeNamespace::Table => false,
            };
            if modeled {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODP shape attribute value: {error}"))
                })?
                .into_owned();
            attributes.push(DrawingAttribute::new(
                namespace,
                String::from_utf8(local_name.to_vec()).map_err(|_err| {
                    Error::InvalidFormat("non-UTF-8 ODP shape attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    #[test]
    fn harvest_matches_fresh_scan_and_preserves_document_order() {
        let xml = shape_element_xml(
            "d:foo=\"early\" p:class=\"title\" d:name=\"n\" svg:x=\"1\" svg:y=\"2\" \
             svg:width=\"3\" svg:height=\"4\" d:style-name=\"s\" d:layer=\"l\" d:z-index=\"7\" \
             d:transform=\"t\" p:placeholder=\"true\" p:user-transformed=\"false\" \
             svg:viewBox=\"0 0 1 1\" dr3d:direction=\"dir\" table:end-cell-address=\"A1\" \
             text:note=\"skip\" plain=\"skip\"",
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        shape_builder_lookups(&mut attributes, &reader);
        let harvested = attributes.drawing_attributes(&reader).unwrap();
        // d:foo comes from the cached prefix, the rest from the continuation;
        // together they must reproduce the fresh scan exactly.
        let expected = vec![
            DrawingAttribute::new(DrawingAttributeNamespace::Drawing, "foo", "early").unwrap(),
            DrawingAttribute::new(DrawingAttributeNamespace::Svg, "viewBox", "0 0 1 1").unwrap(),
            DrawingAttribute::new(DrawingAttributeNamespace::Dr3d, "direction", "dir").unwrap(),
            DrawingAttribute::new(DrawingAttributeNamespace::Table, "end-cell-address", "A1")
                .unwrap(),
        ];
        assert_eq!(harvested, expected);
        assert_eq!(
            harvested,
            oracle_drawing_attributes(&reader, &element).unwrap()
        );
    }

    #[test]
    fn harvest_decodes_entities_and_normalization_like_fresh_scan() {
        let xml = shape_element_xml(
            "d:name=\"n\" d:extra=\"a&amp;b&#x41;c\" svg:viewBox=\"x\ty\" dr3d:d=\"p\"",
        );
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        shape_builder_lookups(&mut attributes, &reader);
        let harvested = attributes.drawing_attributes(&reader).unwrap();
        assert_eq!(
            harvested,
            oracle_drawing_attributes(&reader, &element).unwrap()
        );
        assert_eq!(harvested[0].value(), "a&bAc");
        assert_eq!(harvested[1].value(), "x y");
        assert_eq!(harvested[2].value(), "p");
    }

    #[test]
    fn harvest_after_exhausted_iterator_uses_cache_only() {
        let xml = shape_element_xml("d:foo=\"1\" svg:viewBox=\"v\" table:x=\"t\"");
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        // A miss exhausts the shared iterator; the harvest must then come
        // entirely from the cached prefix and still equal the fresh scan.
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"missing").unwrap(),
            None
        );
        assert_eq!(
            attributes.drawing_attributes(&reader).unwrap(),
            oracle_drawing_attributes(&reader, &element).unwrap()
        );
    }

    #[test]
    fn harvest_malformed_attribute_maps_to_shape_message_like_fresh_scan() {
        // The malformed attribute sits past every lookup target, so only the
        // harvest reaches it: it must surface the drawing-attribute message,
        // identical to the fresh scan.
        let xml = shape_element_xml("d:name=\"n\" d:foo=\"1\" broken");
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("n".to_string())
        );
        let harvested = harvest_error_message(attributes.drawing_attributes(&reader));
        let fresh = harvest_error_message(oracle_drawing_attributes(&reader, &element));
        assert_eq!(harvested, fresh);
        assert!(harvested.starts_with("invalid ODP shape attribute:"));
    }

    #[test]
    fn lookup_reaching_malformed_first_keeps_xml_message() {
        // The malformed attribute precedes the lookup target: the lookup
        // errors first with the generic XML message, and `shape_builder`
        // returns before any harvest — the drawing-attribute message never
        // surfaces for this input.
        let xml = shape_element_xml("broken d:name=\"x\" d:foo=\"1\"");
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        let message = invalid_format_message(attributes.get(&reader, DRAW_NAMESPACE, b"name"));
        assert_eq!(
            message,
            invalid_format_message(Parser::get_attr(&reader, &element, DRAW_NAMESPACE, b"name"))
        );
        assert!(message.starts_with("invalid XML attribute:"));
    }

    #[test]
    fn harvest_duplicate_attribute_matches_fresh_scan() {
        // The first d:name matches the lookup before the duplicate is
        // reached; only the harvest's continuation detects the duplicate, and
        // it must report the identical error the fresh scan produces.
        let xml = shape_element_xml("d:name=\"x\" d:name=\"y\"");
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        assert_eq!(
            attributes.get(&reader, DRAW_NAMESPACE, b"name").unwrap(),
            Some("x".to_string())
        );
        let harvested = harvest_error_message(attributes.drawing_attributes(&reader));
        let fresh = harvest_error_message(oracle_drawing_attributes(&reader, &element));
        assert_eq!(harvested, fresh);
        assert!(harvested.starts_with("invalid ODP shape attribute:"));
    }

    #[test]
    fn harvest_skips_unknown_prefix_and_foreign_namespaces() {
        let xml = shape_element_xml("u:skip=\"v\" text:note=\"n\" xml:id=\"i\" d:foo=\"1\"");
        let mut reader = NsReader::from_str(&xml);
        let mut buf = Vec::new();
        let element = first_element(&mut reader, &mut buf);
        let mut attributes = ElementAttrs::new(&element);
        shape_builder_lookups(&mut attributes, &reader);
        let harvested = attributes.drawing_attributes(&reader).unwrap();
        assert_eq!(
            harvested,
            vec![DrawingAttribute::new(DrawingAttributeNamespace::Drawing, "foo", "1").unwrap()]
        );
        assert_eq!(
            harvested,
            oracle_drawing_attributes(&reader, &element).unwrap()
        );
    }
}
