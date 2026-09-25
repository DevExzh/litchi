//! Inert `[MS-OWEXML]` custom-function and background-runtime metadata.
//!
//! The schema defines these values in the web-extension namespace, but leaves
//! their placement in an OfficeArt extension entry rather than adding them to
//! the `CT_OsfWebExtension` sequence.  This module owns the small typed
//! vocabulary while retaining the surrounding extension XML byte-for-byte
//! until a caller explicitly changes one of the typed values.

use super::super::codec::{escape_attr, invalid, limit};
use super::super::{
    DRAWINGML_NAMESPACE, Result, STRICT_DRAWINGML_NAMESPACE, WEB_EXTENSION_NAMESPACE,
};
use super::{Limits, MAX_WEB_EXTENSION_ITEMS};
use crate::Error;
use crate::xml::attributes::BytesStartExt as _;
use litchi_core::xml::ReaderOrigin;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

/// The observed Office custom-function extension URI used when authoring a
/// new `a:ext` entry.  The local `[MS-OWEXML]` schema constrains the payload
/// vocabulary but does not assign the OfficeArt `uri` value.
pub(in crate::web) const DEFAULT_CUSTOM_FUNCTIONS_EXTENSION_URI: &str =
    "{D87F86FE-615C-45B5-9D79-34F1136793EB}";

/// `CT_ContainsCustomFunctions` from `[MS-OWEXML]` §2.2.11.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ContainsCustomFunctions {
    value: Option<bool>,
}

impl ContainsCustomFunctions {
    /// Create the flag, retaining whether the optional `val` attribute was
    /// authored.  An omitted value has the schema default `false`.
    #[must_use]
    pub const fn new(value: Option<bool>) -> Self {
        Self { value }
    }

    /// Return the schema-effective Boolean value.
    #[must_use]
    pub fn value(&self) -> bool {
        self.value.unwrap_or(false)
    }

    /// Return the explicitly authored value, if `val` was present.
    #[must_use]
    pub const fn explicit_value(&self) -> Option<bool> {
        self.value
    }

    /// Set the optional `val` attribute, retaining the omitted/default state
    /// when `None` is supplied.
    pub fn set_value(&mut self, value: Option<bool>) -> &mut Self {
        self.value = value;
        self
    }
}

/// `CT_BackgroundAppData` from `[MS-OWEXML]` §2.2.12.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BackgroundAppData {
    state: i32,
    runtime_id: String,
}

impl BackgroundAppData {
    /// Create validated startup-state metadata.  Runtime behavior is never
    /// started or contacted by this model.
    ///
    /// # Errors
    ///
    /// Returns an error when `runtime_id` contains a character forbidden by
    /// XML 1.0.
    pub fn new(state: i32, runtime_id: impl Into<String>) -> Result<Self> {
        let runtime_id = runtime_id.into();
        validate_xml_string("background runtimeId", &runtime_id)?;
        Ok(Self { state, runtime_id })
    }

    #[must_use]
    pub const fn state(&self) -> i32 {
        self.state
    }

    #[must_use]
    pub fn runtime_id(&self) -> &str {
        &self.runtime_id
    }

    /// Replace the startup state.
    pub const fn set_state(&mut self, state: i32) -> &mut Self {
        self.state = state;
        self
    }

    /// Replace the inert runtime identifier.
    ///
    /// # Errors
    ///
    /// Returns an error when `runtime_id` contains a character forbidden by
    /// XML 1.0.  The previous value remains unchanged on error.
    pub fn set_runtime_id(&mut self, runtime_id: impl Into<String>) -> Result<&mut Self> {
        let runtime_id = runtime_id.into();
        validate_xml_string("background runtimeId", &runtime_id)?;
        self.runtime_id = runtime_id;
        Ok(self)
    }
}

/// `CT_CustomFunctionList` from `[MS-OWEXML]` §2.2.13.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CustomFunctionList {
    ids: Vec<String>,
}

impl CustomFunctionList {
    #[must_use]
    pub const fn new() -> Self {
        Self { ids: Vec::new() }
    }

    #[must_use]
    pub fn ids(&self) -> &[String] {
        &self.ids
    }

    /// Append one `customFunctionIds` string in document order.
    ///
    /// # Errors
    ///
    /// Returns an error for invalid XML characters or when the bounded item
    /// limit is reached.
    pub fn push_id(&mut self, id: impl Into<String>) -> Result<&mut Self> {
        self.push_id_with_limit(id, MAX_WEB_EXTENSION_ITEMS)
    }

    fn push_id_with_limit(&mut self, id: impl Into<String>, maximum: usize) -> Result<&mut Self> {
        if self.ids.len() >= maximum {
            return limit(
                "custom function IDs",
                maximum,
                self.ids.len().saturating_add(1),
            );
        }
        let id = id.into();
        validate_xml_string("custom function ID", &id)?;
        self.ids
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "custom function IDs",
                source,
            })?;
        self.ids.push(id);
        Ok(self)
    }

    pub fn clear(&mut self) -> &mut Self {
        self.ids.clear();
        self
    }
}

/// Typed inert payloads found beneath an OfficeArt extension entry.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct CustomFunctions {
    contains_custom_functions: Option<ContainsCustomFunctions>,
    background_app_data: Option<BackgroundAppData>,
    custom_function_list: Option<CustomFunctionList>,
}

impl CustomFunctions {
    #[must_use]
    pub const fn new() -> Self {
        Self {
            contains_custom_functions: None,
            background_app_data: None,
            custom_function_list: None,
        }
    }

    #[must_use]
    pub const fn contains_custom_functions(&self) -> Option<&ContainsCustomFunctions> {
        self.contains_custom_functions.as_ref()
    }

    #[must_use]
    pub const fn background_app_data(&self) -> Option<&BackgroundAppData> {
        self.background_app_data.as_ref()
    }

    #[must_use]
    pub const fn custom_function_list(&self) -> Option<&CustomFunctionList> {
        self.custom_function_list.as_ref()
    }

    pub fn set_contains_custom_functions(
        &mut self,
        value: Option<ContainsCustomFunctions>,
    ) -> &mut Self {
        self.contains_custom_functions = value;
        self
    }

    pub fn set_background_app_data(&mut self, value: Option<BackgroundAppData>) -> &mut Self {
        self.background_app_data = value;
        self
    }

    pub fn set_custom_function_list(&mut self, value: Option<CustomFunctionList>) -> &mut Self {
        self.custom_function_list = value;
        self
    }

    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.contains_custom_functions.is_none()
            && self.background_app_data.is_none()
            && self.custom_function_list.is_none()
    }
}

pub(in crate::web) fn validate_custom_functions(
    value: &CustomFunctions,
    limits: &Limits,
) -> Result<()> {
    let mut string_bytes = 0usize;
    if let Some(background) = &value.background_app_data {
        validate_xml_string("background runtimeId", &background.runtime_id)?;
        string_bytes = string_bytes
            .checked_add(background.runtime_id.len())
            .ok_or(Error::Limit {
                resource: "web extension decoded string bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
    }
    if let Some(list) = &value.custom_function_list {
        if list.ids.len() > limits.items {
            return limit("custom function IDs", limits.items, list.ids.len());
        }
        for id in &list.ids {
            validate_xml_string("custom function ID", id)?;
            string_bytes = string_bytes.checked_add(id.len()).ok_or(Error::Limit {
                resource: "web extension decoded string bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
        }
    }
    if string_bytes > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            string_bytes,
        );
    }
    Ok(())
}

pub(in crate::web) fn parse_custom_functions_with_limits(
    xml: &[u8],
    limits: &Limits,
) -> Result<Option<CustomFunctions>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut extension_depth = None;
    let mut list_depth = None;
    let mut active_id_depth = None;
    let mut active_id = String::new();
    let mut active_known_depth = None;
    let mut ignored_depth = None;
    let mut result = CustomFunctions::new();
    let mut found = false;
    let mut string_bytes = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        match event {
            Event::Start(element) => {
                let parent_depth = depth;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| Error::Invalid("custom-function XML depth overflow".into()))?;
                if ignored_depth.is_some() {
                    // A `customFunctionIds` value is an xsd:string.  Foreign
                    // markup may be retained as opaque source, but a nested
                    // element in the web-extension namespace is malformed
                    // structure, even when it is hidden below a foreign
                    // wrapper.
                    if active_id_depth.is_some_and(|id| parent_depth >= id)
                        && is_web_namespace(&namespace)
                    {
                        return invalid(
                            "customFunctionIds contains a nested web-extension element".into(),
                        );
                    }
                    continue;
                }
                if extension_depth.is_none()
                    && depth > 1
                    && is_drawingml_namespace(&namespace)
                    && element.local_name().as_ref() == b"ext"
                {
                    extension_depth = Some(depth);
                } else if extension_depth == Some(parent_depth) && is_web_namespace(&namespace) {
                    match element.local_name().as_ref() {
                        b"containsCustomFunctions" => {
                            if result.contains_custom_functions.is_some() {
                                return invalid("duplicate containsCustomFunctions payload".into());
                            }
                            result.contains_custom_functions = Some(parse_contains(
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                            )?);
                            active_known_depth = Some(depth);
                            found = true;
                        },
                        b"backgroundAppData" => {
                            if result.background_app_data.is_some() {
                                return invalid("duplicate backgroundAppData payload".into());
                            }
                            result.background_app_data = Some(parse_background(
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                                limits,
                                &mut string_bytes,
                            )?);
                            active_known_depth = Some(depth);
                            found = true;
                        },
                        b"customFunctionList" => {
                            if result.custom_function_list.is_some() {
                                return invalid("duplicate customFunctionList payload".into());
                            }
                            reject_no_attributes(&element, "customFunctionList")?;
                            result.custom_function_list = Some(CustomFunctionList::new());
                            list_depth = Some(depth);
                            found = true;
                        },
                        _ => {},
                    }
                } else if list_depth == Some(parent_depth)
                    && is_web_namespace(&namespace)
                    && element.local_name().as_ref() == b"customFunctionIds"
                {
                    if active_id_depth.is_some() {
                        return invalid("nested customFunctionIds payload".into());
                    }
                    reject_no_attributes(&element, "customFunctionIds")?;
                    let list = result
                        .custom_function_list
                        .as_ref()
                        .ok_or_else(|| Error::Invalid("customFunctionIds has no list".into()))?;
                    if list.ids.len() >= limits.items {
                        return limit(
                            "custom function IDs",
                            limits.items,
                            list.ids.len().saturating_add(1),
                        );
                    }
                    active_id_depth = Some(depth);
                    active_id.clear();
                } else if active_id_depth.is_some_and(|id| parent_depth >= id) {
                    if is_web_namespace(&namespace) {
                        return invalid(
                            "customFunctionIds contains a nested web-extension element".into(),
                        );
                    }
                    // Preserve foreign descendants as opaque source markup.
                    // The source-span rewrite path either retains them or
                    // refuses an operation that would discard them.
                    ignored_depth = Some(depth);
                } else if active_known_depth.is_some_and(|known| parent_depth >= known)
                    || list_depth.is_some_and(|list| parent_depth >= list)
                {
                    // Preserve forward-compatible descendants as opaque source
                    // markup.  The source-span rewrite path either retains them
                    // or refuses an operation that would discard them.
                    ignored_depth = Some(depth);
                }
            },
            Event::Empty(element) => {
                let parent_depth = depth;
                if ignored_depth.is_some() {
                    if active_id_depth.is_some_and(|id| parent_depth >= id)
                        && is_web_namespace(&namespace)
                    {
                        return invalid(
                            "customFunctionIds contains a nested web-extension element".into(),
                        );
                    }
                    continue;
                }
                if extension_depth == Some(parent_depth) && is_web_namespace(&namespace) {
                    match element.local_name().as_ref() {
                        b"containsCustomFunctions" => {
                            if result.contains_custom_functions.is_some() {
                                return invalid("duplicate containsCustomFunctions payload".into());
                            }
                            result.contains_custom_functions = Some(parse_contains(
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                            )?);
                            found = true;
                        },
                        b"backgroundAppData" => {
                            if result.background_app_data.is_some() {
                                return invalid("duplicate backgroundAppData payload".into());
                            }
                            result.background_app_data = Some(parse_background(
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                                limits,
                                &mut string_bytes,
                            )?);
                            found = true;
                        },
                        b"customFunctionList" => {
                            if result.custom_function_list.is_some() {
                                return invalid("duplicate customFunctionList payload".into());
                            }
                            reject_no_attributes(&element, "customFunctionList")?;
                            result.custom_function_list = Some(CustomFunctionList::new());
                            found = true;
                        },
                        _ => {},
                    }
                } else if list_depth == Some(parent_depth)
                    && is_web_namespace(&namespace)
                    && element.local_name().as_ref() == b"customFunctionIds"
                {
                    reject_no_attributes(&element, "customFunctionIds")?;
                    let list = result
                        .custom_function_list
                        .as_mut()
                        .ok_or_else(|| Error::Invalid("customFunctionIds has no list".into()))?;
                    push_custom_id(list, String::new(), limits)?;
                } else if active_id_depth.is_some_and(|id| parent_depth >= id)
                    && is_web_namespace(&namespace)
                {
                    return invalid(
                        "customFunctionIds contains a nested web-extension element".into(),
                    );
                } else if active_known_depth.is_some_and(|known| parent_depth >= known)
                    || list_depth.is_some_and(|list| parent_depth >= list)
                    || active_id_depth.is_some_and(|id| parent_depth >= id)
                {
                    // Opaque descendants are intentionally retained by the
                    // source-span rewrite path; no semantic value is inferred.
                }
            },
            Event::Text(text) => {
                if active_id_depth == Some(depth) {
                    preflight_text_bytes(text.as_ref(), string_bytes, limits)?;
                }
                append_custom_text(
                    &mut active_id,
                    active_id_depth,
                    depth,
                    active_known_depth,
                    list_depth,
                    text.xml10_content()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .as_ref(),
                    true,
                    limits,
                    &mut string_bytes,
                )?;
            },
            Event::CData(text) => {
                if active_id_depth == Some(depth) {
                    preflight_text_bytes(text.as_ref(), string_bytes, limits)?;
                }
                append_custom_text(
                    &mut active_id,
                    active_id_depth,
                    depth,
                    active_known_depth,
                    list_depth,
                    text.xml10_content()
                        .map_err(|error| Error::Xml(error.to_string()))?
                        .as_ref(),
                    false,
                    limits,
                    &mut string_bytes,
                )?;
            },
            Event::GeneralRef(reference) => {
                let reference = reference
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
                let value = resolve_xml_reference(reference.as_ref())?;
                append_custom_text(
                    &mut active_id,
                    active_id_depth,
                    depth,
                    active_known_depth,
                    list_depth,
                    &value,
                    false,
                    limits,
                    &mut string_bytes,
                )?;
            },
            Event::End(element) => {
                if let Some(opaque_depth) = ignored_depth {
                    if opaque_depth == depth {
                        ignored_depth = None;
                    }
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid("custom-function XML nesting underflow".into())
                    })?;
                    continue;
                }
                if active_id_depth == Some(depth) {
                    if element.local_name().as_ref() != b"customFunctionIds" {
                        return invalid("customFunctionIds end tag mismatch".into());
                    }
                    let list = result
                        .custom_function_list
                        .as_mut()
                        .ok_or_else(|| Error::Invalid("customFunctionIds has no list".into()))?;
                    push_custom_id(list, std::mem::take(&mut active_id), limits)?;
                    active_id_depth = None;
                }
                if active_known_depth == Some(depth) {
                    active_known_depth = None;
                }
                if list_depth == Some(depth) {
                    list_depth = None;
                }
                if extension_depth == Some(depth) {
                    extension_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::Invalid("custom-function XML nesting underflow".into())
                })?;
            },
            Event::Eof => break,
            Event::Comment(_) | Event::Decl(_) | Event::PI(_) | Event::DocType(_) => {},
        }
    }
    if depth != 0
        || active_id_depth.is_some()
        || active_known_depth.is_some()
        || list_depth.is_some()
        || ignored_depth.is_some()
    {
        return invalid("unterminated custom-function payload".into());
    }
    if found { Ok(Some(result)) } else { Ok(None) }
}

fn append_custom_text(
    active_id: &mut String,
    active_id_depth: Option<usize>,
    depth: usize,
    active_known_depth: Option<usize>,
    list_depth: Option<usize>,
    text: &str,
    decode_entities: bool,
    limits: &Limits,
    string_bytes: &mut usize,
) -> Result<()> {
    if active_id_depth == Some(depth) {
        let text = if decode_entities {
            quick_xml::escape::unescape(text).map_err(|error| Error::Xml(error.to_string()))?
        } else {
            std::borrow::Cow::Borrowed(text)
        };
        validate_xml_string("custom function ID", text.as_ref())?;
        let next_len = active_id
            .len()
            .checked_add(text.len())
            .ok_or(Error::Limit {
                resource: "custom function ID bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
        if next_len > limits.string_bytes {
            return limit("custom function ID bytes", limits.string_bytes, next_len);
        }
        let total = (*string_bytes)
            .checked_add(text.len())
            .ok_or(Error::Limit {
                resource: "web extension decoded string bytes",
                max: limits.string_bytes,
                actual: usize::MAX,
            })?;
        if total > limits.string_bytes {
            return limit(
                "web extension decoded string bytes",
                limits.string_bytes,
                total,
            );
        }
        active_id
            .try_reserve(text.len())
            .map_err(|source| Error::Allocation {
                resource: "custom function ID",
                source,
            })?;
        active_id.push_str(&text);
        *string_bytes = total;
    } else if active_known_depth == Some(depth)
        || (list_depth == Some(depth) && !text.chars().all(is_xml_space))
    {
        if !text.chars().all(is_xml_space) {
            return invalid("custom-function payload contains unexpected text".into());
        }
    }
    Ok(())
}

fn push_custom_id(list: &mut CustomFunctionList, id: String, limits: &Limits) -> Result<()> {
    if list.ids.len() >= limits.items {
        return limit(
            "custom function IDs",
            limits.items,
            list.ids.len().saturating_add(1),
        );
    }
    list.push_id_with_limit(id, limits.items).map(|_| ())
}

fn reject_no_attributes(element: &BytesStart<'_>, name: &str) -> Result<()> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if !is_namespace_attribute(attribute.key.as_ref()) {
            return invalid(format!("{name} has an unexpected attribute"));
        }
    }
    Ok(())
}

fn parse_contains(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
) -> Result<ContainsCustomFunctions> {
    let mut value = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if is_namespace_attribute(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound)
            || attribute.key.local_name().as_ref() != b"val"
        {
            return invalid("containsCustomFunctions has an unexpected attribute".into());
        }
        if value.is_some() {
            return invalid("duplicate containsCustomFunctions val attribute".into());
        }
        let raw = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        value = Some(parse_bool(trim_xml_space(raw.as_ref()))?);
    }
    Ok(ContainsCustomFunctions::new(value))
}

fn parse_background(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: quick_xml::encoding::Decoder,
    limits: &Limits,
    string_bytes: &mut usize,
) -> Result<BackgroundAppData> {
    let mut state = None;
    let mut runtime_id = None;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if is_namespace_attribute(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return invalid("backgroundAppData has a namespaced attribute".into());
        }
        if attribute.key.local_name().as_ref() == b"runtimeId" {
            preflight_attribute_bytes(attribute.value.as_ref(), *string_bytes, limits)?;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        match attribute.key.local_name().as_ref() {
            b"state" => {
                if state.is_some() {
                    return invalid("duplicate backgroundAppData state attribute".into());
                }
                state = Some(trim_xml_space(value.as_ref()).parse::<i32>().map_err(|_| {
                    Error::Invalid("backgroundAppData state must be an xsd:int".into())
                })?);
            },
            b"runtimeId" => {
                if runtime_id.is_some() {
                    return invalid("duplicate backgroundAppData runtimeId attribute".into());
                }
                let value = value.into_owned();
                *string_bytes = (*string_bytes)
                    .checked_add(value.len())
                    .ok_or(Error::Limit {
                        resource: "web extension decoded string bytes",
                        max: limits.string_bytes,
                        actual: usize::MAX,
                    })?;
                runtime_id = Some(value);
            },
            _ => return invalid("backgroundAppData has an unexpected attribute".into()),
        }
    }
    BackgroundAppData::new(
        state.ok_or_else(|| Error::Invalid("backgroundAppData requires state".into()))?,
        runtime_id.ok_or_else(|| Error::Invalid("backgroundAppData requires runtimeId".into()))?,
    )
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid XML boolean '{value}'")),
    }
}

/// XML Schema whitespace for the `collapse`/`replace` rules used by the
/// lexical Boolean and integer types.  Rust's `str::trim` also accepts many
/// Unicode separators (for example NBSP), which are not XML S characters and
/// must therefore remain invalid input here.
fn trim_xml_space(value: &str) -> &str {
    value.trim_matches(is_xml_space)
}

fn is_xml_space(character: char) -> bool {
    matches!(character, ' ' | '\t' | '\r' | '\n')
}

fn preflight_attribute_bytes(raw: &[u8], string_bytes: usize, limits: &Limits) -> Result<()> {
    let next = string_bytes.checked_add(raw.len()).ok_or(Error::Limit {
        resource: "web extension decoded string bytes",
        max: limits.string_bytes,
        actual: usize::MAX,
    })?;
    if next > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            next,
        );
    }
    Ok(())
}

fn preflight_text_bytes(raw: &[u8], string_bytes: usize, limits: &Limits) -> Result<()> {
    let next = string_bytes.checked_add(raw.len()).ok_or(Error::Limit {
        resource: "web extension decoded string bytes",
        max: limits.string_bytes,
        actual: usize::MAX,
    })?;
    if next > limits.string_bytes {
        return limit(
            "web extension decoded string bytes",
            limits.string_bytes,
            next,
        );
    }
    Ok(())
}

fn resolve_xml_reference(value: &str) -> Result<String> {
    if let Some(value) = quick_xml::escape::resolve_xml_entity(value) {
        return Ok(value.to_owned());
    }
    let code_point = if let Some(value) = value.strip_prefix("#x") {
        u32::from_str_radix(value, 16).ok()
    } else if let Some(value) = value.strip_prefix('#') {
        value.parse::<u32>().ok()
    } else {
        None
    };
    let Some(character) = code_point.and_then(char::from_u32) else {
        return invalid("custom-function XML has an unknown entity reference".into());
    };
    if !is_xml10_character(character) {
        return invalid("custom-function XML entity resolves to an invalid XML character".into());
    }
    Ok(character.to_string())
}

fn is_namespace_attribute(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn is_web_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == WEB_EXTENSION_NAMESPACE.as_bytes()
    )
}

fn is_drawingml_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE.as_bytes()
                || *value == STRICT_DRAWINGML_NAMESPACE.as_bytes()
    )
}

fn validate_xml_string(label: &str, value: &str) -> Result<()> {
    if value
        .chars()
        .any(|character| !is_xml10_character(character))
    {
        return invalid(format!("{label} contains a character forbidden by XML 1.0"));
    }
    Ok(())
}

fn is_xml10_character(character: char) -> bool {
    matches!(
        character,
        '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}'
            | '\u{E000}'..='\u{FFFD}'
            | '\u{10000}'..='\u{10FFFF}'
    )
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XmlRange {
    start: usize,
    end: usize,
    start_tag_end: usize,
    end_tag_start: Option<usize>,
    empty: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PayloadKind {
    Contains,
    Background,
    List,
}

impl PayloadKind {
    fn from_name(name: &[u8]) -> Option<Self> {
        match name {
            b"containsCustomFunctions" => Some(Self::Contains),
            b"backgroundAppData" => Some(Self::Background),
            b"customFunctionList" => Some(Self::List),
            _ => None,
        }
    }

    const fn index(self) -> usize {
        match self {
            Self::Contains => 0,
            Self::Background => 1,
            Self::List => 2,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Contains => "containsCustomFunctions",
            Self::Background => "backgroundAppData",
            Self::List => "customFunctionList",
        }
    }
}

#[derive(Debug)]
struct IdLayout {
    range: XmlRange,
    content: Vec<XmlRange>,
    opaque: bool,
    opaque_element: bool,
}

#[derive(Debug)]
struct PayloadLayout {
    kind: PayloadKind,
    owner_extension: usize,
    range: XmlRange,
    prefix: String,
    ids: Vec<IdLayout>,
    opaque: bool,
    opaque_element: bool,
}

#[derive(Debug, Default)]
struct XmlLayout {
    root: Option<XmlRange>,
    extension_entries: Vec<XmlRange>,
    payloads: Vec<PayloadLayout>,
}

#[derive(Debug, Clone, Copy)]
enum LayoutFrame {
    Root,
    Extension(usize),
    Payload(usize),
    Id(usize, usize),
    Other,
}

#[derive(Debug)]
enum ScanEvent {
    Start {
        drawingml: bool,
        web: bool,
        local_name: String,
        qname: String,
    },
    Empty {
        drawingml: bool,
        web: bool,
        local_name: String,
        qname: String,
    },
    End,
    Text(bool),
    CData(bool),
    GeneralRef,
    Comment,
    Pi,
    Decl,
    DocType,
    Eof,
}

#[derive(Debug)]
struct Edit {
    start: usize,
    end: usize,
    replacement: String,
}

#[derive(Debug, Clone, Copy)]
struct AttributeRange {
    delete_start: usize,
    value_start: usize,
    value_end: usize,
    delete_end: usize,
}

fn scan_xml_layout(xml: &[u8], limits: &Limits) -> Result<XmlLayout> {
    std::str::from_utf8(xml).map_err(|error| Error::Xml(error.to_string()))?;
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    reader.config_mut().trim_text(false);
    let mut stack = Vec::<LayoutFrame>::new();
    let mut nodes = 0usize;
    let mut layout = XmlLayout::default();
    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("custom-function XML offset overflow".into()))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let event = match event {
            Event::Start(element) => {
                let local_name = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                ScanEvent::Start {
                    drawingml: is_drawingml_namespace(&namespace),
                    web: is_web_namespace(&namespace),
                    local_name,
                    qname,
                }
            },
            Event::Empty(element) => {
                let local_name = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                let qname = std::str::from_utf8(element.name().as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .to_owned();
                ScanEvent::Empty {
                    drawingml: is_drawingml_namespace(&namespace),
                    web: is_web_namespace(&namespace),
                    local_name,
                    qname,
                }
            },
            Event::End(_) => ScanEvent::End,
            Event::Text(text) => {
                ScanEvent::Text(!text.as_ref().iter().all(u8::is_ascii_whitespace))
            },
            Event::CData(text) => {
                ScanEvent::CData(!text.as_ref().iter().all(u8::is_ascii_whitespace))
            },
            Event::GeneralRef(_) => ScanEvent::GeneralRef,
            Event::Comment(_) => ScanEvent::Comment,
            Event::PI(_) => ScanEvent::Pi,
            Event::Decl(_) => ScanEvent::Decl,
            Event::DocType(_) => ScanEvent::DocType,
            Event::Eof => ScanEvent::Eof,
        };
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| Error::Invalid("custom-function XML offset overflow".into()))?;
        match event {
            ScanEvent::Start {
                drawingml,
                web,
                local_name,
                qname,
            } => {
                nodes = nodes.checked_add(1).ok_or(Error::Limit {
                    resource: "custom-function XML nodes",
                    max: limits.nodes,
                    actual: usize::MAX,
                })?;
                if nodes > limits.nodes {
                    return limit("custom-function XML nodes", limits.nodes, nodes);
                }
                if stack.len() >= limits.depth {
                    return limit(
                        "custom-function XML depth",
                        limits.depth,
                        stack.len().saturating_add(1),
                    );
                }
                let frame = if stack.is_empty() {
                    layout.root = Some(XmlRange {
                        start: event_start,
                        end: 0,
                        start_tag_end: event_end,
                        end_tag_start: None,
                        empty: false,
                    });
                    LayoutFrame::Root
                } else if matches!(stack.last(), Some(LayoutFrame::Root))
                    && drawingml
                    && local_name == "ext"
                {
                    if layout.extension_entries.len() >= limits.items {
                        return limit(
                            "custom-function OfficeArt extensions",
                            limits.items,
                            layout.extension_entries.len().saturating_add(1),
                        );
                    }
                    layout.extension_entries.try_reserve(1).map_err(|source| {
                        Error::Allocation {
                            resource: "custom-function OfficeArt extensions",
                            source,
                        }
                    })?;
                    let index = layout.extension_entries.len();
                    layout.extension_entries.push(XmlRange {
                        start: event_start,
                        end: 0,
                        start_tag_end: event_end,
                        end_tag_start: None,
                        empty: false,
                    });
                    LayoutFrame::Extension(index)
                } else if let Some(LayoutFrame::Extension(extension_index)) = stack.last().copied()
                {
                    if web && let Some(kind) = PayloadKind::from_name(local_name.as_bytes()) {
                        if layout.payloads.len() >= limits.items {
                            return limit(
                                "custom-function payloads",
                                limits.items,
                                layout.payloads.len().saturating_add(1),
                            );
                        }
                        let (prefix, _) = qname.split_once(':').unwrap_or(("", qname.as_str()));
                        layout
                            .payloads
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "custom-function payloads",
                                source,
                            })?;
                        let index = layout.payloads.len();
                        layout.payloads.push(PayloadLayout {
                            kind,
                            owner_extension: extension_index,
                            range: XmlRange {
                                start: event_start,
                                end: 0,
                                start_tag_end: event_end,
                                end_tag_start: None,
                                empty: false,
                            },
                            prefix: prefix.to_owned(),
                            ids: Vec::new(),
                            opaque: false,
                            opaque_element: false,
                        });
                        LayoutFrame::Payload(index)
                    } else {
                        LayoutFrame::Other
                    }
                } else if let Some(LayoutFrame::Payload(payload_index)) = stack.last().copied() {
                    if layout.payloads[payload_index].kind == PayloadKind::List
                        && web
                        && local_name == "customFunctionIds"
                    {
                        if layout.payloads[payload_index].ids.len() >= limits.items {
                            return limit(
                                "custom-function IDs",
                                limits.items,
                                layout.payloads[payload_index].ids.len().saturating_add(1),
                            );
                        }
                        layout.payloads[payload_index]
                            .ids
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "custom-function IDs",
                                source,
                            })?;
                        let id_index = layout.payloads[payload_index].ids.len();
                        layout.payloads[payload_index].ids.push(IdLayout {
                            range: XmlRange {
                                start: event_start,
                                end: 0,
                                start_tag_end: event_end,
                                end_tag_start: None,
                                empty: false,
                            },
                            content: Vec::new(),
                            opaque: false,
                            opaque_element: false,
                        });
                        LayoutFrame::Id(payload_index, id_index)
                    } else {
                        mark_opaque(&mut layout, &stack, true);
                        LayoutFrame::Other
                    }
                } else {
                    mark_opaque(&mut layout, &stack, true);
                    LayoutFrame::Other
                };
                stack.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "custom-function XML layout stack",
                    source,
                })?;
                stack.push(frame);
            },
            ScanEvent::Empty {
                drawingml,
                web,
                local_name,
                qname,
            } => {
                nodes = nodes.checked_add(1).ok_or(Error::Limit {
                    resource: "custom-function XML nodes",
                    max: limits.nodes,
                    actual: usize::MAX,
                })?;
                if nodes > limits.nodes {
                    return limit("custom-function XML nodes", limits.nodes, nodes);
                }
                if stack.is_empty() {
                    layout.root = Some(XmlRange {
                        start: event_start,
                        end: event_end,
                        start_tag_end: event_end,
                        end_tag_start: None,
                        empty: true,
                    });
                } else if matches!(stack.last(), Some(LayoutFrame::Root))
                    && drawingml
                    && local_name == "ext"
                {
                    if layout.extension_entries.len() >= limits.items {
                        return limit(
                            "custom-function OfficeArt extensions",
                            limits.items,
                            layout.extension_entries.len().saturating_add(1),
                        );
                    }
                    layout.extension_entries.try_reserve(1).map_err(|source| {
                        Error::Allocation {
                            resource: "custom-function OfficeArt extensions",
                            source,
                        }
                    })?;
                    layout.extension_entries.push(XmlRange {
                        start: event_start,
                        end: event_end,
                        start_tag_end: event_end,
                        end_tag_start: None,
                        empty: true,
                    });
                } else if let Some(LayoutFrame::Extension(extension_index)) = stack.last().copied()
                {
                    if web && let Some(kind) = PayloadKind::from_name(local_name.as_bytes()) {
                        if layout.payloads.len() >= limits.items {
                            return limit(
                                "custom-function payloads",
                                limits.items,
                                layout.payloads.len().saturating_add(1),
                            );
                        }
                        let (prefix, _) = qname.split_once(':').unwrap_or(("", qname.as_str()));
                        layout
                            .payloads
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "custom-function payloads",
                                source,
                            })?;
                        layout.payloads.push(PayloadLayout {
                            kind,
                            owner_extension: extension_index,
                            range: XmlRange {
                                start: event_start,
                                end: event_end,
                                start_tag_end: event_end,
                                end_tag_start: None,
                                empty: true,
                            },
                            prefix: prefix.to_owned(),
                            ids: Vec::new(),
                            opaque: false,
                            opaque_element: false,
                        });
                    }
                } else if let Some(LayoutFrame::Payload(payload_index)) = stack.last().copied()
                    && layout.payloads[payload_index].kind == PayloadKind::List
                    && web
                    && local_name == "customFunctionIds"
                {
                    if layout.payloads[payload_index].ids.len() >= limits.items {
                        return limit(
                            "custom-function IDs",
                            limits.items,
                            layout.payloads[payload_index].ids.len().saturating_add(1),
                        );
                    }
                    layout.payloads[payload_index]
                        .ids
                        .try_reserve(1)
                        .map_err(|source| Error::Allocation {
                            resource: "custom-function IDs",
                            source,
                        })?;
                    layout.payloads[payload_index].ids.push(IdLayout {
                        range: XmlRange {
                            start: event_start,
                            end: event_end,
                            start_tag_end: event_end,
                            end_tag_start: None,
                            empty: true,
                        },
                        content: Vec::new(),
                        opaque: false,
                        opaque_element: false,
                    });
                } else {
                    mark_opaque(&mut layout, &stack, true);
                }
            },
            ScanEvent::End => {
                let Some(frame) = stack.pop() else {
                    return invalid("custom-function XML nesting underflow".into());
                };
                match frame {
                    LayoutFrame::Root => {
                        if let Some(root) = layout.root.as_mut() {
                            root.end = event_end;
                            root.end_tag_start = Some(event_start);
                        }
                    },
                    LayoutFrame::Extension(index) => {
                        let Some(entry) = layout.extension_entries.get_mut(index) else {
                            return invalid(
                                "custom-function extension layout index is invalid".into(),
                            );
                        };
                        entry.end = event_end;
                        entry.end_tag_start = Some(event_start);
                    },
                    LayoutFrame::Payload(index) => {
                        let Some(payload) = layout.payloads.get_mut(index) else {
                            return invalid(
                                "custom-function payload layout index is invalid".into(),
                            );
                        };
                        payload.range.end = event_end;
                        payload.range.end_tag_start = Some(event_start);
                    },
                    LayoutFrame::Id(payload_index, id_index) => {
                        let Some(payload) = layout.payloads.get_mut(payload_index) else {
                            return invalid(
                                "custom-function ID payload layout index is invalid".into(),
                            );
                        };
                        let Some(id) = payload.ids.get_mut(id_index) else {
                            return invalid("custom-function ID layout index is invalid".into());
                        };
                        id.range.end = event_end;
                        id.range.end_tag_start = Some(event_start);
                    },
                    LayoutFrame::Other => {},
                }
            },
            ScanEvent::Text(non_whitespace) | ScanEvent::CData(non_whitespace) => {
                record_layout_text(&mut layout, &stack, event_start, event_end, non_whitespace)?;
            },
            ScanEvent::GeneralRef => {
                record_layout_text(&mut layout, &stack, event_start, event_end, true)?
            },
            ScanEvent::Comment | ScanEvent::Pi => mark_opaque(&mut layout, &stack, false),
            ScanEvent::Decl => {},
            ScanEvent::DocType => return invalid("DTD is forbidden in custom-function XML".into()),
            ScanEvent::Eof => break,
        }
    }
    if !stack.is_empty() || layout.root.is_none() {
        return invalid("invalid custom-function XML layout".into());
    }
    Ok(layout)
}

fn record_layout_text(
    layout: &mut XmlLayout,
    stack: &[LayoutFrame],
    start: usize,
    end: usize,
    non_whitespace: bool,
) -> Result<()> {
    if let Some(LayoutFrame::Id(payload_index, id_index)) = stack.last().copied() {
        let payload = &mut layout.payloads[payload_index];
        let id = &mut payload.ids[id_index];
        id.content
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "custom-function ID text ranges",
                source,
            })?;
        id.content.push(XmlRange {
            start,
            end,
            start_tag_end: start,
            end_tag_start: None,
            empty: false,
        });
    } else if let Some(LayoutFrame::Payload(payload_index)) = stack.last().copied()
        && non_whitespace
    {
        layout.payloads[payload_index].opaque = true;
        layout.payloads[payload_index].opaque_element = true;
    }
    Ok(())
}

fn mark_opaque(layout: &mut XmlLayout, stack: &[LayoutFrame], element: bool) {
    let mut payload_index = None;
    let mut id_index = None;
    for frame in stack.iter().rev() {
        match frame {
            LayoutFrame::Id(payload, id) => {
                payload_index = Some(*payload);
                id_index = Some(*id);
                break;
            },
            LayoutFrame::Payload(payload) => {
                payload_index = Some(*payload);
                break;
            },
            LayoutFrame::Root | LayoutFrame::Extension(_) | LayoutFrame::Other => {},
        }
    }
    if let Some(payload) = payload_index {
        layout.payloads[payload].opaque = true;
        layout.payloads[payload].opaque_element |= element;
        if let Some(id) = id_index {
            layout.payloads[payload].ids[id].opaque = true;
            layout.payloads[payload].ids[id].opaque_element |= element;
        }
    }
}

pub(in crate::web) fn rewrite_custom_functions(
    xml: &[u8],
    old: Option<&CustomFunctions>,
    value: Option<&CustomFunctions>,
    limits: &Limits,
) -> Result<String> {
    if xml.len() > limits.xml_bytes {
        return limit(
            "web extension extLst XML bytes",
            limits.xml_bytes,
            xml.len(),
        );
    }
    if let Some(value) = value {
        validate_custom_functions(value, limits)?;
    }
    if old == value {
        return clone_source(xml);
    }
    let layout = scan_xml_layout(xml, limits)?;
    let root = layout
        .root
        .ok_or_else(|| Error::Invalid("custom-function XML has no root".into()))?;
    let mut edits = Vec::new();
    let mut present = [false; 3];
    for payload in &layout.payloads {
        present[payload.kind.index()] = true;
    }
    for payload in &layout.payloads {
        let desired = value.and_then(|value| value.payload(payload.kind));
        match desired {
            Some(desired) => match payload.kind {
                PayloadKind::Contains => {
                    let contains = desired.contains_custom_functions().ok_or_else(|| {
                        Error::Invalid("custom-function payload kind mismatch".into())
                    })?;
                    rewrite_contains_attributes(xml, payload, contains, &mut edits)?;
                },
                PayloadKind::Background => {
                    let background = desired.background_app_data().ok_or_else(|| {
                        Error::Invalid("custom-function payload kind mismatch".into())
                    })?;
                    rewrite_background_attributes(xml, payload, background, &mut edits)?;
                },
                PayloadKind::List => {
                    let list = desired.custom_function_list().ok_or_else(|| {
                        Error::Invalid("custom-function payload kind mismatch".into())
                    })?;
                    rewrite_list(xml, payload, list, &mut edits, limits)?;
                },
            },
            None => {
                if payload.opaque {
                    return invalid(format!(
                        "cannot remove {} while retaining unknown descendants or comments/PIs",
                        payload.kind.name()
                    ));
                }
                push_edit(
                    &mut edits,
                    Edit {
                        start: payload.range.start,
                        end: payload.range.end,
                        replacement: String::new(),
                    },
                )?;
            },
        }
    }

    let missing = value.map(|value| {
        let mut missing = CustomFunctions::new();
        if !present[PayloadKind::Contains.index()] && value.contains_custom_functions().is_some() {
            missing.set_contains_custom_functions(value.contains_custom_functions().cloned());
        }
        if !present[PayloadKind::Background.index()] && value.background_app_data().is_some() {
            missing.set_background_app_data(value.background_app_data().cloned());
        }
        if !present[PayloadKind::List.index()] && value.custom_function_list().is_some() {
            missing.set_custom_function_list(value.custom_function_list().cloned());
        }
        missing
    });
    if let Some(missing) = missing.filter(|value| !value.is_empty()) {
        let rendered = render_custom_functions(&missing, limits)?;
        let owner_extension = layout
            .payloads
            .first()
            .map(|payload| payload.owner_extension);
        if let Some(extension) =
            owner_extension.and_then(|index| layout.extension_entries.get(index))
        {
            if extension.empty {
                push_edit(
                    &mut edits,
                    Edit {
                        start: extension.start,
                        end: extension.end,
                        replacement: expand_empty_element(
                            &xml[extension.start..extension.end],
                            &rendered,
                        )?,
                    },
                )?;
            } else {
                let position = closing_tag_start(xml, *extension)?;
                push_edit(
                    &mut edits,
                    Edit {
                        start: position,
                        end: position,
                        replacement: rendered,
                    },
                )?;
            }
        } else {
            let insertion = render_extension(missing, limits)?;
            if root.empty {
                push_edit(
                    &mut edits,
                    Edit {
                        start: root.start,
                        end: root.end,
                        replacement: expand_empty_element(&xml[root.start..root.end], &insertion)?,
                    },
                )?;
            } else {
                let position = closing_tag_start(xml, root)?;
                push_edit(
                    &mut edits,
                    Edit {
                        start: position,
                        end: position,
                        replacement: insertion,
                    },
                )?;
            }
        }
    }

    apply_edits(xml, edits, limits.xml_bytes)
}

impl CustomFunctions {
    fn payload(&self, kind: PayloadKind) -> Option<&Self> {
        match kind {
            PayloadKind::Contains if self.contains_custom_functions.is_some() => Some(self),
            PayloadKind::Background if self.background_app_data.is_some() => Some(self),
            PayloadKind::List if self.custom_function_list.is_some() => Some(self),
            _ => None,
        }
    }
}

fn rewrite_contains_attributes(
    xml: &[u8],
    payload: &PayloadLayout,
    value: &ContainsCustomFunctions,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    let range = start_attribute_range(xml, payload.range, "val")?;
    let desired = value
        .explicit_value()
        .map(|value| if value { "true" } else { "false" });
    add_attribute_edit(xml, payload.range, range, "val", desired, edits)
}

fn rewrite_background_attributes(
    xml: &[u8],
    payload: &PayloadLayout,
    value: &BackgroundAppData,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    let state = start_attribute_range(xml, payload.range, "state")?;
    let state_value = value.state.to_string();
    add_attribute_edit(
        xml,
        payload.range,
        state,
        "state",
        Some(&state_value),
        edits,
    )?;
    let runtime = start_attribute_range(xml, payload.range, "runtimeId")?;
    add_attribute_edit(
        xml,
        payload.range,
        runtime,
        "runtimeId",
        Some(&value.runtime_id),
        edits,
    )
}

fn rewrite_list(
    xml: &[u8],
    payload: &PayloadLayout,
    value: &CustomFunctionList,
    edits: &mut Vec<Edit>,
    limits: &Limits,
) -> Result<()> {
    let retained = payload.ids.len().min(value.ids.len());
    for (id, desired) in payload.ids.iter().zip(value.ids.iter()).take(retained) {
        rewrite_id(xml, id, desired, edits)?;
    }
    if payload.ids.len() > value.ids.len() {
        for id in payload.ids.iter().skip(value.ids.len()) {
            if id.opaque {
                return invalid(
                    "cannot remove custom-function ID while retaining unknown descendants or comments/PIs"
                        .into(),
                );
            }
            push_edit(
                edits,
                Edit {
                    start: id.range.start,
                    end: id.range.end,
                    replacement: String::new(),
                },
            )?;
        }
    } else if value.ids.len() > payload.ids.len() {
        let rendered = render_ids(&payload.prefix, &value.ids[payload.ids.len()..], limits)?;
        if payload.range.empty {
            push_edit(
                edits,
                Edit {
                    start: payload.range.start,
                    end: payload.range.end,
                    replacement: expand_empty_element(
                        &xml[payload.range.start..payload.range.end],
                        &rendered,
                    )?,
                },
            )?;
        } else {
            let position = closing_tag_start(xml, payload.range)?;
            push_edit(
                edits,
                Edit {
                    start: position,
                    end: position,
                    replacement: rendered,
                },
            )?;
        }
    }
    Ok(())
}

fn rewrite_id(xml: &[u8], id: &IdLayout, value: &str, edits: &mut Vec<Edit>) -> Result<()> {
    if id.opaque_element {
        return invalid(
            "cannot rewrite custom-function ID with unknown descendant elements".into(),
        );
    }
    if let Some(first) = id.content.first() {
        push_edit(
            edits,
            Edit {
                start: first.start,
                end: first.end,
                replacement: escape_text(value),
            },
        )?;
        for extra in id.content.iter().skip(1) {
            push_edit(
                edits,
                Edit {
                    start: extra.start,
                    end: extra.end,
                    replacement: String::new(),
                },
            )?;
        }
    } else if !value.is_empty() {
        if id.range.empty {
            push_edit(
                edits,
                Edit {
                    start: id.range.start,
                    end: id.range.end,
                    replacement: expand_empty_element(
                        &xml[id.range.start..id.range.end],
                        &escape_text(value),
                    )?,
                },
            )?;
        } else {
            let position = closing_tag_start(xml, id.range)?;
            push_edit(
                edits,
                Edit {
                    start: position,
                    end: position,
                    replacement: escape_text(value),
                },
            )?;
        }
    }
    Ok(())
}

fn start_attribute_range(
    xml: &[u8],
    range: XmlRange,
    name: &str,
) -> Result<Option<AttributeRange>> {
    let raw = &xml[range.start..range.start_tag_end];
    let mut cursor = 1usize;
    while cursor < raw.len()
        && !raw[cursor].is_ascii_whitespace()
        && raw[cursor] != b'>'
        && raw[cursor] != b'/'
    {
        cursor += 1;
    }
    loop {
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= raw.len() || raw[cursor] == b'>' || raw[cursor] == b'/' {
            return Ok(None);
        }
        let whitespace_start = cursor.saturating_sub(1);
        let attr_start = cursor;
        while cursor < raw.len()
            && !raw[cursor].is_ascii_whitespace()
            && raw[cursor] != b'='
            && raw[cursor] != b'>'
            && raw[cursor] != b'/'
        {
            cursor += 1;
        }
        let attr_name = &raw[attr_start..cursor];
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if raw.get(cursor) != Some(&b'=') {
            return invalid("custom-function start tag has an invalid attribute".into());
        }
        cursor += 1;
        while cursor < raw.len() && raw[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *raw
            .get(cursor)
            .ok_or_else(|| Error::Invalid("custom-function attribute has no quote".into()))?;
        if quote != b'\'' && quote != b'"' {
            return invalid("custom-function attribute value is not quoted".into());
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < raw.len() && raw[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        if cursor >= raw.len() {
            return invalid("custom-function attribute has no closing quote".into());
        }
        cursor += 1;
        if attr_name == name.as_bytes() {
            return Ok(Some(AttributeRange {
                delete_start: whitespace_start,
                value_start,
                value_end,
                delete_end: cursor,
            }));
        }
    }
}

fn add_attribute_edit(
    xml: &[u8],
    range: XmlRange,
    found: Option<AttributeRange>,
    name: &str,
    value: Option<&str>,
    edits: &mut Vec<Edit>,
) -> Result<()> {
    match (found, value) {
        (Some(found), Some(value)) => push_edit(
            edits,
            Edit {
                start: range.start + found.value_start,
                end: range.start + found.value_end,
                replacement: escape_text(value),
            },
        )?,
        (Some(found), None) => push_edit(
            edits,
            Edit {
                start: range.start + found.delete_start,
                end: range.start + found.delete_end,
                replacement: String::new(),
            },
        )?,
        (None, Some(value)) => {
            let raw = &xml[range.start..range.start_tag_end];
            let close = raw
                .iter()
                .rposition(|byte| *byte == b'>')
                .ok_or_else(|| Error::Invalid("custom-function tag has no end".into()))?;
            let close = if range.empty && raw.get(close.saturating_sub(1)) == Some(&b'/') {
                close - 1
            } else {
                close
            };
            let mut replacement = String::new();
            replacement.push(' ');
            replacement.push_str(name);
            replacement.push_str("=\"");
            escape_attr(&mut replacement, value);
            replacement.push('"');
            push_edit(
                edits,
                Edit {
                    start: range.start + close,
                    end: range.start + close,
                    replacement,
                },
            )?;
        },
        (None, None) => {},
    }
    Ok(())
}

fn render_custom_functions(value: &CustomFunctions, limits: &Limits) -> Result<String> {
    let capacity = custom_functions_rendered_len(value)?;
    let mut out = bounded_string(capacity, limits.xml_bytes, "custom-function XML")?;
    if let Some(contains) = &value.contains_custom_functions {
        render_contains(&mut out, contains);
    }
    if let Some(background) = &value.background_app_data {
        render_background(&mut out, background);
    }
    if let Some(list) = &value.custom_function_list {
        render_list(&mut out, "we", list, limits)?;
    }
    if out.len() > limits.xml_bytes {
        return limit("custom-function XML bytes", limits.xml_bytes, out.len());
    }
    Ok(out)
}

fn render_extension(value: CustomFunctions, limits: &Limits) -> Result<String> {
    let payload = render_custom_functions(&value, limits)?;
    let capacity = "<a:ext xmlns:a=\""
        .len()
        .checked_add(escaped_xml_bytes(DRAWINGML_NAMESPACE)?)
        .and_then(|length| length.checked_add("\" xmlns:we=\"".len()))
        .and_then(|length| length.checked_add(escaped_xml_bytes(WEB_EXTENSION_NAMESPACE).ok()?))
        .and_then(|length| length.checked_add("\" uri=\"".len()))
        .and_then(|length| {
            length.checked_add(escaped_xml_bytes(DEFAULT_CUSTOM_FUNCTIONS_EXTENSION_URI).ok()?)
        })
        .and_then(|length| length.checked_add("\">".len()))
        .and_then(|length| length.checked_add(payload.len()))
        .and_then(|length| length.checked_add("</a:ext>".len()))
        .ok_or(Error::Limit {
            resource: "custom-function XML bytes",
            max: limits.xml_bytes,
            actual: usize::MAX,
        })?;
    let mut out = bounded_string(capacity, limits.xml_bytes, "custom-function XML")?;
    out.push_str("<a:ext xmlns:a=\"");
    escape_attr(&mut out, DRAWINGML_NAMESPACE);
    out.push_str("\" xmlns:we=\"");
    escape_attr(&mut out, WEB_EXTENSION_NAMESPACE);
    out.push_str("\" uri=\"");
    escape_attr(&mut out, DEFAULT_CUSTOM_FUNCTIONS_EXTENSION_URI);
    out.push_str("\">");
    out.push_str(&payload);
    out.push_str("</a:ext>");
    if out.len() > limits.xml_bytes {
        return limit("custom-function XML bytes", limits.xml_bytes, out.len());
    }
    Ok(out)
}

fn render_contains(out: &mut String, value: &ContainsCustomFunctions) {
    out.push_str("<we:containsCustomFunctions xmlns:we=\"");
    escape_attr(out, WEB_EXTENSION_NAMESPACE);
    if let Some(value) = value.explicit_value() {
        out.push_str("\" val=\"");
        out.push_str(if value { "true" } else { "false" });
    }
    out.push_str("\"/>");
}

fn render_background(out: &mut String, value: &BackgroundAppData) {
    out.push_str("<we:backgroundAppData xmlns:we=\"");
    escape_attr(out, WEB_EXTENSION_NAMESPACE);
    out.push_str("\" state=\"");
    out.push_str(&value.state.to_string());
    out.push_str("\" runtimeId=\"");
    escape_attr(out, &value.runtime_id);
    out.push_str("\"/>");
}

fn render_list(
    out: &mut String,
    prefix: &str,
    value: &CustomFunctionList,
    limits: &Limits,
) -> Result<()> {
    out.push('<');
    out.push_str(prefix);
    out.push_str(":customFunctionList xmlns:");
    out.push_str(prefix);
    out.push_str("=\"");
    escape_attr(out, WEB_EXTENSION_NAMESPACE);
    if value.ids.is_empty() {
        out.push_str("\"/>");
    } else {
        out.push_str("\">");
        out.push_str(&render_ids(prefix, &value.ids, limits)?);
        out.push_str("</");
        out.push_str(prefix);
        out.push_str(":customFunctionList>");
    }
    Ok(())
}

fn render_ids(prefix: &str, ids: &[String], limits: &Limits) -> Result<String> {
    let mut capacity = 0usize;
    for id in ids {
        capacity = capacity
            .checked_add("<we:customFunctionIds>".len())
            .and_then(|length| length.checked_add(escaped_xml_bytes(id).ok()?))
            .and_then(|length| length.checked_add("</we:customFunctionIds>".len()))
            .ok_or(Error::Limit {
                resource: "custom-function XML bytes",
                max: limits.xml_bytes,
                actual: usize::MAX,
            })?;
    }
    let mut out = bounded_string(capacity, limits.xml_bytes, "custom-function XML")?;
    for id in ids {
        out.push('<');
        out.push_str(prefix);
        out.push_str(":customFunctionIds>");
        escape_attr(&mut out, id);
        out.push_str("</");
        out.push_str(prefix);
        out.push_str(":customFunctionIds>");
    }
    if out.len() > limits.xml_bytes {
        return limit("custom-function XML bytes", limits.xml_bytes, out.len());
    }
    Ok(out)
}

fn custom_functions_rendered_len(value: &CustomFunctions) -> Result<usize> {
    let mut length = 0usize;
    if let Some(contains) = &value.contains_custom_functions {
        length = length
            .checked_add("<we:containsCustomFunctions xmlns:we=\"".len())
            .and_then(|value| value.checked_add(escaped_xml_bytes(WEB_EXTENSION_NAMESPACE).ok()?))
            .and_then(|value| {
                value.checked_add(if let Some(explicit) = contains.explicit_value() {
                    "\" val=\"".len() + if explicit { "true" } else { "false" }.len() + 3
                } else {
                    3
                })
            })
            .ok_or(Error::Limit {
                resource: "custom-function XML bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
    }
    if let Some(background) = &value.background_app_data {
        let state_len = background.state.to_string().len();
        length = length
            .checked_add("<we:backgroundAppData xmlns:we=\"".len())
            .and_then(|value| value.checked_add(escaped_xml_bytes(WEB_EXTENSION_NAMESPACE).ok()?))
            .and_then(|value| value.checked_add("\" state=\"".len()))
            .and_then(|value| value.checked_add(state_len))
            .and_then(|value| value.checked_add("\" runtimeId=\"".len()))
            .and_then(|value| value.checked_add(escaped_xml_bytes(&background.runtime_id).ok()?))
            .and_then(|value| value.checked_add(3))
            .ok_or(Error::Limit {
                resource: "custom-function XML bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
    }
    if let Some(list) = &value.custom_function_list {
        length = length
            .checked_add("<we:customFunctionList xmlns:we=\"".len())
            .and_then(|value| value.checked_add(escaped_xml_bytes(WEB_EXTENSION_NAMESPACE).ok()?))
            .and_then(|value| {
                if list.ids.is_empty() {
                    value.checked_add(3)
                } else {
                    let mut ids = 2usize;
                    for id in &list.ids {
                        ids = ids
                            .checked_add("<we:customFunctionIds>".len())?
                            .checked_add(escaped_xml_bytes(id).ok()?)?
                            .checked_add("</we:customFunctionIds>".len())?;
                    }
                    ids.checked_add("</we:customFunctionList>".len())
                        .and_then(|ids| value.checked_add(ids))
                }
            })
            .ok_or(Error::Limit {
                resource: "custom-function XML bytes",
                max: usize::MAX,
                actual: usize::MAX,
            })?;
    }
    Ok(length)
}

fn escaped_xml_bytes(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let escaped = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            '\t' | '\n' | '\r' => 5,
            _ => character.len_utf8(),
        };
        length.checked_add(escaped).ok_or(Error::Limit {
            resource: "custom-function XML bytes",
            max: usize::MAX,
            actual: usize::MAX,
        })
    })
}

fn bounded_string(capacity: usize, maximum: usize, resource: &'static str) -> Result<String> {
    if capacity > maximum {
        return limit(resource, maximum, capacity);
    }
    let mut value = String::new();
    value
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(value)
}

fn clone_source(xml: &[u8]) -> Result<String> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "custom-function XML source",
            source,
        })?;
    bytes.extend_from_slice(xml);
    String::from_utf8(bytes).map_err(|error| Error::Xml(error.to_string()))
}

fn escape_text(value: &str) -> String {
    let mut out = String::new();
    escape_attr(&mut out, value);
    out
}

fn push_edit(edits: &mut Vec<Edit>, edit: Edit) -> Result<()> {
    edits.try_reserve(1).map_err(|source| Error::Allocation {
        resource: "custom-function XML edits",
        source,
    })?;
    edits.push(edit);
    Ok(())
}

fn apply_edits(xml: &[u8], mut edits: Vec<Edit>, maximum: usize) -> Result<String> {
    edits.sort_unstable_by_key(|edit| edit.start);
    let mut output_len = xml.len();
    let mut previous_end = 0usize;
    for edit in &edits {
        if edit.start > edit.end || edit.end > xml.len() || edit.start < previous_end {
            return invalid("overlapping custom-function XML ranges".into());
        }
        output_len = output_len
            .checked_sub(edit.end - edit.start)
            .and_then(|value| value.checked_add(edit.replacement.len()))
            .ok_or(Error::Limit {
                resource: "custom-function XML bytes",
                max: maximum,
                actual: usize::MAX,
            })?;
        previous_end = edit.end;
    }
    if output_len > maximum {
        return limit("custom-function XML bytes", maximum, output_len);
    }
    if edits.is_empty() {
        return clone_source(xml);
    }
    let mut output = Vec::new();
    output
        .try_reserve(output_len)
        .map_err(|source| Error::Allocation {
            resource: "custom-function rewritten XML",
            source,
        })?;
    let mut cursor = 0usize;
    for edit in edits {
        output.extend_from_slice(&xml[cursor..edit.start]);
        output.extend_from_slice(edit.replacement.as_bytes());
        cursor = edit.end;
    }
    output.extend_from_slice(&xml[cursor..]);
    String::from_utf8(output).map_err(|error| Error::Xml(error.to_string()))
}

/// Compare two retained extension sources after removing only the three
/// modeled custom-function payloads.  This is used by the package planner to
/// decide whether a caller changed custom metadata in place and can therefore
/// use a source-span splice.  Unknown descendants and comments/PIs inside a
/// modeled payload make the comparison inconclusive; callers then use the
/// normal bounded writer or report its precise refusal.
pub(in crate::web) fn custom_source_compatible(
    left: &[u8],
    right: &[u8],
    limits: &Limits,
) -> Result<bool> {
    let (left, left_opaque) = custom_source_skeleton(left, limits)?;
    let (right, right_opaque) = custom_source_skeleton(right, limits)?;
    Ok(!left_opaque && !right_opaque && left == right)
}

fn custom_source_skeleton(xml: &[u8], limits: &Limits) -> Result<(Vec<u8>, bool)> {
    if xml.len() > limits.xml_bytes {
        return limit(
            "web extension extLst XML bytes",
            limits.xml_bytes,
            xml.len(),
        );
    }
    let layout = scan_xml_layout(xml, limits)?;
    let opaque = layout.payloads.iter().any(|payload| {
        payload.opaque || payload.ids.iter().any(|id| id.opaque || id.opaque_element)
    });
    let mut edits = Vec::new();
    edits
        .try_reserve(layout.payloads.len())
        .map_err(|source| Error::Allocation {
            resource: "custom-function source skeleton edits",
            source,
        })?;
    for payload in &layout.payloads {
        edits.push(Edit {
            start: payload.range.start,
            end: payload.range.end,
            replacement: String::new(),
        });
    }
    let skeleton = apply_edits(xml, edits, limits.xml_bytes)?.into_bytes();
    Ok((skeleton, opaque))
}

fn closing_tag_start(xml: &[u8], range: XmlRange) -> Result<usize> {
    let start = range.start;
    let end = range.end;
    let relative = xml[start..end]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or_else(|| Error::Invalid("custom-function XML has no closing tag".into()))?;
    Ok(start + relative)
}

fn expand_empty_element(raw: &[u8], payload: &str) -> Result<String> {
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::Invalid("custom-function empty extension has no slash".into()))?;
    let close = raw
        .iter()
        .position(|byte| *byte == b'>')
        .ok_or_else(|| Error::Invalid("custom-function empty extension has no end".into()))?;
    let name_end = raw[1..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'>' || *byte == b'/')
        .map_or(raw.len() - 1, |index| index + 1);
    let name =
        std::str::from_utf8(&raw[1..name_end]).map_err(|error| Error::Xml(error.to_string()))?;
    let mut out = String::new();
    out.push_str(
        std::str::from_utf8(&raw[..slash]).map_err(|error| Error::Xml(error.to_string()))?,
    );
    out.push('>');
    out.push_str(payload);
    out.push_str("</");
    out.push_str(name);
    out.push('>');
    debug_assert_eq!(close + 1, raw.len());
    Ok(out)
}
