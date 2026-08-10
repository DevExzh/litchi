//! Namespace-aware protection parsing and byte-preserving XML patching.

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Key, Policy};
use super::validation::{validate_type, validate_xml_size};
use super::{
    CONFIG_NAMESPACE, CONFIG_NAMESPACE_TEXT, CONFIGURATION_SET, Kind, MAX_DEPTH, MAX_ITEMS,
    MAX_VALUE_BYTES, OFFICE_NAMESPACE, OFFICE_NAMESPACE_TEXT, invalid, xml_error,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Field {
    Forms,
    Bookmarks,
    ReadOnly,
    RedlineKey,
}

impl Field {
    const ALL: [Self; 4] = [
        Self::Forms,
        Self::Bookmarks,
        Self::ReadOnly,
        Self::RedlineKey,
    ];

    const fn name(self) -> &'static str {
        match self {
            Self::Forms => "ProtectForm",
            Self::Bookmarks => "ProtectBookmarks",
            Self::ReadOnly => "LoadReadonly",
            Self::RedlineKey => "RedlineProtectionKey",
        }
    }

    const fn value_type(self) -> &'static str {
        match self {
            Self::Forms | Self::Bookmarks | Self::ReadOnly => "boolean",
            Self::RedlineKey => "base64Binary",
        }
    }

    fn from_name(value: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|field| field.name() == value)
    }
}

#[derive(Clone, Debug)]
struct ItemSpan {
    field: Field,
    start: usize,
    end: usize,
    content_start: Option<usize>,
    content_end: Option<usize>,
    qname: Vec<u8>,
}

#[derive(Clone, Debug)]
struct ContainerSpan {
    start: usize,
    end: usize,
    content_end: Option<usize>,
    qname: Vec<u8>,
}

#[derive(Default)]
struct Sites {
    root: Option<ContainerSpan>,
    settings: Option<ContainerSpan>,
    configuration: Option<ContainerSpan>,
    items: Vec<ItemSpan>,
}

struct Frame {
    local: Vec<u8>,
    qname: Vec<u8>,
    start: usize,
    content_start: usize,
    root: bool,
    settings: bool,
    configuration: bool,
    item: Option<Field>,
}

#[derive(Clone)]
struct Patch {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

pub(crate) fn parse(source: &[u8], kind: Kind) -> Result<Policy> {
    validate_xml_size(source)?;
    let xml = std::str::from_utf8(source)
        .map_err(|_error| Error::InvalidFormat("ODT protection XML is not UTF-8".to_string()))?;
    policy_from_xml(xml, kind)
}

pub(crate) fn parse_flat(source: &[u8]) -> Result<Policy> {
    parse(source, Kind::Flat)
}

pub(crate) fn parse_package(source: &[u8]) -> Result<Policy> {
    parse(source, Kind::Package)
}

pub(crate) fn rewrite(
    source: &[u8],
    kind: Kind,
    expected: &Policy,
    replacement: &Policy,
) -> Result<Vec<u8>> {
    replacement.validate()?;
    let actual = parse(source, kind)?;
    if &actual != expected {
        return invalid(format!(
            "protection transaction source changed: expected {expected:?}, found {actual:?}"
        ));
    }
    if expected == replacement {
        return Ok(source.to_vec());
    }

    let xml = std::str::from_utf8(source)
        .map_err(|_error| Error::InvalidFormat("ODT protection XML is not UTF-8".to_string()))?;
    let sites = scan_sites(xml, kind)?;
    let mut patches = Vec::new();
    let mut insertions = Vec::new();

    for field in Field::ALL {
        let old = value(expected, field);
        let new = value(replacement, field);
        if old == new {
            continue;
        }
        if let Some(item) = sites.items.iter().find(|item| item.field == field) {
            match new {
                Some(value) => patches.push(Patch {
                    start: item.content_start.unwrap_or(item.start),
                    end: item.content_end.unwrap_or(item.end),
                    replacement: if item.content_start.is_some() {
                        render_value(field, value)?
                    } else {
                        expand_empty_item(xml, item, &render_value(field, value)?)?
                    },
                }),
                None => patches.push(Patch {
                    start: item.start,
                    end: item.end,
                    replacement: Vec::new(),
                }),
            }
        } else if let Some(value) = new {
            insertions.push(render_item(field, value)?);
        }
    }

    if !insertions.is_empty() {
        let inserted = insertions.concat();
        if let Some(configuration) = &sites.configuration {
            if let Some(content_end) = configuration.content_end {
                patches.push(Patch {
                    start: content_end,
                    end: content_end,
                    replacement: inserted,
                });
            } else {
                patches.push(Patch {
                    start: configuration.start,
                    end: configuration.end,
                    replacement: expand_empty_container(xml, configuration, &inserted)?,
                });
            }
        } else if let Some(settings) = &sites.settings {
            let set = render_configuration(&inserted);
            if let Some(content_end) = settings.content_end {
                patches.push(Patch {
                    start: content_end,
                    end: content_end,
                    replacement: set,
                });
            } else {
                patches.push(Patch {
                    start: settings.start,
                    end: settings.end,
                    replacement: expand_empty_container(xml, settings, &set)?,
                });
            }
        } else {
            let root = sites.root.as_ref().ok_or_else(|| {
                Error::InvalidFormat("protection XML has no document root".to_string())
            })?;
            let mut settings =
                format!("<office:settings xmlns:office=\"{OFFICE_NAMESPACE_TEXT}\">").into_bytes();
            settings.extend_from_slice(&render_configuration(&inserted));
            settings.extend_from_slice(b"</office:settings>");
            let content_end = root
                .content_end
                .ok_or_else(|| Error::InvalidFormat("protection XML root is empty".to_string()))?;
            patches.push(Patch {
                start: content_end,
                end: content_end,
                replacement: settings,
            });
        }
    }

    apply_patches(source, patches)
}

pub(crate) fn empty_package(policy: &Policy) -> Result<Vec<u8>> {
    policy.validate()?;
    let values = Field::ALL
        .into_iter()
        .filter_map(|field| value(policy, field).map(|value| (field, value)))
        .map(|(field, value)| render_item(field, value))
        .collect::<Result<Vec<_>>>()?
        .concat();
    let mut output = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-settings xmlns:office=\"{OFFICE_NAMESPACE_TEXT}\"><office:settings>"
    )
    .into_bytes();
    output.extend_from_slice(&render_configuration(&values));
    output.extend_from_slice(b"</office:settings></office:document-settings>");
    Ok(output)
}

fn policy_from_xml(xml: &str, kind: Kind) -> Result<Policy> {
    let sites = scan_sites(xml, kind)?;
    let mut policy = Policy::default();
    let mut seen = [false; 4];
    for item in &sites.items {
        let field = item.field;
        let index = field as usize;
        if seen[index] {
            return invalid(format!(
                "configuration set contains duplicate '{}' items",
                field.name()
            ));
        }
        seen[index] = true;
        let text = item_text(xml, item)?;
        match field {
            Field::Forms => policy.forms = Some(parse_boolean(&text, field.name())?),
            Field::Bookmarks => {
                policy.bookmarks = Some(parse_boolean(&text, field.name())?);
            },
            Field::ReadOnly => policy.read_only = Some(parse_boolean(&text, field.name())?),
            Field::RedlineKey => {
                let normalized = text
                    .bytes()
                    .filter(|byte| !byte.is_ascii_whitespace())
                    .collect::<Vec<_>>();
                let value = BASE64.decode(normalized).map_err(|_error| {
                    Error::InvalidFormat(format!(
                        "configuration item '{}' has invalid base64",
                        field.name()
                    ))
                })?;
                policy.redline_key = Some(Key::new(value)?);
            },
        }
    }
    policy.validate()?;
    Ok(policy)
}

fn parse_boolean(value: &str, field: &str) -> Result<bool> {
    match value.trim() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("configuration item '{field}' has invalid boolean")),
    }
}

fn item_text(xml: &str, item: &ItemSpan) -> Result<String> {
    let Some(content_start) = item.content_start else {
        return Ok(String::new());
    };
    let content_end = item
        .content_end
        .ok_or_else(|| Error::InvalidFormat("protection item has no closing span".to_string()))?;
    let content = xml
        .get(content_start..content_end)
        .ok_or_else(|| Error::InvalidFormat("invalid protection item content span".to_string()))?;
    if content.as_bytes().contains(&b'<') {
        return invalid(format!(
            "configuration item '{}' contains nested XML",
            item.field.name()
        ));
    }
    if content.len() > MAX_VALUE_BYTES {
        return invalid(format!(
            "configuration item '{}' exceeds the configured value limit",
            item.field.name()
        ));
    }
    quick_xml::escape::unescape(content)
        .map(std::borrow::Cow::into_owned)
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "configuration item '{}' has invalid escaped text: {error}",
                item.field.name()
            ))
        })
}

fn value(policy: &Policy, field: Field) -> Option<Value<'_>> {
    match field {
        Field::Forms => policy.forms.map(Value::Boolean),
        Field::Bookmarks => policy.bookmarks.map(Value::Boolean),
        Field::ReadOnly => policy.read_only.map(Value::Boolean),
        Field::RedlineKey => policy
            .redline_key
            .as_ref()
            .map(|key| Value::Binary(key.as_bytes())),
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Value<'a> {
    Boolean(bool),
    Binary(&'a [u8]),
}

fn render_value(field: Field, value: Value<'_>) -> Result<Vec<u8>> {
    match (field, value) {
        (Field::Forms | Field::Bookmarks | Field::ReadOnly, Value::Boolean(value)) => {
            Ok(if value { b"true".as_slice() } else { b"false" }.to_vec())
        },
        (Field::RedlineKey, Value::Binary(value)) => Ok(BASE64.encode(value).into_bytes()),
        _ => invalid(format!("typed value does not match '{}'", field.name())),
    }
}

fn render_item(field: Field, value: Value<'_>) -> Result<Vec<u8>> {
    let value = String::from_utf8(render_value(field, value)?).map_err(|_error| {
        Error::InvalidFormat("rendered protection value is not UTF-8".to_string())
    })?;
    Ok(format!(
        "<config:config-item xmlns:config=\"{}\" config:name=\"{}\" config:type=\"{}\">{}</config:config-item>",
        CONFIG_NAMESPACE_TEXT,
        field.name(),
        field.value_type(),
        escape_xml(&value)
    )
    .into_bytes())
}

fn render_configuration(items: &[u8]) -> Vec<u8> {
    let mut output = format!(
        "<config:config-item-set xmlns:config=\"{CONFIG_NAMESPACE_TEXT}\" config:name=\"{CONFIGURATION_SET}\">"
    )
    .into_bytes();
    output.extend_from_slice(items);
    output.extend_from_slice(b"</config:config-item-set>");
    output
}

fn expand_empty_item(xml: &str, item: &ItemSpan, value: &[u8]) -> Result<Vec<u8>> {
    expand_empty_element(xml, item.start, item.end, &item.qname, value)
}

fn expand_empty_container(xml: &str, container: &ContainerSpan, content: &[u8]) -> Result<Vec<u8>> {
    expand_empty_element(
        xml,
        container.start,
        container.end,
        &container.qname,
        content,
    )
}

fn expand_empty_element(
    xml: &str,
    start: usize,
    end: usize,
    qname: &[u8],
    content: &[u8],
) -> Result<Vec<u8>> {
    let source = xml.as_bytes();
    let tag = source
        .get(start..end)
        .ok_or_else(|| Error::InvalidFormat("invalid protection element span".to_string()))?;
    let slash = tag
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| Error::InvalidFormat("empty protection element has no slash".to_string()))?;
    let mut output = tag[..slash].to_vec();
    output.push(b'>');
    output.extend_from_slice(content);
    output.extend_from_slice(b"</");
    output.extend_from_slice(qname);
    output.push(b'>');
    Ok(output)
}

fn apply_patches(source: &[u8], mut patches: Vec<Patch>) -> Result<Vec<u8>> {
    patches.sort_by(|left, right| right.start.cmp(&left.start).then(right.end.cmp(&left.end)));
    let mut output = source.to_vec();
    let mut previous_start = output.len();
    for patch in patches {
        if patch.start > patch.end || patch.end > output.len() || patch.end > previous_start {
            return invalid("overlapping protection XML edit spans");
        }
        output.splice(patch.start..patch.end, patch.replacement);
        previous_start = patch.start;
    }
    Ok(output)
}

fn scan_sites(xml: &str, kind: Kind) -> Result<Sites> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut sites = Sites::default();
    let mut item_count = 0usize;
    let mut root_closed = false;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let (office_namespace, config_namespace) = (
            is_namespace(&namespace, OFFICE_NAMESPACE),
            is_namespace(&namespace, CONFIG_NAMESPACE),
        );
        match event {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if stack.len() >= MAX_DEPTH {
                    return invalid("protection XML exceeds the configured depth limit");
                }
                let root = stack.is_empty();
                if root {
                    let expected = match kind {
                        Kind::Flat => b"document".as_slice(),
                        Kind::Package => b"document-settings".as_slice(),
                    };
                    if !office_namespace || element.local_name().as_ref() != expected {
                        return invalid("unexpected ODT protection document element");
                    }
                }
                let settings = office_namespace
                    && element.local_name().as_ref() == b"settings"
                    && stack.len() == 1;
                let parent_settings = stack.last().is_some_and(|frame| frame.settings);
                let name = attr_value(&element, b"name", reader.decoder())?;
                let configuration = config_namespace
                    && element.local_name().as_ref() == b"config-item-set"
                    && parent_settings
                    && name.as_deref() == Some(CONFIGURATION_SET);
                if configuration && sites.configuration.is_some() {
                    return invalid(format!(
                        "settings contain duplicate '{CONFIGURATION_SET}' configuration sets"
                    ));
                }
                let parent_configuration = stack.last().is_some_and(|frame| frame.configuration);
                let item = if config_namespace
                    && element.local_name().as_ref() == b"config-item"
                    && parent_configuration
                {
                    name.as_deref().and_then(Field::from_name)
                } else {
                    None
                };
                if item.is_some() {
                    item_count = item_count.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("protection item count overflow".to_string())
                    })?;
                    if item_count > MAX_ITEMS {
                        return invalid("protection settings exceed the item limit");
                    }
                    if let Some(field) = item {
                        let actual_type = attr_value(&element, b"type", reader.decoder())?
                            .ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "configuration item '{}' has no config:type",
                                    field.name()
                                ))
                            })?;
                        validate_type(&actual_type, field.value_type(), field.name())?;
                        if sites.items.iter().any(|candidate| candidate.field == field) {
                            return invalid(format!(
                                "configuration set contains duplicate '{}' items",
                                field.name()
                            ));
                        }
                    }
                }
                stack.push(Frame {
                    local: element.local_name().as_ref().to_vec(),
                    qname: element.name().as_ref().to_vec(),
                    start,
                    content_start: end,
                    root,
                    settings,
                    configuration,
                    item,
                });
            },
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                if stack.len() >= MAX_DEPTH {
                    return invalid("protection XML exceeds the configured depth limit");
                }
                let root = stack.is_empty();
                if root {
                    return invalid("protection XML document root cannot be empty");
                }
                let settings = office_namespace
                    && element.local_name().as_ref() == b"settings"
                    && stack.len() == 1;
                let parent_settings = stack.last().is_some_and(|frame| frame.settings);
                let name = attr_value(&element, b"name", reader.decoder())?;
                let configuration = config_namespace
                    && element.local_name().as_ref() == b"config-item-set"
                    && parent_settings
                    && name.as_deref() == Some(CONFIGURATION_SET);
                if configuration {
                    if sites.configuration.is_some() {
                        return invalid(format!(
                            "settings contain duplicate '{CONFIGURATION_SET}' configuration sets"
                        ));
                    }
                    sites.configuration = Some(ContainerSpan {
                        start,
                        end,
                        content_end: None,
                        qname: element.name().as_ref().to_vec(),
                    });
                }
                if settings {
                    sites.settings = Some(ContainerSpan {
                        start,
                        end,
                        content_end: None,
                        qname: element.name().as_ref().to_vec(),
                    });
                }
                if config_namespace
                    && element.local_name().as_ref() == b"config-item"
                    && stack.last().is_some_and(|frame| frame.configuration)
                    && let Some(field) = name.as_deref().and_then(Field::from_name)
                {
                    let actual_type =
                        attr_value(&element, b"type", reader.decoder())?.ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "configuration item '{}' has no config:type",
                                field.name()
                            ))
                        })?;
                    validate_type(&actual_type, field.value_type(), field.name())?;
                    if sites.items.iter().any(|candidate| candidate.field == field) {
                        return invalid(format!(
                            "configuration set contains duplicate '{}' items",
                            field.name()
                        ));
                    }
                    sites.items.push(ItemSpan {
                        field,
                        start,
                        end,
                        content_start: None,
                        content_end: None,
                        qname: element.name().as_ref().to_vec(),
                    });
                }
            },
            Event::End(element) => {
                let end = reader.buffer_position() as usize;
                let close_start = event_start(xml, end)?;
                let frame = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("protection XML stack underflow".to_string())
                })?;
                if frame.local != element.local_name().as_ref() {
                    return invalid("protection XML closing element does not match its opener");
                }
                if frame.root {
                    if root_closed {
                        return invalid("protection XML has multiple roots");
                    }
                    root_closed = true;
                    sites.root = Some(ContainerSpan {
                        start: frame.start,
                        end,
                        content_end: Some(close_start),
                        qname: frame.qname.clone(),
                    });
                }
                if frame.settings {
                    sites.settings = Some(ContainerSpan {
                        start: frame.start,
                        end,
                        content_end: Some(close_start),
                        qname: frame.qname.clone(),
                    });
                }
                if frame.configuration {
                    sites.configuration = Some(ContainerSpan {
                        start: frame.start,
                        end,
                        content_end: Some(close_start),
                        qname: frame.qname.clone(),
                    });
                }
                if let Some(field) = frame.item {
                    sites.items.push(ItemSpan {
                        field,
                        start: frame.start,
                        end,
                        content_start: Some(frame.content_start),
                        content_end: Some(close_start),
                        qname: frame.qname,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || !root_closed || sites.root.is_none() {
        return invalid("incomplete protection XML document");
    }
    if matches!(kind, Kind::Package) && sites.settings.is_none() {
        return invalid("package protection XML has no office:settings element");
    }
    Ok(sites)
}

fn attr_value(
    element: &BytesStart<'_>,
    local_name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut result = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid protection XML attribute: {error}"))
        })?;
        let key = attribute.key.as_ref();
        let local = match key.rsplit(|byte| *byte == b':').next() {
            Some(local) => local,
            None => key,
        };
        if local == local_name {
            if result.is_some() {
                return invalid(format!(
                    "duplicate protection XML attribute '{}'",
                    String::from_utf8_lossy(local_name)
                ));
            }
            result = Some(
                attribute
                    .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| {
                        Error::InvalidFormat(format!(
                            "invalid protection XML attribute value: {error}"
                        ))
                    })?
                    .into_owned(),
            );
        }
    }
    Ok(result)
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml.as_bytes()[..end.min(xml.len())]
        .iter()
        .rposition(|byte| *byte == b'<')
        .ok_or_else(|| Error::InvalidFormat("invalid protection XML event span".to_string()))
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}
