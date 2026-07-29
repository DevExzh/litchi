//! Typed writing and lossless mutation for `form:typed-value` controls.

use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

use super::generic_writing::OdfGenericControlMetadata;

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_STRING: usize = 1024 * 1024;
const MAX_REFERENCE: usize = 64 * 1024;
const MAX_AGGREGATE: usize = 16 * 1024 * 1024;
const MAX_FORMS: usize = 4096;
const MAX_CONTROLS: usize = 65_536;

/// The four scalar form controls covered by the ODF numeric-control grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum OdfTypedValueControlKind {
    FormattedText,
    Number,
    Date,
    Time,
}

impl OdfTypedValueControlKind {
    const fn local_name(self) -> &'static str {
        match self {
            Self::FormattedText => "formatted-text",
            Self::Number => "number",
            Self::Date => "date",
            Self::Time => "time",
        }
    }

    fn from_local(local: &[u8]) -> Option<Self> {
        match local {
            b"formatted-text" => Some(Self::FormattedText),
            b"number" => Some(Self::Number),
            b"date" => Some(Self::Date),
            b"time" => Some(Self::Time),
            _ => None,
        }
    }
}

/// An exact XML Schema `double` lexical form.
#[derive(Debug, Clone, PartialEq)]
pub struct OdfFormDouble {
    lexical: String,
    numeric: f64,
}

impl OdfFormDouble {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = value.as_ref();
        if value.is_empty() || value.len() > 256 || !valid_xsd_double(value) {
            return Err(format!("invalid XML Schema double lexical form {value:?}"));
        }
        let numeric = match value {
            "INF" => f64::INFINITY,
            "-INF" => f64::NEG_INFINITY,
            "NaN" => f64::NAN,
            _ => value.parse::<f64>().map_err(|_| {
                format!("double is outside the supported IEEE-754 value space: {value:?}")
            })?,
        };
        Ok(Self {
            lexical: value.into(),
            numeric,
        })
    }
    pub fn as_str(&self) -> &str {
        &self.lexical
    }
    pub fn numeric(&self) -> f64 {
        self.numeric
    }
}

fn valid_xsd_double(value: &str) -> bool {
    if matches!(value, "INF" | "-INF" | "NaN") {
        return true;
    }
    let bytes = value.as_bytes();
    let mut i = usize::from(matches!(bytes.first(), Some(b'+' | b'-')));
    let before = i;
    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    let before_digits = i > before;
    let mut after_digits = false;
    if bytes.get(i) == Some(&b'.') {
        i += 1;
        let after = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        after_digits = i > after;
    }
    if !before_digits && !after_digits {
        return false;
    }
    if matches!(bytes.get(i), Some(b'e' | b'E')) {
        i += 1;
        if matches!(bytes.get(i), Some(b'+' | b'-')) {
            i += 1;
        }
        let exponent = i;
        while bytes.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
        }
        if i == exponent {
            return false;
        }
    }
    i == bytes.len()
}

/// An exact, validated XML Schema `date` lexical form.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfFormDate(String);

impl OdfFormDate {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = value.as_ref();
        if !valid_xsd_date(value) {
            return Err(format!("invalid XML Schema date lexical form {value:?}"));
        }
        Ok(Self(value.into()))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

fn valid_xsd_date(value: &str) -> bool {
    if value.len() > 128 || !value.is_ascii() {
        return false;
    }
    let bytes = value.as_bytes();
    let mut i = usize::from(bytes.first() == Some(&b'-'));
    let year_start = i;
    while bytes.get(i).is_some_and(u8::is_ascii_digit) {
        i += 1;
    }
    if i - year_start < 4 || bytes.get(i) != Some(&b'-') {
        return false;
    }
    let Ok(year) = value[year_start..i].parse::<u64>() else {
        return false;
    };
    if year == 0 {
        return false;
    }
    i += 1;
    if i + 2 > bytes.len() || !bytes[i..i + 2].iter().all(u8::is_ascii_digit) {
        return false;
    }
    let month = (bytes[i] - b'0') * 10 + bytes[i + 1] - b'0';
    i += 2;
    if bytes.get(i) != Some(&b'-') {
        return false;
    }
    i += 1;
    if i + 2 > bytes.len() || !bytes[i..i + 2].iter().all(u8::is_ascii_digit) {
        return false;
    }
    let day = (bytes[i] - b'0') * 10 + bytes[i + 1] - b'0';
    i += 2;
    let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
    let days = match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if leap => 29,
        2 => 28,
        _ => return false,
    };
    if day == 0 || day > days {
        return false;
    }
    valid_timezone(&bytes[i..])
}

fn valid_timezone(bytes: &[u8]) -> bool {
    if bytes.is_empty() || bytes == b"Z" {
        return true;
    }
    if bytes.len() != 6
        || !matches!(bytes[0], b'+' | b'-')
        || bytes[3] != b':'
        || !bytes[1..3].iter().all(u8::is_ascii_digit)
        || !bytes[4..6].iter().all(u8::is_ascii_digit)
    {
        return false;
    }
    let hours = (bytes[1] - b'0') * 10 + bytes[2] - b'0';
    let minutes = (bytes[4] - b'0') * 10 + bytes[5] - b'0';
    hours < 14 && minutes < 60 || hours == 14 && minutes == 0
}

/// A kind-specific typed lower or upper bound.
#[derive(Debug, Clone, PartialEq)]
pub enum OdfTypedValueBound {
    Text(String),
    Number(OdfFormDouble),
    Date(OdfFormDate),
    Time(OdfTypedValueDuration),
}

impl OdfTypedValueBound {
    fn as_str(&self) -> &str {
        match self {
            Self::Text(v) => v,
            Self::Number(v) => v.as_str(),
            Self::Date(v) => v.as_str(),
            Self::Time(v) => v.as_str(),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfTypedValueControl {
    pub kind: OdfTypedValueControlKind,
    pub name: String,
    pub xml_id: String,
    pub metadata: OdfGenericControlMetadata,
    pub input_required: Option<bool>,
    pub current_value: Option<String>,
    pub disabled: Option<bool>,
    pub max_length: Option<OdfTypedValueNonNegativeInteger>,
    pub printable: Option<bool>,
    pub readonly: Option<bool>,
    pub tab_index: Option<OdfTypedValueNonNegativeInteger>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub value: Option<String>,
    pub linked_cell: Option<String>,
    pub repeat: Option<bool>,
    pub delay_for_repeat: Option<OdfTypedValueDuration>,
    pub spin_button: Option<bool>,
    pub validation: Option<bool>,
    pub decimal_accuracy: Option<OdfTypedValueNonNegativeInteger>,
    pub max_value: Option<OdfTypedValueBound>,
    pub min_value: Option<OdfTypedValueBound>,
}

impl OdfTypedValueControl {
    pub fn new(
        kind: OdfTypedValueControlKind,
        name: impl Into<String>,
        xml_id: impl Into<String>,
    ) -> Self {
        Self {
            kind,
            name: name.into(),
            xml_id: xml_id.into(),
            metadata: OdfGenericControlMetadata::default(),
            input_required: None,
            current_value: None,
            disabled: None,
            max_length: None,
            printable: None,
            readonly: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            value: None,
            linked_cell: None,
            repeat: None,
            delay_for_repeat: None,
            spin_button: None,
            validation: None,
            decimal_accuracy: None,
            max_value: None,
            min_value: None,
        }
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        typed_value_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct OdfTypedValueForm {
    pub name: String,
    pub controls: Vec<OdfTypedValueControl>,
    pub apply_filter: Option<bool>,
}

impl OdfTypedValueForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }

    pub fn add_control(&mut self, control: OdfTypedValueControl) -> Result<()> {
        validate_control(&control)?;
        if self
            .controls
            .iter()
            .any(|existing| existing.name == control.name)
        {
            return invalid(format!(
                "duplicate typed-value control name '{}'",
                control.name
            ));
        }
        if self
            .controls
            .iter()
            .any(|existing| existing.xml_id == control.xml_id)
        {
            return invalid(format!(
                "duplicate typed-value control xml:id '{}'",
                control.xml_id
            ));
        }
        if self.controls.len() >= MAX_CONTROLS {
            return invalid("too many typed-value controls");
        }
        self.controls.push(control);
        Ok(())
    }

    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        validate_controls(&self.controls)?;
        let mut out = format!(
            r#"<form:form xmlns:form="{FORM}" xmlns:xforms="{XFORMS}" form:name="{}""#,
            escape(&self.name)
        );
        push_bool(&mut out, "form:apply-filter", self.apply_filter);
        if self.controls.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            for control in &self.controls {
                out.push_str(&control.to_xml_fragment()?);
            }
            out.push_str("</form:form>");
        }
        Ok(out)
    }
}

pub fn typed_value_controls(xml: &str) -> Result<Vec<OdfTypedValueControl>> {
    Ok(scan(xml)?
        .controls
        .into_iter()
        .map(|item| item.control)
        .collect())
}

pub fn insert_typed_value_control_xml(
    xml: &str,
    form_index: usize,
    control: &OdfTypedValueControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    if scan.ids.iter().any(|id| id == &control.xml_id) {
        return invalid(format!("duplicate xml:id '{}'", control.xml_id));
    }
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match form.site.clone() {
        Site::Paired { close_start } => apply(xml, close_start..close_start, &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, start, end, &qname, &fragment),
    }
}

pub fn replace_typed_value_control_xml(
    xml: &str,
    index: usize,
    replacement: &OdfTypedValueControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("typed-value control {index} is out of bounds"))
    })?;
    reject_duplicate(&scan.forms[old.form], replacement, Some(&old.control))?;
    if replacement.xml_id != old.control.xml_id
        && scan.ids.iter().any(|id| id == &replacement.xml_id)
    {
        return invalid(format!("duplicate xml:id '{}'", replacement.xml_id));
    }
    apply(
        xml,
        old.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}

pub fn remove_typed_value_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("typed-value control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn typed_value_xml(value: &OdfTypedValueControl) -> Result<String> {
    validate_control(value)?;
    let mut out = format!(
        r#"<form:{} form:name="{}" xml:id="{}""#,
        value.kind.local_name(),
        escape(&value.name),
        escape(&value.xml_id)
    );
    push_string(&mut out, "form:id", value.metadata.form_id.as_deref());
    push_string(
        &mut out,
        "form:control-implementation",
        value.metadata.control_implementation.as_deref(),
    );
    push_string(
        &mut out,
        "xforms:bind",
        value.metadata.xforms_bind.as_deref(),
    );
    push_bool(&mut out, "form:input-required", value.input_required);
    push_string(
        &mut out,
        "form:current-value",
        value.current_value.as_deref(),
    );
    push_bool(&mut out, "form:disabled", value.disabled);
    push_string(
        &mut out,
        "form:max-length",
        value
            .max_length
            .as_ref()
            .map(OdfTypedValueNonNegativeInteger::as_str),
    );
    push_bool(&mut out, "form:printable", value.printable);
    push_bool(&mut out, "form:readonly", value.readonly);
    push_string(
        &mut out,
        "form:tab-index",
        value
            .tab_index
            .as_ref()
            .map(OdfTypedValueNonNegativeInteger::as_str),
    );
    push_bool(&mut out, "form:tab-stop", value.tab_stop);
    push_string(&mut out, "form:title", value.title.as_deref());
    push_string(&mut out, "form:value", value.value.as_deref());
    push_string(&mut out, "form:linked-cell", value.linked_cell.as_deref());
    push_bool(&mut out, "form:repeat", value.repeat);
    push_string(
        &mut out,
        "form:delay-for-repeat",
        value
            .delay_for_repeat
            .as_ref()
            .map(OdfTypedValueDuration::as_str),
    );
    push_bool(&mut out, "form:spin-button", value.spin_button);
    push_bool(&mut out, "form:validation", value.validation);
    push_string(
        &mut out,
        "form:decimal-accuracy",
        value
            .decimal_accuracy
            .as_ref()
            .map(OdfTypedValueNonNegativeInteger::as_str),
    );
    push_string(
        &mut out,
        "form:max-value",
        value.max_value.as_ref().map(OdfTypedValueBound::as_str),
    );
    push_string(
        &mut out,
        "form:min-value",
        value.min_value.as_ref().map(OdfTypedValueBound::as_str),
    );
    out.push_str("/>");
    Ok(out)
}

fn validate_control(value: &OdfTypedValueControl) -> Result<()> {
    validate_name("typed-value control name", &value.name)?;
    validate_xml_id(&value.xml_id)?;
    if let Some(form_id) = value.metadata.form_id.as_deref() {
        validate_xml_id(form_id)?;
    }
    validate_optional(
        "control implementation",
        value.metadata.control_implementation.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional(
        "XForms bind",
        value.metadata.xforms_bind.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_optional("typed-value title", value.title.as_deref(), MAX_STRING)?;
    validate_optional("typed-value value", value.value.as_deref(), MAX_STRING)?;
    validate_optional(
        "typed-value current value",
        value.current_value.as_deref(),
        MAX_STRING,
    )?;
    validate_optional(
        "typed-value linked cell",
        value.linked_cell.as_deref(),
        MAX_REFERENCE,
    )?;
    validate_bound(value.kind, value.min_value.as_ref())?;
    validate_bound(value.kind, value.max_value.as_ref())?;
    if value.validation.is_some() && value.kind != OdfTypedValueControlKind::FormattedText {
        return invalid("form:validation is only valid on form:formatted-text");
    }
    if value.decimal_accuracy.is_some() && value.kind != OdfTypedValueControlKind::Number {
        return invalid("form:decimal-accuracy is only valid on form:number");
    }
    Ok(())
}

fn validate_bound(
    kind: OdfTypedValueControlKind,
    bound: Option<&OdfTypedValueBound>,
) -> Result<()> {
    let Some(bound) = bound else { return Ok(()) };
    let valid = matches!(
        (kind, bound),
        (
            OdfTypedValueControlKind::FormattedText,
            OdfTypedValueBound::Text(_)
        ) | (
            OdfTypedValueControlKind::Number,
            OdfTypedValueBound::Number(_)
        ) | (OdfTypedValueControlKind::Date, OdfTypedValueBound::Date(_))
            | (OdfTypedValueControlKind::Time, OdfTypedValueBound::Time(_))
    );
    if !valid {
        return invalid("typed bound does not match its form control kind");
    }
    validate_string("typed-value bound", bound.as_str(), MAX_STRING)
}

fn validate_controls(controls: &[OdfTypedValueControl]) -> Result<()> {
    if controls.len() > MAX_CONTROLS {
        return invalid("too many typed-value controls");
    }
    let mut names = Vec::<&str>::new();
    let mut ids = Vec::<&str>::new();
    let mut aggregate = 0usize;
    for control in controls {
        validate_control(control)?;
        if names.contains(&control.name.as_str()) {
            return invalid(format!(
                "duplicate typed-value control name '{}'",
                control.name
            ));
        }
        if ids.contains(&control.xml_id.as_str()) {
            return invalid(format!(
                "duplicate typed-value control xml:id '{}'",
                control.xml_id
            ));
        }
        names.push(&control.name);
        ids.push(&control.xml_id);
        aggregate = aggregate.saturating_add(control_size(control));
        if aggregate > MAX_AGGREGATE {
            return invalid("typed-value control strings exceed 16 MiB");
        }
    }
    Ok(())
}

fn control_size(value: &OdfTypedValueControl) -> usize {
    [&value.name, &value.xml_id]
        .iter()
        .fold(0usize, |sum, value| sum.saturating_add(value.len()))
        .saturating_add(value.metadata.form_id.as_ref().map_or(0, String::len))
        .saturating_add(
            value
                .metadata
                .control_implementation
                .as_ref()
                .map_or(0, String::len),
        )
        .saturating_add(value.metadata.xforms_bind.as_ref().map_or(0, String::len))
        .saturating_add(value.title.as_ref().map_or(0, String::len))
        .saturating_add(value.value.as_ref().map_or(0, String::len))
        .saturating_add(value.current_value.as_ref().map_or(0, String::len))
        .saturating_add(value.linked_cell.as_ref().map_or(0, String::len))
        .saturating_add(
            value
                .delay_for_repeat
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .max_value
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
        .saturating_add(
            value
                .min_value
                .as_ref()
                .map_or(0, |value| value.as_str().len()),
        )
}

#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: OdfTypedValueControl,
}
#[derive(Clone)]
struct FormLocation {
    site: Site,
    controls: Vec<ControlLocation>,
}
#[derive(Clone)]
enum Site {
    Paired {
        close_start: usize,
    },
    Empty {
        start: usize,
        end: usize,
        qname: String,
    },
}
struct Scan {
    forms: Vec<FormLocation>,
    controls: Vec<ControlLocation>,
    ids: Vec<String>,
}
struct Open {
    local: Vec<u8>,
    form: Option<usize>,
    control: Option<usize>,
}
#[derive(Clone)]
struct Attr {
    namespace: Option<String>,
    local: String,
    value: String,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("typed-value form XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut form_stack = Vec::<usize>::new();
    let mut forms = Vec::<FormLocation>::new();
    let mut controls = Vec::<ControlLocation>::new();
    let mut ids = Vec::<String>::new();
    let mut aggregate = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid typed-value form XML: {error}"))
            })?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                if stack.len() >= 128 {
                    return invalid("typed-value form XML nesting exceeds 128 levels");
                }
                track_xml_id(&reader, element, &mut ids)?;
                let local = element.local_name().as_ref().to_vec();
                let mut form = None;
                let mut control = None;
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many forms");
                    }
                    form = Some(forms.len());
                    forms.push(FormLocation {
                        site: Site::Paired { close_start: 0 },
                        controls: Vec::new(),
                    });
                    form_stack.push(form.unwrap());
                } else if namespace.as_deref() == Some(FORM)
                    && OdfTypedValueControlKind::from_local(&local).is_some()
                {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("typed-value controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("typed-value control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(
                        &reader,
                        element,
                        OdfTypedValueControlKind::from_local(&local).unwrap(),
                    )?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("typed-value control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many typed-value controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: owner,
                        control: parsed,
                    });
                } else if is_active(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the typed-value mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in typed-value control");
                }
                stack.push(Open {
                    local,
                    form,
                    control,
                });
            },
            Event::Empty(ref element) => {
                track_xml_id(&reader, element, &mut ids)?;
                let local = element.local_name().as_ref().to_vec();
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, element)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many forms");
                    }
                    forms.push(FormLocation {
                        site: Site::Empty {
                            start: previous,
                            end,
                            qname: qname(element)?,
                        },
                        controls: Vec::new(),
                    });
                } else if namespace.as_deref() == Some(FORM)
                    && OdfTypedValueControlKind::from_local(&local).is_some()
                {
                    if stack.iter().any(|open| open.control.is_some()) {
                        return invalid("typed-value controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("typed-value control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(
                        &reader,
                        element,
                        OdfTypedValueControlKind::from_local(&local).unwrap(),
                    )?;
                    aggregate = aggregate.saturating_add(control_size(&parsed));
                    if aggregate > MAX_AGGREGATE {
                        return invalid("typed-value control strings exceed 16 MiB");
                    }
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many typed-value controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: owner,
                        control: parsed,
                    });
                } else if is_active(namespace.as_deref(), local.as_slice())
                    && !form_stack.is_empty()
                {
                    return invalid(
                        "event and macro content is outside the typed-value mutation API",
                    );
                } else if stack.iter().any(|open| open.control.is_some())
                    && !is_property(namespace.as_deref(), local.as_slice())
                {
                    return invalid("unexpected child element in typed-value control");
                }
            },
            Event::Text(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid typed-value control text: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("typed-value controls cannot contain character data");
                    }
                },
            Event::CData(text)
                if stack.iter().any(|open| open.control.is_some()) => {
                    let decoded = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid typed-value control CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return invalid("typed-value controls cannot contain CDATA");
                    }
                },
            Event::GeneralRef(_) if stack.iter().any(|open| open.control.is_some()) => {
                return invalid("typed-value controls cannot contain entity references");
            },
            Event::End(ref element) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("typed-value form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != element.local_name().as_ref() {
                    return invalid("mismatched typed-value form XML elements");
                }
                if let Some(index) = open.control {
                    controls[index].span.end = end;
                }
                if let Some(index) = open.form {
                    forms[index].site = Site::Paired {
                        close_start: previous,
                    };
                    form_stack.pop();
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in typed-value form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed typed-value form XML elements");
    }
    for control in &controls {
        forms[control.form].controls.push(control.clone());
    }
    for form in &forms {
        validate_controls(
            &form
                .controls
                .iter()
                .map(|item| item.control.clone())
                .collect::<Vec<_>>(),
        )?;
    }
    Ok(Scan {
        forms,
        controls,
        ids,
    })
}

fn parse_control(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: OdfTypedValueControlKind,
) -> Result<OdfTypedValueControl> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, allowed_attributes(kind))?;
    let mut value = OdfTypedValueControl::new(
        kind,
        required(&attrs, FORM, "name")?,
        required(&attrs, XML, "id")?,
    );
    value.metadata = OdfGenericControlMetadata {
        form_id: optional(&attrs, FORM, "id"),
        control_implementation: optional(&attrs, FORM, "control-implementation"),
        xforms_bind: optional(&attrs, XFORMS, "bind"),
    };
    value.input_required = optional_bool(&attrs, FORM, "input-required")?;
    value.current_value = optional(&attrs, FORM, "current-value");
    value.disabled = optional_bool(&attrs, FORM, "disabled")?;
    value.max_length = optional(&attrs, FORM, "max-length")
        .map(OdfTypedValueNonNegativeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.printable = optional_bool(&attrs, FORM, "printable")?;
    value.readonly = optional_bool(&attrs, FORM, "readonly")?;
    value.tab_index = optional(&attrs, FORM, "tab-index")
        .map(OdfTypedValueNonNegativeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
    value.title = optional(&attrs, FORM, "title");
    value.value = optional(&attrs, FORM, "value");
    value.linked_cell = optional(&attrs, FORM, "linked-cell");
    value.repeat = optional_bool(&attrs, FORM, "repeat")?;
    value.delay_for_repeat = optional(&attrs, FORM, "delay-for-repeat")
        .map(OdfTypedValueDuration::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.spin_button = optional_bool(&attrs, FORM, "spin-button")?;
    value.validation = optional_bool(&attrs, FORM, "validation")?;
    value.decimal_accuracy = optional(&attrs, FORM, "decimal-accuracy")
        .map(OdfTypedValueNonNegativeInteger::new)
        .transpose()
        .map_err(Error::InvalidFormat)?;
    value.max_value = optional(&attrs, FORM, "max-value")
        .map(|raw| parse_bound(kind, raw))
        .transpose()?;
    value.min_value = optional(&attrs, FORM, "min-value")
        .map(|raw| parse_bound(kind, raw))
        .transpose()?;
    validate_control(&value)?;
    Ok(value)
}

const FORMATTED_TEXT_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "current-value"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "linked-cell"),
    (FORM, "repeat"),
    (FORM, "delay-for-repeat"),
    (FORM, "spin-button"),
    (FORM, "max-value"),
    (FORM, "min-value"),
    (FORM, "validation"),
];
const NUMBER_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "current-value"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "linked-cell"),
    (FORM, "repeat"),
    (FORM, "delay-for-repeat"),
    (FORM, "spin-button"),
    (FORM, "max-value"),
    (FORM, "min-value"),
    (FORM, "decimal-accuracy"),
];
const DATE_TIME_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (XFORMS, "bind"),
    (FORM, "current-value"),
    (FORM, "disabled"),
    (FORM, "max-length"),
    (FORM, "printable"),
    (FORM, "readonly"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (FORM, "value"),
    (FORM, "linked-cell"),
    (FORM, "repeat"),
    (FORM, "delay-for-repeat"),
    (FORM, "spin-button"),
    (FORM, "max-value"),
    (FORM, "min-value"),
];
fn allowed_attributes(kind: OdfTypedValueControlKind) -> &'static [(&'static str, &'static str)] {
    match kind {
        OdfTypedValueControlKind::FormattedText => FORMATTED_TEXT_ATTRS,
        OdfTypedValueControlKind::Number => NUMBER_ATTRS,
        OdfTypedValueControlKind::Date | OdfTypedValueControlKind::Time => DATE_TIME_ATTRS,
    }
}
fn parse_bound(kind: OdfTypedValueControlKind, raw: String) -> Result<OdfTypedValueBound> {
    match kind {
        OdfTypedValueControlKind::FormattedText => {
            validate_string("formatted-text bound", &raw, MAX_STRING)?;
            Ok(OdfTypedValueBound::Text(raw))
        },
        OdfTypedValueControlKind::Number => OdfFormDouble::new(raw)
            .map(OdfTypedValueBound::Number)
            .map_err(Error::InvalidFormat),
        OdfTypedValueControlKind::Date => OdfFormDate::new(raw)
            .map(OdfTypedValueBound::Date)
            .map_err(Error::InvalidFormat),
        OdfTypedValueControlKind::Time => OdfTypedValueDuration::new(raw)
            .map(OdfTypedValueBound::Time)
            .map_err(Error::InvalidFormat),
    }
}

const FORM_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (FORM, "control-implementation"),
    (FORM, "service-name"),
    (FORM, "allow-deletes"),
    (FORM, "allow-inserts"),
    (FORM, "allow-updates"),
    (FORM, "apply-filter"),
    (FORM, "command"),
    (FORM, "command-type"),
    (FORM, "datasource"),
    (FORM, "detail-fields"),
    (FORM, "enctype"),
    (FORM, "escape-processing"),
    (FORM, "filter"),
    (FORM, "ignore-result"),
    (FORM, "master-fields"),
    (FORM, "method"),
    (FORM, "navigation-mode"),
    (FORM, "order"),
    (FORM, "tab-cycle"),
    (OFFICE, "target-frame"),
    (XLINK, "type"),
    (XLINK, "href"),
    (XLINK, "actuate"),
];

fn is_property(namespace: Option<&str>, local: &[u8]) -> bool {
    namespace == Some(FORM)
        && matches!(
            local,
            b"properties" | b"property" | b"list-property" | b"list-value"
        )
}
fn is_active(namespace: Option<&str>, local: &[u8]) -> bool {
    (namespace == Some(OFFICE) && matches!(local, b"event-listeners" | b"events" | b"scripts"))
        || namespace == Some(SCRIPT)
}

fn validate_form(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    let attrs = attributes(reader, element)?;
    validate_allowed(&attrs, FORM_ATTRS)?;
    validate_name("form name", &required(&attrs, FORM, "name")?)?;
    for attr in &attrs {
        if attr.namespace.as_deref() == Some(FORM)
            && matches!(
                attr.local.as_str(),
                "allow-deletes"
                    | "allow-inserts"
                    | "allow-updates"
                    | "apply-filter"
                    | "escape-processing"
                    | "ignore-result"
            )
        {
            parse_bool(&attr.value, &attr.local)?;
        }
    }
    Ok(())
}

fn attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut out = Vec::new();
    for attr in element.attributes().with_checks(true) {
        let attr = attr.map_err(|error| {
            Error::InvalidFormat(format!("invalid typed-value control attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        if namespace.as_deref() == Some("http://www.w3.org/2000/xmlns/")
            || attr.key.as_ref() == b"xmlns"
        {
            continue;
        }
        let value = attr
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "invalid typed-value control attribute value: {error}"
                ))
            })?
            .into_owned();
        out.push(Attr {
            namespace,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            value,
        });
    }
    Ok(out)
}

fn track_xml_id(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    ids: &mut Vec<String>,
) -> Result<()> {
    if let Some(id) = optional(&attributes(reader, element)?, XML, "id") {
        validate_xml_id(&id)?;
        if ids.iter().any(|existing| existing == &id) {
            return invalid(format!("duplicate xml:id '{id}'"));
        }
        ids.push(id);
    }
    Ok(())
}
fn validate_allowed(attrs: &[Attr], allowed: &[(&str, &str)]) -> Result<()> {
    for attr in attrs {
        if !allowed.iter().any(|(namespace, local)| {
            attr.namespace.as_deref() == Some(*namespace) && attr.local == *local
        }) {
            return invalid(format!(
                "unexpected typed-value control attribute '{}'",
                attr.local
            ));
        }
    }
    Ok(())
}
fn optional(attrs: &[Attr], namespace: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|attr| attr.namespace.as_deref() == Some(namespace) && attr.local == local)
        .map(|attr| attr.value.clone())
}
fn required(attrs: &[Attr], namespace: &str, local: &str) -> Result<String> {
    optional(attrs, namespace, local)
        .ok_or_else(|| Error::InvalidFormat(format!("missing required attribute {local}")))
}
fn optional_bool(attrs: &[Attr], namespace: &str, local: &str) -> Result<Option<bool>> {
    optional(attrs, namespace, local)
        .map(|value| parse_bool(&value, local))
        .transpose()
}
fn parse_bool(value: &str, local: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid boolean '{value}' for {local}")),
    }
}
fn validate_name(label: &str, value: &str) -> Result<()> {
    validate_string(label, value, MAX_STRING)?;
    if value.is_empty() {
        invalid(format!("{label} cannot be empty"))
    } else {
        Ok(())
    }
}
fn validate_xml_id(value: &str) -> Result<()> {
    validate_name("typed-value control xml:id", value)?;
    let mut chars = value.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid typed-value control xml:id '{value}'"));
    }
    Ok(())
}
fn validate_optional(label: &str, value: Option<&str>, limit: usize) -> Result<()> {
    if let Some(value) = value {
        validate_string(label, value, limit)?;
    }
    Ok(())
}
fn validate_string(label: &str, value: &str, limit: usize) -> Result<()> {
    if value.len() > limit {
        return invalid(format!("{label} exceeds {limit} bytes"));
    }
    if value.chars().any(
        |ch| !matches!(ch as u32, 9 | 10 | 13 | 32..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF),
    ) {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn reject_duplicate(
    form: &FormLocation,
    replacement: &OdfTypedValueControl,
    current: Option<&OdfTypedValueControl>,
) -> Result<()> {
    for item in &form.controls {
        if current.is_some_and(|value| value.xml_id == item.control.xml_id) {
            continue;
        }
        if item.control.name == replacement.name {
            return invalid(format!(
                "duplicate typed-value control name '{}'",
                replacement.name
            ));
        }
        if item.control.xml_id == replacement.xml_id {
            return invalid(format!(
                "duplicate typed-value control xml:id '{}'",
                replacement.xml_id
            ));
        }
    }
    Ok(())
}
fn push_string(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push_str(&format!(r#" {name}="{}""#, escape(value)));
    }
}
fn push_bool(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push_str(&format!(
            r#" {name}="{}""#,
            if value { "true" } else { "false" }
        ));
    }
}
fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
fn bind_fragment(mut value: String) -> String {
    if value.contains("form:") && !value.contains("xmlns:form=") {
        value = value.replacen(' ', &format!(" xmlns:form=\"{FORM}\" "), 1);
    }
    if value.contains("xforms:") && !value.contains("xmlns:xforms=") {
        value = value.replacen(' ', &format!(" xmlns:xforms=\"{XFORMS}\" "), 1);
    }
    value
}
fn qname(element: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(element.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid typed-value form element name".to_string()))
}
fn apply(xml: &str, span: Range<usize>, replacement: &str) -> Result<String> {
    let mut out = String::with_capacity(xml.len() - span.len() + replacement.len());
    out.push_str(&xml[..span.start]);
    out.push_str(replacement);
    out.push_str(&xml[span.end..]);
    Ok(out)
}
fn expand_empty(xml: &str, start: usize, end: usize, qname: &str, content: &str) -> Result<String> {
    let raw = &xml[start..end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| Error::InvalidFormat("invalid empty form element".to_string()))?;
    apply(
        xml,
        start..end,
        &format!("{}>{content}</{qname}>", &raw[..slash]),
    )
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

use std::fmt;
use std::str::FromStr;

const MAX_TYPED_VALUE_INTEGER_DIGITS: usize = 4096;
const MAX_TYPED_VALUE_DURATION_LEN: usize = 256;

fn canonical_integer(value: &str) -> std::result::Result<String, String> {
    if value.is_empty() || value.len() > MAX_TYPED_VALUE_INTEGER_DIGITS + 1 {
        return Err("integer lexical form is empty or exceeds the safety limit".into());
    }
    let (negative, digits) = match value.as_bytes()[0] {
        b'+' => (false, &value[1..]),
        b'-' => (true, &value[1..]),
        _ => (false, value),
    };
    if digits.is_empty()
        || digits.len() > MAX_TYPED_VALUE_INTEGER_DIGITS
        || !digits.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(format!("invalid XML Schema integer lexical form {value:?}"));
    }
    let digits = digits.trim_start_matches('0');
    if digits.is_empty() {
        return Ok("0".into());
    }
    Ok(if negative {
        format!("-{digits}")
    } else {
        digits.into()
    })
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfTypedValueNonNegativeInteger(String);

impl OdfTypedValueNonNegativeInteger {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = canonical_integer(value.as_ref())?;
        if value.starts_with('-') {
            return Err("nonNegativeInteger cannot be negative".into());
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
impl fmt::Display for OdfTypedValueNonNegativeInteger {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}
impl FromStr for OdfTypedValueNonNegativeInteger {
    type Err = String;
    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfTypedValueDuration(String);

impl OdfTypedValueDuration {
    pub fn new(value: impl AsRef<str>) -> std::result::Result<Self, String> {
        let value = value.as_ref();
        if !valid_xsd_duration(value) {
            return Err(format!(
                "invalid XML Schema duration lexical form {value:?}"
            ));
        }
        Ok(Self(value.into()))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl fmt::Display for OdfTypedValueDuration {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(&self.0)
    }
}

impl FromStr for OdfTypedValueDuration {
    type Err = String;

    fn from_str(value: &str) -> std::result::Result<Self, Self::Err> {
        Self::new(value)
    }
}

fn valid_xsd_duration(value: &str) -> bool {
    if value.is_empty() || value.len() > MAX_TYPED_VALUE_DURATION_LEN || !value.is_ascii() {
        return false;
    }
    let bytes = value.as_bytes();
    let mut offset = usize::from(bytes[0] == b'-');
    if bytes.get(offset) != Some(&b'P') {
        return false;
    }
    offset += 1;
    let mut in_time = false;
    let mut saw_component = false;
    let mut saw_time_component = false;
    let mut last_order = 0_u8;
    while offset < bytes.len() {
        if bytes[offset] == b'T' {
            if in_time {
                return false;
            }
            in_time = true;
            offset += 1;
            continue;
        }
        let digits_start = offset;
        while offset < bytes.len() && bytes[offset].is_ascii_digit() {
            offset += 1;
        }
        if offset == digits_start {
            return false;
        }
        let mut fractional = false;
        if bytes.get(offset) == Some(&b'.') {
            fractional = true;
            offset += 1;
            let fraction_start = offset;
            while offset < bytes.len() && bytes[offset].is_ascii_digit() {
                offset += 1;
            }
            if offset == fraction_start {
                return false;
            }
        }
        let Some(&designator) = bytes.get(offset) else {
            return false;
        };
        offset += 1;
        let order = match (in_time, designator) {
            (false, b'Y') => 1,
            (false, b'M') => 2,
            (false, b'D') => 3,
            (true, b'H') => 4,
            (true, b'M') => 5,
            (true, b'S') => 6,
            _ => return false,
        };
        if order <= last_order || (fractional && designator != b'S') {
            return false;
        }
        last_order = order;
        saw_component = true;
        saw_time_component |= in_time;
    }
    saw_component && (!in_time || saw_time_component)
}
#[cfg(test)]
mod tests {
    use super::*;

    fn document(controls: &str) -> String {
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:form="{FORM}" xmlns:xforms="{XFORMS}"><office:body><office:text><form:form form:name="Values">{controls}</form:form><office:tail/></office:text></office:body></office:document-content>"#
        )
    }

    #[test]
    fn rng_and_producer_shaped_controls_parse_all_typed_attributes() {
        let xml = document(concat!(
            r#"<form:formatted-text form:name="Formatted" xml:id="fmt" form:value="12.00" form:current-value="11.00" form:max-value="99.99" form:min-value="0.00" form:validation="true" form:max-length="12" form:spin-button="true" form:repeat="true" form:delay-for-repeat="PT0.5S"/>"#,
            r#"<form:number form:name="Number" xml:id="num" form:value="50" form:current-value="49.5" form:min-value="-INF" form:max-value="1.25E3" form:decimal-accuracy="+0002" form:linked-cell="Sheet1.A1"/>"#,
            r#"<form:date form:name="Date" xml:id="date" form:value="2026-07-19" form:current-value="2026-07-18" form:min-value="1900-01-01Z" form:max-value="9999-12-31+14:00" form:spin-button="true"/>"#,
            r#"<form:time form:name="Time" xml:id="time" form:value="PT12H" form:current-value="PT11H" form:min-value="PT0S" form:max-value="PT23H59M59.999S" form:tab-index="0004" xforms:bind="time-bind"/>"#,
        ));
        let controls = typed_value_controls(&xml).unwrap();
        assert_eq!(
            controls.iter().map(|c| c.kind).collect::<Vec<_>>(),
            [
                OdfTypedValueControlKind::FormattedText,
                OdfTypedValueControlKind::Number,
                OdfTypedValueControlKind::Date,
                OdfTypedValueControlKind::Time,
            ]
        );
        assert_eq!(controls[0].max_length.as_ref().unwrap().as_str(), "12");
        assert_eq!(controls[1].decimal_accuracy.as_ref().unwrap().as_str(), "2");
        assert!(
            matches!(controls[1].min_value, Some(OdfTypedValueBound::Number(ref v)) if v.numeric().is_infinite())
        );
        assert!(
            matches!(controls[2].max_value, Some(OdfTypedValueBound::Date(ref v)) if v.as_str() == "9999-12-31+14:00")
        );
        assert!(
            matches!(controls[3].max_value, Some(OdfTypedValueBound::Time(ref v)) if v.as_str() == "PT23H59M59.999S")
        );
        assert_eq!(
            controls[3].metadata.xforms_bind.as_deref(),
            Some("time-bind")
        );
        assert!(
            controls[0]
                .to_xml_fragment()
                .unwrap()
                .starts_with("<form:formatted-text ")
        );
        assert!(
            controls[1]
                .to_xml_fragment()
                .unwrap()
                .starts_with("<form:number ")
        );
    }

    #[test]
    fn lexical_domains_and_kind_specific_attributes_are_strict() {
        for valid in ["0", "+1.25", "-.5E-2", "INF", "-INF", "NaN"] {
            assert!(OdfFormDouble::new(valid).is_ok(), "{valid}");
        }
        for invalid in ["", " inf", "Infinity", ".", "1e", "1_0"] {
            assert!(OdfFormDouble::new(invalid).is_err(), "{invalid}");
        }
        for valid in ["2024-02-29", "-0001-01-01Z", "2026-07-19+08:00"] {
            assert!(OdfFormDate::new(valid).is_ok(), "{valid}");
        }
        for invalid in ["0000-01-01", "2023-02-29", "2026-13-01", "2026-01-01+14:01"] {
            assert!(OdfFormDate::new(invalid).is_err(), "{invalid}");
        }

        let wrong_number =
            document(r#"<form:number form:name="N" xml:id="n" form:max-value="one"/>"#);
        assert!(typed_value_controls(&wrong_number).is_err());
        let wrong_date =
            document(r#"<form:date form:name="D" xml:id="d" form:min-value="2023-02-29"/>"#);
        assert!(typed_value_controls(&wrong_date).is_err());
        let wrong_attr =
            document(r#"<form:time form:name="T" xml:id="t" form:decimal-accuracy="2"/>"#);
        assert!(typed_value_controls(&wrong_attr).is_err());
        let foreign =
            document(r#"<form:number xmlns:e="urn:evil" form:name="N" xml:id="n" e:value="1"/>"#);
        assert!(typed_value_controls(&foreign).is_err());
        let active = document(
            r#"<form:date form:name="D" xml:id="d"><office:event-listeners/></form:date>"#,
        );
        assert!(typed_value_controls(&active).is_err());
        assert!(
            typed_value_controls(&format!(
                "<!DOCTYPE x>{}",
                document(r#"<form:time form:name="T" xml:id="t"/>"#)
            ))
            .is_err()
        );

        let mut mismatch = OdfTypedValueControl::new(OdfTypedValueControlKind::Date, "D", "d");
        mismatch.max_value = Some(OdfTypedValueBound::Number(OdfFormDouble::new("1").unwrap()));
        assert!(mismatch.to_xml_fragment().is_err());
    }

    #[test]
    fn lossless_mutation_builder_and_mutable_facades_round_trip() {
        let xml = document(r#"<form:number form:name="Old" xml:id="old" form:value="7"/>"#);
        let inserted =
            OdfTypedValueControl::new(OdfTypedValueControlKind::Date, "Inserted", "inserted");
        let updated = insert_typed_value_control_xml(&xml, 0, &inserted).unwrap();
        assert!(updated.contains("<office:tail/>"));
        let replacement =
            OdfTypedValueControl::new(OdfTypedValueControlKind::Time, "Replacement", "replacement");
        let updated = replace_typed_value_control_xml(&updated, 0, &replacement).unwrap();
        assert_eq!(
            typed_value_controls(&remove_typed_value_control_xml(&updated, 1).unwrap()).unwrap(),
            std::slice::from_ref(&replacement)
        );

        let mut form = OdfTypedValueForm::new("TypedValues");
        form.add_control(OdfTypedValueControl::new(
            OdfTypedValueControlKind::FormattedText,
            "Initial",
            "initial",
        ))
        .unwrap();
        let mut builder = crate::DocumentBuilder::new();
        builder.add_typed_value_form(&form).unwrap();
        let document = crate::Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = crate::MutableDocument::from_document(document).unwrap();
        assert_eq!(mutable.typed_value_controls().unwrap().len(), 1);
        mutable.insert_typed_value_control(0, &inserted).unwrap();
        assert_eq!(
            mutable
                .replace_typed_value_control(0, &replacement)
                .unwrap()
                .name,
            "Initial"
        );
        assert_eq!(
            mutable.remove_typed_value_control(1).unwrap().name,
            "Inserted"
        );
    }
}
