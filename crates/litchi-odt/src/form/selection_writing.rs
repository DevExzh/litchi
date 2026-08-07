use std::ops::Range;

use litchi_core::{Error, Result};
use quick_xml::NsReader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;

const FORM: &str = "urn:oasis:names:tc:opendocument:xmlns:form:1.0";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const XML: &str = "http://www.w3.org/XML/1998/namespace";
const SCRIPT: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XFORMS: &str = "http://www.w3.org/2002/xforms";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_STRING: usize = 1_048_576;
const MAX_SOURCE: usize = 65_536;
const MAX_FORMS: usize = 16_384;
const MAX_CONTROLS: usize = 65_536;
const MAX_ENTRIES: usize = 65_536;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListSourceType {
    Table,
    Query,
    Sql,
    SqlPassThrough,
    ValueList,
    TableFields,
}
impl ListSourceType {
    fn token(self) -> &'static str {
        match self {
            Self::Table => "table",
            Self::Query => "query",
            Self::Sql => "sql",
            Self::SqlPassThrough => "sql-pass-through",
            Self::ValueList => "value-list",
            Self::TableFields => "table-fields",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "table" => Ok(Self::Table),
            "query" => Ok(Self::Query),
            "sql" => Ok(Self::Sql),
            "sql-pass-through" => Ok(Self::SqlPassThrough),
            "value-list" => Ok(Self::ValueList),
            "table-fields" => Ok(Self::TableFields),
            _ => invalid(format!("invalid list source type '{value}'")),
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListLinkageType {
    Selection,
    SelectionIndices,
}
impl ListLinkageType {
    fn token(self) -> &'static str {
        match self {
            Self::Selection => "selection",
            Self::SelectionIndices => "selection-indices",
        }
    }
    fn parse(value: &str) -> Result<Self> {
        match value {
            "selection" => Ok(Self::Selection),
            "selection-indices" => Ok(Self::SelectionIndices),
            _ => invalid(format!("invalid list linkage type '{value}'")),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComboItem {
    pub label: Option<String>,
    pub text: String,
}
impl ComboItem {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            label: None,
            text: text.into(),
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_entry(self.label.as_deref(), None, &self.text)?;
        let mut out = "<form:item".to_string();
        push_string(&mut out, "form:label", self.label.as_deref());
        if self.text.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            out.push_str(&escape(&self.text));
            out.push_str("</form:item>");
        }
        Ok(out)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListOption {
    pub label: Option<String>,
    pub value: Option<String>,
    pub selected: Option<bool>,
    pub current_selected: Option<bool>,
    pub text: String,
}
impl ListOption {
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            label: None,
            value: None,
            selected: None,
            current_selected: None,
            text: text.into(),
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_entry(self.label.as_deref(), self.value.as_deref(), &self.text)?;
        let mut out = "<form:option".to_string();
        push_string(&mut out, "form:label", self.label.as_deref());
        push_string(&mut out, "form:value", self.value.as_deref());
        push_bool(&mut out, "form:selected", self.selected);
        push_bool(&mut out, "form:current-selected", self.current_selected);
        if self.text.is_empty() {
            out.push_str("/>");
        } else {
            out.push('>');
            out.push_str(&escape(&self.text));
            out.push_str("</form:option>");
        }
        Ok(out)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ComboboxControl {
    pub name: String,
    pub xml_id: String,
    pub value: Option<String>,
    pub current_value: Option<String>,
    pub disabled: Option<bool>,
    pub dropdown: Option<bool>,
    pub max_length: Option<u64>,
    pub printable: Option<bool>,
    pub readonly: Option<bool>,
    pub size: Option<u64>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub convert_empty_to_null: Option<bool>,
    pub data_field: Option<String>,
    pub list_source: Option<String>,
    pub list_source_type: Option<ListSourceType>,
    pub linked_cell: Option<String>,
    pub source_cell_range: Option<String>,
    pub auto_complete: Option<bool>,
    pub items: Vec<ComboItem>,
}
impl ComboboxControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            value: None,
            current_value: None,
            disabled: None,
            dropdown: None,
            max_length: None,
            printable: None,
            readonly: None,
            size: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            convert_empty_to_null: None,
            data_field: None,
            list_source: None,
            list_source_type: None,
            linked_cell: None,
            source_cell_range: None,
            auto_complete: None,
            items: Vec::new(),
        }
    }
    pub fn add_item(&mut self, item: ComboItem) -> Result<()> {
        validate_entry(item.label.as_deref(), None, &item.text)?;
        if self.items.len() >= MAX_ENTRIES {
            return invalid("too many combobox items");
        }
        self.items.push(item);
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        combobox_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ListboxControl {
    pub name: String,
    pub xml_id: String,
    pub disabled: Option<bool>,
    pub dropdown: Option<bool>,
    pub printable: Option<bool>,
    pub size: Option<u64>,
    pub tab_index: Option<u64>,
    pub tab_stop: Option<bool>,
    pub title: Option<String>,
    pub bound_column: Option<String>,
    pub data_field: Option<String>,
    pub list_source: Option<String>,
    pub list_source_type: Option<ListSourceType>,
    pub linked_cell: Option<String>,
    pub list_linkage_type: Option<ListLinkageType>,
    pub source_cell_range: Option<String>,
    pub multiple: Option<bool>,
    pub options: Vec<ListOption>,
}
impl ListboxControl {
    pub fn new(name: impl Into<String>, xml_id: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            xml_id: xml_id.into(),
            disabled: None,
            dropdown: None,
            printable: None,
            size: None,
            tab_index: None,
            tab_stop: None,
            title: None,
            bound_column: None,
            data_field: None,
            list_source: None,
            list_source_type: None,
            linked_cell: None,
            list_linkage_type: None,
            source_cell_range: None,
            multiple: None,
            options: Vec::new(),
        }
    }
    pub fn add_option(&mut self, option: ListOption) -> Result<()> {
        validate_entry(
            option.label.as_deref(),
            option.value.as_deref(),
            &option.text,
        )?;
        if self.options.len() >= MAX_ENTRIES {
            return invalid("too many listbox options");
        }
        self.options.push(option);
        if let Err(error) = validate_listbox(self) {
            self.options.pop();
            return Err(error);
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        listbox_xml(self)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SelectionControl {
    Combobox(ComboboxControl),
    Listbox(ListboxControl),
}
impl From<ComboboxControl> for SelectionControl {
    fn from(value: ComboboxControl) -> Self {
        Self::Combobox(value)
    }
}
impl From<ListboxControl> for SelectionControl {
    fn from(value: ListboxControl) -> Self {
        Self::Listbox(value)
    }
}
impl SelectionControl {
    pub fn name(&self) -> &str {
        match self {
            Self::Combobox(v) => &v.name,
            Self::Listbox(v) => &v.name,
        }
    }
    pub fn xml_id(&self) -> &str {
        match self {
            Self::Combobox(v) => &v.xml_id,
            Self::Listbox(v) => &v.xml_id,
        }
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        match self {
            Self::Combobox(v) => combobox_xml(v),
            Self::Listbox(v) => listbox_xml(v),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SelectionForm {
    pub name: String,
    pub controls: Vec<SelectionControl>,
    pub apply_filter: Option<bool>,
}
impl SelectionForm {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            controls: Vec::new(),
            apply_filter: None,
        }
    }
    pub fn add_control(&mut self, control: impl Into<SelectionControl>) -> Result<()> {
        let control = control.into();
        validate_control(&control)?;
        if self.controls.iter().any(|v| v.name() == control.name()) {
            return invalid(format!(
                "duplicate selection control name '{}'",
                control.name()
            ));
        }
        if self.controls.iter().any(|v| v.xml_id() == control.xml_id()) {
            return invalid(format!(
                "duplicate selection control xml:id '{}'",
                control.xml_id()
            ));
        }
        self.controls.push(control);
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_name("form name", &self.name)?;
        let mut out = format!(
            r#"<form:form xmlns:form="{FORM}" form:name="{}""#,
            escape(&self.name)
        );
        push_bool(&mut out, "form:apply-filter", self.apply_filter);
        if self.controls.is_empty() {
            out.push_str("/>");
            return Ok(out);
        }
        out.push('>');
        for c in &self.controls {
            out.push_str(&c.to_xml_fragment()?);
        }
        out.push_str("</form:form>");
        Ok(out)
    }
}

pub fn selection_controls(xml: &str) -> Result<Vec<SelectionControl>> {
    Ok(scan(xml)?.controls.into_iter().map(|v| v.control).collect())
}
pub fn insert_selection_control_xml(
    xml: &str,
    form_index: usize,
    control: &SelectionControl,
) -> Result<String> {
    validate_control(control)?;
    let scan = scan(xml)?;
    let form = scan
        .forms
        .get(form_index)
        .ok_or_else(|| Error::InvalidFormat(format!("form {form_index} is out of bounds")))?;
    reject_duplicate(form, control, None)?;
    let fragment = bind_fragment(control.to_xml_fragment()?);
    match &form.site {
        Site::Paired { close_start } => apply(xml, (*close_start)..(*close_start), &fragment),
        Site::Empty { start, end, qname } => expand_empty(xml, *start, *end, qname, &fragment),
    }
}
pub fn replace_selection_control_xml(
    xml: &str,
    index: usize,
    replacement: &SelectionControl,
) -> Result<String> {
    validate_control(replacement)?;
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("selection control {index} is out of bounds"))
    })?;
    reject_duplicate(&scan.forms[old.form], replacement, Some(&old.control))?;
    apply(
        xml,
        old.span.clone(),
        &bind_fragment(replacement.to_xml_fragment()?),
    )
}
pub fn remove_selection_control_xml(xml: &str, index: usize) -> Result<String> {
    let scan = scan(xml)?;
    let old = scan.controls.get(index).ok_or_else(|| {
        Error::InvalidFormat(format!("selection control {index} is out of bounds"))
    })?;
    apply(xml, old.span.clone(), "")
}

fn combobox_xml(v: &ComboboxControl) -> Result<String> {
    validate_combobox(v)?;
    let mut out = format!(
        r#"<form:combobox form:name="{}" xml:id="{}""#,
        escape(&v.name),
        escape(&v.xml_id)
    );
    push_string(&mut out, "form:value", v.value.as_deref());
    push_string(&mut out, "form:current-value", v.current_value.as_deref());
    push_bool(&mut out, "form:disabled", v.disabled);
    push_bool(&mut out, "form:dropdown", v.dropdown);
    push_u64(&mut out, "form:max-length", v.max_length);
    push_bool(&mut out, "form:printable", v.printable);
    push_bool(&mut out, "form:readonly", v.readonly);
    push_u64(&mut out, "form:size", v.size);
    push_u64(&mut out, "form:tab-index", v.tab_index);
    push_bool(&mut out, "form:tab-stop", v.tab_stop);
    push_string(&mut out, "form:title", v.title.as_deref());
    push_bool(
        &mut out,
        "form:convert-empty-to-null",
        v.convert_empty_to_null,
    );
    push_string(&mut out, "form:data-field", v.data_field.as_deref());
    push_string(&mut out, "form:list-source", v.list_source.as_deref());
    if let Some(t) = v.list_source_type {
        push_string(&mut out, "form:list-source-type", Some(t.token()));
    }
    push_string(&mut out, "form:linked-cell", v.linked_cell.as_deref());
    push_string(
        &mut out,
        "form:source-cell-range",
        v.source_cell_range.as_deref(),
    );
    push_bool(&mut out, "form:auto-complete", v.auto_complete);
    if v.items.is_empty() {
        out.push_str("/>");
    } else {
        out.push('>');
        for item in &v.items {
            out.push_str(&item.to_xml_fragment()?);
        }
        out.push_str("</form:combobox>");
    }
    Ok(out)
}
fn listbox_xml(v: &ListboxControl) -> Result<String> {
    validate_listbox(v)?;
    let mut out = format!(
        r#"<form:listbox form:name="{}" xml:id="{}""#,
        escape(&v.name),
        escape(&v.xml_id)
    );
    push_bool(&mut out, "form:disabled", v.disabled);
    push_bool(&mut out, "form:dropdown", v.dropdown);
    push_bool(&mut out, "form:printable", v.printable);
    push_u64(&mut out, "form:size", v.size);
    push_u64(&mut out, "form:tab-index", v.tab_index);
    push_bool(&mut out, "form:tab-stop", v.tab_stop);
    push_string(&mut out, "form:title", v.title.as_deref());
    push_string(&mut out, "form:bound-column", v.bound_column.as_deref());
    push_string(&mut out, "form:data-field", v.data_field.as_deref());
    push_string(&mut out, "form:list-source", v.list_source.as_deref());
    if let Some(t) = v.list_source_type {
        push_string(&mut out, "form:list-source-type", Some(t.token()));
    }
    push_string(&mut out, "form:linked-cell", v.linked_cell.as_deref());
    if let Some(t) = v.list_linkage_type {
        push_string(&mut out, "form:list-linkage-type", Some(t.token()));
    }
    push_string(
        &mut out,
        "form:source-cell-range",
        v.source_cell_range.as_deref(),
    );
    push_bool(&mut out, "form:multiple", v.multiple);
    if v.options.is_empty() {
        out.push_str("/>");
    } else {
        out.push('>');
        for option in &v.options {
            out.push_str(&option.to_xml_fragment()?);
        }
        out.push_str("</form:listbox>");
    }
    Ok(out)
}
fn validate_control(v: &SelectionControl) -> Result<()> {
    match v {
        SelectionControl::Combobox(v) => validate_combobox(v),
        SelectionControl::Listbox(v) => validate_listbox(v),
    }
}
fn validate_combobox(v: &ComboboxControl) -> Result<()> {
    validate_identity(&v.name, &v.xml_id)?;
    for (label, value) in [
        ("combobox value", v.value.as_deref()),
        ("combobox current value", v.current_value.as_deref()),
        ("combobox title", v.title.as_deref()),
        ("combobox data field", v.data_field.as_deref()),
        ("combobox linked cell", v.linked_cell.as_deref()),
        ("combobox source cell range", v.source_cell_range.as_deref()),
    ] {
        validate_optional(label, value)?;
    }
    validate_source(v.list_source.as_deref())?;
    validate_entries(v.items.iter().map(|i| (&i.label, None, &i.text)))
}
fn validate_listbox(v: &ListboxControl) -> Result<()> {
    validate_identity(&v.name, &v.xml_id)?;
    for (label, value) in [
        ("listbox title", v.title.as_deref()),
        ("listbox bound column", v.bound_column.as_deref()),
        ("listbox data field", v.data_field.as_deref()),
        ("listbox linked cell", v.linked_cell.as_deref()),
        ("listbox source cell range", v.source_cell_range.as_deref()),
    ] {
        validate_optional(label, value)?;
    }
    validate_source(v.list_source.as_deref())?;
    validate_entries(
        v.options
            .iter()
            .map(|i| (&i.label, i.value.as_ref(), &i.text)),
    )?;
    if v.multiple != Some(true) {
        if v.options
            .iter()
            .filter(|o| o.selected == Some(true))
            .count()
            > 1
        {
            return invalid("single-selection listbox has multiple selected options");
        }
        if v.options
            .iter()
            .filter(|o| o.current_selected == Some(true))
            .count()
            > 1
        {
            return invalid("single-selection listbox has multiple current selections");
        }
    }
    Ok(())
}
fn validate_entries<'a, I>(entries: I) -> Result<()>
where
    I: IntoIterator<Item = (&'a Option<String>, Option<&'a String>, &'a String)>,
{
    let mut count = 0usize;
    let mut total = 0usize;
    for (label, value, text) in entries {
        count += 1;
        if count > MAX_ENTRIES {
            return invalid("too many selection entries");
        }
        validate_entry(label.as_deref(), value.map(String::as_str), text)?;
        total = total
            .checked_add(
                text.len() + label.as_ref().map_or(0, String::len) + value.map_or(0, String::len),
            )
            .ok_or_else(|| {
                Error::InvalidFormat("selection entry content is too large".to_string())
            })?;
        if total > MAX_STRING {
            return invalid("selection entry content exceeds 1 MiB");
        }
    }
    Ok(())
}
fn validate_entry(label: Option<&str>, value: Option<&str>, text: &str) -> Result<()> {
    validate_optional("selection entry label", label)?;
    validate_optional("selection entry value", value)?;
    validate_string("selection entry text", text)
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
struct FormLocation {
    site: Site,
    controls: Vec<ControlLocation>,
}
#[derive(Clone)]
struct ControlLocation {
    span: Range<usize>,
    form: usize,
    control: SelectionControl,
}
struct Scan {
    forms: Vec<FormLocation>,
    controls: Vec<ControlLocation>,
}
struct Open {
    local: Vec<u8>,
    form: Option<usize>,
    control: Option<usize>,
    entry: Option<EntryOpen>,
}
struct EntryOpen {
    control: usize,
    kind: EntryKind,
    text: String,
}
enum EntryKind {
    Item {
        label: Option<String>,
    },
    Option {
        label: Option<String>,
        value: Option<String>,
        selected: Option<bool>,
        current_selected: Option<bool>,
    },
}
struct Attr {
    namespace: Option<String>,
    local: String,
    value: String,
}

fn scan(xml: &str) -> Result<Scan> {
    if xml.len() > MAX_XML {
        return invalid("selection form XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut previous = 0usize;
    let mut stack = Vec::<Open>::new();
    let mut form_stack = Vec::<usize>::new();
    let mut forms = Vec::<FormLocation>::new();
    let mut controls = Vec::<ControlLocation>::new();
    let mut ids = Vec::<String>::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|e| Error::InvalidFormat(format!("invalid selection form XML: {e}")))?;
        let namespace = match namespace {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref e) => {
                if stack.len() >= 128 {
                    return invalid("selection form XML nesting exceeds 128 levels");
                }
                track_xml_id(&reader, e, &mut ids)?;
                let local = e.local_name().as_ref().to_vec();
                let mut form = None;
                let mut control = None;
                let mut entry = None;
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, e)?;
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
                    && matches!(local.as_slice(), b"combobox" | b"listbox")
                {
                    if stack.iter().any(|o| o.control.is_some()) {
                        return invalid("selection controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("selection control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, e, local.as_slice())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many selection controls");
                    }
                    control = Some(controls.len());
                    controls.push(ControlLocation {
                        span: previous..0,
                        form: owner,
                        control: parsed,
                    });
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"item" | b"option")
                {
                    let owner = stack.last().and_then(|o| o.control).ok_or_else(|| {
                        Error::InvalidFormat(
                            "selection entry must be a direct child of its control".to_string(),
                        )
                    })?;
                    entry = Some(parse_entry(&reader, e, local.as_slice(), owner)?);
                } else if stack.iter().any(|o| o.entry.is_some()) {
                    return invalid("selection item/option cannot contain child elements");
                } else if !form_stack.is_empty()
                    && ((namespace.as_deref() == Some(OFFICE)
                        && matches!(local.as_slice(), b"event-listeners" | b"events"))
                        || (namespace.as_deref() == Some(SCRIPT)
                            && matches!(local.as_slice(), b"event-listener" | b"event")))
                {
                    return invalid(
                        "event and macro content is outside the selection-control mutation API",
                    );
                } else if stack.iter().any(|o| o.control.is_some())
                    && !(namespace.as_deref() == Some(FORM)
                        && matches!(
                            local.as_slice(),
                            b"properties" | b"property" | b"list-property" | b"list-value"
                        ))
                {
                    return invalid("unexpected child element in selection control");
                }
                stack.push(Open {
                    local,
                    form,
                    control,
                    entry,
                });
            },
            Event::Empty(ref e) => {
                track_xml_id(&reader, e, &mut ids)?;
                let local = e.local_name().as_ref().to_vec();
                if namespace.as_deref() == Some(FORM) && local == b"form" {
                    validate_form(&reader, e)?;
                    if forms.len() >= MAX_FORMS {
                        return invalid("too many forms");
                    }
                    forms.push(FormLocation {
                        site: Site::Empty {
                            start: previous,
                            end,
                            qname: qname(e)?,
                        },
                        controls: Vec::new(),
                    });
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"combobox" | b"listbox")
                {
                    if stack.iter().any(|o| o.control.is_some()) {
                        return invalid("selection controls cannot be nested");
                    }
                    let owner = *form_stack.last().ok_or_else(|| {
                        Error::InvalidFormat("selection control has no form owner".to_string())
                    })?;
                    let parsed = parse_control(&reader, e, local.as_slice())?;
                    if controls.len() >= MAX_CONTROLS {
                        return invalid("too many selection controls");
                    }
                    controls.push(ControlLocation {
                        span: previous..end,
                        form: owner,
                        control: parsed,
                    });
                } else if namespace.as_deref() == Some(FORM)
                    && matches!(local.as_slice(), b"item" | b"option")
                {
                    let owner = stack.last().and_then(|o| o.control).ok_or_else(|| {
                        Error::InvalidFormat(
                            "selection entry must be a direct child of its control".to_string(),
                        )
                    })?;
                    finish_entry(
                        &mut controls,
                        parse_entry(&reader, e, local.as_slice(), owner)?,
                    )?;
                } else if stack.iter().any(|o| o.entry.is_some()) {
                    return invalid("selection item/option cannot contain child elements");
                } else if !form_stack.is_empty()
                    && ((namespace.as_deref() == Some(OFFICE)
                        && matches!(local.as_slice(), b"event-listeners" | b"events"))
                        || (namespace.as_deref() == Some(SCRIPT)
                            && matches!(local.as_slice(), b"event-listener" | b"event")))
                {
                    return invalid(
                        "event and macro content is outside the selection-control mutation API",
                    );
                } else if stack.iter().any(|o| o.control.is_some())
                    && !(namespace.as_deref() == Some(FORM)
                        && matches!(
                            local.as_slice(),
                            b"properties" | b"property" | b"list-property" | b"list-value"
                        ))
                {
                    return invalid("unexpected child element in selection control");
                }
            },
            Event::Text(text) => {
                if let Some(entry) = stack.iter_mut().rev().find_map(|o| o.entry.as_mut()) {
                    let decoded = text.decode().map_err(|e| {
                        Error::InvalidFormat(format!("invalid selection entry text: {e}"))
                    })?;
                    let unescaped = quick_xml::escape::unescape(&decoded).map_err(|e| {
                        Error::InvalidFormat(format!("invalid selection entry text: {e}"))
                    })?;
                    entry.text.push_str(&unescaped);
                    if entry.text.len() > MAX_STRING {
                        return invalid("selection entry text exceeds 1 MiB");
                    }
                }
            },
            Event::CData(text) => {
                if let Some(entry) = stack.iter_mut().rev().find_map(|o| o.entry.as_mut()) {
                    entry.text.push_str(&text.decode().map_err(|e| {
                        Error::InvalidFormat(format!("invalid selection entry CDATA: {e}"))
                    })?);
                    if entry.text.len() > MAX_STRING {
                        return invalid("selection entry text exceeds 1 MiB");
                    }
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(entry) = stack.iter_mut().rev().find_map(|o| o.entry.as_mut()) {
                    if let Some(ch) = reference.resolve_char_ref().map_err(|e| {
                        Error::InvalidFormat(format!(
                            "invalid selection entry character reference: {e}"
                        ))
                    })? {
                        entry.text.push(ch);
                    } else {
                        let name = reference.decode().map_err(|e| {
                            Error::InvalidFormat(format!(
                                "invalid selection entry entity reference: {e}"
                            ))
                        })?;
                        let value = quick_xml::escape::resolve_predefined_entity(&name)
                            .ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "unsupported selection entry entity reference '&{name};'"
                                ))
                            })?;
                        entry.text.push_str(value);
                    }
                    if entry.text.len() > MAX_STRING {
                        return invalid("selection entry text exceeds 1 MiB");
                    }
                }
            },
            Event::End(ref e) => {
                let open = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("selection form XML stack underflow".to_string())
                })?;
                if open.local.as_slice() != e.local_name().as_ref() {
                    return invalid("mismatched selection form XML elements");
                }
                if let Some(entry) = open.entry {
                    finish_entry(&mut controls, entry)?;
                }
                if let Some(i) = open.control {
                    controls[i].span.end = end;
                    validate_control(&controls[i].control)?;
                }
                if let Some(i) = open.form {
                    forms[i].site = Site::Paired {
                        close_start: previous,
                    };
                    form_stack.pop();
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in selection form XML"),
            Event::Eof => break,
            _ => {},
        }
        previous = end;
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed selection form XML elements");
    }
    for item in &controls {
        forms[item.form].controls.push(item.clone());
    }
    Ok(Scan { forms, controls })
}

fn parse_control(
    reader: &NsReader<&[u8]>,
    e: &BytesStart<'_>,
    local: &[u8],
) -> Result<SelectionControl> {
    let attrs = attributes(reader, e)?;
    validate_allowed(
        &attrs,
        if local == b"combobox" {
            COMBO_ATTRS
        } else {
            LIST_ATTRS
        },
    )?;
    let name = required(&attrs, FORM, "name")?;
    let id = required(&attrs, XML, "id")?;
    if local == b"combobox" {
        let mut v = ComboboxControl::new(name, id);
        v.value = optional(&attrs, FORM, "value");
        v.current_value = optional(&attrs, FORM, "current-value");
        v.disabled = optional_bool(&attrs, FORM, "disabled")?;
        v.dropdown = optional_bool(&attrs, FORM, "dropdown")?;
        v.max_length = optional_u64(&attrs, FORM, "max-length")?;
        v.printable = optional_bool(&attrs, FORM, "printable")?;
        v.readonly = optional_bool(&attrs, FORM, "readonly")?;
        v.size = optional_u64(&attrs, FORM, "size")?;
        v.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        v.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        v.title = optional(&attrs, FORM, "title");
        v.convert_empty_to_null = optional_bool(&attrs, FORM, "convert-empty-to-null")?;
        v.data_field = optional(&attrs, FORM, "data-field");
        v.list_source = optional(&attrs, FORM, "list-source");
        v.list_source_type = optional(&attrs, FORM, "list-source-type")
            .map(|s| ListSourceType::parse(&s))
            .transpose()?;
        v.linked_cell = optional(&attrs, FORM, "linked-cell");
        v.source_cell_range = optional(&attrs, FORM, "source-cell-range");
        v.auto_complete = optional_bool(&attrs, FORM, "auto-complete")?;
        validate_combobox(&v)?;
        Ok(v.into())
    } else {
        let mut v = ListboxControl::new(name, id);
        v.disabled = optional_bool(&attrs, FORM, "disabled")?;
        v.dropdown = optional_bool(&attrs, FORM, "dropdown")?;
        v.printable = optional_bool(&attrs, FORM, "printable")?;
        v.size = optional_u64(&attrs, FORM, "size")?;
        v.tab_index = optional_u64(&attrs, FORM, "tab-index")?;
        v.tab_stop = optional_bool(&attrs, FORM, "tab-stop")?;
        v.title = optional(&attrs, FORM, "title");
        v.bound_column = optional(&attrs, FORM, "bound-column");
        v.data_field = optional(&attrs, FORM, "data-field");
        v.list_source = optional(&attrs, FORM, "list-source");
        v.list_source_type = optional(&attrs, FORM, "list-source-type")
            .map(|s| ListSourceType::parse(&s))
            .transpose()?;
        v.linked_cell = optional(&attrs, FORM, "linked-cell");
        v.list_linkage_type = optional(&attrs, FORM, "list-linkage-type")
            .map(|s| ListLinkageType::parse(&s))
            .transpose()?;
        v.source_cell_range = optional(&attrs, FORM, "source-cell-range");
        v.multiple = optional_bool(&attrs, FORM, "multiple")?;
        validate_listbox(&v)?;
        Ok(v.into())
    }
}
fn parse_entry(
    reader: &NsReader<&[u8]>,
    e: &BytesStart<'_>,
    local: &[u8],
    control: usize,
) -> Result<EntryOpen> {
    let attrs = attributes(reader, e)?;
    validate_allowed(
        &attrs,
        if local == b"item" {
            ITEM_ATTRS
        } else {
            OPTION_ATTRS
        },
    )?;
    let kind = if local == b"item" {
        EntryKind::Item {
            label: optional(&attrs, FORM, "label"),
        }
    } else {
        EntryKind::Option {
            label: optional(&attrs, FORM, "label"),
            value: optional(&attrs, FORM, "value"),
            selected: optional_bool(&attrs, FORM, "selected")?,
            current_selected: optional_bool(&attrs, FORM, "current-selected")?,
        }
    };
    Ok(EntryOpen {
        control,
        kind,
        text: String::new(),
    })
}
fn finish_entry(controls: &mut [ControlLocation], entry: EntryOpen) -> Result<()> {
    let control = controls
        .get_mut(entry.control)
        .ok_or_else(|| Error::InvalidFormat("selection entry owner is invalid".to_string()))?;
    match (&mut control.control, entry.kind) {
        (SelectionControl::Combobox(v), EntryKind::Item { label }) => v.add_item(ComboItem {
            label,
            text: entry.text,
        })?,
        (
            SelectionControl::Listbox(v),
            EntryKind::Option {
                label,
                value,
                selected,
                current_selected,
            },
        ) => v.add_option(ListOption {
            label,
            value,
            selected,
            current_selected,
            text: entry.text,
        })?,
        _ => return invalid("form:item/form:option has the wrong selection control owner"),
    }
    Ok(())
}

const COMBO_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (FORM, "disabled"),
    (FORM, "dropdown"),
    (FORM, "printable"),
    (FORM, "size"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (XFORMS, "bind"),
    (FORM, "value"),
    (FORM, "current-value"),
    (FORM, "max-length"),
    (FORM, "readonly"),
    (FORM, "convert-empty-to-null"),
    (FORM, "data-field"),
    (FORM, "list-source"),
    (FORM, "list-source-type"),
    (FORM, "linked-cell"),
    (FORM, "source-cell-range"),
    (FORM, "auto-complete"),
];
const LIST_ATTRS: &[(&str, &str)] = &[
    (FORM, "name"),
    (XML, "id"),
    (FORM, "id"),
    (FORM, "control-implementation"),
    (FORM, "input-required"),
    (FORM, "disabled"),
    (FORM, "dropdown"),
    (FORM, "printable"),
    (FORM, "size"),
    (FORM, "tab-index"),
    (FORM, "tab-stop"),
    (FORM, "title"),
    (XFORMS, "bind"),
    (FORM, "bound-column"),
    (FORM, "data-field"),
    (FORM, "list-source"),
    (FORM, "list-source-type"),
    (FORM, "linked-cell"),
    (FORM, "list-linkage-type"),
    (FORM, "source-cell-range"),
    (FORM, "multiple"),
    (FORM, "xforms-list-source"),
];
const ITEM_ATTRS: &[(&str, &str)] = &[(FORM, "label")];
const OPTION_ATTRS: &[(&str, &str)] = &[
    (FORM, "label"),
    (FORM, "value"),
    (FORM, "selected"),
    (FORM, "current-selected"),
];
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
fn validate_form(reader: &NsReader<&[u8]>, e: &BytesStart<'_>) -> Result<()> {
    let attrs = attributes(reader, e)?;
    validate_allowed(&attrs, FORM_ATTRS)?;
    validate_name("form name", &required(&attrs, FORM, "name")?)?;
    for a in &attrs {
        if a.namespace.as_deref() == Some(FORM)
            && matches!(
                a.local.as_str(),
                "allow-deletes"
                    | "allow-inserts"
                    | "allow-updates"
                    | "apply-filter"
                    | "escape-processing"
                    | "ignore-result"
            )
        {
            parse_bool(&a.value, &a.local)?;
        }
    }
    Ok(())
}
fn attributes(reader: &NsReader<&[u8]>, e: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut out = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(|e| Error::InvalidFormat(format!("invalid selection attribute: {e}")))?;
        let (ns, local) = reader.resolver().resolve_attribute(a.key);
        let ns = match ns {
            ResolveResult::Bound(uri) => Some(String::from_utf8_lossy(uri.as_ref()).into_owned()),
            _ => None,
        };
        if ns.as_deref() == Some("http://www.w3.org/2000/xmlns/") || a.key.as_ref() == b"xmlns" {
            continue;
        }
        let value = a
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|e| Error::InvalidFormat(format!("invalid selection attribute value: {e}")))?
            .into_owned();
        out.push(Attr {
            namespace: ns,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            value,
        });
    }
    Ok(out)
}
fn track_xml_id(reader: &NsReader<&[u8]>, e: &BytesStart<'_>, ids: &mut Vec<String>) -> Result<()> {
    if let Some(id) = optional(&attributes(reader, e)?, XML, "id") {
        validate_xml_id(&id)?;
        if ids.iter().any(|v| v == &id) {
            return invalid(format!("duplicate xml:id '{id}'"));
        }
        ids.push(id);
    }
    Ok(())
}
fn validate_allowed(attrs: &[Attr], allowed: &[(&str, &str)]) -> Result<()> {
    for a in attrs {
        if !allowed
            .iter()
            .any(|(ns, local)| a.namespace.as_deref() == Some(*ns) && a.local == *local)
        {
            return invalid(format!(
                "unsupported selection attribute '{}:{}'",
                a.namespace.as_deref().unwrap_or(""),
                a.local
            ));
        }
    }
    Ok(())
}
fn optional(attrs: &[Attr], ns: &str, local: &str) -> Option<String> {
    attrs
        .iter()
        .find(|a| a.namespace.as_deref() == Some(ns) && a.local == local)
        .map(|a| a.value.clone())
}
fn required(attrs: &[Attr], ns: &str, local: &str) -> Result<String> {
    optional(attrs, ns, local)
        .ok_or_else(|| Error::InvalidFormat(format!("missing required attribute {local}")))
}
fn optional_bool(attrs: &[Attr], ns: &str, local: &str) -> Result<Option<bool>> {
    optional(attrs, ns, local)
        .map(|v| parse_bool(&v, local))
        .transpose()
}
fn parse_bool(v: &str, local: &str) -> Result<bool> {
    match v {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid boolean '{v}' for {local}")),
    }
}
fn optional_u64(attrs: &[Attr], ns: &str, local: &str) -> Result<Option<u64>> {
    optional(attrs, ns, local)
        .map(|v| {
            v.parse::<u64>().map_err(|_| {
                Error::InvalidFormat(format!("invalid non-negative integer '{v}' for {local}"))
            })
        })
        .transpose()
}
fn validate_identity(name: &str, id: &str) -> Result<()> {
    validate_name("selection control name", name)?;
    validate_xml_id(id)
}
fn validate_name(label: &str, v: &str) -> Result<()> {
    validate_string(label, v)?;
    if v.is_empty() {
        invalid(format!("{label} cannot be empty"))
    } else {
        Ok(())
    }
}
fn validate_xml_id(v: &str) -> Result<()> {
    validate_name("selection control xml:id", v)?;
    let mut chars = v.chars();
    let first = chars.next().unwrap();
    if !(first == '_' || first.is_ascii_alphabetic())
        || !chars.all(|c| c == '_' || c == '-' || c == '.' || c.is_ascii_alphanumeric())
    {
        return invalid(format!("invalid selection control xml:id '{v}'"));
    }
    Ok(())
}
fn validate_optional(label: &str, v: Option<&str>) -> Result<()> {
    if let Some(v) = v {
        validate_string(label, v)?;
    }
    Ok(())
}
fn validate_string(label: &str, v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        return invalid(format!("{label} exceeds 1 MiB"));
    }
    if v.chars().any(|c| matches!(c as u32,0..=8|11|12|14..=31)) {
        return invalid(format!("{label} contains invalid XML characters"));
    }
    Ok(())
}
fn validate_source(v: Option<&str>) -> Result<()> {
    if let Some(v) = v {
        if v.len() > MAX_SOURCE {
            return invalid("selection source exceeds 64 KiB");
        }
        validate_string("selection source", v)?;
    }
    Ok(())
}
fn push_string(out: &mut String, name: &str, v: Option<&str>) {
    if let Some(v) = v {
        out.push_str(&format!(r#" {name}="{}""#, escape(v)));
    }
}
fn push_bool(out: &mut String, name: &str, v: Option<bool>) {
    if let Some(v) = v {
        out.push_str(&format!(
            r#" {name}="{}""#,
            if v { "true" } else { "false" }
        ));
    }
}
fn push_u64(out: &mut String, name: &str, v: Option<u64>) {
    if let Some(v) = v {
        out.push_str(&format!(r#" {name}="{v}""#));
    }
}
fn escape(v: &str) -> String {
    v.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}
fn bind_fragment(mut v: String) -> String {
    if v.contains("form:") && !v.contains("xmlns:form=") {
        v = v.replacen(' ', &format!(" xmlns:form=\"{FORM}\" "), 1);
    }
    v
}
fn reject_duplicate(
    form: &FormLocation,
    replacement: &SelectionControl,
    current: Option<&SelectionControl>,
) -> Result<()> {
    for item in &form.controls {
        if item.control.name() == replacement.name()
            && current.is_none_or(|v| v.name() != item.control.name())
        {
            return invalid(format!(
                "duplicate selection control name '{}'",
                replacement.name()
            ));
        }
        if item.control.xml_id() == replacement.xml_id()
            && current.is_none_or(|v| v.xml_id() != item.control.xml_id())
        {
            return invalid(format!(
                "duplicate selection control xml:id '{}'",
                replacement.xml_id()
            ));
        }
    }
    Ok(())
}
fn qname(e: &BytesStart<'_>) -> Result<String> {
    String::from_utf8(e.name().as_ref().to_vec())
        .map_err(|_| Error::InvalidFormat("invalid selection element name".to_string()))
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

#[cfg(test)]
mod tests {
    use super::*;
    const ROOT: &str = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:f="urn:oasis:names:tc:opendocument:xmlns:form:1.0"><o:body><o:text><o:forms>"#;
    const END: &str = "</o:forms></o:text></o:body></o:document-content>";
    #[test]
    fn canonical_controls_and_children_round_trip() {
        let mut combo = ComboboxControl::new("Choice & edit", "combo_1");
        combo.value = Some("A < B".into());
        combo.auto_complete = Some(true);
        let mut item = ComboItem::new("alpha & beta");
        item.label = Some("Alpha".into());
        combo.add_item(item).unwrap();
        let mut list = ListboxControl::new("Choice", "list_1");
        list.multiple = Some(false);
        let mut option = ListOption::new("one");
        option.value = Some("1".into());
        option.selected = Some(true);
        list.add_option(option).unwrap();
        let mut form = SelectionForm::new("Main");
        form.add_control(combo).unwrap();
        form.add_control(list).unwrap();
        let xml = format!("{ROOT}{}{END}", form.to_xml_fragment().unwrap());
        let parsed = selection_controls(&xml).unwrap();
        assert_eq!(parsed.len(), 2);
        let SelectionControl::Combobox(parsed_combo) = &parsed[0] else {
            panic!("expected combobox")
        };
        assert_eq!(parsed_combo.items[0].text, "alpha & beta");
    }
    #[test]
    fn aliases_lossless_mutation_and_empty_form_expansion() {
        let xml = format!(
            r#"{ROOT}<f:form f:name="Main"><f:listbox f:name="Old" xml:id="old"><f:properties><f:property f:property-name="Keep" o:value-type="void"/></f:properties><f:option f:label="One" f:selected="true">one</f:option></f:listbox><!--keep--><f:text f:name="Text" xml:id="text"/></f:form>{END}"#
        );
        let control: SelectionControl = ComboboxControl::new("Combo", "combo").into();
        let inserted = insert_selection_control_xml(&xml, 0, &control).unwrap();
        assert!(inserted.contains("<!--keep-->") && inserted.contains("form:combobox"));
        let replacement: SelectionControl = ListboxControl::new("New", "new").into();
        let replaced = replace_selection_control_xml(&inserted, 0, &replacement).unwrap();
        assert!(replaced.contains("<!--keep-->") && replaced.contains("f:text"));
        let removed = remove_selection_control_xml(&replaced, 1).unwrap();
        assert_eq!(selection_controls(&removed).unwrap().len(), 1);
        let empty = format!(r#"{ROOT}<f:form f:name="Empty"/>{END}"#);
        assert!(
            insert_selection_control_xml(&empty, 0, &control)
                .unwrap()
                .contains("</f:form>")
        );
    }
    #[test]
    fn hostile_input_bounds_events_and_cardinality_are_rejected() {
        assert!(ComboboxControl::new("C", "1bad").to_xml_fragment().is_err());
        let wrong = format!(
            r#"{ROOT}<f:form f:name="Main"><o:listbox f:name="L" xml:id="l"/></f:form>{END}"#
        );
        assert!(selection_controls(&wrong).unwrap().is_empty());
        let attr = format!(
            r#"{ROOT}<f:form f:name="Main"><f:listbox f:name="L" xml:id="l" o:size="2"/></f:form>{END}"#
        );
        assert!(selection_controls(&attr).is_err());
        let child = format!(
            r#"{ROOT}<f:form f:name="Main"><f:combobox f:name="C" xml:id="c"><f:option/></f:combobox></f:form>{END}"#
        );
        assert!(selection_controls(&child).is_err());
        let event = format!(r#"{ROOT}<f:form f:name="Main"><o:event-listeners/></f:form>{END}"#);
        assert!(selection_controls(&event).is_err());
        let duplicate = format!(
            r#"{ROOT}<o:p xml:id="same"/><f:form f:name="Main"><f:listbox f:name="L" xml:id="same"/></f:form>{END}"#
        );
        assert!(selection_controls(&duplicate).is_err());
        let mut list = ListboxControl::new("L", "l");
        for text in ["one", "two"] {
            let mut o = ListOption::new(text);
            o.selected = Some(true);
            list.options.push(o);
        }
        assert!(list.to_xml_fragment().is_err());
        let mut combo = ComboboxControl::new("C", "c");
        combo.list_source = Some("x".repeat(MAX_SOURCE + 1));
        assert!(combo.to_xml_fragment().is_err());
    }
    #[test]
    fn unexpected_children_are_rejected_and_failed_add_is_atomic() {
        let foreign = format!(
            r#"{ROOT}<f:form f:name="Main"><f:listbox f:name="L" xml:id="l"><o:p/></f:listbox></f:form>{END}"#
        );
        assert!(selection_controls(&foreign).is_err());
        let nested = format!(
            r#"{ROOT}<f:form f:name="Main"><f:listbox f:name="L" xml:id="l"><f:listbox f:name="Nested" xml:id="nested"/></f:listbox></f:form>{END}"#
        );
        assert!(selection_controls(&nested).is_err());
        let mut list = ListboxControl::new("L", "l");
        let mut first = ListOption::new("one");
        first.selected = Some(true);
        list.add_option(first).unwrap();
        let mut second = ListOption::new("two");
        second.selected = Some(true);
        assert!(list.add_option(second).is_err());
        assert_eq!(list.options.len(), 1);
    }
    #[test]
    fn libreoffice_odfpy_odfdo_and_inline_shapes_parse() {
        let lo = include_str!(
            "../../../../test-data/libreoffice-core/vcl/qa/cppunit/pdfexport/data/tdf159817.fodt"
        );
        let parsed = selection_controls(lo).unwrap();
        assert!(parsed.iter().any(|v|matches!(v,SelectionControl::Listbox(l)if l.list_source_type==Some(ListSourceType::Sql))));
        let producer = format!(
            r#"{ROOT}<f:form f:name="Producers"><f:combobox f:name="odfpy" xml:id="combo" f:dropdown="true"><f:item f:label="A">alpha</f:item></f:combobox><f:listbox f:name="odfdo" xml:id="list" f:multiple="true" f:list-linkage-type="selection-indices"><f:option f:label="A" f:value="1" f:selected="true">alpha</f:option><f:option f:label="B" f:value="2" f:current-selected="true">beta</f:option></f:listbox></f:form>{END}"#
        );
        assert_eq!(selection_controls(&producer).unwrap().len(), 2);
    }
    #[test]
    fn builder_mutable_package_round_trip() {
        use crate::{Builder, Document, mutable::MutableDocument};
        let mut form = SelectionForm::new("Main");
        form.add_control(ComboboxControl::new("Combo", "combo"))
            .unwrap();
        let mut builder = Builder::new();
        builder.add_selection_form(&form).unwrap();
        builder.add_paragraph("body").unwrap();
        let doc = Document::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutableDocument::from_document(doc).unwrap();
        assert_eq!(mutable.selection_controls().unwrap().len(), 1);
        let list: SelectionControl = ListboxControl::new("List", "list").into();
        mutable.insert_selection_control(0, &list).unwrap();
        let replacement: SelectionControl = ComboboxControl::new("Other", "other").into();
        assert_eq!(
            mutable
                .replace_selection_control(0, &replacement)
                .unwrap()
                .name(),
            "Combo"
        );
        assert_eq!(mutable.remove_selection_control(1).unwrap().name(), "List");
        assert_eq!(mutable.selection_controls().unwrap(), [replacement]);
    }
}
