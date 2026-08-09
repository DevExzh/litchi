//! Source-checked stored-query package transactions.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

use crate::Database;

const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const MAX_QUERY_VALUE_BYTES: usize = 1024 * 1024;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

/// A source-bound, single-query edit over one immutable database snapshot.
pub struct Edit<'source> {
    source: &'source Database,
    change: Option<QueryChange>,
}

impl<'source> Edit<'source> {
    pub(crate) const fn new(source: &'source Database) -> Self {
        Self {
            source,
            change: None,
        }
    }

    /// Replaces the inert command text stored for one exactly named query.
    ///
    /// The command is never parsed or executed. A second call for the same
    /// query updates the staged value; selecting another query is refused.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or ambiguous query, a value over the
    /// semantic limit, or a second query selector in the same transaction.
    pub fn set_query_command(&mut self, name: &str, value: impl Into<String>) -> Result<()> {
        let command = value.into();
        validate_value(&command, "command")?;
        let staged = self.ensure_query(name)?;
        staged.after_command = command;
        self.discard_noop();
        Ok(())
    }

    /// Sets or removes the stored `db:escape-processing` declaration.
    ///
    /// This only changes descriptive metadata and never enables execution.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent or ambiguous query or a second query
    /// selector in the same transaction.
    pub fn set_query_escape_processing(&mut self, name: &str, value: Option<bool>) -> Result<()> {
        let staged = self.ensure_query(name)?;
        staged.after_escape_processing = value;
        self.discard_noop();
        Ok(())
    }

    fn discard_noop(&mut self) {
        if self.change.as_ref().is_some_and(QueryChange::is_noop) {
            self.change = None;
        }
    }

    fn ensure_query(&mut self, name: &str) -> Result<&mut QueryChange> {
        if self
            .change
            .as_ref()
            .is_some_and(|change| change.name != name)
        {
            return invalid("an ODB transaction supports one stored-query edit");
        }
        if self.change.is_none() {
            let catalog = self.source.catalog()?;
            let query = catalog.query(name)?.ok_or_else(|| {
                Error::InvalidFormat(format!("ODB query '{name}' does not exist"))
            })?;
            self.change = Some(QueryChange {
                name: name.to_owned(),
                before_command: query.command().to_owned(),
                after_command: query.command().to_owned(),
                before_escape_processing: query.escape_processing(),
                after_escape_processing: query.escape_processing(),
            });
        }
        self.change
            .as_mut()
            .ok_or_else(|| Error::InvalidFormat("ODB staged query disappeared".to_string()))
    }

    /// Atomically rebuilds, reopens, and semantically verifies the candidate.
    ///
    /// # Errors
    ///
    /// Returns an error when the source cannot be losslessly rebuilt, the XML
    /// edit is not addressable, or complete package readback fails.
    pub fn commit(self) -> Result<Commit> {
        let Some(change) = self.change else {
            return Ok(Commit::unchanged(self.source.clone()));
        };
        let site = locate_query(self.source.content_xml(), &change.name)?;
        let rewritten = rewrite_query_tag(self.source.content_xml(), &site, &change)?;
        let snapshot = Database {
            package: self.source.package.rebuild_with_content(&rewritten)?,
        };
        let catalog = snapshot.catalog()?;
        let query = catalog.query(&change.name)?.ok_or_else(|| {
            Error::InvalidFormat("ODB edited query disappeared during readback".to_string())
        })?;
        if query.command() != change.after_command
            || query.escape_processing() != change.after_escape_processing
        {
            return invalid("ODB package edit failed semantic readback");
        }
        Ok(Commit {
            patch: Patch {
                source: self.source.clone(),
                target: snapshot.clone(),
                change: Some(change),
            },
            snapshot,
            changed: true,
        })
    }
}

/// One reversible stored-query metadata operation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryChange {
    name: String,
    before_command: String,
    after_command: String,
    before_escape_processing: Option<bool>,
    after_escape_processing: Option<bool>,
}

impl QueryChange {
    /// Returns the exact producer-visible query name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the command expected in the source snapshot.
    #[must_use]
    pub fn before_command(&self) -> &str {
        &self.before_command
    }

    /// Returns the staged inert command text.
    #[must_use]
    pub fn after_command(&self) -> &str {
        &self.after_command
    }

    /// Returns the source escape-processing declaration.
    #[must_use]
    pub const fn before_escape_processing(&self) -> Option<bool> {
        self.before_escape_processing
    }

    /// Returns the staged escape-processing declaration.
    #[must_use]
    pub const fn after_escape_processing(&self) -> Option<bool> {
        self.after_escape_processing
    }

    fn is_noop(&self) -> bool {
        self.before_escape_processing == self.after_escape_processing
            && self.before_command == self.after_command
    }
}

/// A committed immutable database and its source-checked reversible patch.
pub struct Commit {
    snapshot: Database,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Database) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                change: None,
            },
            snapshot,
            changed: false,
        }
    }

    /// Returns whether package bytes changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Returns the committed immutable database snapshot.
    #[must_use]
    pub const fn database(&self) -> &Database {
        &self.snapshot
    }

    /// Returns the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes this commit into its published database snapshot.
    #[must_use]
    pub fn into_database(self) -> Database {
        self.snapshot
    }
}

/// A byte-exact source-checked reversible ODB stored-query patch.
#[derive(Clone)]
pub struct Patch {
    source: Database,
    target: Database,
    change: Option<QueryChange>,
}

impl Patch {
    /// Returns whether the patch authorizes this exact source artifact.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Database) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub fn apply(&self, source: &Database) -> Result<Database> {
        if !self.is_applicable_to(source) {
            return invalid("ODB patch source does not match its expected snapshot");
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic change, if this patch is not an exact no-op.
    #[must_use]
    pub const fn change(&self) -> Option<&QueryChange> {
        self.change.as_ref()
    }

    /// Returns the patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            change: self.change.as_ref().map(|change| QueryChange {
                name: change.name.clone(),
                before_command: change.after_command.clone(),
                after_command: change.before_command.clone(),
                before_escape_processing: change.after_escape_processing,
                after_escape_processing: change.before_escape_processing,
            }),
        }
    }
}

struct QuerySite {
    tag: Range<usize>,
    command_qname: String,
    escape_processing_qname: Option<String>,
    database_prefix: Option<String>,
}

fn locate_query(source: &str, wanted: &str) -> Result<QuerySite> {
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut selected = None;
    loop {
        let (namespace, raw_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODB edit XML: {error}")))?;
        let is_database =
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE);
        let event = raw_event.into_owned();
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidFormat("ODB edit XML position exceeds this platform".to_string())
        })?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if is_database && element.local_name().as_ref() == b"query" =>
            {
                let name = required_db_attribute(&reader, &element, b"name")?;
                if name == wanted {
                    if selected.is_some() {
                        return invalid("ODB query edit selector is ambiguous");
                    }
                    let start = source[..end].rfind('<').ok_or_else(|| {
                        Error::InvalidFormat("ODB query start tag is missing".to_string())
                    })?;
                    let tag = start..end;
                    if source.as_bytes().get(end.wrapping_sub(1)) != Some(&b'>') {
                        return invalid("ODB query start tag has no closing delimiter");
                    }
                    let command_qname = attribute_qname(&reader, &element, b"command")?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODB query command attribute is missing".to_string(),
                            )
                        })?;
                    let escape_processing_qname =
                        attribute_qname(&reader, &element, b"escape-processing")?;
                    let element_name = element.name();
                    let raw_name =
                        std::str::from_utf8(element_name.as_ref()).map_err(|_error| {
                            Error::InvalidFormat("ODB query name is not UTF-8".to_string())
                        })?;
                    let database_prefix = raw_name
                        .rsplit_once(':')
                        .map(|(prefix, _)| prefix.to_owned());
                    selected = Some(QuerySite {
                        tag,
                        command_qname,
                        escape_processing_qname,
                        database_prefix,
                    });
                }
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODB edit XML"),
            Event::Eof => break,
            Event::Start(_)
            | Event::Empty(_)
            | Event::End(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    selected.ok_or_else(|| Error::InvalidFormat("ODB query edit site was not found".to_string()))
}

fn rewrite_query_tag(source: &str, site: &QuerySite, change: &QueryChange) -> Result<String> {
    let raw = source
        .get(site.tag.clone())
        .ok_or_else(|| Error::InvalidFormat("ODB query tag span is invalid".to_string()))?;
    let command = quick_xml::escape::escape(&change.after_command);
    let mut tag = replace_attribute(raw, &site.command_qname, &command)?;
    tag = match (
        site.escape_processing_qname.as_deref(),
        change.after_escape_processing,
    ) {
        (Some(name), Some(value)) => replace_attribute(&tag, name, bool_text(value))?,
        (Some(name), None) => remove_attribute(&tag, name)?,
        (None, Some(value)) => {
            let prefix = site.database_prefix.as_deref().ok_or_else(|| {
                Error::Unsupported(
                    "ODB cannot add a namespaced query attribute to a default-namespace element"
                        .to_string(),
                )
            })?;
            insert_attribute(
                &tag,
                &format!("{prefix}:escape-processing"),
                bool_text(value),
            )?
        },
        (None, None) => tag,
    };
    let output_size = source
        .len()
        .checked_sub(site.tag.end - site.tag.start)
        .and_then(|size| size.checked_add(tag.len()))
        .ok_or_else(|| Error::InvalidFormat("ODB edited content size overflow".to_string()))?;
    if output_size > MAX_OUTPUT_BYTES {
        return invalid("ODB edited content exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve_exact(output_size)
        .map_err(|allocation| Error::Allocation {
            resource: "ODB edited content",
            source: allocation,
        })?;
    output.push_str(&source[..site.tag.start]);
    output.push_str(&tag);
    output.push_str(&source[site.tag.end..]);
    Ok(output)
}

fn validate_value(value: &str, kind: &str) -> Result<()> {
    if value.len() > MAX_QUERY_VALUE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODB query {kind} exceeds the byte limit"
        )));
    }
    if value.chars().any(|character| {
        let scalar = u32::from(character);
        scalar == 0
            || scalar == 0xFFFE
            || scalar == 0xFFFF
            || (scalar < 0x20 && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(Error::InvalidFormat(format!(
            "ODB query {kind} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}

fn required_db_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<String> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid ODB attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE)
            && name.as_ref() == local
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map(std::borrow::Cow::into_owned)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODB attribute value: {error}"))
                });
        }
    }
    invalid("ODB query is missing its required name")
}

fn attribute_qname(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<Option<String>> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid ODB attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE)
            && name.as_ref() == local
        {
            return std::str::from_utf8(attribute.key.as_ref())
                .map(str::to_owned)
                .map(Some)
                .map_err(|_error| {
                    Error::InvalidFormat("ODB attribute name is not UTF-8".to_string())
                });
        }
    }
    Ok(None)
}

fn bool_text(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

fn replace_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let (_, span) = find_attribute(tag, name)?
        .ok_or_else(|| Error::InvalidFormat("ODB query attribute disappeared".to_string()))?;
    Ok(format!(
        "{}{}{}",
        &tag[..span.start],
        value,
        &tag[span.end..]
    ))
}

fn remove_attribute(tag: &str, name: &str) -> Result<String> {
    let (span, _) = find_attribute(tag, name)?
        .ok_or_else(|| Error::InvalidFormat("ODB query attribute disappeared".to_string()))?;
    Ok(format!("{}{}", &tag[..span.start], &tag[span.end..]))
}

fn insert_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let position = if tag.ends_with("/>") {
        tag.len() - 2
    } else if tag.ends_with('>') {
        tag.len() - 1
    } else {
        return invalid("ODB query start tag has no closing delimiter");
    };
    Ok(format!(
        "{} {}=\"{}\"{}",
        &tag[..position],
        name,
        value,
        &tag[position..]
    ))
}

fn find_attribute(tag: &str, wanted: &str) -> Result<Option<(Range<usize>, Range<usize>)>> {
    let bytes = tag.as_bytes();
    let mut cursor = 1usize;
    while cursor < bytes.len()
        && !bytes[cursor].is_ascii_whitespace()
        && bytes[cursor] != b'>'
        && bytes[cursor] != b'/'
    {
        cursor += 1;
    }
    while cursor < bytes.len() {
        let attribute_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] == b'>' || bytes[cursor] == b'/' {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len() && !bytes[cursor].is_ascii_whitespace() && bytes[cursor] != b'='
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] != b'=' {
            return invalid("ODB query attribute is malformed");
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes.get(cursor).ok_or_else(|| {
            Error::InvalidFormat("ODB query attribute value is missing".to_string())
        })?;
        if quote != b'\'' && quote != b'\"' {
            return invalid("ODB query attribute value is not quoted");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor == bytes.len() {
            return invalid("ODB query attribute value is unterminated");
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok(Some((attribute_start..cursor, value_start..value_end)));
        }
    }
    Ok(None)
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_owned()))
}
