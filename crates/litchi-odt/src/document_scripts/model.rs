//! Typed semantic model for inert ODF document script declarations.

use litchi_core::Result;
use litchi_odf_common::XmlStreamLimits;

use super::{
    MAX_LISTENER_COUNT, MAX_SCRIPT_COUNT,
    codec::{
        FragmentContainer, checked_text_bytes, emitted_namespace_count, validate_empty_script,
        validate_fragment, validate_required_value, validate_script_listener,
        validate_static_output,
    },
    invalid,
};

/// One inert `office:script` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EmbeddedScript {
    /// Required script language identifier.
    pub language: String,
    /// The exact inner XML payload. It is never interpreted or executed.
    pub content_xml: String,
}

/// The target stored by a `script:event-listener`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ScriptBinding {
    /// An inert macro name. Litchi never invokes it.
    MacroName(String),
    /// An inert linked script reference. Litchi never resolves it.
    Linked { href: String },
}

/// One typed `script:event-listener` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScriptEventListener {
    pub event_name: String,
    pub language: String,
    pub binding: ScriptBinding,
}

/// One child of the document-level `office:event-listeners` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum EventListener {
    Script(ScriptEventListener),
    /// A presentation listener preserved as inert XML for lossless round trips.
    PresentationXml(String),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct NamespaceDeclaration {
    pub(super) prefix: Option<String>,
    pub(super) uri: String,
}
/// Semantic contents of an ODF `office:scripts` element.
#[derive(Debug, Clone, Default)]
pub struct Scripts {
    pub scripts: Vec<EmbeddedScript>,
    pub event_listeners: Vec<EventListener>,
    pub(super) namespace_declarations: Vec<NamespaceDeclaration>,
}

impl PartialEq for Scripts {
    fn eq(&self, other: &Self) -> bool {
        self.scripts == other.scripts && self.event_listeners == other.event_listeners
    }
}

impl Eq for Scripts {}

impl Scripts {
    /// Validate resource limits, required values, and preserved XML fragments.
    pub fn validate(&self) -> Result<()> {
        self.validate_for_output(&XmlStreamLimits::owner_default_value())
    }

    pub(super) fn validate_for_output(&self, limits: &XmlStreamLimits) -> Result<()> {
        if self.scripts.len() > MAX_SCRIPT_COUNT {
            return invalid(format!(
                "office:scripts exceeds the {MAX_SCRIPT_COUNT} script limit"
            ));
        }
        if self.event_listeners.len() > MAX_LISTENER_COUNT {
            return invalid(format!(
                "office:scripts exceeds the {MAX_LISTENER_COUNT} event-listener limit"
            ));
        }

        let mut text_bytes = 0usize;
        validate_static_output(&self.namespace_declarations, limits)?;
        // The complete emitted fragment is parsed as one document: outer
        // office:scripts start/end plus the final EOF event.
        let mut stream_events = 3usize;
        let mut stream_attributes = emitted_namespace_count(&self.namespace_declarations);
        for script in &self.scripts {
            validate_required_value(&script.language, "script:language")?;
            text_bytes = checked_text_bytes(text_bytes, script.language.len())?;
            text_bytes = checked_text_bytes(text_bytes, script.content_xml.len())?;
            let (fragment_events, fragment_attributes) = if script.content_xml.is_empty() {
                validate_empty_script(&script.language, &self.namespace_declarations, limits)?;
                (0, 0)
            } else {
                let report = validate_fragment(
                    &script.content_xml,
                    &self.namespace_declarations,
                    limits,
                    FragmentContainer::Script {
                        language: &script.language,
                    },
                )?;
                (report.events, report.attributes)
            };
            let script_events = if script.content_xml.is_empty() {
                1
            } else {
                2usize
                    .checked_add(fragment_events)
                    .ok_or_else(|| invalid_error("office:scripts event count overflow"))?
            };
            stream_events =
                checked_stream_count(stream_events, script_events, limits.max_events(), "events")?;
            stream_attributes = checked_stream_count(
                stream_attributes,
                1usize
                    .checked_add(fragment_attributes)
                    .ok_or_else(|| invalid_error("office:scripts attribute count overflow"))?,
                limits.max_attributes(),
                "attributes",
            )?;
        }

        for listener in &self.event_listeners {
            match listener {
                EventListener::Script(listener) => {
                    validate_required_value(&listener.event_name, "script:event-name")?;
                    validate_required_value(&listener.language, "script:language")?;
                    text_bytes = checked_text_bytes(text_bytes, listener.event_name.len())?;
                    text_bytes = checked_text_bytes(text_bytes, listener.language.len())?;
                    let value = match &listener.binding {
                        ScriptBinding::MacroName(value) => value,
                        ScriptBinding::Linked { href } => href,
                    };
                    validate_required_value(value, "script event target")?;
                    text_bytes = checked_text_bytes(text_bytes, value.len())?;
                    let attributes = match &listener.binding {
                        ScriptBinding::MacroName(_) => 3,
                        ScriptBinding::Linked { .. } => 5,
                    };
                    validate_script_listener(listener, &self.namespace_declarations, limits)?;
                    stream_events =
                        checked_stream_count(stream_events, 1, limits.max_events(), "events")?;
                    stream_attributes = checked_stream_count(
                        stream_attributes,
                        attributes,
                        limits.max_attributes(),
                        "attributes",
                    )?;
                },
                EventListener::PresentationXml(xml) => {
                    text_bytes = checked_text_bytes(text_bytes, xml.len())?;
                    let report = validate_fragment(
                        xml,
                        &self.namespace_declarations,
                        limits,
                        FragmentContainer::Presentation,
                    )?;
                    let (fragment_events, fragment_attributes) = (report.events, report.attributes);
                    stream_events = checked_stream_count(
                        stream_events,
                        fragment_events,
                        limits.max_events(),
                        "events",
                    )?;
                    stream_attributes = checked_stream_count(
                        stream_attributes,
                        fragment_attributes,
                        limits.max_attributes(),
                        "attributes",
                    )?;
                },
            }
        }
        if !self.event_listeners.is_empty() {
            stream_events = checked_stream_count(stream_events, 2, limits.max_events(), "events")?;
        }
        if stream_events > limits.max_events() || stream_attributes > limits.max_attributes() {
            return invalid("office:scripts exceeds the configured XML stream limits");
        }
        Ok(())
    }

    /// Serialize a namespace-complete, deterministic `office:scripts` element.
    pub fn to_xml(&self, limits: &XmlStreamLimits) -> Result<String> {
        super::codec::write_scripts(self, limits)
    }
}

fn checked_stream_count(
    current: usize,
    additional: usize,
    limit: usize,
    dimension: &str,
) -> Result<usize> {
    let observed = current
        .checked_add(additional)
        .ok_or_else(|| invalid_error("office:scripts stream count overflow"))?;
    if observed > limit {
        return invalid(format!(
            "office:scripts exceeds the configured XML {dimension} limit"
        ));
    }
    Ok(observed)
}

fn invalid_error(message: &str) -> litchi_core::Error {
    litchi_core::Error::InvalidFormat(message.to_string())
}
