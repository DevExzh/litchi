//! Typed semantic model for inert ODF document script declarations.

use litchi_core::Result;

use super::{
    MAX_LISTENER_COUNT, MAX_SCRIPT_COUNT,
    codec::{checked_text_bytes, validate_fragment, validate_required_value},
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
        for script in &self.scripts {
            validate_required_value(&script.language, "script:language")?;
            text_bytes = checked_text_bytes(text_bytes, script.language.len())?;
            text_bytes = checked_text_bytes(text_bytes, script.content_xml.len())?;
            validate_fragment(&script.content_xml, &self.namespace_declarations)?;
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
                },
                EventListener::PresentationXml(xml) => {
                    text_bytes = checked_text_bytes(text_bytes, xml.len())?;
                    validate_fragment(xml, &self.namespace_declarations)?;
                },
            }
        }
        Ok(())
    }

    /// Serialize a namespace-complete, deterministic `office:scripts` element.
    pub fn to_xml(&self) -> Result<String> {
        super::codec::write_scripts(self)
    }

    pub(super) fn total_payload_bytes(&self) -> usize {
        self.scripts
            .iter()
            .map(|script| script.language.len() + script.content_xml.len())
            .chain(self.event_listeners.iter().map(|listener| match listener {
                EventListener::Script(listener) => {
                    listener.event_name.len()
                        + listener.language.len()
                        + match &listener.binding {
                            ScriptBinding::MacroName(value) => value.len(),
                            ScriptBinding::Linked { href } => href.len(),
                        }
                },
                EventListener::PresentationXml(xml) => xml.len(),
            }))
            .sum()
    }
}
