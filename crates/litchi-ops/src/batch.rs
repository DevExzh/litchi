//! OfficeCLI-compatible command inputs for the unified batch executor.
//!
//! Bounded transport decoding is separate from [`crate::File::batch`] execution.
//! The integrated command subset uses immutable transactions and strict field
//! admission; decoding alone does not resolve or edit documents.

use std::borrow::Cow;
use std::fmt;

use serde::de::{self, Deserialize, Deserializer, MapAccess, Visitor};
use serde::ser::{Serialize, SerializeMap, Serializer};

use crate::props::Text;
use crate::{Code, Error as OpsError, Props, Result};

mod decode;
mod dump;
mod execute;
pub use decode::{Error, SourceLimitKind, SourceLimits, SourceLimitsBuilder, from_bytes};
pub use dump::Dump;
pub(crate) use dump::support as dump_support;
pub use execute::{Applied, BatchCommit, ItemResult};

use crate::transaction::EditOptions;
#[cfg(feature = "docx")]
pub(crate) use execute::{has_text_controls, inline_text_props};

/// Whether a batch stops and discards its candidate on failure or skips the
/// failing item and continues. Execution is atomic unless explicitly changed.
///
/// ```
/// use litchi_ops::BatchMode;
/// assert_eq!(BatchMode::default(), BatchMode::Atomic);
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum BatchMode {
    /// The first failure discards every change in the batch.
    #[default]
    Atomic,
    /// Keep successful items and report each failed item in input order.
    BestEffort,
}

/// What a command does with a property its target does not support.
///
/// The library verbs ([`crate::File::set`], [`crate::File::add`], closure
/// transactions) always refuse. Only the command channel
/// ([`crate::File::apply_with`], [`crate::File::batch_with`]) can choose
/// [`Unsupported::Warn`], OfficeCLI's "apply the rest and warn", so that
/// OfficeCLI scripts replay (record 0015 Q2/Δ1, record 0269).
///
/// ```
/// use litchi_ops::batch::Unsupported;
/// assert_eq!(Unsupported::default(), Unsupported::Refuse);
/// ```
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum Unsupported {
    /// Refuse the whole command with `unsupported_property`.
    #[default]
    Refuse,
    /// Skip each property the target refuses with `unsupported_property`,
    /// apply the remaining properties atomically, and report every skipped
    /// property in [`Applied::warnings`]. Any other failure (an invalid
    /// value, a missing path, a resource limit) still fails the command.
    ///
    /// A command that fails after skipping reports the failure of the
    /// reduced command with the skipped properties attached
    /// ([`OpsError::warnings`]; record 0271). The code stays the reduced
    /// command's own, because that failure is real (the command, as it can
    /// be applied, is incomplete or invalid), and the list says which
    /// skipped keys it lost: an `add` whose data came only from skipped
    /// keys fails with `missing_property` *and* names those keys, so the
    /// caller is never told it omitted something it did supply.
    Warn,
}

/// Command-channel options of [`crate::File::apply_with`] and
/// [`crate::File::batch_with`]: the batch mode, the property policy and the
/// transaction's [`EditOptions`]. The default is an atomic batch that refuses unsupported
/// properties under default limits, which is what [`crate::File::batch`]
/// does.
///
/// ```
/// use litchi_ops::BatchMode;
/// use litchi_ops::batch::{Options, Unsupported};
///
/// // OfficeCLI's batch behaviour: keep going, apply what is supported.
/// let options = Options::new()
///     .mode(BatchMode::BestEffort)
///     .unsupported(Unsupported::Warn);
/// # let _ = options;
/// ```
#[derive(Debug, Clone, Copy, Default)]
#[non_exhaustive]
pub struct Options {
    mode: BatchMode,
    unsupported: Unsupported,
    edit: EditOptions,
}

impl Options {
    /// The defaults: atomic, refusing, default limits.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the batch mode. A single command ([`crate::File::apply_with`])
    /// ignores it.
    #[must_use]
    pub const fn mode(mut self, mode: BatchMode) -> Self {
        self.mode = mode;
        self
    }

    /// Set what a command does with a property its target does not support.
    #[must_use]
    pub const fn unsupported(mut self, unsupported: Unsupported) -> Self {
        self.unsupported = unsupported;
        self
    }

    /// Set the transaction's [`EditOptions`] (its semantic limits).
    #[must_use]
    pub const fn edit_options(mut self, options: EditOptions) -> Self {
        self.edit = options;
        self
    }
}

/// One OfficeCLI command payload, borrowing unescaped text when possible.
///
/// All 21 wire keys (including the `op` alias for `command`) are accepted,
/// ASCII case-insensitively. Serialization uses canonical spellings and omits
/// absent optional fields. `props` accepts an object or a `"key=value"` array
/// through [`Props`]; explicit property nulls remain clear operations.
///
/// Unknown and repeated top-level keys are refused, including `command` plus
/// `op`. This strict policy prevents silent typo loss and order-dependent
/// alias overrides. Optional JSON nulls mean absent fields. Unknown command
/// names are retained so command dispatch can report `unsupported_command`;
/// decoding alone does not claim support for a command, dump version, raw
/// operation, numeric range, or format property.
///
/// The signed 32-bit numeric fields preserve OfficeCLI's wire representation;
/// non-integers and integers outside that range are malformed transport.
/// Negative integers are retained for command-specific range admission.
/// The executor converts positions to checked typed coordinates. Callers
/// decoding untrusted input choose the [`SourceLimits`] that [`from_bytes`]
/// requires, which bound the bytes and item count before retaining decoded
/// payloads.
/// `Debug` omits all authored values. Use [`Self::into_owned`] to detach from
/// the input buffer before storing or sending a payload.
///
/// ```
/// use litchi_ops::{BatchItem, Props};
///
/// let mut item = BatchItem::new("set");
/// item.path = Some("/body/p[1]/r[1]".into());
/// item.props = Some(Props::new().with("font.bold", true));
/// item.validate()?;
/// let owned: BatchItem<'static> = item.into_owned();
/// assert_eq!(owned.command, "set");
/// # Ok::<(), litchi_ops::Error>(())
/// ```
#[derive(Clone, Default, PartialEq)]
#[non_exhaustive]
pub struct BatchItem<'a> {
    /// Command name. `op` is accepted as a wire alias, never emitted.
    pub command: Cow<'a, str>,
    /// Element path or mutation target.
    pub path: Option<Cow<'a, str>>,
    /// Second element of a swap.
    pub path2: Option<Cow<'a, str>>,
    /// Parent of a new element.
    pub parent: Option<Cow<'a, str>>,
    /// Format-specific element type to add.
    pub kind: Option<Cow<'a, str>>,
    /// Source of a move or copy.
    pub from: Option<Cow<'a, str>>,
    /// Destination of a move, or the legacy second swap path.
    pub to: Option<Cow<'a, str>>,
    /// Zero-based wire position; checked when the command is admitted.
    pub index: Option<i32>,
    /// Insert immediately after this element.
    pub after: Option<Cow<'a, str>>,
    /// Insert immediately before this element.
    pub before: Option<Cow<'a, str>>,
    /// Command-specific mode.
    pub mode: Option<Cow<'a, str>>,
    /// Requested projection depth; command admission interprets its range.
    pub depth: Option<i32>,
    /// Part addressed by a raw command.
    pub part: Option<Cow<'a, str>>,
    /// Raw XML selector.
    pub xpath: Option<Cow<'a, str>>,
    /// Raw edit action.
    pub action: Option<Cow<'a, str>>,
    /// Raw replacement XML.
    pub xml: Option<Cow<'a, str>>,
    /// Ordered property assignments, preserving scalar types and repeated keys.
    pub props: Option<Props<'a>>,
    /// Scoped selector for a command.
    pub selector: Option<Cow<'a, str>>,
    /// Command-specific text.
    pub text: Option<Cow<'a, str>>,
    /// Dump newline-semantics version; retained without implicit text rewriting.
    pub dump_version: Option<i32>,
}

impl<'a> BatchItem<'a> {
    /// Construct a payload with this command and no optional fields.
    #[must_use]
    pub fn new(command: impl Into<Cow<'a, str>>) -> Self {
        Self {
            command: command.into(),
            ..Self::default()
        }
    }

    /// Validate the payload envelope without resolving or executing it.
    ///
    /// # Errors
    /// Returns [`Code::CorruptBatch`] if the required command is empty or
    /// whitespace-only. Command-specific requirements belong to dispatch.
    pub fn validate(&self) -> Result<()> {
        if self.command.trim().is_empty() {
            return Err(OpsError::new(
                Code::CorruptBatch,
                "a batch item requires a command",
            ));
        }
        Ok(())
    }

    /// Detach all borrowed values, preserving ordered property assignments.
    #[must_use]
    pub fn into_owned(self) -> BatchItem<'static> {
        BatchItem {
            command: Cow::Owned(self.command.into_owned()),
            path: self.path.map(|value| Cow::Owned(value.into_owned())),
            path2: self.path2.map(|value| Cow::Owned(value.into_owned())),
            parent: self.parent.map(|value| Cow::Owned(value.into_owned())),
            kind: self.kind.map(|value| Cow::Owned(value.into_owned())),
            from: self.from.map(|value| Cow::Owned(value.into_owned())),
            to: self.to.map(|value| Cow::Owned(value.into_owned())),
            index: self.index,
            after: self.after.map(|value| Cow::Owned(value.into_owned())),
            before: self.before.map(|value| Cow::Owned(value.into_owned())),
            mode: self.mode.map(|value| Cow::Owned(value.into_owned())),
            depth: self.depth,
            part: self.part.map(|value| Cow::Owned(value.into_owned())),
            xpath: self.xpath.map(|value| Cow::Owned(value.into_owned())),
            action: self.action.map(|value| Cow::Owned(value.into_owned())),
            xml: self.xml.map(|value| Cow::Owned(value.into_owned())),
            props: self.props.map(Props::into_owned),
            selector: self.selector.map(|value| Cow::Owned(value.into_owned())),
            text: self.text.map(|value| Cow::Owned(value.into_owned())),
            dump_version: self.dump_version,
        }
    }
}

impl<'a> From<&BatchItem<'a>> for BatchItem<'a> {
    fn from(item: &BatchItem<'a>) -> Self {
        item.clone()
    }
}

impl fmt::Debug for BatchItem<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("BatchItem").finish_non_exhaustive()
    }
}

impl Serialize for BatchItem<'_> {
    fn serialize<S: Serializer>(&self, serializer: S) -> std::result::Result<S::Ok, S::Error> {
        self.validate()
            .map_err(|_| serde::ser::Error::custom("a batch item requires a command"))?;
        let mut map = serializer.serialize_map(None)?;
        map.serialize_entry("command", &self.command)?;
        if let Some(value) = &self.path {
            map.serialize_entry("path", value)?;
        }
        if let Some(value) = &self.path2 {
            map.serialize_entry("path2", value)?;
        }
        if let Some(value) = &self.parent {
            map.serialize_entry("parent", value)?;
        }
        if let Some(value) = &self.kind {
            map.serialize_entry("type", value)?;
        }
        if let Some(value) = &self.from {
            map.serialize_entry("from", value)?;
        }
        if let Some(value) = &self.to {
            map.serialize_entry("to", value)?;
        }
        if let Some(value) = &self.index {
            map.serialize_entry("index", value)?;
        }
        if let Some(value) = &self.after {
            map.serialize_entry("after", value)?;
        }
        if let Some(value) = &self.before {
            map.serialize_entry("before", value)?;
        }
        if let Some(value) = &self.mode {
            map.serialize_entry("mode", value)?;
        }
        if let Some(value) = &self.depth {
            map.serialize_entry("depth", value)?;
        }
        if let Some(value) = &self.part {
            map.serialize_entry("part", value)?;
        }
        if let Some(value) = &self.xpath {
            map.serialize_entry("xpath", value)?;
        }
        if let Some(value) = &self.action {
            map.serialize_entry("action", value)?;
        }
        if let Some(value) = &self.xml {
            map.serialize_entry("xml", value)?;
        }
        if let Some(value) = &self.props {
            map.serialize_entry("props", value)?;
        }
        if let Some(value) = &self.selector {
            map.serialize_entry("selector", value)?;
        }
        if let Some(value) = &self.text {
            map.serialize_entry("text", value)?;
        }
        if let Some(value) = &self.dump_version {
            map.serialize_entry("dumpVersion", value)?;
        }
        map.end()
    }
}

impl<'de> Deserialize<'de> for BatchItem<'de> {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> std::result::Result<Self, D::Error> {
        // A decoder's unexpected-value diagnostics can contain authored text.
        // Keep this public payload boundary content-free even for wrong types,
        // malformed escapes, and nested property failures.
        deserializer
            .deserialize_map(ItemVisitor)
            .map_err(|_| de::Error::custom("invalid batch item envelope"))
    }
}

struct ItemVisitor;

impl<'de> Visitor<'de> for ItemVisitor {
    type Value = BatchItem<'de>;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a batch item object")
    }

    fn visit_map<A: MapAccess<'de>>(
        self,
        mut map: A,
    ) -> std::result::Result<Self::Value, A::Error> {
        let mut item = BatchItem::default();
        let mut seen = 0_u32;
        while let Some(Text(key)) = map.next_key::<Text<'de>>()? {
            let field =
                field_index(&key).ok_or_else(|| de::Error::custom("unknown batch item field"))?;
            let bit = 1_u32 << field;
            if seen & bit != 0 {
                return Err(de::Error::custom("duplicate batch item field"));
            }
            seen |= bit;
            match field {
                0 => item.command = map.next_value::<Text<'de>>()?.0,
                1 => item.path = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                2 => item.path2 = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                3 => item.parent = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                4 => item.kind = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                5 => item.from = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                6 => item.to = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                7 => item.index = map.next_value()?,
                8 => item.after = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                9 => item.before = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                10 => item.mode = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                11 => item.depth = map.next_value()?,
                12 => item.part = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                13 => item.xpath = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                14 => item.action = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                15 => item.xml = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                16 => item.props = map.next_value()?,
                17 => item.selector = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                18 => item.text = map.next_value::<Option<Text<'de>>>()?.map(|value| value.0),
                19 => item.dump_version = map.next_value()?,
                _ => return Err(de::Error::custom("unregistered batch item field")),
            }
        }
        item.validate()
            .map_err(|_| de::Error::custom("a batch item requires a command"))?;
        Ok(item)
    }
}

fn field_index(key: &str) -> Option<usize> {
    // Small static transport vocabulary; no per-item map or lowercase buffer.
    if key.eq_ignore_ascii_case("op") {
        return Some(0);
    }
    const FIELDS: &[&str] = &[
        "command",
        "path",
        "path2",
        "parent",
        "type",
        "from",
        "to",
        "index",
        "after",
        "before",
        "mode",
        "depth",
        "part",
        "xpath",
        "action",
        "xml",
        "props",
        "selector",
        "text",
        "dumpVersion",
    ];
    FIELDS
        .iter()
        .position(|field| key.eq_ignore_ascii_case(field))
}

#[cfg(test)]
mod tests;
