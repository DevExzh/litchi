//! Source-checked, durable protection XML transactions.

use std::collections::BTreeSet;

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use litchi_core::{Error, Result};
use serde_json::{Value, json};

use super::codec;
use super::model::Policy;
use super::{Key, Kind, invalid};

/// A staged protection-policy edit against one immutable XML source.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Vec<u8>,
    kind: Kind,
    before: Policy,
    current: Policy,
}

impl Transaction {
    /// Start an edit against package `settings.xml` bytes.
    pub(crate) fn package(source: impl AsRef<[u8]>) -> Result<Self> {
        Self::new(source.as_ref(), Kind::Package)
    }

    /// Start an edit against a flat `OpenDocument` XML source.
    pub(crate) fn flat(source: impl AsRef<[u8]>) -> Result<Self> {
        Self::new(source.as_ref(), Kind::Flat)
    }

    fn new(source: &[u8], kind: Kind) -> Result<Self> {
        let before = codec::parse(source, kind)?;
        Ok(Self {
            source: source.to_vec(),
            kind,
            current: before.clone(),
            before,
        })
    }

    /// Return the policy currently staged in this transaction.
    pub fn policy(&self) -> &Policy {
        &self.current
    }

    /// Replace the complete staged policy after validation.
    pub fn set(&mut self, policy: Policy) -> Result<()> {
        policy.validate()?;
        self.current = policy;
        Ok(())
    }

    /// Stage a form-protection toggle or clear it with `None`.
    pub fn set_forms(&mut self, value: Option<bool>) {
        self.current.forms = value;
    }

    /// Stage a bookmark-protection toggle or clear it with `None`.
    pub fn set_bookmarks(&mut self, value: Option<bool>) {
        self.current.bookmarks = value;
    }

    /// Stage a read-only loading hint or clear it with `None`.
    pub fn set_read_only(&mut self, value: Option<bool>) {
        self.current.read_only = value;
    }

    /// Stage tracked-change protection digest material or clear it with `None`.
    pub fn set_redline_key(&mut self, value: Option<Key>) -> Result<()> {
        if let Some(key) = &value {
            Key::new(key.as_bytes().to_vec())?;
        }
        self.current.redline_key = value;
        Ok(())
    }

    /// Commit the staged edit without mutating the source bytes.
    pub fn commit(self) -> Result<Commit> {
        let xml = codec::rewrite(&self.source, self.kind, &self.before, &self.current)?;
        let after = codec::parse(&xml, self.kind)?;
        if after != self.current {
            return invalid("protection transaction did not round-trip its staged policy");
        }
        Ok(Commit {
            patch: Patch {
                source: self.source.clone(),
                target: xml.clone(),
                kind: self.kind,
                before: self.before.clone(),
                after: after.clone(),
            },
            source: self.source,
            xml,
            before: self.before,
            after,
        })
    }
}

/// The immutable result of a successful protection transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    patch: Patch,
    source: Vec<u8>,
    xml: Vec<u8>,
    before: Policy,
    after: Policy,
}

impl Commit {
    /// Return the exact-source, reversible semantic patch produced by the edit.
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Return the source policy observed when the transaction began.
    pub fn before(&self) -> &Policy {
        &self.before
    }

    /// Return the policy represented by the committed XML.
    pub fn after(&self) -> &Policy {
        &self.after
    }

    /// Borrow the committed XML bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return committed XML bytes, consuming the commit.
    pub fn into_bytes(self) -> Vec<u8> {
        self.xml
    }

    /// Whether the transaction left the XML source byte-for-byte unchanged.
    pub fn is_unchanged(&self) -> bool {
        self.source == self.xml
    }
}

/// One semantic protection field that can conflict during a three-way merge.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum Field {
    /// Classic form-control protection.
    Forms,
    /// Bookmark protection.
    Bookmarks,
    /// Producer read-only loading hint.
    ReadOnly,
    /// Inert tracked-change digest material.
    RedlineKey,
}

/// Which branch supplies a conflicting value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Resolution {
    /// Keep the left branch value.
    Left,
    /// Keep the right branch value.
    Right,
}

/// An exact-source, reversible protection patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: Vec<u8>,
    target: Vec<u8>,
    kind: Kind,
    before: Policy,
    after: Policy,
}

impl Patch {
    /// Whether the patch changes its XML source.
    pub fn is_empty(&self) -> bool {
        self.source == self.target
    }

    /// Whether the patch applies to these exact source bytes.
    pub fn is_applicable_to(&self, source: impl AsRef<[u8]>) -> bool {
        self.source == source.as_ref()
    }

    /// Apply the patch after checking its complete source artifact.
    pub fn apply(&self, source: impl AsRef<[u8]>) -> Result<Commit> {
        if !self.is_applicable_to(&source) {
            return invalid("protection patch source snapshot does not match");
        }
        let before = codec::parse(source.as_ref(), self.kind)?;
        if before != self.before {
            return invalid("protection patch source policy does not match");
        }
        let after = codec::parse(&self.target, self.kind)?;
        if after != self.after {
            return invalid("protection patch target policy does not match");
        }
        Ok(Commit {
            patch: self.clone(),
            source: self.source.clone(),
            xml: self.target.clone(),
            before,
            after,
        })
    }

    /// Return the patch that restores the exact accepted XML source.
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            kind: self.kind,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Convert this patch into a bounded cross-process representation.
    pub fn durable(&self) -> Result<DurablePatch> {
        DurablePatch::new(self.clone())
    }

    /// Prepare the final policy for semantic replay onto another XML owner.
    pub fn transfer(&self) -> Transfer {
        Transfer {
            policy: self.after.clone(),
        }
    }

    /// Build a field-aware three-way merge plan for two patches from one source.
    pub fn merge(left: &Self, right: &Self) -> Result<MergePlan> {
        if left.kind != right.kind || left.source != right.source || left.before != right.before {
            return invalid("protection merge branches do not share one source");
        }
        let mut after = left.before.clone();
        let mut conflicts = BTreeSet::new();
        merge_value(
            Field::Forms,
            &left.before.forms,
            &left.after.forms,
            &right.after.forms,
            &mut after.forms,
            &mut conflicts,
        );
        merge_value(
            Field::Bookmarks,
            &left.before.bookmarks,
            &left.after.bookmarks,
            &right.after.bookmarks,
            &mut after.bookmarks,
            &mut conflicts,
        );
        merge_value(
            Field::ReadOnly,
            &left.before.read_only,
            &left.after.read_only,
            &right.after.read_only,
            &mut after.read_only,
            &mut conflicts,
        );
        merge_value(
            Field::RedlineKey,
            &left.before.redline_key,
            &left.after.redline_key,
            &right.after.redline_key,
            &mut after.redline_key,
            &mut conflicts,
        );
        Ok(MergePlan {
            source: left.source.clone(),
            kind: left.kind,
            before: left.before.clone(),
            left: left.after.clone(),
            right: right.after.clone(),
            after,
            conflicts,
        })
    }
}

/// A detached final protection policy for cross-document replay.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Transfer {
    policy: Policy,
}

impl Transfer {
    /// Borrow the inert policy retained by this transfer.
    pub fn policy(&self) -> &Policy {
        &self.policy
    }

    /// Apply the retained policy to package `settings.xml` bytes.
    pub fn apply_package(&self, source: impl AsRef<[u8]>) -> Result<Commit> {
        let mut transaction = Transaction::package(source)?;
        transaction.set(self.policy.clone())?;
        transaction.commit()
    }

    /// Apply the retained policy to a flat `OpenDocument` XML source.
    pub fn apply_flat(&self, source: impl AsRef<[u8]>) -> Result<Commit> {
        let mut transaction = Transaction::flat(source)?;
        transaction.set(self.policy.clone())?;
        transaction.commit()
    }
}

/// A field-aware protection merge with explicit conflict resolution.
#[derive(Clone, Debug)]
pub struct MergePlan {
    source: Vec<u8>,
    kind: Kind,
    before: Policy,
    left: Policy,
    right: Policy,
    after: Policy,
    conflicts: BTreeSet<Field>,
}

impl MergePlan {
    /// Return unresolved semantic fields in deterministic order.
    pub fn conflicts(&self) -> impl ExactSizeIterator<Item = Field> + '_ {
        self.conflicts.iter().copied()
    }

    /// Resolve one conflicting field from a named branch.
    pub fn resolve(&mut self, field: Field, resolution: Resolution) -> Result<()> {
        if !self.conflicts.remove(&field) {
            return invalid("protection merge field is not an unresolved conflict");
        }
        let source = match resolution {
            Resolution::Left => &self.left,
            Resolution::Right => &self.right,
        };
        match field {
            Field::Forms => self.after.forms = source.forms,
            Field::Bookmarks => self.after.bookmarks = source.bookmarks,
            Field::ReadOnly => self.after.read_only = source.read_only,
            Field::RedlineKey => self.after.redline_key.clone_from(&source.redline_key),
        }
        Ok(())
    }

    /// Materialize the merged exact-source patch after all conflicts are resolved.
    pub fn finish(self) -> Result<Patch> {
        if !self.conflicts.is_empty() {
            return invalid("protection merge still has unresolved conflicts");
        }
        let target = codec::rewrite(&self.source, self.kind, &self.before, &self.after)?;
        Ok(Patch {
            source: self.source,
            target,
            kind: self.kind,
            before: self.before,
            after: self.after,
        })
    }
}

fn merge_value<T: Clone + Eq>(
    field: Field,
    base: &T,
    left: &T,
    right: &T,
    output: &mut T,
    conflicts: &mut BTreeSet<Field>,
) {
    if left == right || right == base {
        output.clone_from(left);
    } else if left == base {
        output.clone_from(right);
    } else {
        conflicts.insert(field);
    }
}

const DURABLE_FORMAT: &str = "litchi.odt.protection.v1";
const MAX_DURABLE_BYTES: usize = 192 * 1024 * 1024;

/// Bounded deterministic-JSON representation of a protection patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DurablePatch {
    patch: Patch,
}

impl DurablePatch {
    fn new(patch: Patch) -> Result<Self> {
        validate_patch(&patch)?;
        Ok(Self { patch })
    }

    /// Parse and validate a canonical durable protection patch.
    pub fn from_deterministic_json(bytes: &[u8]) -> Result<Self> {
        if bytes.len() > MAX_DURABLE_BYTES {
            return invalid("durable protection patch exceeds its byte limit");
        }
        let value: Value = serde_json::from_slice(bytes)
            .map_err(|error| Error::InvalidFormat(format!("invalid protection patch: {error}")))?;
        let patch = patch_from_value(&value)?;
        let durable = Self::new(patch)?;
        if durable.to_deterministic_json()? != bytes {
            return invalid("durable protection patch is not canonical JSON");
        }
        Ok(durable)
    }

    /// Serialize this semantic patch as canonical deterministic JSON.
    pub fn to_deterministic_json(&self) -> Result<Vec<u8>> {
        let bytes = serde_json::to_vec(&patch_value(&self.patch)).map_err(|error| {
            Error::InvalidFormat(format!("protection patch write failed: {error}"))
        })?;
        if bytes.len() > MAX_DURABLE_BYTES {
            return invalid("durable protection patch exceeds its byte limit");
        }
        Ok(bytes)
    }

    /// Apply after checking the complete source XML artifact.
    pub fn apply(&self, source: impl AsRef<[u8]>) -> Result<Commit> {
        self.patch.apply(source)
    }

    /// Return the durable patch that restores the exact source XML.
    pub fn inverse(&self) -> Self {
        Self {
            patch: self.patch.inverse(),
        }
    }
}

fn validate_patch(patch: &Patch) -> Result<()> {
    if patch.source.len() > super::MAX_XML_BYTES || patch.target.len() > super::MAX_XML_BYTES {
        return invalid("protection patch XML exceeds its byte limit");
    }
    let before = codec::parse(&patch.source, patch.kind)?;
    let after = codec::parse(&patch.target, patch.kind)?;
    if before != patch.before || after != patch.after {
        return invalid("protection patch semantic payload does not match its XML");
    }
    if codec::rewrite(&patch.source, patch.kind, &patch.before, &patch.after)? != patch.target {
        return invalid("protection patch target is not its semantic replay result");
    }
    Ok(())
}

fn patch_value(patch: &Patch) -> Value {
    json!({
        "after": policy_value(&patch.after),
        "before": policy_value(&patch.before),
        "format": DURABLE_FORMAT,
        "kind": match patch.kind { Kind::Flat => "flat", Kind::Package => "package" },
        "operation": "protection.set",
        "source": BASE64.encode(&patch.source),
        "target": BASE64.encode(&patch.target),
    })
}

fn patch_from_value(value: &Value) -> Result<Patch> {
    let object = value
        .as_object()
        .ok_or_else(|| Error::InvalidFormat("protection patch must be an object".to_string()))?;
    if object.len() != 7
        || object.get("format").and_then(Value::as_str) != Some(DURABLE_FORMAT)
        || object.get("operation").and_then(Value::as_str) != Some("protection.set")
    {
        return invalid("unknown protection patch envelope");
    }
    let kind = match object.get("kind").and_then(Value::as_str) {
        Some("flat") => Kind::Flat,
        Some("package") => Kind::Package,
        _ => return invalid("invalid protection patch XML kind"),
    };
    Ok(Patch {
        source: decode_xml(object.get("source"), "source")?,
        target: decode_xml(object.get("target"), "target")?,
        kind,
        before: policy_from_value(object.get("before"))?,
        after: policy_from_value(object.get("after"))?,
    })
}

fn decode_xml(value: Option<&Value>, name: &str) -> Result<Vec<u8>> {
    let encoded = value
        .and_then(Value::as_str)
        .ok_or_else(|| Error::InvalidFormat(format!("protection patch {name} is missing")))?;
    let bytes = BASE64.decode(encoded).map_err(|error| {
        Error::InvalidFormat(format!("invalid protection patch {name}: {error}"))
    })?;
    if bytes.len() > super::MAX_XML_BYTES {
        return invalid("protection patch XML exceeds its byte limit");
    }
    Ok(bytes)
}

fn policy_value(policy: &Policy) -> Value {
    json!({
        "bookmarks": policy.bookmarks,
        "forms": policy.forms,
        "read_only": policy.read_only,
        "redline_key": policy.redline_key.as_ref().map(|key| BASE64.encode(key.as_bytes())),
    })
}

fn policy_from_value(value: Option<&Value>) -> Result<Policy> {
    let object = value
        .and_then(Value::as_object)
        .ok_or_else(|| Error::InvalidFormat("protection patch policy is missing".to_string()))?;
    if object.len() != 4 {
        return invalid("protection patch policy has unknown fields");
    }
    let redline_key = match object.get("redline_key") {
        Some(Value::Null) => None,
        Some(Value::String(value)) => Some(Key::new(BASE64.decode(value).map_err(|error| {
            Error::InvalidFormat(format!("invalid protection patch key: {error}"))
        })?)?),
        _ => return invalid("invalid protection patch redline key"),
    };
    let policy = Policy {
        forms: optional_bool(object.get("forms"), "forms")?,
        bookmarks: optional_bool(object.get("bookmarks"), "bookmarks")?,
        read_only: optional_bool(object.get("read_only"), "read_only")?,
        redline_key,
    };
    policy.validate()?;
    Ok(policy)
}

fn optional_bool(value: Option<&Value>, name: &str) -> Result<Option<bool>> {
    match value {
        Some(Value::Null) => Ok(None),
        Some(Value::Bool(value)) => Ok(Some(*value)),
        _ => invalid(format!("invalid protection patch {name}")),
    }
}
