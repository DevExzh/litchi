use std::{
    collections::HashMap,
    convert::Infallible,
    io::Write,
    path::Path,
    sync::{Arc, OnceLock},
};

use super::error::check;
use super::validation::{decode, encode_name};
use super::{
    Change, EditOptions, Error, Flags, LimitKind, Limits, Metadata, Name, OpenOptions, Result,
};
use crate::{
    Record,
    prop::{Id, Prop, Props, Value},
};

pub use litchi_core::edit::Diagnostics;
pub type Commit<'data> = litchi_core::edit::Commit<Package<'data>, Patch<'data>>;

struct Parsed<'data> {
    record: Record<'data>,
    properties: Props<'data>,
}

impl<'data> Parsed<'data> {
    fn parse(record: &Record<'data>, limits: Limits) -> Result<Self> {
        check(
            LimitKind::PropertyBytes,
            record.data().len(),
            limits.property_bytes_value(),
        )?;
        check(
            LimitKind::Properties,
            usize::from(record.instance()),
            limits.properties_value(),
        )?;
        let properties = Props::parse(record, record.limits())?;
        decode(&properties, limits)?;
        Ok(Self {
            record: record.clone(),
            properties,
        })
    }
}

/// A read-only borrowed projection of one admitted picture property record.
pub struct View<'data> {
    inner: ViewInner<'data>,
    limits: Limits,
}

enum ViewInner<'data> {
    Parsed(Parsed<'data>),
}

impl<'data> View<'data> {
    /// Admits a raw Opt-family record under explicitly supplied limits.
    pub fn parse(record: &Record<'data>, limits: Limits) -> Result<Self> {
        Ok(Self {
            inner: ViewInner::Parsed(Parsed::parse(record, limits)?),
            limits,
        })
    }

    fn parsed(&self) -> &Parsed<'data> {
        match &self.inner {
            ViewInner::Parsed(parsed) => parsed,
        }
    }

    #[must_use]
    pub fn record(&self) -> &Record<'data> {
        &self.parsed().record
    }

    #[must_use]
    pub fn properties(&self) -> &Props<'data> {
        &self.parsed().properties
    }

    pub fn metadata(&self) -> Result<Option<Metadata<'data>>> {
        decode(self.properties(), self.limits)
    }
}

#[derive(Clone)]
struct Entry {
    opid: u16,
    value: i32,
    complex: Option<Arc<[u8]>>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Header {
    kind: u16,
    version: u8,
    count: u16,
    body_len: u32,
}

impl Header {
    fn bytes(self) -> [u8; 8] {
        let mut bytes = [0; 8];
        bytes[..2].copy_from_slice(&((self.count << 4) | u16::from(self.version)).to_le_bytes());
        bytes[2..4].copy_from_slice(&self.kind.to_le_bytes());
        bytes[4..].copy_from_slice(&self.body_len.to_le_bytes());
        bytes
    }
}

// Wire maxima fit in these fields; public options retain ergonomic usize values.
#[derive(Clone, Copy, PartialEq, Eq)]
struct Policy {
    bytes: u32,
    units: u32,
    properties: u16,
}
impl Policy {
    #[allow(
        clippy::expect_used,
        reason = "public Limits constructors enforce the wire maxima"
    )]
    fn from_limits(limits: Limits) -> Self {
        Self {
            bytes: u32::try_from(limits.property_bytes_value()).expect("validated wire byte limit"),
            units: u32::try_from(limits.name_units_value()).expect("validated wire name limit"),
            properties: u16::try_from(limits.properties_value())
                .expect("validated wire count limit"),
        }
    }
    fn limits(self) -> Limits {
        Limits::from_admitted(
            self.bytes as usize,
            usize::from(self.properties),
            self.units as usize,
        )
    }
}

struct Base<'data> {
    policy: Policy,
    data: BaseData<'data>,
}
enum BaseData<'data> {
    Borrowed {
        record: Record<'data>,
        properties: Option<Box<Props<'data>>>,
    },
    Owned(Box<OwnedBase>),
}
struct OwnedBase {
    header: Header,
    entries: Vec<Entry>,
    name: Option<usize>,
    flags: Option<usize>,
    #[allow(
        clippy::box_collection,
        reason = "keep the unused contiguous-view cache small; allocation occurs only on explicit view"
    )]
    bytes: OnceLock<Box<Vec<u8>>>,
}
struct Deferred<'data> {
    base: Arc<Base<'data>>,
    // None inherits the original descriptor; Some(None) clears its value.
    name: Option<Option<Arc<[u8]>>>,
    flags: Option<u32>,
    name_last: bool,
    policy: Policy,
    #[allow(
        clippy::box_collection,
        reason = "the out-of-line view cache saves retained bytes on every unrendered commit"
    )]
    bytes: OnceLock<Box<Vec<u8>>>,
}
#[derive(Clone)]
enum State<'data> {
    Opened(Arc<Base<'data>>),
    Edited(Arc<Deferred<'data>>),
}
#[derive(Clone, Copy, PartialEq, Eq)]
struct EntryRef<'data> {
    opid: u16,
    value: i32,
    complex: Option<&'data [u8]>,
}
impl<'data> EntryRef<'data> {
    fn from_prop(prop: &Prop<'data>) -> Self {
        Self {
            opid: prop.raw_opid(),
            value: prop.raw_value(),
            complex: match prop.value() {
                Value::Simple(_) => None,
                Value::Complex(bytes) => Some(bytes),
                Value::Array(array) => Some(array.raw_data()),
            },
        }
    }
    fn from_entry(entry: &'data Entry) -> Self {
        Self {
            opid: entry.opid,
            value: entry.value,
            complex: entry.complex.as_deref(),
        }
    }
    #[allow(
        clippy::expect_used,
        reason = "only admitted UTF-16 payloads reach this encoder"
    )]
    fn with_name(self, name: Option<&'data [u8]>) -> Self {
        Self {
            opid: (self.opid & !0x8000) | if name.is_some() { 0x8000 } else { 0 },
            value: name.map_or(0, |bytes| {
                i32::try_from(bytes.len()).expect("admitted name length")
            }),
            complex: name,
        }
    }
}
enum BaseEntries<'data> {
    Parsed(std::slice::Iter<'data, Prop<'data>>),
    Owned(std::slice::Iter<'data, Entry>),
    Raw {
        data: &'data [u8],
        count: usize,
        index: usize,
        complex_offset: usize,
    },
}
impl<'data> Iterator for BaseEntries<'data> {
    type Item = EntryRef<'data>;
    fn next(&mut self) -> Option<Self::Item> {
        match self {
            Self::Parsed(entries) => entries.next().map(EntryRef::from_prop),
            Self::Owned(entries) => entries.next().map(EntryRef::from_entry),
            Self::Raw {
                data,
                count,
                index,
                complex_offset,
            } => {
                if *index == *count {
                    return None;
                }
                let offset = *index * 6;
                *index += 1;
                let opid = u16::from_le_bytes([data[offset], data[offset + 1]]);
                let value = i32::from_le_bytes([
                    data[offset + 2],
                    data[offset + 3],
                    data[offset + 4],
                    data[offset + 5],
                ]);
                let complex = if opid & 0x8000 != 0 {
                    let end = *complex_offset + value.cast_unsigned() as usize;
                    let bytes = &data[*complex_offset..end];
                    *complex_offset = end;
                    Some(bytes)
                } else {
                    None
                };
                Some(EntryRef {
                    opid,
                    value,
                    complex,
                })
            },
        }
    }
}
impl Base<'_> {
    fn header(&self) -> Header {
        match &self.data {
            BaseData::Borrowed { record, .. } => Header {
                kind: record.raw_kind(),
                version: record.version(),
                count: record.instance(),
                body_len: record.len(),
            },
            BaseData::Owned(owned) => owned.header,
        }
    }
    fn entries(&self) -> BaseEntries<'_> {
        match &self.data {
            BaseData::Borrowed {
                properties: Some(properties),
                ..
            } => BaseEntries::Parsed(properties.properties.iter()),
            BaseData::Borrowed {
                record,
                properties: None,
            } => {
                let count = usize::from(record.instance());
                BaseEntries::Raw {
                    data: record.data(),
                    count,
                    index: 0,
                    complex_offset: count * 6,
                }
            },
            BaseData::Owned(owned) => BaseEntries::Owned(owned.entries.iter()),
        }
    }
    fn selected(&self, id: Id) -> Option<EntryRef<'_>> {
        match &self.data {
            BaseData::Borrowed {
                properties: Some(properties),
                ..
            } => properties.prop(id).map(EntryRef::from_prop),
            BaseData::Borrowed {
                properties: None, ..
            } => self.entries().find(|entry| entry.opid & 0x3fff == id.raw()),
            BaseData::Owned(owned) => {
                let index = if id == Id::PictureFileName {
                    owned.name
                } else {
                    owned.flags
                };
                index
                    .and_then(|index| owned.entries.get(index))
                    .map(EntryRef::from_entry)
            },
        }
    }
}
struct Entries<'data> {
    base: BaseEntries<'data>,
    name: Option<Option<&'data [u8]>>,
    flags: Option<u32>,
    name_last: bool,
}
impl<'data> Iterator for Entries<'data> {
    type Item = EntryRef<'data>;
    fn next(&mut self) -> Option<Self::Item> {
        if let Some(mut entry) = self.base.next() {
            if entry.opid & 0x3fff == Id::PictureFileName.raw() {
                if let Some(name) = self.name.take() {
                    entry = entry.with_name(name);
                }
            } else if entry.opid & 0x3fff == Id::BlipFlags.raw() {
                if let Some(flags) = self.flags.take() {
                    entry.value = flags.cast_signed();
                }
            }
            return Some(entry);
        }
        if self.name_last {
            if let Some(flags) = self.flags.take() {
                return Some(EntryRef {
                    opid: Id::BlipFlags.raw(),
                    value: flags.cast_signed(),
                    complex: None,
                });
            }
        }
        if let Some(name) = self.name.take() {
            return Some(
                EntryRef {
                    opid: Id::PictureFileName.raw(),
                    value: 0,
                    complex: None,
                }
                .with_name(name),
            );
        }
        self.flags.take().map(|flags| EntryRef {
            opid: Id::BlipFlags.raw(),
            value: flags.cast_signed(),
            complex: None,
        })
    }
}
impl<'data> State<'data> {
    fn base(&self) -> &Arc<Base<'data>> {
        match self {
            Self::Opened(base) => base,
            Self::Edited(edited) => &edited.base,
        }
    }
    fn name_override(&self) -> Option<&Option<Arc<[u8]>>> {
        match self {
            Self::Opened(_) => None,
            Self::Edited(edited) => edited.name.as_ref(),
        }
    }
    fn flags_override(&self) -> Option<u32> {
        match self {
            Self::Opened(_) => None,
            Self::Edited(edited) => edited.flags,
        }
    }
    fn name_last(&self) -> bool {
        match self {
            Self::Opened(_) => false,
            Self::Edited(edited) => edited.name_last,
        }
    }
    fn policy(&self) -> Policy {
        match self {
            Self::Opened(base) => base.policy,
            Self::Edited(edited) => edited.policy,
        }
    }
    fn name_bytes(&self) -> Option<&[u8]> {
        match self.name_override() {
            Some(name) => name.as_deref(),
            None => self
                .base()
                .selected(Id::PictureFileName)
                .and_then(|entry| entry.complex),
        }
    }
    fn raw_flags(&self) -> Option<u32> {
        self.flags_override().or_else(|| {
            self.base()
                .selected(Id::BlipFlags)
                .map(|entry| entry.value.cast_unsigned())
        })
    }
    #[allow(
        clippy::expect_used,
        reason = "commit validates record extent before constructing a deferred state"
    )]
    fn header(&self) -> Header {
        let mut header = self.base().header();
        if let Some(name) = self.name_override() {
            let old = self.base().selected(Id::PictureFileName);
            if old.is_none() {
                header.count += 1;
                header.body_len += 6;
            }
            let old_len = old.and_then(|entry| entry.complex).map_or(0, <[u8]>::len);
            let new_len = name.as_ref().map_or(0, |bytes| bytes.len());
            header.body_len = u32::try_from(header.body_len as usize - old_len + new_len)
                .expect("admitted record length");
        }
        if self.flags_override().is_some() && self.base().selected(Id::BlipFlags).is_none() {
            header.count += 1;
            header.body_len += 6;
        }
        header
    }
    fn entries(&self) -> Entries<'_> {
        Entries {
            base: self.base().entries(),
            name: self.name_override().map(|name| name.as_deref()),
            flags: self.flags_override(),
            name_last: self.name_last(),
        }
    }
    fn same_content(&self, other: &Self) -> bool {
        Arc::ptr_eq(self.base(), other.base())
            && self.flags_override() == other.flags_override()
            && self.name_last() == other.name_last()
            && match (self.name_override(), other.name_override()) {
                (None, None) | (Some(None), Some(None)) => true,
                (Some(Some(a)), Some(Some(b))) => Arc::ptr_eq(a, b),
                _ => false,
            }
    }
    fn with_policy(&self, policy: Policy) -> Self {
        if self.policy() == policy {
            return self.clone();
        }
        Self::Edited(Arc::new(Deferred {
            base: Arc::clone(self.base()),
            name: self.name_override().cloned(),
            flags: self.flags_override(),
            name_last: self.name_last(),
            policy,
            bytes: OnceLock::new(),
        }))
    }
    fn emit<E>(
        &self,
        write: impl FnMut(&[u8]) -> std::result::Result<(), E>,
    ) -> std::result::Result<(), E> {
        let mut output = Output {
            write,
            bytes: [0; 1024],
            used: 0,
        };
        output.append(&self.header().bytes())?;
        for entry in self.entries() {
            let mut descriptor = [0; 6];
            descriptor[..2].copy_from_slice(&entry.opid.to_le_bytes());
            descriptor[2..].copy_from_slice(&entry.value.to_le_bytes());
            output.append(&descriptor)?;
        }
        for entry in self.entries() {
            if let Some(data) = entry.complex {
                output.append(data)?;
            }
        }
        output.flush()
    }
    fn render(&self) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(self.header().body_len as usize + 8);
        match self.emit(|chunk| {
            bytes.extend_from_slice(chunk);
            Ok::<_, Infallible>(())
        }) {
            Ok(()) => bytes,
            Err(never) => match never {},
        }
    }
    fn check(&self, limits: Limits) -> Result<()> {
        let header = self.header();
        check(
            LimitKind::PropertyBytes,
            header.body_len as usize,
            limits.property_bytes_value(),
        )?;
        check(
            LimitKind::Properties,
            usize::from(header.count),
            limits.properties_value(),
        )?;
        if let Some(name) = self.name_bytes() {
            check(
                LimitKind::NameUnits,
                name.len() / 2 - 1,
                limits.name_units_value(),
            )?;
        }
        Ok(())
    }
}

/// An immutable standalone Opt-family picture package.
/// Opening borrows its input; publication shares untouched complex payloads.
#[derive(Clone)]
pub struct Package<'data> {
    state: State<'data>,
}
impl PartialEq for Package<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.state.same_content(&other.state)
            || self.state.header() == other.state.header()
                && self.state.entries().eq(other.state.entries())
    }
}
impl Eq for Package<'_> {}
impl<'data> Package<'data> {
    pub fn open(bytes: &'data [u8]) -> Result<Self> {
        Self::open_with_options(bytes, OpenOptions::default())
    }
    pub fn open_with_options(bytes: &'data [u8], options: OpenOptions) -> Result<Self> {
        let (record, consumed) = Record::parse(bytes, 0, crate::Limits::default())?;
        if consumed != bytes.len() {
            return Err(crate::Error::TrailingData { offset: consumed }.into());
        }
        let limits = options.limits_value();
        let parsed = Parsed::parse(&record, limits)?;
        // Keep the index for large tables; validated tiny tables use a bounded
        // raw scan and do not retain a second descriptor allocation.
        let properties = if parsed.properties.len() > crate::prop::LINEAR_PROPERTIES {
            Some(Box::new(parsed.properties))
        } else {
            drop(parsed.properties);
            None
        };
        Ok(Self {
            state: State::Opened(Arc::new(Base {
                policy: Policy::from_limits(limits),
                data: BaseData::Borrowed {
                    record: parsed.record,
                    properties,
                },
            })),
        })
    }
    /// Detaches borrowed payloads, sharing bytes already owned by an edit.
    #[must_use]
    pub fn into_owned(self) -> Package<'static> {
        own_package(&self, &mut HashMap::new(), &mut HashMap::new())
    }
    /// Requests a contiguous raw projection; rendering an edited record is lazy.
    pub fn view(&self) -> Result<View<'_>> {
        let cache = match &self.state {
            State::Opened(base) => match &base.data {
                BaseData::Borrowed { record, .. } => {
                    return View::parse(record, self.state.policy().limits());
                },
                BaseData::Owned(owned) => &owned.bytes,
            },
            State::Edited(edited) => &edited.bytes,
        };
        let bytes = cache.get_or_init(|| Box::new(self.state.render()));
        let (record, _) = Record::parse(bytes, 0, crate::Limits::default())?;
        View::parse(&record, self.state.policy().limits())
    }
    pub fn metadata(&self) -> Result<Option<Metadata<'_>>> {
        let name = self.state.name_bytes().map(Name::from_raw);
        let flags = self.state.raw_flags();
        if name.is_none() && flags.is_none() {
            return Ok(None);
        }
        Ok(Some(Metadata::new(
            name,
            Flags::from_raw(flags.unwrap_or(0))?,
        )))
    }
    #[must_use]
    pub fn edit(&self) -> Edit<'data> {
        self.edit_with_options(EditOptions::default())
    }
    #[must_use]
    pub fn edit_with_options(&self, options: EditOptions) -> Edit<'data> {
        let admitted = self.state.policy().limits();
        Edit {
            source: self.clone(),
            limits: options
                .limits_value()
                .unwrap_or(admitted)
                .constrained_by(admitted),
            name: None,
            flags: None,
        }
    }
    pub fn apply(&self, patch: &Patch<'data>) -> Result<Commit<'data>> {
        if self != &patch.before {
            return Err(Error::PatchConflict);
        }
        patch.after.state.check(self.state.policy().limits())?;
        let package = Self {
            state: patch.after.state.with_policy(self.state.policy()),
        };
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
    pub fn write_to(&self, mut writer: impl Write) -> Result<()> {
        self.state
            .emit(|chunk| writer.write_all(chunk))
            .map_err(Error::from)
    }
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.write_to(std::fs::File::create(path)?)
    }
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.state.render()
    }
}

/// Exact content-bound replay and inverse with borrowed modeled value changes.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch<'data> {
    before: Package<'data>,
    after: Package<'data>,
}
impl Patch<'_> {
    #[must_use]
    pub fn into_owned(self) -> Patch<'static> {
        let mut bases = HashMap::new();
        let mut copied = HashMap::new();
        Patch {
            before: own_package(&self.before, &mut bases, &mut copied),
            after: own_package(&self.after, &mut bases, &mut copied),
        }
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.state.same_content(&self.after.state)
    }
    #[must_use]
    pub fn name(&self) -> Option<Change<&[u8]>> {
        let before = self.before.state.name_bytes();
        let after = self.after.state.name_bytes();
        (before != after).then(|| Change::new(before, after))
    }
    #[must_use]
    pub fn flags(&self) -> Option<Change<u32>> {
        let before = self.before.state.raw_flags();
        let after = self.after.state.raw_flags();
        (before != after).then(|| Change::new(before, after))
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}
fn own_package(
    package: &Package<'_>,
    bases: &mut HashMap<usize, Arc<Base<'static>>>,
    copied: &mut HashMap<(usize, usize), Arc<[u8]>>,
) -> Package<'static> {
    let original = package.state.base();
    let base = Arc::clone(
        bases
            .entry(Arc::as_ptr(original) as usize)
            .or_insert_with(|| {
                let entries = match &original.data {
                    BaseData::Owned(owned) => owned.entries.clone(),
                    BaseData::Borrowed { .. } => original
                        .entries()
                        .map(|entry| Entry {
                            opid: entry.opid,
                            value: entry.value,
                            complex: entry.complex.map(|bytes| {
                                Arc::clone(
                                    copied
                                        .entry((bytes.as_ptr() as usize, bytes.len()))
                                        .or_insert_with(|| Arc::from(bytes)),
                                )
                            }),
                        })
                        .collect(),
                };
                let name = entries
                    .iter()
                    .position(|entry| entry.opid & 0x3fff == Id::PictureFileName.raw());
                let flags = entries
                    .iter()
                    .position(|entry| entry.opid & 0x3fff == Id::BlipFlags.raw());
                Arc::new(Base {
                    policy: original.policy,
                    data: BaseData::Owned(Box::new(OwnedBase {
                        header: original.header(),
                        entries,
                        name,
                        flags,
                        bytes: OnceLock::new(),
                    })),
                })
            }),
    );
    let state = match &package.state {
        State::Opened(_) => State::Opened(base),
        State::Edited(edited) => State::Edited(Arc::new(Deferred {
            base,
            name: edited.name.clone(),
            flags: edited.flags,
            name_last: edited.name_last,
            policy: edited.policy,
            bytes: OnceLock::new(),
        })),
    };
    Package { state }
}

/// Independent pending changes; dropping the edit leaves its source untouched.
pub struct Edit<'data> {
    source: Package<'data>,
    limits: Limits,
    name: Option<Option<Vec<u8>>>,
    flags: Option<Flags>,
}
impl<'data> Edit<'data> {
    pub fn set_name(&mut self, text: &str) -> Result<&mut Self> {
        self.name = Some(Some(encode_name(text, self.limits)?));
        Ok(self)
    }
    pub fn clear_name(&mut self) -> &mut Self {
        self.name = Some(None);
        self
    }
    pub fn set_flags(&mut self, flags: Flags) -> &mut Self {
        self.flags = Some(flags);
        self
    }
    pub fn commit(self) -> Result<Commit<'data>> {
        let before_name = self.source.state.name_bytes();
        let before_flags = self.source.state.raw_flags();
        let name_changed = self
            .name
            .as_ref()
            .is_some_and(|name| name.as_deref() != before_name);
        let flags_changed = self
            .flags
            .is_some_and(|flags| Some(flags.raw()) != before_flags);
        let policy = Policy::from_limits(self.limits);
        if !name_changed && !flags_changed {
            self.source.state.check(self.limits)?;
            let package = Package {
                state: self.source.state.with_policy(policy),
            };
            let patch = Patch {
                before: self.source,
                after: package.clone(),
            };
            return Ok(Commit::new(package, patch, Diagnostics::UNCHANGED));
        }
        let base = self.source.state.base();
        let original_name = base
            .selected(Id::PictureFileName)
            .and_then(|entry| entry.complex);
        let name = if name_changed {
            let after = self.name.flatten();
            if base.selected(Id::PictureFileName).is_some() && after.as_deref() == original_name {
                None
            } else {
                Some(after.map(Arc::from))
            }
        } else {
            self.source.state.name_override().cloned()
        };
        let original_flags = base
            .selected(Id::BlipFlags)
            .map(|entry| entry.value.cast_unsigned());
        let flags = if flags_changed {
            self.flags
                .map(Flags::raw)
                .filter(|value| Some(*value) != original_flags)
        } else {
            self.source.state.flags_override()
        };
        let added_name = name.is_some() && base.selected(Id::PictureFileName).is_none();
        let added_flags = flags.is_some() && original_flags.is_none();
        let name_last = added_name
            && added_flags
            && (self.source.state.name_last()
                || self.source.state.name_override().is_none()
                    && self.source.state.flags_override().is_some());
        // Validate arithmetic before constructing a state whose header is infallible.
        let mut count = usize::from(base.header().count);
        let mut body_len = base.header().body_len as usize;
        if let Some(after) = &name {
            let old = base.selected(Id::PictureFileName);
            if old.is_none() {
                count += 1;
                body_len += 6;
            }
            body_len -= old.and_then(|entry| entry.complex).map_or(0, <[u8]>::len);
            body_len = body_len
                .checked_add(after.as_ref().map_or(0, |bytes| bytes.len()))
                .ok_or(crate::Error::ArithmeticOverflow {
                    context: "picture property body length",
                })?;
        }
        if flags.is_some() && original_flags.is_none() {
            count += 1;
            body_len = body_len
                .checked_add(6)
                .ok_or(crate::Error::ArithmeticOverflow {
                    context: "picture property body length",
                })?;
        }
        check(LimitKind::Properties, count, self.limits.properties_value())?;
        check(
            LimitKind::PropertyBytes,
            body_len,
            self.limits.property_bytes_value(),
        )?;
        let package = Package {
            state: State::Edited(Arc::new(Deferred {
                base: Arc::clone(base),
                name,
                flags,
                name_last,
                policy,
                bytes: OnceLock::new(),
            })),
        };
        package.state.check(self.limits)?;
        let patch = Patch {
            before: self.source,
            after: package.clone(),
        };
        Ok(Commit::new(package, patch, Diagnostics::CHANGED))
    }
}
impl std::fmt::Debug for Package<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Package")
            .field("header", &self.state.header())
            .field("limits", &self.state.policy().limits())
            .finish_non_exhaustive()
    }
}
impl std::fmt::Debug for View<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("View")
            .field("record", self.record())
            .field("property_count", &self.properties().len())
            .finish_non_exhaustive()
    }
}
impl std::fmt::Debug for Patch<'_> {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("Patch")
            .field("is_noop", &self.is_noop())
            .field("name_changed", &self.name().is_some())
            .field("flags_changed", &self.flags().is_some())
            .finish_non_exhaustive()
    }
}
// Batch small descriptors/names on the stack, avoiding one syscall per field.
// Large payloads bypass this buffer and are passed directly to the sink.
struct Output<F> {
    write: F,
    bytes: [u8; 1024],
    used: usize,
}

impl<E, F: FnMut(&[u8]) -> std::result::Result<(), E>> Output<F> {
    fn append(&mut self, bytes: &[u8]) -> std::result::Result<(), E> {
        if bytes.len() > self.bytes.len() - self.used {
            self.flush()?;
        }
        if bytes.len() >= self.bytes.len() {
            return (self.write)(bytes);
        }
        self.bytes[self.used..self.used + bytes.len()].copy_from_slice(bytes);
        self.used += bytes.len();
        Ok(())
    }

    fn flush(&mut self) -> std::result::Result<(), E> {
        if self.used != 0 {
            (self.write)(&self.bytes[..self.used])?;
            self.used = 0;
        }
        Ok(())
    }
}

#[cfg(test)]
mod storage_tests {
    #![allow(clippy::expect_used, reason = "test assertions panic on failure")]
    use super::*;
    #[test]
    fn large_opaque_payload_is_shared_and_publication_stays_deferred() {
        let opaque = vec![0xa5; 2 * 1024 * 1024];
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0x13_u16.to_le_bytes());
        bytes.extend_from_slice(&crate::RecordKind::TertiaryOpt.raw().to_le_bytes());
        bytes.extend_from_slice(&u32::try_from(opaque.len() + 6).expect("size").to_le_bytes());
        bytes.extend_from_slice(&0xc600_u16.to_le_bytes());
        bytes.extend_from_slice(&i32::try_from(opaque.len()).expect("size").to_le_bytes());
        bytes.extend_from_slice(&opaque);
        let source = Package::open(&bytes).expect("source");
        let unchanged = source.edit().commit().expect("no-op");
        assert!(source.state.same_content(&unchanged.package().state));
        let mut edit = source.edit();
        edit.set_name("new").expect("name");
        let committed = edit.commit().expect("commit");
        let State::Edited(edited) = &committed.package().state else {
            panic!("deferred");
        };
        assert!(Arc::ptr_eq(&edited.base, source.state.base()));
        assert!(edited.bytes.get().is_none());
        assert_eq!(
            edited
                .base
                .entries()
                .next()
                .expect("opaque")
                .complex
                .expect("payload")
                .as_ptr(),
            bytes[14..].as_ptr()
        );
        let detached = committed.patch().clone().into_owned();
        assert!(Arc::ptr_eq(
            detached.before.state.base(),
            detached.after.state.base()
        ));
        let BaseData::Owned(owned) = &detached.before.state.base().data else {
            panic!("owned");
        };
        let entries = &owned.entries;
        assert_eq!(entries[0].complex.as_deref(), Some(opaque.as_slice()));
        let twice = detached.after.clone().into_owned();
        let BaseData::Owned(owned_twice) = &twice.state.base().data else {
            panic!("owned");
        };
        let twice_entries = &owned_twice.entries;
        assert!(Arc::ptr_eq(
            entries[0].complex.as_ref().expect("opaque"),
            twice_entries[0].complex.as_ref().expect("opaque")
        ));
        committed.package().metadata().expect("metadata");
        committed
            .package()
            .write_to(std::io::sink())
            .expect("stream");
        assert!(edited.bytes.get().is_none());
        let mut next = committed.package().edit();
        next.set_flags(Flags::from_raw(1).expect("flags"));
        let next = next.commit().expect("next");
        let State::Edited(next_edit) = &next.package().state else {
            panic!("deferred");
        };
        assert!(Arc::ptr_eq(&edited.base, &next_edit.base));
        assert!(Arc::ptr_eq(
            edited
                .name
                .as_ref()
                .expect("changed")
                .as_ref()
                .expect("name"),
            next_edit
                .name
                .as_ref()
                .expect("changed")
                .as_ref()
                .expect("name")
        ));
        next.package()
            .write_to(std::io::sink())
            .expect("stream chain");
        assert!(edited.bytes.get().is_none());
        assert!(next_edit.bytes.get().is_none());
        let view = committed.package().view().expect("contiguous view");
        assert_eq!(view.record().kind(), crate::RecordKind::TertiaryOpt);
        assert!(edited.bytes.get().is_some());
        assert_eq!(
            view.properties()
                .prop(Id::unknown(0x600).expect("id"))
                .expect("opaque")
                .raw_opid(),
            0xc600
        );
    }
}
