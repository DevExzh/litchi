//! `OfficeArt` shape topology, budget, and record validation.

use crate::prop::{Anchor, Id, Props};
use crate::{Container, Error, Limit, Record, RecordKind, Result};

use super::model::{Flags, Kind, Native};

#[derive(Debug)]
pub(crate) struct Budget {
    limits: crate::Limits,
    records: u32,
}

impl Budget {
    pub(crate) const fn new(limits: crate::Limits) -> Self {
        Self { limits, records: 0 }
    }

    pub(crate) fn visit(&mut self) -> Result<()> {
        self.records = self.records.checked_add(1).ok_or(Error::LimitExceeded {
            limit: Limit::Records,
            maximum: self.limits.max_records,
        })?;
        if self.records > self.limits.max_records {
            return Err(Error::LimitExceeded {
                limit: Limit::Records,
                maximum: self.limits.max_records,
            });
        }
        Ok(())
    }

    pub(crate) fn depth(&self, depth: u16) -> Result<()> {
        if depth > self.limits.max_depth {
            return Err(Error::LimitExceeded {
                limit: Limit::Depth,
                maximum: u32::from(self.limits.max_depth),
            });
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
pub(crate) enum Role {
    Patriarch,
    Root,
    Member,
    Standalone,
}

#[derive(Debug)]
pub(crate) struct Meta<'data> {
    pub(crate) sp: Option<Record<'data>>,
    pub(crate) spgr: Option<Record<'data>>,
    pub(crate) primary: Option<Props<'data>>,
    pub(crate) secondary: Option<Props<'data>>,
    pub(crate) tertiary: Option<Props<'data>>,
    pub(crate) child_anchor: Option<Record<'data>>,
    pub(crate) client_anchor: Option<Record<'data>>,
    pub(crate) client_data: Option<Record<'data>>,
    pub(crate) textbox: Option<Record<'data>>,
}

impl Meta<'_> {
    pub(crate) const fn new() -> Self {
        Self {
            sp: None,
            spgr: None,
            primary: None,
            secondary: None,
            tertiary: None,
            child_anchor: None,
            client_anchor: None,
            client_data: None,
            textbox: None,
        }
    }
}

pub(crate) fn scan_meta<'data>(
    container: &Container<'data>,
    budget: &mut Budget,
) -> Result<Meta<'data>> {
    validate_container_header(container, RecordKind::SpContainer)?;
    let mut meta = Meta::new();
    for child in container.children() {
        let child = child?;
        budget.visit()?;
        match child.kind() {
            RecordKind::Sp => {
                validate_atom(&child, RecordKind::Sp, 2, None, 8)?;
                insert(&mut meta.sp, child)?;
            },
            RecordKind::Spgr => {
                validate_atom(&child, RecordKind::Spgr, 1, Some(0), 16)?;
                insert(&mut meta.spgr, child)?;
            },
            RecordKind::Opt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.primary, props)?;
            },
            RecordKind::SecondaryOpt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.secondary, props)?;
            },
            RecordKind::TertiaryOpt => {
                let props = Props::parse(&child)?;
                insert(&mut meta.tertiary, props)?;
            },
            RecordKind::ChildAnchor => {
                validate_atom(&child, RecordKind::ChildAnchor, 0, Some(0), 16)?;
                insert(&mut meta.child_anchor, child)?;
            },
            RecordKind::ClientAnchor => {
                validate_atom_kind(&child, RecordKind::ClientAnchor, 0, Some(0))?;
                insert(&mut meta.client_anchor, child)?;
            },
            RecordKind::ClientData => insert(&mut meta.client_data, child)?,
            RecordKind::ClientTextbox => insert(&mut meta.textbox, child)?,
            RecordKind::Unknown(_) => {},
            _ => {
                return Err(Error::MalformedShape {
                    reason: "SpContainer contains an invalid child record",
                });
            },
        }
    }
    Ok(meta)
}

pub(crate) fn validate_meta(
    meta: &Meta<'_>,
    group: bool,
    role: Role,
) -> Result<(u32, Native, Flags, Option<Anchor>)> {
    let sp = meta.sp.as_ref().ok_or(Error::MalformedShape {
        reason: "SpContainer has no shape atom",
    })?;
    let data: &[u8; 8] = sp.data().try_into().map_err(|_err| Error::MalformedShape {
        reason: "shape atom payload is not eight bytes",
    })?;
    let id = u32::from_le_bytes(data[..4].try_into().map_err(|_err| Error::MalformedShape {
        reason: "shape identifier is truncated",
    })?);
    let raw_flags =
        u32::from_le_bytes(data[4..].try_into().map_err(|_err| Error::MalformedShape {
            reason: "shape flags are truncated",
        })?);
    let native = Native::from_raw(sp.instance());
    let flags = Flags::from_bits_retain(raw_flags);

    if flags.contains(Flags::GROUP) != group {
        return Err(Error::MalformedShape {
            reason: "shape GROUP flag disagrees with its container topology",
        });
    }
    if group != meta.spgr.is_some() {
        return Err(Error::MalformedShape {
            reason: "group shape must contain exactly one Spgr atom",
        });
    }
    if group && native != Native::FREEFORM {
        return Err(Error::MalformedShape {
            reason: "group shape must use the non-primitive native kind",
        });
    }

    let patriarch = matches!(role, Role::Patriarch);
    if flags.contains(Flags::PATRIARCH) != patriarch && !matches!(role, Role::Standalone) {
        return Err(Error::MalformedShape {
            reason: "shape PATRIARCH flag disagrees with its container topology",
        });
    }
    let expected_child = match role {
        Role::Patriarch | Role::Root => Some(false),
        Role::Member => Some(true),
        Role::Standalone => None,
    };
    if expected_child.is_some_and(|expected| flags.contains(Flags::CHILD) != expected) {
        return Err(Error::MalformedShape {
            reason: "shape CHILD flag disagrees with its group membership",
        });
    }

    if meta.child_anchor.is_some() && meta.client_anchor.is_some() {
        return Err(Error::MalformedShape {
            reason: "shape contains both child and host anchors",
        });
    }
    let has_anchor = meta.child_anchor.is_some() || meta.client_anchor.is_some();
    // Word emits a direct background shape whose fHaveAnchor bit is set even
    // though its host-owned anchor is omitted. Keep the structural checks for
    // every user shape, but accept this producer-specific, non-visible sentinel.
    let word_background_sentinel =
        flags.contains(Flags::BACKGROUND | Flags::HAVE_ANCHOR) && !has_anchor;
    if flags.contains(Flags::HAVE_ANCHOR) != has_anchor && !word_background_sentinel {
        return Err(Error::MalformedShape {
            reason: "shape HAVE_ANCHOR flag disagrees with its anchor records",
        });
    }
    if flags.contains(Flags::CHILD) {
        if meta.client_anchor.is_some() {
            return Err(Error::MalformedShape {
                reason: "group child uses a host ClientAnchor",
            });
        }
    } else if meta.child_anchor.is_some() {
        return Err(Error::MalformedShape {
            reason: "non-child shape uses a ChildAnchor",
        });
    }

    let anchor = meta
        .child_anchor
        .as_ref()
        .map(|record| {
            Anchor::from_child_anchor(record).ok_or(Error::MalformedShape {
                reason: "child anchor payload is not sixteen bytes",
            })
        })
        .transpose()?;

    Ok((id, native, flags, anchor))
}

pub(crate) fn detect(native: Native, props: &Props<'_>) -> Kind {
    match native.raw() {
        0 if props.has(Id::Vertices) => Kind::Polygon,
        0 => Kind::AutoShape,
        1 => Kind::Rectangle,
        3 => Kind::Ellipse,
        20 => Kind::Line,
        32..=40 => Kind::Connector,
        41..=52 | 61..=63 | 106 | 178..=189 => Kind::Callout,
        75 => Kind::Picture,
        202 => Kind::TextBox,
        2..=201 => Kind::AutoShape,
        _ => Kind::Unknown,
    }
}

fn insert<T>(slot: &mut Option<T>, value: T) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(Error::MalformedShape {
            reason: "shape container contains a duplicate singleton record",
        });
    }
    Ok(())
}

pub(crate) fn validate_container_header(container: &Container<'_>, kind: RecordKind) -> Result<()> {
    validate_container_record(container.record(), kind)
}

pub(crate) fn validate_container_record(record: &Record<'_>, kind: RecordKind) -> Result<()> {
    if record.kind() != kind
        || record.raw_kind() != kind.raw()
        || record.version() != 0x0F
        || record.instance() != 0
    {
        return Err(Error::MalformedShape {
            reason: "OfficeArt container header is invalid",
        });
    }
    Ok(())
}

pub(crate) fn validate_atom(
    record: &Record<'_>,
    kind: RecordKind,
    version: u8,
    instance: Option<u16>,
    len: u32,
) -> Result<()> {
    validate_atom_kind(record, kind, version, instance)?;
    if record.len() != len || usize::try_from(len).ok() != Some(record.data().len()) {
        return Err(Error::MalformedShape {
            reason: "OfficeArt atom payload length is invalid",
        });
    }
    Ok(())
}

pub(crate) fn validate_atom_kind(
    record: &Record<'_>,
    kind: RecordKind,
    version: u8,
    instance: Option<u16>,
) -> Result<()> {
    if record.kind() != kind
        || record.raw_kind() != kind.raw()
        || record.version() != version
        || instance.is_some_and(|expected| record.instance() != expected)
    {
        return Err(Error::MalformedShape {
            reason: "OfficeArt atom header is invalid",
        });
    }
    Ok(())
}

pub(crate) fn next_depth(depth: u16) -> Result<u16> {
    depth.checked_add(1).ok_or(Error::LimitExceeded {
        limit: Limit::Depth,
        maximum: u32::from(u16::MAX),
    })
}

pub(crate) fn coordinate(data: &[u8], offset: usize) -> Result<i32> {
    let end = offset.checked_add(4).ok_or(Error::ArithmeticOverflow {
        context: "group coordinate extent",
    })?;
    let bytes = data.get(offset..end).ok_or(Error::MalformedShape {
        reason: "group coordinate atom payload is truncated",
    })?;
    let bytes: [u8; 4] = bytes.try_into().map_err(|_err| Error::MalformedShape {
        reason: "group coordinate atom payload is truncated",
    })?;
    Ok(i32::from_le_bytes(bytes))
}
