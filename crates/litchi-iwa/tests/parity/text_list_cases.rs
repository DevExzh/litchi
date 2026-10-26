//! Paragraph-list parity (wave W1-D, record 0374): the host `text::Draft`
//! list families against `litchi_iwa_text_wire::list` published by
//! [`text_glue`](crate::text_glue). New list styles get random metadata
//! UUIDs on the host, so a change that creates one is tier B; in-place
//! rewrites, collapses, level and numbering changes are tier A.

use std::collections::BTreeMap;

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_common::color::{RgbColorSpace, Rgba};
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::list::{
    ParagraphList, ParagraphListBullet, ParagraphListBulletBaselineOffset,
    ParagraphListBulletGeometry, ParagraphListBulletScale, ParagraphListIndentation,
    ParagraphListLabelColor, ParagraphListLabelIndent, ParagraphListLevel,
    ParagraphListNumberFormat, ParagraphListNumberPunctuation, ParagraphListNumberScale,
    ParagraphListNumberSequence, ParagraphListNumberTiering, ParagraphListNumbering,
    ParagraphListPlacement, ParagraphListStart, ParagraphListTextGap,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::list::{self, Change, Plan, Property, PropertyKind};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::text_glue::{Workspace, new_object};
use crate::text_style_cases::FIXTURES;

const LIMITS: WireLimits = WireLimits::DEFAULT;

#[derive(Clone, Debug)]
pub enum Operation {
    SetList(ParagraphList),
    ResetList,
    SetLists(Vec<ParagraphListPlacement>),
    SetLevel(TextPosition, ParagraphListLevel),
    ResetLevel(TextPosition),
    SetNumbering(TextPosition, ParagraphListNumbering),
    Property(TextPosition, Change),
}

struct Styles {
    payloads: BTreeMap<u64, Vec<u8>>,
    entries: BTreeMap<u64, String>,
    references: BTreeMap<u64, usize>,
    parents: Vec<u64>,
}

impl Styles {
    fn load(workspace: &Workspace) -> TestResult<Self> {
        let mut payloads = BTreeMap::new();
        let mut entries = BTreeMap::new();
        let mut parents = Vec::new();
        for located in workspace.messages(list::MESSAGE_TYPE)? {
            if let Ok(decoded) = tswp::ListStyleArchive::decode(located.payload.as_slice()) {
                parents.extend(decoded.super_.parent.map(|parent| parent.identifier));
            }
            entries.insert(located.object, located.entry);
            payloads.insert(located.object, located.payload);
        }
        let mut references = BTreeMap::new();
        for storage in workspace.storages()? {
            let decoded = tswp::StorageArchive::decode(storage.payload.as_slice())?;
            for entry in decoded
                .table_list_style
                .iter()
                .flat_map(|table| &table.entries)
            {
                if let Some(reference) = &entry.object {
                    *references.entry(reference.identifier).or_default() += 1;
                }
            }
        }
        Ok(Self {
            payloads,
            entries,
            references,
            parents,
        })
    }

    fn exclusive(&self, style: u64) -> bool {
        self.references.get(&style) == Some(&1) && !self.parents.contains(&style)
    }

    fn entry(&self, style: u64) -> TestResult<String> {
        Ok(self.entries.get(&style).ok_or("list style entry")?.clone())
    }
}

fn apply(
    workspace: &mut Workspace,
    storage: u64,
    paragraph: TextPosition,
    plan: Plan,
) -> TestResult<bool> {
    let styles = Styles::load(workspace)?;
    let located = workspace.storage(storage)?;
    match plan {
        Plan::Unchanged => return Ok(false),
        Plan::Rewrite { style, payload } => {
            workspace.replace_object_message(
                &styles.entry(style)?,
                style,
                list::MESSAGE_TYPE,
                payload,
            )?;
        },
        Plan::Derive {
            parent,
            stylesheet,
            payload,
        } => {
            let style_entry = styles.entry(parent)?;
            let created = workspace.next_identifier()?;
            let object = new_object(
                created,
                list::MESSAGE_TYPE,
                payload,
                &list::MESSAGE_VERSION,
                &[parent],
                Vec::new(),
            )?;
            workspace.insert_style_variation(&style_entry, stylesheet, parent, object)?;
            workspace.register_private_style(&located.entry, &style_entry, created)?;
            let located = workspace.storage(storage)?;
            let old: Vec<u64> = list::boundaries(&located.payload, LIMITS)?
                .into_iter()
                .map(|entry| entry.1)
                .collect();
            let new = list::boundaries_with_style(&located.payload, paragraph, created, LIMITS)?;
            let payload = list::replace_boundaries(&located.payload, &new, LIMITS)?;
            let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
            workspace.replace_with_reference_sequence(&located, payload, &old, &new)?;
            workspace.raise_watermark(created)?;
        },
        Plan::Collapse {
            style,
            parent,
            stylesheet,
        } => {
            let style_entry = styles.entry(style)?;
            let old: Vec<u64> = list::boundaries(&located.payload, LIMITS)?
                .into_iter()
                .map(|entry| entry.1)
                .collect();
            let new = list::boundaries_with_style(&located.payload, paragraph, parent, LIMITS)?;
            let payload = list::replace_boundaries(&located.payload, &new, LIMITS)?;
            let new: Vec<u64> = new.into_iter().map(|entry| entry.1).collect();
            workspace.replace_with_reference_sequence(&located, payload, &old, &new)?;
            workspace.remove_style_variation(&style_entry, stylesheet, parent, style)?;
            workspace.unregister_private_style(
                &located.entry,
                &style_entry,
                style,
                Some(parent),
            )?;
            workspace.release_watermark(&[style])?;
        },
        other => return Err(format!("unmapped plan {other:?}").into()),
    }
    Ok(true)
}

/// The stylesheet's preset object for `preset`, created when missing.
fn preset_object(
    workspace: &mut Workspace,
    owner_entry: &str,
    style_entry: &str,
    stylesheet: u64,
    preset: ParagraphList,
) -> TestResult<u64> {
    let candidates: Vec<(u64, Vec<u8>)> = workspace
        .messages(list::MESSAGE_TYPE)?
        .into_iter()
        .filter(|located| located.entry == style_entry)
        .map(|located| (located.object, located.payload))
        .collect();
    let existing = list::find_preset(
        candidates
            .iter()
            .map(|(identifier, payload)| (*identifier, payload.as_slice())),
        stylesheet,
        preset,
        LIMITS,
    )?;
    if let Some(identifier) = existing {
        workspace.register_style_reference(owner_entry, style_entry, identifier)?;
        return Ok(identifier);
    }
    let created = workspace.next_identifier()?;
    let object = new_object(
        created,
        list::MESSAGE_TYPE,
        list::preset_payload(stylesheet, preset, LIMITS)?,
        &list::MESSAGE_VERSION,
        &[],
        Vec::new(),
    )?;
    workspace.insert_private_style(style_entry, stylesheet, object)?;
    workspace.register_private_style(owner_entry, style_entry, created)?;
    workspace.raise_watermark(created)?;
    Ok(created)
}

fn set_lists(
    workspace: &mut Workspace,
    storage: u64,
    placements: &[ParagraphListPlacement],
) -> TestResult {
    let located = workspace.storage(storage)?;
    let normalized = list::normalize(&located.payload, placements, LIMITS)?;
    let styles = Styles::load(workspace)?;
    if list::placements(&located.payload, &styles.payloads, LIMITS)? == normalized {
        return Ok(());
    }
    let boundaries = list::boundaries(&located.payload, LIMITS)?;
    let first = boundaries.first().ok_or("empty list table")?.1;
    let style_entry = styles.entry(first)?;
    let stylesheet = list::stylesheet(&styles.payloads, first, LIMITS)?;
    let mut slots: [Option<u64>; 3] = [None; 3];
    let slot = |preset: ParagraphList| match preset {
        ParagraphList::None => 0,
        ParagraphList::Bullet => 1,
        ParagraphList::Numbered => 2,
    };
    for placement in &normalized {
        if slots[slot(placement.list)].is_none() {
            slots[slot(placement.list)] = Some(preset_object(
                workspace,
                &located.entry,
                &style_entry,
                stylesheet,
                placement.list,
            )?);
        }
    }
    let new: Vec<(u32, u64)> = normalized
        .iter()
        .map(|placement| {
            (
                placement.paragraph.utf16_index(),
                slots[slot(placement.list)].unwrap_or_default(),
            )
        })
        .collect();
    let located = workspace.storage(storage)?;
    let old: Vec<u64> = boundaries.iter().map(|entry| entry.1).collect();
    let payload = list::replace_boundaries(&located.payload, &new, LIMITS)?;
    let new_styles: Vec<u64> = new.iter().map(|entry| entry.1).collect();
    workspace.replace_with_reference_sequence(&located, payload, &old, &new_styles)?;
    let mut unique = old;
    unique.sort_unstable();
    unique.dedup();
    for style in unique {
        workspace.unregister_owner_reference_if_unused(&located.entry, &style_entry, style)?;
    }
    Ok(())
}

fn set_list(workspace: &mut Workspace, storage: u64, preset: ParagraphList) -> TestResult {
    let located = workspace.storage(storage)?;
    let style = list::uniform_style(&located.payload, LIMITS)?;
    let styles = Styles::load(workspace)?;
    if list::preset(&styles.payloads, style, LIMITS).ok() == Some(preset) {
        return Ok(());
    }
    let style_entry = styles.entry(style)?;
    let stylesheet = list::stylesheet(&styles.payloads, style, LIMITS)?;
    let target = preset_object(workspace, &located.entry, &style_entry, stylesheet, preset)?;
    let located = workspace.storage(storage)?;
    let payload = list::replace_boundaries(&located.payload, &[(0, target)], LIMITS)?;
    workspace.replace_with_reference_sequence(&located, payload, &[style], &[target])?;
    workspace.unregister_owner_reference_if_unused(&located.entry, &style_entry, style)?;
    Ok(())
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    match operation {
        Operation::SetList(preset) => set_list(&mut workspace, storage, *preset)?,
        Operation::ResetList => {
            let located = workspace.storage(storage)?;
            let style = list::uniform_style(&located.payload, LIMITS)?;
            let styles = Styles::load(&workspace)?;
            if list::preset(&styles.payloads, style, LIMITS).ok() != Some(ParagraphList::None) {
                set_list(&mut workspace, storage, ParagraphList::None)?;
            }
        },
        Operation::SetLists(placements) => set_lists(&mut workspace, storage, placements)?,
        Operation::SetLevel(paragraph, level) => {
            let located = workspace.storage(storage)?;
            if let Some(payload) = list::with_level(&located.payload, *paragraph, *level, LIMITS)? {
                workspace.replace_with_references(&located, payload, None, None)?;
            }
        },
        Operation::ResetLevel(paragraph) => {
            let located = workspace.storage(storage)?;
            if let Some(payload) = list::with_level(
                &located.payload,
                *paragraph,
                ParagraphListLevel::ZERO,
                LIMITS,
            )? {
                workspace.replace_with_references(&located, payload, None, None)?;
            }
        },
        Operation::SetNumbering(paragraph, numbering) => {
            let located = workspace.storage(storage)?;
            let styles = Styles::load(&workspace)?;
            if let Some(payload) = list::with_numbering(
                &located.payload,
                &styles.payloads,
                *paragraph,
                *numbering,
                LIMITS,
            )? {
                workspace.replace_with_references(&located, payload, None, None)?;
            }
        },
        Operation::Property(paragraph, change) => {
            let located = workspace.storage(storage)?;
            let styles = Styles::load(&workspace)?;
            let plan = list::plan(
                &located.payload,
                &styles.payloads,
                *paragraph,
                change,
                |style| styles.exclusive(style),
                LIMITS,
            )?;
            let changed = apply(&mut workspace, storage, *paragraph, plan)?;
            if let Change::Reset(kind) = change
                && changed
            {
                let located = workspace.storage(storage)?;
                let styles = Styles::load(&workspace)?;
                let plan = list::plan_collapse(
                    &located.payload,
                    &styles.payloads,
                    *paragraph,
                    *kind,
                    |style| styles.exclusive(style),
                    LIMITS,
                )?;
                apply(&mut workspace, storage, *paragraph, plan)?;
            }
        },
    }
    workspace.into_bytes()
}

#[allow(clippy::too_many_lines, reason = "one arm per host method")]
pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation.clone() {
        Operation::SetList(preset) => draft.set_paragraph_list(id, preset)?,
        Operation::ResetList => {
            draft.reset_paragraph_list(id)?;
        },
        Operation::SetLists(placements) => draft.set_paragraph_lists(id, &placements)?,
        Operation::SetLevel(paragraph, level) => {
            draft.set_paragraph_list_level(id, paragraph, level)?
        },
        Operation::ResetLevel(paragraph) => {
            draft.reset_paragraph_list_level(id, paragraph)?;
        },
        Operation::SetNumbering(paragraph, numbering) => {
            draft.set_paragraph_list_numbering(id, paragraph, numbering)?;
        },
        Operation::Property(paragraph, Change::Set(property)) => match property {
            Property::Bullet(value) => draft.set_paragraph_list_bullet(id, paragraph, &value)?,
            Property::BulletGeometry(value) => {
                draft.set_paragraph_list_bullet_geometry(id, paragraph, value)?;
            },
            Property::Indentation(value) => {
                draft.set_paragraph_list_indentation(id, paragraph, value)?
            },
            Property::LabelColor(value) => {
                draft.set_paragraph_list_label_color(id, paragraph, value)?
            },
            Property::NumberFormat(value) => {
                draft.set_paragraph_list_number_format(id, paragraph, value)?;
            },
            Property::NumberScale(value) => {
                draft.set_paragraph_list_number_scale(id, paragraph, value)?
            },
            Property::NumberTiering(value) => {
                draft.set_paragraph_list_number_tiering(id, paragraph, value)?;
            },
            other => return Err(format!("unmapped property {other:?}").into()),
        },
        Operation::Property(paragraph, Change::Reset(kind)) => {
            match kind {
                PropertyKind::Bullet => draft.reset_paragraph_list_bullet(id, paragraph)?,
                PropertyKind::BulletGeometry => {
                    draft.reset_paragraph_list_bullet_geometry(id, paragraph)?
                },
                PropertyKind::Indentation => {
                    draft.reset_paragraph_list_indentation(id, paragraph)?
                },
                PropertyKind::LabelColor => {
                    draft.reset_paragraph_list_label_color(id, paragraph)?
                },
                PropertyKind::NumberFormat => {
                    draft.reset_paragraph_list_number_format(id, paragraph)?
                },
                PropertyKind::NumberScale => {
                    draft.reset_paragraph_list_number_scale(id, paragraph)?
                },
                PropertyKind::NumberTiering => {
                    draft.reset_paragraph_list_number_tiering(id, paragraph)?
                },
                other => return Err(format!("unmapped kind {other:?}").into()),
            };
        },
        Operation::Property(_, other) => return Err(format!("unmapped change {other:?}").into()),
    }
    Ok(draft.to_bytes()?)
}

/// Everything the list families read for one storage, through one side.
pub type Reads = Vec<String>;

fn paragraphs(bytes: &[u8], storage: u64) -> TestResult<Vec<TextPosition>> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    Ok(list::paragraph_starts(&located.payload, LIMITS)?
        .into_iter()
        .map(TextPosition::from_utf16_code_units)
        .collect())
}

pub fn host_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    let mut reads = vec![
        format!("{:?}", draft.paragraph_list(id).ok()),
        format!("{:?}", draft.paragraph_lists(id).ok()),
        format!("{:?}", draft.paragraph_list_levels(id).ok()),
    ];
    for paragraph in paragraphs(bytes, storage)? {
        reads.push(format!(
            "{:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            draft.paragraph_list_level(id, paragraph).ok(),
            draft.paragraph_list_numbering(id, paragraph).ok(),
            draft.paragraph_list_bullet(id, paragraph).ok(),
            draft.paragraph_list_bullet_geometry(id, paragraph).ok(),
            draft.paragraph_list_indentation(id, paragraph).ok(),
            draft.paragraph_list_label_color(id, paragraph).ok(),
            draft.paragraph_list_number_format(id, paragraph).ok(),
            draft.paragraph_list_number_scale(id, paragraph).ok(),
            draft.paragraph_list_number_tiering(id, paragraph).ok(),
            paragraph,
        ));
    }
    Ok(reads)
}

pub fn leaf_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let styles = Styles::load(&workspace)?;
    let payload = located.payload.as_slice();
    let map = &styles.payloads;
    let uniform = list::uniform_style(payload, LIMITS)
        .ok()
        .and_then(|style| list::preset(map, style, LIMITS).ok());
    let mut reads = vec![
        format!("{uniform:?}"),
        format!("{:?}", list::placements(payload, map, LIMITS).ok()),
        format!("{:?}", list::levels(payload, LIMITS).ok()),
    ];
    for paragraph in paragraphs(bytes, storage)? {
        let read = |kind| list::property(payload, map, paragraph, kind, LIMITS).ok();
        let bullet = match read(PropertyKind::Bullet) {
            Some(Property::Bullet(value)) => Some(value),
            _ => None,
        };
        let geometry = match read(PropertyKind::BulletGeometry) {
            Some(Property::BulletGeometry(value)) => Some(value),
            _ => None,
        };
        let indentation = match read(PropertyKind::Indentation) {
            Some(Property::Indentation(value)) => Some(value),
            _ => None,
        };
        let color = match read(PropertyKind::LabelColor) {
            Some(Property::LabelColor(value)) => Some(value),
            _ => None,
        };
        let format = match read(PropertyKind::NumberFormat) {
            Some(Property::NumberFormat(value)) => Some(value),
            _ => None,
        };
        let scale = match read(PropertyKind::NumberScale) {
            Some(Property::NumberScale(value)) => Some(value),
            _ => None,
        };
        let tiering = match read(PropertyKind::NumberTiering) {
            Some(Property::NumberTiering(value)) => Some(value),
            _ => None,
        };
        reads.push(format!(
            "{:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?} {:?}",
            list::level(payload, paragraph, LIMITS).ok(),
            list::numbering(payload, paragraph, LIMITS).ok(),
            bullet,
            geometry,
            indentation,
            color,
            format,
            scale,
            tiering,
            paragraph,
        ));
    }
    Ok(reads)
}

fn run(
    name: &'static str,
    format: Format,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    operation: &Operation,
) -> TestResult {
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_route(source, storage, operation))
        .leaf(|source| Ok(LeafRun::output(leaf_route(source, storage, operation)?)))
        .read("host list reads", move |bytes| host_reads(bytes, storage))
        .read("leaf list reads", move |bytes| leaf_reads(bytes, storage))
        .known_fields(|known| {
            known.with(
                list::MESSAGE_TYPE,
                "TSWP.ListStyleArchive",
                decode_as::<tswp::ListStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

/// Storages with a list table, a paragraph-data table and text.
fn targets(bytes: &[u8], limit: usize) -> TestResult<Vec<u64>> {
    let draft = Draft::from_bytes(bytes)?;
    let mut found = Vec::new();
    for storage in draft.storages()? {
        let id: u64 = storage.id.to_string().parse()?;
        if storage.storage.text().is_empty() {
            continue;
        }
        let id_typed = storage.id;
        if draft.paragraph_lists(id_typed).is_err()
            || draft.paragraph_list_levels(id_typed).is_err()
        {
            continue;
        }
        found.push(id);
        if found.len() == limit {
            break;
        }
    }
    Ok(found)
}

fn with(input: &Fixture, storage: u64, operation: &Operation) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(&input.bytes, storage, operation)?,
    })
}

fn samples() -> TestResult<Vec<Property>> {
    let color = Rgba::new(0.8, 0.2, 0.1, 1.0, RgbColorSpace::Srgb)?;
    Ok(vec![
        Property::Bullet(ParagraphListBullet::new("✓")?),
        Property::BulletGeometry(ParagraphListBulletGeometry::new(
            ParagraphListBulletScale::from_ratio(1.5)?,
            ParagraphListBulletBaselineOffset::from_points(2.0)?,
        )),
        Property::Indentation(ParagraphListIndentation::new(
            ParagraphListLabelIndent::from_points(30.0)?,
            ParagraphListTextGap::from_em(1.5)?,
        )),
        Property::LabelColor(ParagraphListLabelColor::Explicit(color)),
        Property::NumberFormat(ParagraphListNumberFormat::affixed(
            ParagraphListNumberSequence::RomanUppercase,
            ParagraphListNumberPunctuation::RightParenthesis,
        )),
        Property::NumberScale(ParagraphListNumberScale::from_ratio(1.25)?),
        Property::NumberTiering(ParagraphListNumberTiering::Tiered),
    ])
}

const fn preset_for(kind: PropertyKind) -> ParagraphList {
    match kind {
        PropertyKind::Bullet | PropertyKind::BulletGeometry | PropertyKind::LabelColor => {
            ParagraphList::Bullet
        },
        _ => ParagraphList::Numbered,
    }
}

#[test]
fn text_list_preset_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
                run(
                    "text_list_preset_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    &Operation::SetList(preset),
                )
                .map_err(|e| format!("{path} storage {storage} {preset:?}: {e}"))?;
                let base = with(&input, storage, &Operation::SetList(preset))?;
                for operation in [
                    Operation::ResetList,
                    Operation::SetList(ParagraphList::None),
                ] {
                    run(
                        "text_list_preset_parity",
                        format,
                        Tier::B,
                        &base,
                        storage,
                        &operation,
                    )
                    .map_err(|e| format!("{path} storage {storage} {operation:?}: {e}"))?;
                }
                cases += 3;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 12, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_placements_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            let starts = paragraphs(&input.bytes, storage)?;
            let mut placements = vec![ParagraphListPlacement::new(
                TextPosition::ZERO,
                ParagraphList::Numbered,
            )];
            if let Some(second) = starts.get(1)
                && starts.len() > 2
            {
                placements.push(ParagraphListPlacement::new(*second, ParagraphList::Bullet));
            }
            run(
                "text_list_placements_parity",
                format,
                Tier::B,
                &input,
                storage,
                &Operation::SetLists(placements),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 4, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_level_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            let base = with(&input, storage, &Operation::SetList(ParagraphList::Bullet))?;
            let level = ParagraphListLevel::new(2)?;
            run(
                "text_list_level_parity",
                format,
                Tier::A,
                &base,
                storage,
                &Operation::SetLevel(TextPosition::ZERO, level),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            let leveled = with(
                &base,
                storage,
                &Operation::SetLevel(TextPosition::ZERO, level),
            )?;
            run(
                "text_list_level_parity",
                format,
                Tier::A,
                &leveled,
                storage,
                &Operation::ResetLevel(TextPosition::ZERO),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 2;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_numbering_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            let base = with(
                &input,
                storage,
                &Operation::SetList(ParagraphList::Numbered),
            )?;
            let restart = ParagraphListNumbering::StartAt(ParagraphListStart::new(5)?);
            run(
                "text_list_numbering_parity",
                format,
                Tier::A,
                &base,
                storage,
                &Operation::SetNumbering(TextPosition::ZERO, restart),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            let restarted = with(
                &base,
                storage,
                &Operation::SetNumbering(TextPosition::ZERO, restart),
            )?;
            run(
                "text_list_numbering_parity",
                format,
                Tier::A,
                &restarted,
                storage,
                &Operation::SetNumbering(
                    TextPosition::ZERO,
                    ParagraphListNumbering::StartAt(ParagraphListStart::new(2)?),
                ),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 2;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 8, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_property_set_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 1)? {
            for property in samples()? {
                let base = with(
                    &input,
                    storage,
                    &Operation::SetList(preset_for(property.kind())),
                )?;
                run(
                    "text_list_property_set_parity",
                    format,
                    Tier::B,
                    &base,
                    storage,
                    &Operation::Property(TextPosition::ZERO, Change::Set(property.clone())),
                )
                .map_err(|e| format!("{path} storage {storage} {property:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 28, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_property_rewrite_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 1)? {
            for property in samples()? {
                let base = with(
                    &input,
                    storage,
                    &Operation::SetList(preset_for(property.kind())),
                )?;
                let set = with(
                    &base,
                    storage,
                    &Operation::Property(TextPosition::ZERO, Change::Set(property.clone())),
                )?;
                let again = list_second(&property)?;
                run(
                    "text_list_property_rewrite_parity",
                    format,
                    Tier::A,
                    &set,
                    storage,
                    &Operation::Property(TextPosition::ZERO, Change::Set(again)),
                )
                .map_err(|e| format!("{path} storage {storage} {property:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 28, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_property_reset_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 1)? {
            for property in samples()? {
                let base = with(
                    &input,
                    storage,
                    &Operation::SetList(preset_for(property.kind())),
                )?;
                let set = with(
                    &base,
                    storage,
                    &Operation::Property(TextPosition::ZERO, Change::Set(property.clone())),
                )?;
                run(
                    "text_list_property_reset_parity",
                    format,
                    Tier::A,
                    &set,
                    storage,
                    &Operation::Property(TextPosition::ZERO, Change::Reset(property.kind())),
                )
                .map_err(|e| format!("{path} storage {storage} {property:?}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 28, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_list_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 3)? {
            let mut inputs = vec![input.bytes.clone()];
            for preset in [ParagraphList::Bullet, ParagraphList::Numbered] {
                let base = with(&input, storage, &Operation::SetList(preset))?;
                inputs.push(base.bytes.clone());
                for property in samples()? {
                    if preset_for(property.kind()) == preset {
                        inputs.push(
                            with(
                                &base,
                                storage,
                                &Operation::Property(TextPosition::ZERO, Change::Set(property)),
                            )?
                            .bytes,
                        );
                    }
                }
            }
            for bytes in &inputs {
                assert_eq!(
                    host_reads(bytes, storage)?,
                    leaf_reads(bytes, storage)?,
                    "{path} storage {storage}"
                );
                reads += 1;
            }
        }
    }
    assert!(reads >= 40, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}

fn list_second(property: &Property) -> TestResult<Property> {
    let color = Rgba::new(0.1, 0.6, 0.3, 1.0, RgbColorSpace::DisplayP3)?;
    Ok(match property {
        Property::Bullet(_) => Property::Bullet(ParagraphListBullet::new("➤")?),
        Property::BulletGeometry(_) => Property::BulletGeometry(ParagraphListBulletGeometry::new(
            ParagraphListBulletScale::from_ratio(0.75)?,
            ParagraphListBulletBaselineOffset::from_points(-3.0)?,
        )),
        Property::Indentation(_) => Property::Indentation(ParagraphListIndentation::new(
            ParagraphListLabelIndent::from_points(12.0)?,
            ParagraphListTextGap::from_em(0.5)?,
        )),
        Property::LabelColor(_) => Property::LabelColor(ParagraphListLabelColor::Explicit(color)),
        Property::NumberFormat(_) => Property::NumberFormat(ParagraphListNumberFormat::Circled),
        Property::NumberScale(_) => {
            Property::NumberScale(ParagraphListNumberScale::from_ratio(0.8)?)
        },
        Property::NumberTiering(_) => Property::NumberTiering(ParagraphListNumberTiering::Flat),
        other => return Err(format!("no second sample for {other:?}").into()),
    })
}
