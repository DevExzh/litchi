//! Plain-text Drop Cap parity (wave W1-D, record 0374): the host
//! `text::Draft` Drop Cap family against `litchi_iwa_text_wire::drop_cap`
//! published by [`text_glue`](crate::text_glue). A new variation gets a
//! random metadata UUID on the host, so creation is tier B; rewrites and
//! removals are tier A.

use std::collections::BTreeMap;

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::drop_cap::{
    CharacterCount, DropCap, LineCount, Outdent, Padding, RaisedLines, Wrap,
};
use litchi_iwa_text::position::TextPosition;
use litchi_iwa_text_wire::drop_cap::{self, Change, Plan};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::text_glue::{Workspace, new_object};

const LIMITS: WireLimits = WireLimits::DEFAULT;

/// Native-saved fixtures whose themes carry a Drop Cap base style.
const FIXTURES: [(&str, Format); 3] = [
    (
        "iwork/keynote/chart-grid-authored-native.key",
        Format::Keynote,
    ),
    (
        "iwork/pages/chart-grid-authored-native.pages",
        Format::Pages,
    ),
    (
        "iwork/numbers/chart-grid-authored-native.numbers",
        Format::Numbers,
    ),
];

pub fn host_drop_caps(bytes: &[u8], storage: u64) -> TestResult<Vec<(TextPosition, DropCap)>> {
    Ok(Draft::from_bytes(bytes)?
        .paragraph_drop_caps(TextStorageId::try_from(storage)?)?
        .into_iter()
        .map(|placement| (placement.paragraph, placement.drop_cap))
        .collect())
}

struct Styles {
    payloads: BTreeMap<u64, Vec<u8>>,
    entries: BTreeMap<u64, String>,
    exclusive: BTreeMap<u64, bool>,
}

fn styles(workspace: &Workspace) -> TestResult<Styles> {
    let mut payloads = BTreeMap::new();
    let mut entries = BTreeMap::new();
    let mut parents = Vec::new();
    for located in workspace.messages(drop_cap::MESSAGE_TYPE)? {
        if let Ok(decoded) = tswp::DropCapStyleArchive::decode(located.payload.as_slice()) {
            parents.extend(decoded.super_.parent.map(|parent| parent.identifier));
        }
        entries.insert(located.object, located.entry);
        payloads.insert(located.object, located.payload);
    }
    let mut references = BTreeMap::<u64, usize>::new();
    for storage in workspace.storages()? {
        let decoded = tswp::StorageArchive::decode(storage.payload.as_slice())?;
        for entry in decoded
            .table_drop_cap_style
            .iter()
            .flat_map(|table| &table.entries)
        {
            if let Some(reference) = &entry.object {
                *references.entry(reference.identifier).or_default() += 1;
            }
        }
    }
    let exclusive = payloads
        .keys()
        .map(|style| {
            (
                *style,
                references.get(style) == Some(&1) && !parents.contains(style),
            )
        })
        .collect();
    Ok(Styles {
        payloads,
        entries,
        exclusive,
    })
}

/// Every Drop Cap style preset the package's themes list.
fn presets(workspace: &Workspace) -> TestResult<Vec<u64>> {
    let mut presets = Vec::new();
    for (_, archive) in workspace.archives()? {
        for object in &archive.objects {
            for message in &object.messages {
                let data = message.data.as_slice();
                let length = data.len().max(1);
                let options = litchi_iwa_protos::theme_codec::DecodeOptions::new(
                    length,
                    length.saturating_mul(4),
                    length.saturating_mul(16),
                    64,
                );
                let Ok(theme) = litchi_iwa_protos::theme_codec::decode_theme_archive(data, options)
                else {
                    continue;
                };
                if let Some(text) = theme.text_presets()
                    && let Ok(decoded) = tswp::ThemePresetsArchive::decode(text)
                {
                    presets.extend(
                        decoded
                            .dropcap_style_presets
                            .iter()
                            .map(|reference| reference.identifier)
                            .filter(|identifier| *identifier != 0),
                    );
                }
            }
        }
    }
    Ok(presets)
}

/// The stylesheet's Drop Cap base: the theme preset that is a root style of
/// the stylesheet (the host's `base_style`).
fn base(workspace: &Workspace, stylesheet: u64) -> TestResult<u64> {
    let presets = presets(workspace)?;
    let mut found = Vec::new();
    for located in workspace.messages(drop_cap::MESSAGE_TYPE)? {
        let decoded = tswp::DropCapStyleArchive::decode(located.payload.as_slice())?;
        if presets.contains(&located.object)
            && decoded.super_.parent.is_none()
            && decoded.super_.is_variation.is_none()
            && decoded.super_.stylesheet.map(|r| r.identifier) == Some(stylesheet)
        {
            found.push(located.object);
        }
    }
    match found.as_slice() {
        [base] => Ok(*base),
        _ => Err(format!("{} Drop Cap base candidates", found.len()).into()),
    }
}

pub fn leaf_drop_caps(bytes: &[u8], storage: u64) -> TestResult<Vec<(TextPosition, DropCap)>> {
    let workspace = Workspace::open(bytes)?;
    let located = workspace.storage(storage)?;
    let placements = drop_cap::placements(&located.payload, LIMITS)?;
    if placements
        .entries()
        .iter()
        .all(|(_, style)| style.is_none())
    {
        return Ok(Vec::new());
    }
    let base = base(&workspace, placements.stylesheet())?;
    let styles = styles(&workspace)?;
    let mut found = Vec::new();
    for (paragraph, style) in placements.entries() {
        let Some(style) = style.filter(|style| *style != base) else {
            continue;
        };
        let payload = styles
            .payloads
            .get(&style)
            .ok_or("missing Drop Cap style")?;
        found.push((*paragraph, drop_cap::model(payload, LIMITS)?.drop_cap()));
    }
    Ok(found)
}

pub fn host_route(
    source: &[u8],
    storage: u64,
    paragraph: TextPosition,
    change: Change,
) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match change {
        Change::Set(value) => draft.set_paragraph_drop_cap(id, paragraph, value)?,
        Change::Remove => {
            draft.remove_paragraph_drop_cap(id, paragraph)?;
        },
        other => return Err(format!("unmapped change {other:?}").into()),
    }
    Ok(draft.to_bytes()?)
}

pub fn leaf_route(
    source: &[u8],
    storage: u64,
    paragraph: TextPosition,
    change: Change,
) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    let located = workspace.storage(storage)?;
    let placements = drop_cap::placements(&located.payload, LIMITS)?;
    let stylesheet = placements.stylesheet();
    let base = base(&workspace, stylesheet)?;
    let styles = styles(&workspace)?;
    let base_entry = styles.entries.get(&base).ok_or("base entry")?.clone();
    let plan = drop_cap::plan(
        &located.payload,
        paragraph,
        base,
        |style| styles.payloads.get(&style).map(Vec::as_slice),
        |style| styles.exclusive.get(&style).copied().unwrap_or(false),
        change,
        LIMITS,
    )?;
    match plan {
        Plan::Unchanged => {},
        Plan::Rewrite { style, payload } => {
            let entry = styles.entries.get(&style).ok_or("style entry")?.clone();
            workspace.replace_object_message(&entry, style, drop_cap::MESSAGE_TYPE, payload)?;
        },
        Plan::Derive { previous, payload } => {
            let created = workspace.next_identifier()?;
            let object = new_object(
                created,
                drop_cap::MESSAGE_TYPE,
                payload,
                &drop_cap::MESSAGE_VERSION,
                &[base],
                Vec::new(),
            )?;
            workspace.insert_style_variation(&base_entry, stylesheet, base, object)?;
            let located = workspace.storage(storage)?;
            let rewrite =
                drop_cap::point(&located.payload, paragraph, previous, Some(created), LIMITS)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            workspace.replace_with_table_references(
                &located,
                rewrite.into_payload(),
                added,
                removed,
                drop_cap::TABLE_FIELD,
            )?;
            workspace.register_private_style(&located.entry, &base_entry, created)?;
            if let Some(previous) = previous {
                workspace.unregister_owner_reference_if_unused(
                    &located.entry,
                    &base_entry,
                    previous,
                )?;
            }
            workspace.raise_watermark(created)?;
        },
        Plan::Detach {
            style,
            parent,
            exclusive,
        } => {
            let entry = styles.entries.get(&style).ok_or("style entry")?.clone();
            let rewrite = drop_cap::point(&located.payload, paragraph, Some(style), None, LIMITS)?;
            let (added, removed) = (rewrite.added(), rewrite.removed());
            workspace.replace_with_table_references(
                &located,
                rewrite.into_payload(),
                added,
                removed,
                drop_cap::TABLE_FIELD,
            )?;
            if exclusive {
                workspace.remove_style_variation(&entry, stylesheet, parent, style)?;
                workspace.unregister_private_style(&located.entry, &entry, style, None)?;
                workspace.release_watermark(&[style])?;
            } else {
                workspace.unregister_owner_reference_if_unused(&located.entry, &entry, style)?;
            }
        },
        other => return Err(format!("unmapped plan {other:?}").into()),
    }
    workspace.into_bytes()
}

fn run(
    name: &'static str,
    format: Format,
    tier: Tier,
    input: &Fixture,
    storage: u64,
    paragraph: TextPosition,
    change: Change,
) -> TestResult {
    let mut case = Case::new(name, format, tier, Evidence::Queued, input)
        .host(|source| host_route(source, storage, paragraph, change))
        .leaf(|source| {
            Ok(LeafRun::output(leaf_route(
                source, storage, paragraph, change,
            )?))
        })
        .read("host drop caps", move |bytes| {
            host_drop_caps(bytes, storage)
        })
        .read("leaf drop caps", move |bytes| {
            leaf_drop_caps(bytes, storage)
        })
        .known_fields(|known| {
            known.with(
                drop_cap::MESSAGE_TYPE,
                "TSWP.DropCapStyleArchive",
                decode_as::<tswp::DropCapStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

fn samples() -> TestResult<[DropCap; 2]> {
    Ok([
        DropCap::new(LineCount::new(3)?, CharacterCount::new(1)?),
        DropCap::new(LineCount::new(4)?, CharacterCount::new(2)?)
            .with_raised_lines(RaisedLines::new(1)?)
            .with_wrap(Wrap::Contour)
            .with_padding(Padding::from_points(6.0)?)
            .with_outdent(Outdent::from_ratio(0.25)?),
    ])
}

/// Storages whose first paragraph can carry a Drop Cap on both routes.
fn targets(bytes: &[u8], limit: usize) -> TestResult<Vec<u64>> {
    let draft = Draft::from_bytes(bytes)?;
    let mut found = Vec::new();
    for storage in draft.storages()? {
        let id: u64 = storage.id.to_string().parse()?;
        let text = storage.storage.text();
        if text.is_empty() || text.starts_with('\n') {
            continue;
        }
        if host_drop_caps(bytes, id).is_err() || leaf_drop_caps(bytes, id).is_err() {
            continue;
        }
        let Ok(probe) = host_route(bytes, id, TextPosition::ZERO, Change::Set(samples()?[0]))
        else {
            continue;
        };
        if probe.as_slice() == bytes {
            continue;
        }
        found.push(id);
        if found.len() == limit {
            break;
        }
    }
    Ok(found)
}

fn with_drop_cap(input: &Fixture, storage: u64) -> TestResult<Fixture> {
    Ok(Fixture {
        path: input.path,
        bytes: host_route(
            &input.bytes,
            storage,
            TextPosition::ZERO,
            Change::Set(samples()?[0]),
        )?,
    })
}

#[test]
fn text_drop_cap_set_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            for value in samples()? {
                run(
                    "text_drop_cap_set_parity",
                    format,
                    Tier::B,
                    &input,
                    storage,
                    TextPosition::ZERO,
                    Change::Set(value),
                )
                .map_err(|e| format!("{path} storage {storage}: {e}"))?;
                cases += 1;
            }
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 6, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_drop_cap_rewrite_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            let base = with_drop_cap(&input, storage)?;
            run(
                "text_drop_cap_rewrite_parity",
                format,
                Tier::A,
                &base,
                storage,
                TextPosition::ZERO,
                Change::Set(samples()?[1]),
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 3, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_drop_cap_remove_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 2)? {
            let base = with_drop_cap(&input, storage)?;
            run(
                "text_drop_cap_remove_parity",
                format,
                Tier::A,
                &base,
                storage,
                TextPosition::ZERO,
                Change::Remove,
            )
            .map_err(|e| format!("{path} storage {storage}: {e}"))?;
            cases += 1;
        }
        let _ = Evidence::Queued;
    }
    assert!(cases >= 3, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_drop_cap_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for storage in targets(&input.bytes, 4)? {
            let base = with_drop_cap(&input, storage)?;
            for bytes in [&input.bytes, &base.bytes] {
                assert_eq!(
                    host_drop_caps(bytes, storage)?,
                    leaf_drop_caps(bytes, storage)?
                );
                reads += 1;
            }
        }
    }
    assert!(reads >= 4, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
