//! Named paragraph-style parity (wave W1-D, record 0374): the host
//! `text::Draft` named-style families against
//! `litchi_iwa_text_wire::paragraph_style::named` published by
//! [`text_glue`](crate::text_glue). Creating a style gives it a random
//! metadata UUID on the host, so creation is tier B; apply, rename, delete
//! and redefine of existing styles are tier A.

use std::collections::BTreeMap;

use litchi_iwa::text::{Draft, TextStorageId};
use litchi_iwa_archive::iwa::{ArchiveObject, MessageInfo, RawMessage};
use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::tswp;
use litchi_iwa_text::paragraph::format::Alignment;
use litchi_iwa_text::paragraph::style::raw::native_id;
use litchi_iwa_text::paragraph::style::{ParagraphStyleId, ParagraphStyleName};
use litchi_iwa_text_wire::paragraph_style::named::{self, ApplyPlan, Catalog};
use litchi_iwa_text_wire::paragraph_style::{self, Change, Property};
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::known_fields::decode_as;
use crate::text_glue::{
    Located, PARAGRAPH_STYLE_TYPE, STYLESHEET_TYPE, Workspace, exclusive, paragraph_styles,
    sync_field_info,
};
use crate::text_style_cases::{FIXTURES, host_change};

const LIMITS: WireLimits = WireLimits::DEFAULT;
const MANAGED_PATHS: [[u32; 2]; 4] = [[11, 46], [11, 44], [11, 21], [11, 26]];

#[derive(Clone, Debug)]
pub enum Operation {
    Apply(ParagraphStyleId),
    Create(ParagraphStyleId, &'static str),
    Rename(ParagraphStyleId, &'static str),
    Delete(ParagraphStyleId),
    DeleteApplied(ParagraphStyleId, ParagraphStyleId),
    Redefine,
}

fn name(text: &str) -> TestResult<ParagraphStyleName> {
    Ok(ParagraphStyleName::new(text)?)
}

struct Context {
    storage: Located,
    first: u64,
    styles: BTreeMap<u64, Vec<u8>>,
    entries: BTreeMap<u64, String>,
    catalog: Catalog,
    themes: Vec<(Located, Vec<u64>)>,
}

impl Context {
    fn load(workspace: &Workspace, storage: u64) -> TestResult<Self> {
        let storage = workspace.storage(storage)?;
        let first = paragraph_style::uniform_style(&storage.payload, LIMITS)?;
        let (styles, entries) = paragraph_styles(workspace)?;
        let stylesheet = named::stylesheet(&styles, first, LIMITS)?;
        let sheet = workspace.message(stylesheet, STYLESHEET_TYPE)?;
        let mut themes = Vec::new();
        let mut presets = Vec::new();
        for message_type in named::THEME_MESSAGE_TYPES {
            for theme in workspace.messages(message_type)? {
                if let Some(listed) = named::theme_presets(&theme.payload, stylesheet, LIMITS)? {
                    presets.extend_from_slice(&listed);
                    themes.push((theme, listed));
                }
            }
        }
        let catalog = named::catalog(&styles, stylesheet, &sheet.payload, &presets, LIMITS)?;
        for style in catalog.styles() {
            if entries.get(&native_id(style.id())) != Some(&sheet.entry) {
                return Err("named style is not stored with its stylesheet".into());
            }
        }
        Ok(Self {
            storage,
            first,
            styles,
            entries,
            catalog,
            themes,
        })
    }

    fn entry(&self, style: u64) -> TestResult<String> {
        Ok(self.entries.get(&style).ok_or("style entry")?.clone())
    }

    fn themes_with(&self, style: u64) -> TestResult<Vec<Located>> {
        let found: Vec<Located> = self
            .themes
            .iter()
            .filter(|(_, presets)| presets.contains(&style))
            .map(|(theme, _)| theme.clone())
            .collect();
        if found.is_empty() {
            return Err("no theme lists the paragraph style".into());
        }
        Ok(found)
    }
}

/// The host's `is_unreferenced` after the storage was repointed: no other
/// storage's paragraph-style table and no paragraph style outside
/// `removed` names `style`.
fn unreferenced(
    tables: &[(u64, Vec<u64>)],
    parents: &[(u64, u64)],
    storage: u64,
    style: u64,
    removed: &[u64],
) -> bool {
    !tables
        .iter()
        .any(|(owner, styles)| *owner != storage && styles.contains(&style))
        && !parents
            .iter()
            .any(|(child, parent)| *parent == style && !removed.contains(child))
}

fn apply(workspace: &mut Workspace, storage: u64, target: ParagraphStyleId) -> TestResult {
    let context = Context::load(workspace, storage)?;
    let mut tables = Vec::new();
    for located in workspace.storages()? {
        let decoded = tswp::StorageArchive::decode(located.payload.as_slice())?;
        let styles = decoded
            .table_para_style
            .iter()
            .flat_map(|table| &table.entries)
            .filter_map(|entry| entry.object.as_ref().map(|reference| reference.identifier))
            .collect();
        tables.push((located.object, styles));
    }
    let mut parents = Vec::new();
    for (&style, payload) in &context.styles {
        let decoded = tswp::ParagraphStyleArchive::decode(payload.as_slice())?;
        if let Some(parent) = decoded.super_.parent {
            parents.push((style, parent.identifier));
        }
    }
    let plan = named::plan_apply(
        &context.styles,
        &context.catalog,
        context.first,
        target,
        exclusive(workspace, context.first)?,
        |style, removed| unreferenced(&tables, &parents, storage, style, removed),
        LIMITS,
    )?;
    let owner = context.storage.entry.clone();
    match plan {
        ApplyPlan::Unchanged => {},
        ApplyPlan::Reference { from, to } => {
            let repointed = paragraph_style::repoint(&context.storage.payload, from, to, LIMITS)?;
            workspace.replace(&context.storage, repointed, Some((from, to)))?;
            workspace.register_style_reference(&owner, &context.entry(to)?, to)?;
            workspace.unregister_owner_reference_if_unused(&owner, &context.entry(from)?, from)?;
        },
        ApplyPlan::Replace { from, to, removals } => {
            let repointed = paragraph_style::repoint(&context.storage.payload, from, to, LIMITS)?;
            workspace.replace(&context.storage, repointed, Some((from, to)))?;
            for removal in removals {
                let entry = context.entry(removal.style)?;
                workspace.remove_style_variation(
                    &entry,
                    removal.stylesheet,
                    removal.parent,
                    removal.style,
                )?;
                workspace.unregister_private_style(&owner, &entry, removal.style, Some(to))?;
                workspace.release_watermark(&[removal.style])?;
            }
        },
        other => return Err(format!("unmapped plan {other:?}").into()),
    }
    Ok(())
}

fn replace_message(
    workspace: &mut Workspace,
    located: &Located,
    payload: Vec<u8>,
    edit: impl FnOnce(&mut MessageInfo),
) -> TestResult {
    workspace.update(&located.entry, |archive| {
        let object = archive
            .object_mut(located.object)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("object disappeared".to_owned()))?;
        object.replace_message(
            located.index,
            RawMessage {
                type_: located.message_type,
                data: payload,
            },
            Default::default(),
        )?;
        edit(&mut object.archive_info.message_infos.as_mut_slice()[located.index]);
        Ok(())
    })
}

fn drop_reference(info: &mut MessageInfo, target: u64) {
    info.object_references
        .retain(|&reference| reference != target);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|&reference| reference != target);
    }
}

fn create(
    workspace: &mut Workspace,
    storage: u64,
    source: ParagraphStyleId,
    style_name: &ParagraphStyleName,
) -> TestResult {
    let context = Context::load(workspace, storage)?;
    context.catalog.plan_create(source, style_name)?;
    let source_id = native_id(source);
    let stylesheet = context.catalog.stylesheet();
    let sheet = workspace.message(stylesheet, STYLESHEET_TYPE)?;
    let themes = context.themes_with(source_id)?;
    let created = workspace.next_identifier()?;
    let identifier = named::generated_identifier(created);
    let payload = named::named_copy(
        context.styles.get(&source_id).ok_or("source style")?,
        style_name,
        &identifier,
        LIMITS,
    )?;
    let sheet_payload = named::stylesheet_with_named(&sheet.payload, created, &identifier, LIMITS)?;
    let source_message = workspace.message(source_id, PARAGRAPH_STYLE_TYPE)?;
    let archive = workspace
        .archives()?
        .into_iter()
        .find(|(entry, _)| *entry == sheet.entry)
        .ok_or("stylesheet archive")?
        .1;
    let original = archive.object(source_id).ok_or("source object")?;
    let mut object = ArchiveObject::new(created, original.messages.clone(), Default::default())?;
    object.archive_info.message_infos = original.archive_info.message_infos.clone();
    object.archive_info.should_merge = original.archive_info.should_merge;
    object.replace_message(
        source_message.index,
        RawMessage {
            type_: PARAGRAPH_STYLE_TYPE,
            data: payload,
        },
        Default::default(),
    )?;
    workspace.update(&sheet.entry, |archive| {
        let sheet_object = archive
            .object_mut(stylesheet)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("stylesheet disappeared".to_owned()))?;
        sheet_object.replace_message(
            sheet.index,
            RawMessage {
                type_: STYLESHEET_TYPE,
                data: sheet_payload,
            },
            Default::default(),
        )?;
        sheet_object.archive_info.message_infos.as_mut_slice()[sheet.index]
            .object_references
            .push(created);
        archive.insert_object(object, Default::default())?;
        Ok(())
    })?;
    for theme in &themes {
        let payload = named::theme_with_preset(&theme.payload, stylesheet, created, LIMITS)?;
        replace_message(workspace, theme, payload, |info| {
            info.object_references.push(created);
        })?;
    }
    if let Some(style_component) = workspace.component(&sheet.entry)? {
        workspace.add_object_uuid(style_component, created)?;
        for theme in &themes {
            if let Some(theme_component) = workspace.component(&theme.entry)?
                && theme_component != style_component
            {
                workspace.add_external_reference(theme_component, style_component, created)?;
            }
        }
    }
    workspace.raise_watermark(created)
}

fn rename(
    workspace: &mut Workspace,
    storage: u64,
    target: ParagraphStyleId,
    style_name: &ParagraphStyleName,
) -> TestResult {
    let context = Context::load(workspace, storage)?;
    if context.catalog.plan_rename(target, style_name)?.is_none() {
        return Ok(());
    }
    let located = workspace.message(native_id(target), PARAGRAPH_STYLE_TYPE)?;
    let payload = named::renamed(&located.payload, style_name, LIMITS)?;
    workspace.replace(&located, payload, None)
}

fn delete(workspace: &mut Workspace, storage: u64, target: ParagraphStyleId) -> TestResult {
    let context = Context::load(workspace, storage)?;
    context.catalog.plan_delete(target)?;
    let id = native_id(target);
    let stylesheet = context.catalog.stylesheet();
    let sheet = workspace.message(stylesheet, STYLESHEET_TYPE)?;
    let themes = context.themes_with(id)?;
    for (entry, archive) in workspace.archives()? {
        for object in &archive.objects {
            let object_id = object.archive_info.identifier.ok_or("object identifier")?;
            for (message, info) in object
                .messages
                .iter()
                .zip(object.archive_info.message_infos.iter())
            {
                if !info.object_references.contains(&id) {
                    continue;
                }
                let is_sheet = object_id == stylesheet && message.type_ == STYLESHEET_TYPE;
                let is_theme = themes.iter().any(|theme| {
                    theme.entry == entry
                        && theme.object == object_id
                        && theme.message_type == message.type_
                });
                if !is_sheet && !is_theme {
                    return Err("the paragraph style is still referenced".into());
                }
            }
        }
    }
    for theme in &themes {
        let payload = named::theme_without_preset(&theme.payload, stylesheet, id, LIMITS)?;
        replace_message(workspace, theme, payload, |info| drop_reference(info, id))?;
    }
    let sheet_payload = named::stylesheet_without_named(&sheet.payload, id, LIMITS)?;
    workspace.update(&sheet.entry, |archive| {
        let sheet_object = archive
            .object_mut(stylesheet)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("stylesheet disappeared".to_owned()))?;
        sheet_object.replace_message(
            sheet.index,
            RawMessage {
                type_: STYLESHEET_TYPE,
                data: sheet_payload,
            },
            Default::default(),
        )?;
        drop_reference(
            &mut sheet_object.archive_info.message_infos.as_mut_slice()[sheet.index],
            id,
        );
        archive
            .remove_object(id)
            .ok_or_else(|| litchi_iwa::Error::InvalidFormat("style disappeared".to_owned()))?;
        Ok(())
    })?;
    if let Some(component) = workspace.component(&sheet.entry)? {
        workspace.remove_external_references_to(component, id)?;
        workspace.remove_object_uuid(component, id)?;
    }
    workspace.release_watermark(&[id])
}

fn delete_applied(
    workspace: &mut Workspace,
    storage: u64,
    target: ParagraphStyleId,
    replacement: ParagraphStyleId,
) -> TestResult {
    let context = Context::load(workspace, storage)?;
    let applied = named::applied(&context.styles, &context.catalog, context.first, LIMITS)?;
    context
        .catalog
        .plan_replacement(&applied, target, replacement)?;
    apply(workspace, storage, replacement)?;
    delete(workspace, storage, target)
}

fn redefine(workspace: &mut Workspace, storage: u64) -> TestResult {
    let context = Context::load(workspace, storage)?;
    let applied = named::applied(&context.styles, &context.catalog, context.first, LIMITS)?;
    let redefinition = named::redefine(&context.styles, &context.catalog, context.first, LIMITS)?;
    let located = workspace.message(redefinition.style(), PARAGRAPH_STYLE_TYPE)?;
    let managed = redefinition.managed_fields();
    let flags = [
        managed.text_fill(),
        managed.text_stroke(),
        managed.text_shadow(),
        managed.text_background(),
    ];
    let (previous, following) = (redefinition.previous_following(), redefinition.following());
    replace_message(
        workspace,
        &located,
        redefinition.payload().to_vec(),
        |info| {
            if let Some(previous) = previous {
                info.object_references
                    .retain(|&reference| reference != previous);
            }
            if let Some(following) = following
                && !info.object_references.contains(&following)
            {
                info.object_references.push(following);
            }
            for (present, path) in flags.into_iter().zip(MANAGED_PATHS) {
                sync_field_info(&mut info.field_infos, present, path);
            }
        },
    )?;
    apply(workspace, storage, applied.style().id())
}

pub fn leaf_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut workspace = Workspace::open(source)?;
    match operation {
        Operation::Apply(target) => apply(&mut workspace, storage, *target)?,
        Operation::Create(source, text) => create(&mut workspace, storage, *source, &name(text)?)?,
        Operation::Rename(target, text) => rename(&mut workspace, storage, *target, &name(text)?)?,
        Operation::Delete(target) => delete(&mut workspace, storage, *target)?,
        Operation::DeleteApplied(target, replacement) => {
            delete_applied(&mut workspace, storage, *target, *replacement)?;
        },
        Operation::Redefine => redefine(&mut workspace, storage)?,
    }
    workspace.into_bytes()
}

pub fn host_route(source: &[u8], storage: u64, operation: &Operation) -> TestResult<Vec<u8>> {
    let mut draft = Draft::from_bytes(source)?;
    let id = TextStorageId::try_from(storage)?;
    match operation {
        Operation::Apply(target) => {
            draft.apply_named_paragraph_style(id, *target)?;
        },
        Operation::Create(source, text) => {
            draft.create_named_paragraph_style(id, *source, name(text)?)?;
        },
        Operation::Rename(target, text) => {
            draft.rename_named_paragraph_style(id, *target, name(text)?)?;
        },
        Operation::Delete(target) => {
            draft.delete_named_paragraph_style(id, *target)?;
        },
        Operation::DeleteApplied(target, replacement) => {
            draft.delete_applied_named_paragraph_style_with_replacement(
                id,
                *target,
                *replacement,
            )?;
        },
        Operation::Redefine => {
            draft.redefine_applied_named_paragraph_style(id)?;
        },
    }
    Ok(draft.to_bytes()?)
}

pub type Reads = Vec<String>;

pub fn host_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let draft = Draft::from_bytes(bytes)?;
    let id = TextStorageId::try_from(storage)?;
    Ok(vec![
        format!("{:?}", draft.named_paragraph_styles(id).ok()),
        format!(
            "{:?}",
            draft
                .applied_named_paragraph_style(id)
                .ok()
                .map(|applied| (applied.style().clone(), applied.has_overrides()))
        ),
    ])
}

pub fn leaf_reads(bytes: &[u8], storage: u64) -> TestResult<Reads> {
    let workspace = Workspace::open(bytes)?;
    let context = Context::load(&workspace, storage).ok();
    let applied = context.as_ref().and_then(|context| {
        named::applied(&context.styles, &context.catalog, context.first, LIMITS).ok()
    });
    Ok(vec![
        format!(
            "{:?}",
            context
                .as_ref()
                .map(|context| context.catalog.styles().to_vec())
        ),
        format!(
            "{:?}",
            applied.map(|applied| (applied.style().clone(), applied.has_overrides()))
        ),
    ])
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
        .read("host named-style reads", move |bytes| {
            host_reads(bytes, storage)
        })
        .read("leaf named-style reads", move |bytes| {
            leaf_reads(bytes, storage)
        })
        .known_fields(|known| {
            known.with(
                paragraph_style::MESSAGE_TYPE,
                "TSWP.ParagraphStyleArchive",
                decode_as::<tswp::ParagraphStyleArchive>,
            )
        });
    if tier == Tier::B {
        case = case.leaf_structural_metadata();
    }
    case.run()?;
    Ok(())
}

/// Storages with a resolvable named style, with their catalog.
pub fn targets(
    bytes: &[u8],
    limit: usize,
) -> TestResult<Vec<(u64, Vec<ParagraphStyleId>, ParagraphStyleId)>> {
    let draft = Draft::from_bytes(bytes)?;
    let mut found = Vec::new();
    for storage in draft.storages()? {
        let id: u64 = storage.id.to_string().parse()?;
        let (Ok(styles), Ok(applied)) = (
            draft.named_paragraph_styles(storage.id),
            draft.applied_named_paragraph_style(storage.id),
        ) else {
            continue;
        };
        if styles.len() < 2 {
            continue;
        }
        found.push((
            id,
            styles.iter().map(|style| style.id()).collect(),
            applied.style().id(),
        ));
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

pub fn overridden(input: &Fixture, storage: u64) -> TestResult<Fixture> {
    for alignment in [Alignment::Right, Alignment::Center] {
        if let Ok(bytes) = host_change(
            &input.bytes,
            storage,
            &Change::Set(Property::Alignment(alignment)),
        ) {
            let draft = Draft::from_bytes(&bytes)?;
            if draft
                .applied_named_paragraph_style(TextStorageId::try_from(storage)?)?
                .has_overrides()
            {
                return Ok(Fixture {
                    path: input.path,
                    bytes,
                });
            }
        }
    }
    Err("no override applies".into())
}

#[test]
fn text_named_paragraph_style_apply_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, styles, applied) in targets(&input.bytes, 2)? {
            let derived = overridden(&input, storage)?;
            for &target in styles.iter().take(3) {
                for base in [&input, &derived] {
                    if target == applied && std::ptr::eq(base, &input) {
                        continue;
                    }
                    run(
                        "text_named_paragraph_style_apply_parity",
                        format,
                        Tier::A,
                        base,
                        storage,
                        &Operation::Apply(target),
                    )
                    .map_err(|e| format!("{path} storage {storage} apply {target:?}: {e}"))?;
                    cases += 1;
                }
            }
            run(
                "text_named_paragraph_style_apply_parity",
                format,
                Tier::A,
                &derived,
                storage,
                &Operation::Apply(applied),
            )
            .map_err(|e| format!("{path} storage {storage} reapply: {e}"))?;
            cases += 1;
        }
    }
    let _ = Evidence::Queued;
    assert!(cases >= 20, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_named_paragraph_style_create_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, _, applied) in targets(&input.bytes, 2)? {
            run(
                "text_named_paragraph_style_create_parity",
                format,
                Tier::B,
                &input,
                storage,
                &Operation::Create(applied, "Litchi Heading"),
            )
            .map_err(|e| format!("{path} storage {storage} create: {e}"))?;
            cases += 1;
        }
    }
    let _ = Evidence::Queued;
    assert!(cases >= 6, "only {cases} cases ran");
    Ok(())
}

/// The fixture with a created copy of the applied style, and its id.
fn with_created(
    input: &Fixture,
    storage: u64,
    applied: ParagraphStyleId,
) -> TestResult<(Fixture, ParagraphStyleId)> {
    let created = with(input, storage, &Operation::Create(applied, "Litchi Copy"))?;
    let draft = Draft::from_bytes(&created.bytes)?;
    let style = draft
        .named_paragraph_styles(TextStorageId::try_from(storage)?)?
        .into_iter()
        .find(|style| style.name().as_str() == "Litchi Copy")
        .ok_or("created style")?;
    Ok((created, style.id()))
}

#[test]
fn text_named_paragraph_style_rename_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, styles, _) in targets(&input.bytes, 2)? {
            for &target in styles.iter().take(2) {
                run(
                    "text_named_paragraph_style_rename_parity",
                    format,
                    Tier::A,
                    &input,
                    storage,
                    &Operation::Rename(target, "Litchi Renamed"),
                )
                .map_err(|e| format!("{path} storage {storage} rename {target:?}: {e}"))?;
                cases += 1;
            }
        }
    }
    let _ = Evidence::Queued;
    assert!(cases >= 12, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_named_paragraph_style_delete_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, _, applied) in targets(&input.bytes, 2)? {
            let (created, copy) = with_created(&input, storage, applied)?;
            run(
                "text_named_paragraph_style_delete_parity",
                format,
                Tier::A,
                &created,
                storage,
                &Operation::Delete(copy),
            )
            .map_err(|e| format!("{path} storage {storage} delete: {e}"))?;
            let applied_copy = with(&created, storage, &Operation::Apply(copy))?;
            run(
                "text_named_paragraph_style_delete_parity",
                format,
                Tier::A,
                &applied_copy,
                storage,
                &Operation::DeleteApplied(copy, applied),
            )
            .map_err(|e| format!("{path} storage {storage} delete applied: {e}"))?;
            cases += 2;
        }
    }
    let _ = Evidence::Queued;
    assert!(cases >= 12, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_named_paragraph_style_redefine_parity() -> TestResult {
    let mut cases = 0;
    for (path, format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, _, _) in targets(&input.bytes, 2)? {
            let derived = overridden(&input, storage)?;
            run(
                "text_named_paragraph_style_redefine_parity",
                format,
                Tier::A,
                &derived,
                storage,
                &Operation::Redefine,
            )
            .map_err(|e| format!("{path} storage {storage} redefine: {e}"))?;
            cases += 1;
        }
    }
    let _ = Evidence::Queued;
    assert!(cases >= 6, "only {cases} cases ran");
    Ok(())
}

#[test]
fn text_named_paragraph_style_read_parity() -> TestResult {
    let mut reads = 0;
    for (path, _format) in FIXTURES {
        let input = fixture(path)?;
        for (storage, _, applied) in targets(&input.bytes, 3)? {
            let derived = overridden(&input, storage)?;
            let (created, _) = with_created(&input, storage, applied)?;
            for bytes in [&input.bytes, &derived.bytes, &created.bytes] {
                assert_eq!(
                    host_reads(bytes, storage)?,
                    leaf_reads(bytes, storage)?,
                    "{path} storage {storage}"
                );
                reads += 1;
            }
        }
    }
    assert!(reads >= 18, "only {reads} reads ran");
    let _ = (Tier::A, Evidence::Queued);
    Ok(())
}
