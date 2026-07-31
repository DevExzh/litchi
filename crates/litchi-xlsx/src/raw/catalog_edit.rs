//! Narrow, lossless surgery for the workbook sheet catalog.
//!
//! Only directly modeled sheet state/order, positional workbook-view fields,
//! and sheet-local defined-name scopes are regenerated. All untouched workbook
//! bytes remain exact.

use std::collections::{HashMap, HashSet};

use litchi_core::xml::escape_xml;
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Error, Result, TabEditBlock, invalid};
use crate::raw::namespace::{
    STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name, relationship_attribute_value,
};

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
/// `[MS-OE376]` section 2.1.622(c) requires `activeTab` in 0..=32,766.
pub(crate) const MAX_ACTIVE_TAB: usize = 32_766;
/// `[MS-OE376]` section 2.1.613(a) limits `<sheet>` to 32,767 occurrences.
pub(crate) const MAX_SHEETS: usize = 32_767;
/// `[MS-OE376]` section 2.1.612(b) requires `sheetId` in 1..=65,534.
pub(crate) const MAX_SHEET_ID: u32 = 65_534;
/// `[MS-OE376]` section 2.1.612(c) limits relationship IDs to 255 characters.
const MAX_RELATIONSHIP_ID_CHARS: usize = 255;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Dialect {
    Transitional,
    Strict,
}

impl Dialect {
    pub(crate) const fn worksheet_namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        }
    }
}

/// Recognized sheet states that are safe to author.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum State {
    Visible,
    Hidden,
    VeryHidden,
}

impl State {
    const fn attribute(self) -> Option<&'static str> {
        match self {
            Self::Visible => None,
            Self::Hidden => Some("hidden"),
            Self::VeryHidden => Some("veryHidden"),
        }
    }
}

/// One borrowed semantic tab change. Physical relationship IDs never escape
/// this low-level boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Tab<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_id: &'a str,
    pub(crate) state: State,
}

/// One checked semantic sheet-name change.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Rename<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_id: &'a str,
    pub(crate) name: &'a str,
}

/// Semantic active-tab target. The physical workbook view remains private to
/// this low-level boundary.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Active<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
}

/// One physical catalog record synthesized below the semantic facade.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Create<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) sheet_id: u32,
    pub(crate) relationship_id: &'a str,
    pub(crate) state: State,
}

/// Final relationship order plus semantic error context. Relationship IDs are
/// borrowed only inside this physical rewrite boundary.
#[derive(Debug)]
pub(crate) struct Order<'a> {
    pub(crate) sheet: &'a str,
    pub(crate) position: usize,
    pub(crate) relationship_ids: Vec<&'a str>,
    pub(crate) local_scopes: usize,
}

/// Move-only workbook rewrite plan.
#[derive(Debug)]
pub(crate) struct Plan<'a> {
    pub(crate) tabs: Vec<Tab<'a>>,
    pub(crate) renames: Vec<Rename<'a>>,
    /// A replacement for the first workbook view's active tab. `None` leaves
    /// its active sheet unchanged unless an order edit remaps its position.
    pub(crate) active: Option<Active<'a>>,
    /// Final sheet relationship order. `None` leaves order-dependent fields
    /// byte-exact.
    pub(crate) order: Option<Order<'a>>,
}

#[derive(Debug, Clone, Copy)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Box<str>,
    value: Box<str>,
}

#[derive(Debug, Clone)]
struct Tag {
    name: Box<str>,
    attributes: Box<[Attribute]>,
}

#[derive(Debug)]
struct Slot {
    span: Span,
    tag_end: usize,
    close_start: usize,
    tag: Tag,
    empty: bool,
}

#[derive(Debug)]
struct SheetSlot {
    relationship_id: Box<str>,
    slot: Slot,
}

#[derive(Debug)]
struct ViewSlot {
    slot: Slot,
    active: Option<usize>,
    first: Option<u32>,
}

#[derive(Debug)]
struct DefinedNameSlot {
    slot: Slot,
    local_sheet_id: Option<usize>,
}

#[derive(Debug)]
struct Container {
    slot: Slot,
    payload: bool,
}

#[derive(Debug)]
struct Layout {
    root: Tag,
    dialect: Dialect,
    sheets: Container,
    sheet_slots: Box<[SheetSlot]>,
    book_views: Option<Container>,
    workbook_views: Box<[ViewSlot]>,
    defined_names: Option<Container>,
    defined_name_slots: Box<[DefinedNameSlot]>,
    protected: bool,
    alternate_content: bool,
    alternate_dependencies: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Kind {
    Workbook,
    Sheets,
    Sheet,
    BookViews,
    WorkbookView,
    DefinedNames,
    DefinedName,
    Other,
}

#[derive(Debug, Clone, Copy)]
struct Frame {
    kind: Kind,
    alternate_content: bool,
}

#[derive(Debug)]
struct Pending {
    start: usize,
    tag_end: usize,
    tag: Tag,
}

#[derive(Debug, Default)]
struct Scanner {
    root_seen: bool,
    root: Option<Tag>,
    dialect: Option<Dialect>,
    sheets: Option<Container>,
    pending_sheets: Option<Pending>,
    sheet_slots: Vec<SheetSlot>,
    pending_sheet: Option<(Pending, Box<str>)>,
    sheets_payload: bool,
    book_views: Option<Container>,
    pending_book_views: Option<Pending>,
    workbook_views: Vec<ViewSlot>,
    pending_workbook_view: Option<(Pending, Option<usize>, Option<u32>)>,
    book_views_payload: bool,
    defined_names: Option<Container>,
    pending_defined_names: Option<Pending>,
    defined_name_slots: Vec<DefinedNameSlot>,
    pending_defined_name: Option<(Pending, Option<usize>)>,
    defined_names_payload: bool,
    protected: bool,
    alternate_content: bool,
    alternate_dependencies: bool,
}

#[derive(Debug)]
struct Replacement {
    span: Span,
    bytes: Vec<u8>,
}

const FIRST_SHEET_SENTINEL: u32 = 4_294_967_286;

struct OrderMap {
    old_to_new: Vec<usize>,
    new_to_old: Vec<usize>,
}

/// Rewrite recognized tab state/order and their positional dependencies. The
/// caller reparses and verifies the semantic result before publish.
pub(crate) fn rewrite(content: &[u8], plan: Plan<'_>) -> Result<Vec<u8>> {
    if plan.tabs.is_empty()
        && plan.renames.is_empty()
        && plan.active.is_none()
        && plan.order.is_none()
    {
        return Ok(content.to_vec());
    }
    let layout = scan(content)?;
    let Plan {
        tabs,
        renames,
        active,
        order,
    } = plan;
    let first = tabs.first().copied();
    let first_rename = renames.first().copied();
    let order_context = order.as_ref().map(|order| (order.sheet, order.position));
    let context = order_context
        .or_else(|| active.map(|active| (active.sheet, active.position)))
        .or_else(|| first.map(|tab| (tab.sheet, tab.position)))
        .or_else(|| first_rename.map(|rename| (rename.sheet, rename.position)));
    if layout.protected && (!tabs.is_empty() || !renames.is_empty() || order.is_some()) {
        return Err(block(context, TabEditBlock::ProtectedWorkbook));
    }
    if !renames.is_empty() && (layout.sheets.payload || layout.alternate_dependencies) {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if let Some(active) = active
        && active.position > MAX_ACTIVE_TAB
    {
        return Err(block(
            Some((active.sheet, active.position)),
            TabEditBlock::ActiveTabLimit,
        ));
    }

    let order_map = order
        .as_ref()
        .map(|order| validate_order(&layout, order))
        .transpose()?;

    let mut replacements = Vec::new();
    replacements
        .try_reserve(
            tabs.len()
                .saturating_add(renames.len())
                .saturating_add(order_map.as_ref().map_or(0, |_| layout.sheet_slots.len()))
                .saturating_add(layout.workbook_views.len())
                .saturating_add(layout.defined_name_slots.len()),
        )
        .map_err(|error| invalid(format!("cannot reserve workbook edit plan: {error}")))?;
    sheet_replacements(
        content,
        &layout,
        &tabs,
        &renames,
        order_map.as_ref(),
        &mut replacements,
    )?;

    if let Some(order_map) = order_map.as_ref() {
        view_replacements(
            content,
            &layout,
            order_map,
            active,
            context,
            &mut replacements,
        )?;
        defined_name_replacements(content, &layout, order_map, &mut replacements)?;
    } else if let Some(active) = active {
        replacements.push(active_replacement(
            content,
            &layout,
            active.position,
            Some((active.sheet, active.position)),
        )?);
    }
    replacements.sort_unstable_by_key(|replacement| replacement.span.start);
    if replacements
        .windows(2)
        .any(|pair| pair[0].span.end > pair[1].span.start)
    {
        return Err(invalid("overlapping workbook edit replacements"));
    }

    let output_len = replacements
        .iter()
        .try_fold(content.len(), |size, replacement| {
            let removed = replacement.span.end.checked_sub(replacement.span.start)?;
            size.checked_sub(removed)?
                .checked_add(replacement.bytes.len())
        })
        .ok_or_else(|| invalid("workbook edit output size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|error| invalid(format!("cannot reserve workbook edit output: {error}")))?;
    let mut cursor = 0usize;
    for replacement in replacements {
        output.extend_from_slice(&content[cursor..replacement.span.start]);
        output.extend_from_slice(&replacement.bytes);
        cursor = replacement.span.end;
    }
    output.extend_from_slice(&content[cursor..]);
    Ok(output)
}

/// Append one checked sheet catalog entry while preserving every existing
/// element byte. Positional dependencies do not move for an append; activation
/// is intentionally handled by the ordinary view rewriter after insertion.
pub(crate) fn append(content: &[u8], create: Create<'_>) -> Result<Vec<u8>> {
    let layout = scan(content)?;
    let context = Some((create.sheet, create.position));
    if layout.protected {
        return Err(block(context, TabEditBlock::ProtectedWorkbook));
    }
    if layout.sheets.payload || layout.alternate_dependencies || layout.alternate_content {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if layout.sheet_slots.len() >= MAX_SHEETS {
        return Err(block(context, TabEditBlock::SheetLimit));
    }
    if create.position != layout.sheet_slots.len() {
        return Err(invalid("new worksheet position is not the catalog tail"));
    }
    if !(1..=MAX_SHEET_ID).contains(&create.sheet_id) {
        return Err(invalid("new worksheet native sheet ID is out of range"));
    }
    if create.relationship_id.is_empty()
        || create.relationship_id.chars().count() > MAX_RELATIONSHIP_ID_CHARS
    {
        return Err(invalid("new worksheet relationship ID is out of range"));
    }
    if layout
        .sheet_slots
        .iter()
        .any(|sheet| sheet.relationship_id.as_ref() == create.relationship_id)
    {
        return Err(invalid(
            "new worksheet relationship ID already exists in the catalog",
        ));
    }

    let sheet_name = layout.sheet_slots.first().map_or_else(
        || sibling_name(&layout.sheets.slot.tag.name, "sheet"),
        |sheet| sheet.slot.tag.name.to_string(),
    );
    let relationship_name = layout
        .sheet_slots
        .first()
        .and_then(relationship_attribute_name)
        .map(str::to_owned)
        .or_else(|| relationship_attribute_from_namespaces(&layout.root))
        .or_else(|| relationship_attribute_from_namespaces(&layout.sheets.slot.tag))
        .ok_or_else(|| block(context, TabEditBlock::MarkupCompatibility))?;
    let tag = Tag {
        name: sheet_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut created = Vec::new();
    let mut attributes = vec![
        ("name", create.sheet.to_owned()),
        ("sheetId", create.sheet_id.to_string()),
        (
            relationship_name.as_str(),
            create.relationship_id.to_owned(),
        ),
    ];
    if let Some(state) = create.state.attribute() {
        attributes.push(("state", state.to_owned()));
    }
    write_tag(&mut created, &tag, true, &[], &attributes);

    if layout.sheets.slot.empty {
        let mut replacement = Vec::new();
        write_tag(&mut replacement, &layout.sheets.slot.tag, false, &[], &[]);
        replacement.extend_from_slice(&created);
        write_close(&mut replacement, &layout.sheets.slot.tag.name);
        let mut output = Vec::new();
        output
            .try_reserve_exact(
                content
                    .len()
                    .checked_sub(layout.sheets.slot.span.end - layout.sheets.slot.span.start)
                    .and_then(|size| size.checked_add(replacement.len()))
                    .ok_or_else(|| invalid("workbook append output size overflow"))?,
            )
            .map_err(|error| invalid(format!("cannot reserve workbook append output: {error}")))?;
        output.extend_from_slice(&content[..layout.sheets.slot.span.start]);
        output.extend_from_slice(&replacement);
        output.extend_from_slice(&content[layout.sheets.slot.span.end..]);
        return Ok(output);
    }

    let at = layout.sheets.slot.close_start;
    let mut output = Vec::new();
    output
        .try_reserve_exact(
            content
                .len()
                .checked_add(created.len())
                .ok_or_else(|| invalid("workbook append output size overflow"))?,
        )
        .map_err(|error| invalid(format!("cannot reserve workbook append output: {error}")))?;
    output.extend_from_slice(&content[..at]);
    output.extend_from_slice(&created);
    output.extend_from_slice(&content[at..]);
    Ok(output)
}

pub(crate) fn dialect(content: &[u8]) -> Result<Dialect> {
    Ok(scan(content)?.dialect)
}

fn relationship_attribute_name(sheet: &SheetSlot) -> Option<&str> {
    let mut found = sheet.slot.tag.attributes.iter().filter(|attribute| {
        attribute.value.as_ref() == sheet.relationship_id.as_ref()
            && attribute
                .name
                .rsplit_once(':')
                .is_some_and(|(_, local)| local == "id")
    });
    let name = found.next()?.name.as_ref();
    found.next().is_none().then_some(name)
}

fn relationship_attribute_from_namespaces(root: &Tag) -> Option<String> {
    root.attributes.iter().find_map(|attribute| {
        let prefix = attribute.name.strip_prefix("xmlns:")?;
        matches!(
            attribute.value.as_ref(),
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                | "http://purl.oclc.org/ooxml/officeDocument/relationships"
        )
        .then(|| format!("{prefix}:id"))
    })
}

fn validate_order(layout: &Layout, order: &Order<'_>) -> Result<OrderMap> {
    let context = Some((order.sheet, order.position));
    if layout.sheets.payload
        || layout
            .book_views
            .as_ref()
            .is_some_and(|views| views.payload)
        || layout
            .defined_names
            .as_ref()
            .is_some_and(|names| names.payload)
    {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if layout.alternate_dependencies
        || (layout.alternate_content && layout.workbook_views.is_empty())
        || layout
            .defined_name_slots
            .iter()
            .filter(|name| name.local_sheet_id.is_some())
            .count()
            != order.local_scopes
    {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if order.relationship_ids.len() != layout.sheet_slots.len() {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }

    let mut direct = HashMap::new();
    direct
        .try_reserve(layout.sheet_slots.len())
        .map_err(|error| invalid(format!("cannot reserve sheet-order index: {error}")))?;
    for (old, slot) in layout.sheet_slots.iter().enumerate() {
        if direct.insert(slot.relationship_id.as_ref(), old).is_some() {
            return Err(invalid(format!(
                "duplicate direct workbook sheet relationship '{}' during reorder",
                slot.relationship_id
            )));
        }
    }

    let mut seen = HashSet::new();
    seen.try_reserve(order.relationship_ids.len())
        .map_err(|error| invalid(format!("cannot reserve sheet-order validation: {error}")))?;
    let mut old_to_new = Vec::new();
    old_to_new
        .try_reserve_exact(order.relationship_ids.len())
        .map_err(|error| {
            invalid(format!(
                "cannot reserve reverse sheet-order mapping: {error}"
            ))
        })?;
    old_to_new.resize(order.relationship_ids.len(), 0usize);
    let mut new_to_old = Vec::new();
    new_to_old
        .try_reserve_exact(order.relationship_ids.len())
        .map_err(|error| invalid(format!("cannot reserve sheet-order mapping: {error}")))?;
    for (new, relationship_id) in order.relationship_ids.iter().copied().enumerate() {
        if !seen.insert(relationship_id) {
            return Err(invalid(format!(
                "sheet reorder repeats relationship '{relationship_id}'"
            )));
        }
        let Some(&old) = direct.get(relationship_id) else {
            return Err(block(context, TabEditBlock::MarkupCompatibility));
        };
        old_to_new[old] = new;
        new_to_old.push(old);
    }
    Ok(OrderMap {
        old_to_new,
        new_to_old,
    })
}

fn sheet_replacements(
    source: &[u8],
    layout: &Layout,
    tabs: &[Tab<'_>],
    renames: &[Rename<'_>],
    order: Option<&OrderMap>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    #[derive(Clone, Copy)]
    struct Update<'a> {
        sheet: &'a str,
        position: usize,
        state: Option<State>,
        name: Option<&'a str>,
    }

    let mut updates = HashMap::<&str, Update<'_>>::new();
    updates
        .try_reserve(tabs.len().saturating_add(renames.len()))
        .map_err(|error| invalid(format!("cannot reserve tab update index: {error}")))?;
    for tab in tabs {
        let update = updates.entry(tab.relationship_id).or_insert(Update {
            sheet: tab.sheet,
            position: tab.position,
            state: None,
            name: None,
        });
        if update.state.replace(tab.state).is_some() {
            return Err(invalid(format!(
                "duplicate tab state for relationship '{}'",
                tab.relationship_id
            )));
        }
    }
    for rename in renames {
        let update = updates.entry(rename.relationship_id).or_insert(Update {
            sheet: rename.sheet,
            position: rename.position,
            state: None,
            name: None,
        });
        if update.name.replace(rename.name).is_some() {
            return Err(invalid(format!(
                "duplicate tab rename for relationship '{}'",
                rename.relationship_id
            )));
        }
    }

    if let Some(order) = order {
        for (new, old) in order.new_to_old.iter().copied().enumerate() {
            let destination = &layout.sheet_slots[new];
            let selected = &layout.sheet_slots[old];
            let update = updates.remove(selected.relationship_id.as_ref());
            if new == old && update.is_none() {
                continue;
            }
            let bytes = update.map_or_else(
                || source[selected.slot.span.start..selected.slot.span.end].to_vec(),
                |update| sheet_replacement(source, &selected.slot, update.state, update.name),
            );
            replacements.push(Replacement {
                span: destination.slot.span,
                bytes,
            });
        }
    } else {
        for found in &layout.sheet_slots {
            let Some(update) = updates.remove(found.relationship_id.as_ref()) else {
                continue;
            };
            replacements.push(Replacement {
                span: found.slot.span,
                bytes: sheet_replacement(source, &found.slot, update.state, update.name),
            });
        }
    }
    if let Some(update) = updates.values().next() {
        return Err(Error::TabEditBlocked {
            sheet: update.sheet.to_owned(),
            position: update.position,
            reason: TabEditBlock::MarkupCompatibility,
        });
    }
    Ok(())
}

fn sheet_replacement(
    source: &[u8],
    slot: &Slot,
    state: Option<State>,
    name: Option<&str>,
) -> Vec<u8> {
    let mut removed = Vec::new();
    let mut appended = Vec::new();
    if let Some(state) = state {
        removed.push("state");
        if let Some(value) = state.attribute() {
            appended.push(("state", value.to_owned()));
        }
    }
    if let Some(name) = name {
        removed.push("name");
        appended.push(("name", name.to_owned()));
    }
    rewrite_slot(source, slot, &removed, &appended)
}

fn view_replacements(
    source: &[u8],
    layout: &Layout,
    order: &OrderMap,
    active: Option<Active<'_>>,
    context: Option<(&str, usize)>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    if layout.workbook_views.is_empty() {
        if let Some(active) = active {
            replacements.push(active_replacement(
                source,
                layout,
                active.position,
                context,
            )?);
        }
        return Ok(());
    }
    for (index, view) in layout.workbook_views.iter().enumerate() {
        let old_active = view.active.unwrap_or(0);
        let Some(&mapped_active) = order.old_to_new.get(old_active) else {
            return Err(block(context, TabEditBlock::ViewIndex));
        };
        let desired_active = if index == 0 {
            active.map_or(mapped_active, |active| active.position)
        } else {
            mapped_active
        };
        if desired_active > MAX_ACTIVE_TAB {
            return Err(block(context, TabEditBlock::ActiveTabLimit));
        }

        let old_first = view.first.unwrap_or(0);
        let desired_first = if old_first == FIRST_SHEET_SENTINEL {
            old_first
        } else {
            let old =
                usize::try_from(old_first).map_err(|_| block(context, TabEditBlock::ViewIndex))?;
            let mapped = order
                .old_to_new
                .get(old)
                .copied()
                .ok_or_else(|| block(context, TabEditBlock::ViewIndex))?;
            u32::try_from(mapped).map_err(|_| block(context, TabEditBlock::ViewIndex))?
        };

        let active_changed = view
            .active
            .map_or(desired_active != 0, |old| old != desired_active);
        let first_changed = view
            .first
            .map_or(desired_first != 0, |old| old != desired_first);
        if !active_changed && !first_changed {
            continue;
        }
        let mut removed = Vec::new();
        let mut appended = Vec::new();
        if active_changed {
            removed.push("activeTab");
            appended.push(("activeTab", desired_active.to_string()));
        }
        if first_changed {
            removed.push("firstSheet");
            appended.push(("firstSheet", desired_first.to_string()));
        }
        replacements.push(Replacement {
            span: view.slot.span,
            bytes: rewrite_slot(source, &view.slot, &removed, &appended),
        });
    }
    Ok(())
}

fn defined_name_replacements(
    source: &[u8],
    layout: &Layout,
    order: &OrderMap,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    for name in &layout.defined_name_slots {
        let Some(old) = name.local_sheet_id else {
            continue;
        };
        let Some(&new) = order.old_to_new.get(old) else {
            return Err(invalid(
                "defined-name scope exceeds the workbook sheet order during reorder",
            ));
        };
        if old == new {
            continue;
        }
        replacements.push(Replacement {
            span: name.slot.span,
            bytes: rewrite_slot(
                source,
                &name.slot,
                &["localSheetId"],
                &[("localSheetId", new.to_string())],
            ),
        });
    }
    Ok(())
}

fn block(context: Option<(&str, usize)>, reason: TabEditBlock) -> Error {
    context.map_or_else(
        || invalid("workbook catalog rewrite has no associated tab change"),
        |(sheet, position)| Error::TabEditBlocked {
            sheet: sheet.to_owned(),
            position,
            reason,
        },
    )
}

fn active_replacement(
    source: &[u8],
    layout: &Layout,
    active: usize,
    context: Option<(&str, usize)>,
) -> Result<Replacement> {
    let appended = [("activeTab", active.to_string())];
    if let Some(view) = layout.workbook_views.first() {
        return Ok(Replacement {
            span: view.slot.span,
            bytes: rewrite_slot(source, &view.slot, &["activeTab"], &appended),
        });
    }
    if layout.alternate_content {
        return Err(block(context, TabEditBlock::MarkupCompatibility));
    }
    if let Some(book_views) = &layout.book_views {
        if book_views.payload {
            return Err(block(context, TabEditBlock::MarkupCompatibility));
        }
        let name = sibling_name(&book_views.slot.tag.name, "workbookView");
        let view = Tag {
            name: name.into_boxed_str(),
            attributes: Box::new([]),
        };
        let mut bytes = Vec::new();
        if book_views.slot.empty {
            write_tag(&mut bytes, &book_views.slot.tag, false, &[], &[]);
            write_tag(&mut bytes, &view, true, &[], &appended);
            write_close(&mut bytes, &book_views.slot.tag.name);
        } else {
            bytes.extend_from_slice(
                &source[book_views.slot.span.start..book_views.slot.close_start],
            );
            write_tag(&mut bytes, &view, true, &[], &appended);
            bytes.extend_from_slice(&source[book_views.slot.close_start..book_views.slot.span.end]);
        }
        return Ok(Replacement {
            span: book_views.slot.span,
            bytes,
        });
    }

    let book_views_name = sibling_name(&layout.sheets.slot.tag.name, "bookViews");
    let workbook_view_name = sibling_name(&book_views_name, "workbookView");
    let book_views = Tag {
        name: book_views_name.clone().into_boxed_str(),
        attributes: Box::new([]),
    };
    let workbook_view = Tag {
        name: workbook_view_name.into_boxed_str(),
        attributes: Box::new([]),
    };
    let mut bytes = Vec::new();
    write_tag(&mut bytes, &book_views, false, &[], &[]);
    write_tag(&mut bytes, &workbook_view, true, &[], &appended);
    write_close(&mut bytes, &book_views_name);
    Ok(Replacement {
        span: Span {
            start: layout.sheets.slot.span.start,
            end: layout.sheets.slot.span.start,
        },
        bytes,
    })
}

fn rewrite_slot(
    source: &[u8],
    slot: &Slot,
    removed: &[&str],
    appended: &[(&str, String)],
) -> Vec<u8> {
    let mut output = Vec::new();
    write_tag(&mut output, &slot.tag, slot.empty, removed, appended);
    if !slot.empty {
        output.extend_from_slice(&source[slot.tag_end..slot.span.end]);
    }
    output
}

fn scan(content: &[u8]) -> Result<Layout> {
    let mut reader = NsReader::from_reader(content);
    let mut scanner = Scanner::default();
    let mut stack = Vec::<Frame>::new();
    loop {
        let event_start = position(&reader)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        let event_end = position(&reader)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                let parent = stack.last().map(|frame| frame.kind);
                let alternate_content = stack.last().is_some_and(|frame| frame.alternate_content)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                let kind = scanner.start(
                    parent,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    alternate_content,
                    event_start,
                    event_end,
                )?;
                stack.push(Frame {
                    kind,
                    alternate_content,
                });
            },
            Event::Empty(element) => {
                let alternate_content = stack.last().is_some_and(|frame| frame.alternate_content)
                    || is_mce_name(&namespace, &element, b"AlternateContent");
                scanner.empty(
                    stack.last().map(|frame| frame.kind),
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    alternate_content,
                    Span {
                        start: event_start,
                        end: event_end,
                    },
                )?;
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("workbook edit scan has an unmatched closing tag"))?;
                scanner.finish(frame, event_start, event_end)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("workbook edit scan ended inside an element"));
    }
    scanner.finish_layout()
}

impl Scanner {
    #[allow(clippy::too_many_arguments)]
    fn start(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        alternate_content: bool,
        start: usize,
        end: usize,
    ) -> Result<Kind> {
        if parent.is_none() {
            if self.root_seen || !is_spreadsheetml_name(namespace, element.name(), b"workbook") {
                return Err(invalid("workbook edit requires one SpreadsheetML root"));
            }
            self.root_seen = true;
            self.root = Some(tag(element, decoder)?);
            self.dialect = Some(
                matches!(
                    namespace,
                    ResolveResult::Bound(Namespace(value))
                        if *value == STRICT_SPREADSHEETML_NAMESPACE
                )
                .then_some(Dialect::Strict)
                .unwrap_or(Dialect::Transitional),
            );
            return Ok(Kind::Workbook);
        }
        self.observe_guard(parent, namespace, element, decoder, alternate_content)?;
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"sheets")
        {
            if self.sheets.is_some() || self.pending_sheets.is_some() {
                return Err(invalid("duplicate direct sheets element during edit"));
            }
            self.pending_sheets = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::Sheets);
        }
        if parent == Some(Kind::Sheets)
            && is_spreadsheetml_name(namespace, element.name(), b"sheet")
        {
            if self.pending_sheet.is_some() {
                return Err(invalid("nested direct sheet element during edit"));
            }
            let relationship_id = sheet_relationship(element, decoder, resolver)?;
            self.pending_sheet = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                relationship_id,
            ));
            return Ok(Kind::Sheet);
        }
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"bookViews")
        {
            if self.book_views.is_some() || self.pending_book_views.is_some() {
                return Err(invalid("duplicate direct bookViews element during edit"));
            }
            self.pending_book_views = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::BookViews);
        }
        if parent == Some(Kind::BookViews)
            && is_spreadsheetml_name(namespace, element.name(), b"workbookView")
        {
            if self.pending_workbook_view.is_some() {
                return Err(invalid("nested direct workbookView element during edit"));
            }
            self.pending_workbook_view = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                optional_usize(element, b"activeTab", decoder, "workbook activeTab")?,
                optional_u32(element, b"firstSheet", decoder, "workbook firstSheet")?,
            ));
            return Ok(Kind::WorkbookView);
        }
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            if self.defined_names.is_some() || self.pending_defined_names.is_some() {
                return Err(invalid("duplicate direct definedNames element during edit"));
            }
            self.pending_defined_names = Some(Pending {
                start,
                tag_end: end,
                tag: tag(element, decoder)?,
            });
            return Ok(Kind::DefinedNames);
        }
        if parent == Some(Kind::DefinedNames)
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            if self.pending_defined_name.is_some() {
                return Err(invalid("nested direct definedName element during edit"));
            }
            self.pending_defined_name = Some((
                Pending {
                    start,
                    tag_end: end,
                    tag: tag(element, decoder)?,
                },
                optional_usize(
                    element,
                    b"localSheetId",
                    decoder,
                    "defined name localSheetId",
                )?,
            ));
            return Ok(Kind::DefinedName);
        }
        if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        } else if parent == Some(Kind::Sheets) {
            self.sheets_payload = true;
        } else if parent == Some(Kind::DefinedNames) {
            self.defined_names_payload = true;
        }
        Ok(Kind::Other)
    }

    #[allow(clippy::too_many_arguments)]
    fn empty(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
        alternate_content: bool,
        span: Span,
    ) -> Result<()> {
        self.observe_guard(parent, namespace, element, decoder, alternate_content)?;
        if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"sheets")
        {
            if self.sheets.is_some() || self.pending_sheets.is_some() {
                return Err(invalid("duplicate direct sheets element during edit"));
            }
            self.sheets = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::Sheets)
            && is_spreadsheetml_name(namespace, element.name(), b"sheet")
        {
            self.sheet_slots.push(SheetSlot {
                relationship_id: sheet_relationship(element, decoder, resolver)?,
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
            });
        } else if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"bookViews")
        {
            if self.book_views.is_some() || self.pending_book_views.is_some() {
                return Err(invalid("duplicate direct bookViews element during edit"));
            }
            self.book_views = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::BookViews)
            && is_spreadsheetml_name(namespace, element.name(), b"workbookView")
        {
            self.workbook_views.push(ViewSlot {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                active: optional_usize(element, b"activeTab", decoder, "workbook activeTab")?,
                first: optional_u32(element, b"firstSheet", decoder, "workbook firstSheet")?,
            });
        } else if parent == Some(Kind::Workbook)
            && is_spreadsheetml_name(namespace, element.name(), b"definedNames")
        {
            if self.defined_names.is_some() || self.pending_defined_names.is_some() {
                return Err(invalid("duplicate direct definedNames element during edit"));
            }
            self.defined_names = Some(Container {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                payload: false,
            });
        } else if parent == Some(Kind::DefinedNames)
            && is_spreadsheetml_name(namespace, element.name(), b"definedName")
        {
            self.defined_name_slots.push(DefinedNameSlot {
                slot: Slot {
                    span,
                    tag_end: span.end,
                    close_start: span.end,
                    tag: tag(element, decoder)?,
                    empty: true,
                },
                local_sheet_id: optional_usize(
                    element,
                    b"localSheetId",
                    decoder,
                    "defined name localSheetId",
                )?,
            });
        } else if parent == Some(Kind::BookViews) {
            self.book_views_payload = true;
        } else if parent == Some(Kind::Sheets) {
            self.sheets_payload = true;
        } else if parent == Some(Kind::DefinedNames) {
            self.defined_names_payload = true;
        }
        Ok(())
    }

    fn observe_guard(
        &mut self,
        parent: Option<Kind>,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        alternate_content: bool,
    ) -> Result<()> {
        if is_spreadsheetml_name(namespace, element.name(), b"workbookProtection") {
            self.protected |= optional_bool(element, b"lockStructure", decoder)?.unwrap_or(false);
        }
        if matches!(
            parent,
            Some(Kind::Workbook | Kind::Sheets | Kind::BookViews | Kind::DefinedNames)
        ) && is_mce_name(namespace, element, b"AlternateContent")
        {
            self.alternate_content = true;
        }
        if alternate_content && is_order_dependency_name(namespace, element) {
            self.alternate_dependencies = true;
        }
        Ok(())
    }

    fn finish(&mut self, frame: Frame, close_start: usize, end: usize) -> Result<()> {
        match frame.kind {
            Kind::Sheet => {
                let (pending, relationship_id) = self
                    .pending_sheet
                    .take()
                    .ok_or_else(|| invalid("sheet close without workbook edit state"))?;
                self.sheet_slots.push(SheetSlot {
                    relationship_id,
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                });
            },
            Kind::Sheets => {
                let pending = self
                    .pending_sheets
                    .take()
                    .ok_or_else(|| invalid("sheets close without workbook edit state"))?;
                self.sheets = Some(Container {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    payload: self.sheets_payload,
                });
            },
            Kind::WorkbookView => {
                let (pending, active, first) = self
                    .pending_workbook_view
                    .take()
                    .ok_or_else(|| invalid("workbookView close without edit state"))?;
                self.workbook_views.push(ViewSlot {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    active,
                    first,
                });
            },
            Kind::BookViews => {
                let pending = self
                    .pending_book_views
                    .take()
                    .ok_or_else(|| invalid("bookViews close without workbook edit state"))?;
                self.book_views = Some(Container {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    payload: self.book_views_payload,
                });
            },
            Kind::DefinedName => {
                let (pending, local_sheet_id) = self
                    .pending_defined_name
                    .take()
                    .ok_or_else(|| invalid("definedName close without edit state"))?;
                self.defined_name_slots.push(DefinedNameSlot {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    local_sheet_id,
                });
            },
            Kind::DefinedNames => {
                let pending = self
                    .pending_defined_names
                    .take()
                    .ok_or_else(|| invalid("definedNames close without edit state"))?;
                self.defined_names = Some(Container {
                    slot: Slot {
                        span: Span {
                            start: pending.start,
                            end,
                        },
                        tag_end: pending.tag_end,
                        close_start,
                        tag: pending.tag,
                        empty: false,
                    },
                    payload: self.defined_names_payload,
                });
            },
            _ => {},
        }
        Ok(())
    }

    fn finish_layout(self) -> Result<Layout> {
        if !self.root_seen {
            return Err(invalid("workbook edit scan did not find a root"));
        }
        let sheets = self
            .sheets
            .ok_or_else(|| invalid("tab edits require a direct sheets element"))?;
        let root = self
            .root
            .ok_or_else(|| invalid("workbook edit scan lost its root element"))?;
        let dialect = self
            .dialect
            .ok_or_else(|| invalid("workbook edit scan lost its XML dialect"))?;
        Ok(Layout {
            root,
            dialect,
            sheets,
            sheet_slots: self.sheet_slots.into_boxed_slice(),
            book_views: self.book_views,
            workbook_views: self.workbook_views.into_boxed_slice(),
            defined_names: self.defined_names,
            defined_name_slots: self.defined_name_slots.into_boxed_slice(),
            protected: self.protected,
            alternate_content: self.alternate_content,
            alternate_dependencies: self.alternate_dependencies,
        })
    }
}

fn sheet_relationship(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Box<str>> {
    relationship_attribute_value(element, b"id", decoder, resolver)?
        .filter(|value| !value.is_empty())
        .map(String::into_boxed_str)
        .ok_or_else(|| invalid("direct workbook sheet is missing a relationship ID during edit"))
}

fn optional_bool(element: &BytesStart<'_>, name: &[u8], decoder: Decoder) -> Result<Option<bool>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| match value.as_str() {
            "1" | "true" => Ok(true),
            "0" | "false" => Ok(false),
            _ => Err(invalid(format!(
                "invalid workbook protection boolean '{value}' during edit"
            ))),
        })
        .transpose()
}

fn optional_u32(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<u32>> {
    unqualified_attribute_value(element, name, decoder)?
        .map(|value| {
            value
                .parse::<u32>()
                .map_err(|_| invalid(format!("invalid {description} '{value}' during edit")))
        })
        .transpose()
}

fn optional_usize(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: Decoder,
    description: &str,
) -> Result<Option<usize>> {
    optional_u32(element, name, decoder, description)?
        .map(|value| {
            usize::try_from(value)
                .map_err(|_| invalid(format!("{description} does not fit usize during edit")))
        })
        .transpose()
}

fn write_tag(
    output: &mut Vec<u8>,
    tag: &Tag,
    empty: bool,
    removed: &[&str],
    appended: &[(&str, String)],
) {
    output.extend_from_slice(b"<");
    output.extend_from_slice(tag.name.as_bytes());
    for attribute in &tag.attributes {
        if removed.iter().any(|name| *name == attribute.name.as_ref()) {
            continue;
        }
        output.extend_from_slice(b" ");
        output.extend_from_slice(attribute.name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(&attribute.value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    for (name, value) in appended {
        output.extend_from_slice(b" ");
        output.extend_from_slice(name.as_bytes());
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(escape_xml(value).as_bytes());
        output.extend_from_slice(b"\"");
    }
    if empty {
        output.extend_from_slice(b"/>");
    } else {
        output.extend_from_slice(b">");
    }
}

fn write_close(output: &mut Vec<u8>, name: &str) {
    output.extend_from_slice(b"</");
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b">");
}

fn tag(element: &BytesStart<'_>, decoder: Decoder) -> Result<Tag> {
    let name = std::str::from_utf8(element.name().as_ref())
        .map_err(|error| invalid(format!("workbook element name is not UTF-8: {error}")))?
        .to_owned();
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| invalid(format!("workbook attribute name is not UTF-8: {error}")))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| invalid(error.to_string()))?
            .into_owned();
        attributes.push(Attribute {
            name: name.into_boxed_str(),
            value: value.into_boxed_str(),
        });
    }
    Ok(Tag {
        name: name.into_boxed_str(),
        attributes: attributes.into_boxed_slice(),
    })
}

fn sibling_name(name: &str, local: &str) -> String {
    name.split_once(':').map_or_else(
        || local.to_owned(),
        |(prefix, _)| format!("{prefix}:{local}"),
    )
}

fn is_mce_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>, local: &[u8]) -> bool {
    element.name().local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE)
}

fn is_order_dependency_name(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    [
        b"sheets".as_slice(),
        b"sheet".as_slice(),
        b"bookViews".as_slice(),
        b"workbookView".as_slice(),
        b"definedNames".as_slice(),
        b"definedName".as_slice(),
    ]
    .iter()
    .any(|local| is_spreadsheetml_name(namespace, element.name(), local))
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| invalid("workbook XML position does not fit usize"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::{Visibility, parse_catalog};

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn plan<'a>(tabs: Vec<Tab<'a>>, active: Option<usize>) -> Plan<'a> {
        Plan {
            tabs,
            renames: Vec::new(),
            active: active.map(|position| Active {
                sheet: "Active",
                position,
            }),
            order: None,
        }
    }

    #[test]
    fn rewrites_only_selected_states_and_first_view() {
        let source = format!(
            r#"<?xml version="1.0"?><x:workbook xmlns:x="{S}" xmlns:rel="{R}" xmlns:z="urn:future"><x:bookViews><x:workbookView activeTab="0" z:keep="yes"/><x:workbookView activeTab="0"/></x:bookViews><x:sheets z:container="exact"><x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes"/><x:sheet name="Two" sheetId="2" state="hidden" rel:id="r2"/></x:sheets><x:extLst><z:data value="exact"/></x:extLst></x:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![
                    Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    },
                    Tab {
                        sheet: "Two",
                        position: 1,
                        relationship_id: "r2",
                        state: State::Visible,
                    },
                ],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes" state="hidden"/>"#
        ));
        assert!(text.contains(r#"<x:sheet name="Two" sheetId="2" rel:id="r2"/>"#));
        assert!(text.contains(r#"<x:workbookView z:keep="yes" activeTab="1"/>"#));
        assert!(text.contains(r#"<x:workbookView activeTab="0"/>"#));
        assert!(text.contains(r#"<x:extLst><z:data value="exact"/></x:extLst>"#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.active_sheet_index, 1);
        assert_eq!(catalog.sheets[0].visibility, Visibility::Hidden);
        assert_eq!(catalog.sheets[1].visibility, Visibility::Visible);
    }

    #[test]
    fn composes_name_and_visibility_on_one_lossless_sheet_slot() {
        let source = format!(
            r#"<x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future"><x:sheets><x:sheet name="Data" sheetId="7" r:id="r1" z:keep="exact"/></x:sheets></x:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            Plan {
                tabs: vec![Tab {
                    sheet: "Data",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                renames: vec![Rename {
                    sheet: "Data",
                    position: 0,
                    relationship_id: "r1",
                    name: "Input 2026",
                }],
                active: None,
                order: None,
            },
        )
        .expect("catalog rewrite");
        assert_eq!(
            std::str::from_utf8(&output).expect("UTF-8"),
            format!(
                r#"<x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future"><x:sheets><x:sheet sheetId="7" r:id="r1" z:keep="exact" state="hidden" name="Input 2026"/></x:sheets></x:workbook>"#
            )
        );
    }

    #[test]
    fn inserts_prefixed_book_views_before_sheets() {
        let source = format!(
            r#"<s:workbook xmlns:s="{S}" xmlns:r="{R}"><s:sheets><s:sheet name="One" sheetId="1" state="hidden" r:id="r1"/><s:sheet name="Two" sheetId="2" r:id="r2"/></s:sheets></s:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(
            text.contains(
                r#"<s:bookViews><s:workbookView activeTab="1"/></s:bookViews><s:sheets>"#
            )
        );
        assert_eq!(
            parse_catalog(&output).expect("catalog").active_sheet_index,
            1
        );
    }

    #[test]
    fn expands_an_empty_book_views_container() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews data="kept"/><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        assert!(
            std::str::from_utf8(&output)
                .expect("UTF-8")
                .contains(r#"<bookViews data="kept"><workbookView activeTab="1"/></bookViews>"#)
        );
    }

    #[test]
    fn blocks_structure_protection_but_preserves_unrelated_alternate_content() {
        let protected = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><workbookProtection lockStructure="1"/><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        let tab = Tab {
            sheet: "One",
            position: 0,
            relationship_id: "r1",
            state: State::Hidden,
        };
        assert!(matches!(
            rewrite(protected.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
        let activated = rewrite(protected.as_bytes(), plan(Vec::new(), Some(0)))
            .expect("structure protection permits active-tab selection");
        assert_eq!(
            parse_catalog(&activated)
                .expect("protected catalog")
                .active_sheet_index,
            0
        );

        let mce = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        let rewritten =
            rewrite(mce.as_bytes(), plan(vec![tab], None)).expect("unrelated compatibility XML");
        let text = std::str::from_utf8(&rewritten).expect("UTF-8");
        assert!(text.contains("mc:AlternateContent"));
        assert!(text.contains(r#"state="hidden""#));

        let nested_protection = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback><workbookProtection lockStructure="1"/></mc:Fallback></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(nested_protection.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
    }

    #[test]
    fn blocks_active_view_insertion_beside_alternate_content() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" state="hidden" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    }],
                    Some(1),
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
        assert!(matches!(
            rewrite(source.as_bytes(), plan(Vec::new(), Some(1))),
            Err(Error::TabEditBlocked {
                sheet,
                position: 1,
                reason: TabEditBlock::MarkupCompatibility,
            }) if sheet == "Active"
        ));
    }

    #[test]
    fn blocks_an_effective_sheet_without_a_direct_slot() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "Fallback",
                        position: 1,
                        relationship_id: "mce-rel",
                        state: State::VeryHidden,
                    }],
                    None,
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }

    #[test]
    fn active_tab_limit_is_a_typed_block() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                Plan {
                    tabs: Vec::new(),
                    renames: Vec::new(),
                    active: Some(Active {
                        sheet: "Too Far",
                        position: MAX_ACTIVE_TAB + 1,
                    }),
                    order: None,
                }
            ),
            Err(Error::TabEditBlocked {
                sheet,
                position,
                reason: TabEditBlock::ActiveTabLimit,
            }) if sheet == "Too Far" && position == MAX_ACTIVE_TAB + 1
        ));
    }

    #[test]
    fn reorders_losslessly_and_remaps_every_positional_dependency() {
        let source = format!(
            r#"<?xml version="1.0"?><x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future" xmlns:mc="{mce}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15"><mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="/exact/" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><x:bookViews><x:workbookView activeTab="2" firstSheet="1" z:keep="view-one"/><x:workbookView activeTab="0" firstSheet="{FIRST_SHEET_SENTINEL}" z:keep="view-two"/><x:workbookView z:keep="defaults"/></x:bookViews><x:sheets><x:sheet name="One" sheetId="10" r:id="r1" z:keep="one"/><x:sheet name="Two" sheetId="20" r:id="r2"/><x:sheet name="Three" sheetId="30" state="hidden" r:id="r3"/></x:sheets><x:definedNames><x:definedName name="OneLocal" localSheetId="0">One!$A$1</x:definedName><x:definedName name="ThreeLocal" localSheetId="2">Three!$A$1</x:definedName><x:definedName name="Global">1</x:definedName></x:definedNames><x:customWorkbookViews><x:customWorkbookView name="Exact" guid="{{00000000-0000-0000-0000-000000000001}}" activeSheetId="30"/></x:customWorkbookViews></x:workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        let output = rewrite(
            source.as_bytes(),
            Plan {
                tabs: vec![Tab {
                    sheet: "Two",
                    position: 1,
                    relationship_id: "r2",
                    state: State::VeryHidden,
                }],
                renames: Vec::new(),
                active: None,
                order: Some(Order {
                    sheet: "Three",
                    position: 2,
                    relationship_ids: vec!["r3", "r1", "r2"],
                    local_scopes: 2,
                }),
            },
        )
        .expect("reorder");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        let three = text.find("name=\"Three\"").expect("Three");
        let one = text.find("name=\"One\"").expect("One");
        let two = text.find("name=\"Two\"").expect("Two");
        assert!(three < one && one < two);
        assert!(text.contains(r#"name="One" sheetId="10" r:id="r1" z:keep="one""#));
        assert!(text.contains(r#"name="Two" sheetId="20" r:id="r2" state="veryHidden""#));
        assert!(text.contains(r#"z:keep="view-one" activeTab="0" firstSheet="2""#));
        assert!(text.contains(&format!(
            r#"firstSheet="{FIRST_SHEET_SENTINEL}" z:keep="view-two" activeTab="1""#
        )));
        assert!(text.contains(r#"z:keep="defaults" activeTab="1" firstSheet="1""#));
        assert!(text.contains(r#"name="OneLocal" localSheetId="1""#));
        assert!(text.contains(r#"name="ThreeLocal" localSheetId="0""#));
        assert!(text.contains(r#"activeSheetId="30""#));
        assert!(text.contains(r#"<x15ac:absPath url="/exact/""#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.active_sheet_index, 0);
        assert_eq!(
            catalog
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["Three", "One", "Two"]
        );
        assert_eq!(catalog.defined_names[0].local_sheet_id, Some(1));
        assert_eq!(catalog.defined_names[1].local_sheet_id, Some(0));
    }

    #[test]
    fn reorder_blocks_unmodeled_catalogs_and_invalid_secondary_views() {
        let order = || Order {
            sheet: "Two",
            position: 1,
            relationship_ids: vec!["r2", "r1"],
            local_scopes: 0,
        };
        for source in [
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><future/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView/><future/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets><definedNames><future/></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15"><mc:AlternateContent><mc:Choice Requires="x15"><bookViews><workbookView activeTab="1"/></bookViews></mc:Choice></mc:AlternateContent><bookViews><workbookView/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#,
                mce = String::from_utf8_lossy(MCE)
            ),
        ] {
            assert!(matches!(
                rewrite(
                    source.as_bytes(),
                    Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: None,
                        order: Some(order()),
                    }
                ),
                Err(Error::TabEditBlocked {
                    reason: TabEditBlock::MarkupCompatibility,
                    ..
                })
            ));
        }

        let invalid_view = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView/><workbookView activeTab="9"/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                invalid_view.as_bytes(),
                Plan {
                    tabs: Vec::new(),
                    renames: Vec::new(),
                    active: None,
                    order: Some(order()),
                }
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ViewIndex,
                ..
            })
        ));
    }

    #[test]
    fn append_preserves_existing_bytes_and_uses_the_document_prefixes() {
        let source = format!(
            r#"<?xml version="1.0"?><s:workbook xmlns:s="{S}" xmlns:rel="{R}" xmlns:x="urn:keep" x:exact="yes"><s:bookViews><s:workbookView activeTab="0"/></s:bookViews><s:sheets><s:sheet name="One" sheetId="7" rel:id="tab" x:keep="1"/></s:sheets><x:tail>opaque</x:tail></s:workbook>"#
        );
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "A&B",
                position: 1,
                sheet_id: 1,
                relationship_id: "rId1",
                state: State::Hidden,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<s:sheet name="One" sheetId="7" rel:id="tab" x:keep="1"/><s:sheet name="A&amp;B" sheetId="1" rel:id="rId1" state="hidden"/>"#
        ));
        assert!(text.contains(r#"<x:tail>opaque</x:tail>"#));
        assert!(text.contains(r#"<s:workbook xmlns:s="#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.sheets.len(), 2);
        assert_eq!(catalog.sheets[1].name, "A&B");
        assert!(matches!(catalog.sheets[1].visibility, Visibility::Hidden));
    }

    #[test]
    fn append_expands_an_empty_sheet_container_from_root_namespaces() {
        let source =
            format!(r#"<s:workbook xmlns:s="{S}" xmlns:rel="{R}"><s:sheets/></s:workbook>"#);
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "Only",
                position: 0,
                sheet_id: 9,
                relationship_id: "new",
                state: State::Visible,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(
            text.contains(
                r#"<s:sheets><s:sheet name="Only" sheetId="9" rel:id="new"/></s:sheets>"#
            )
        );
        assert_eq!(parse_catalog(&output).expect("catalog").sheets.len(), 1);
    }

    #[test]
    fn append_expands_an_empty_sheet_container_with_a_local_relationship_prefix() {
        let source =
            format!(r#"<s:workbook xmlns:s="{S}"><s:sheets xmlns:rel="{R}"/></s:workbook>"#);
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "Only",
                position: 0,
                sheet_id: 1,
                relationship_id: "rId1",
                state: State::Visible,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(r#"<s:sheet name="Only" sheetId="1" rel:id="rId1"/>"#));
        assert_eq!(parse_catalog(&output).expect("catalog").sheets.len(), 1);
    }
}
