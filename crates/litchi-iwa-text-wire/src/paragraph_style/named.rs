//! Named paragraph styles: the theme-visible catalog of one stylesheet and
//! the storage-scoped apply, create, rename, delete and redefine verbs of
//! the `litchi-iwa` host's `text::Draft` (record 0374).
//!
//! A named style is a `TSWP.ParagraphStyleArchive` listed in a theme's text
//! presets (`TSWP.ThemePresetsArchive.paragraph_style_presets`, field 7 of
//! theme extension 110) whose theme names the storage's stylesheet as its
//! document stylesheet. It must be registered exactly once in that
//! stylesheet, must not be a variation and must have a name.
//!
//! The engine owns payloads only. The caller (a format package):
//!
//! - finds the stylesheet ([`stylesheet`]), its `TSS.StylesheetArchive`
//!   payload and every theme payload (message types [`THEME_MESSAGE_TYPES`]),
//!   reads their presets with [`theme_presets`] and builds the [`Catalog`];
//! - says whether a style is exclusive to the storage and whether a parent
//!   is still referenced ([`plan_apply`]);
//! - allocates identifiers, clones objects, keeps archive metadata
//!   (object references, external references, object UUIDs, the identifier
//!   watermark) and checks that no other object references a style it
//!   deletes.

use litchi_iwa_common::wire::{
    append_repeated_length_delimited_field, overlay_singular_wire_fields, parse_wire_fields,
    patch_length_delimited_field, patch_nested_length_delimited_field, patch_nested_varint_field,
    patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::theme_codec::{
    DecodeError, DecodeErrorKind, DecodeLimit, DecodeOptions, ThemeArchiveSnapshot,
    ThemeArchiveWrite, ThemeExtension, ThemeReferenceSnapshot, decode_theme_archive,
    rewrite_theme_archive,
};
use litchi_iwa_text::paragraph::style::raw::{from_native_id, native_id};
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
};

use super::{Error, ManagedFields, Refusal, Result, Styles, inherit, overrides};
use crate::native::{self, Budget, NativeError, Writer};

/// Message types of the Keynote, Pages and Numbers theme archives.
pub const THEME_MESSAGE_TYPES: [u32; 3] = [10, 10_001, 12_009];
/// The prefix of the stylesheet identifier a created style is mapped by.
pub const GENERATED_IDENTIFIER_PREFIX: &str = "com.litchi.paragraph-style.";

const PARAGRAPH_PRESETS_FIELD: u32 = 7;
const TEXT_PRESET_FIELDS: std::ops::RangeInclusive<u32> = 1..=8;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_IDENTIFIER_MAP_FIELD: u32 = 2;
const STYLESHEET_PARENT_CHILDREN_FIELD: u32 = 5;
const STYLESHEET_VERSIONED_FIELDS: std::ops::RangeInclusive<u32> = 7..=22;
const VERSIONED_STYLES_FIELD: u32 = 1;
const VERSIONED_IDENTIFIER_MAP_FIELD: u32 = 2;
const VERSIONED_PARENT_CHILDREN_FIELD: u32 = 3;
const ENTRY_IDENTIFIER_FIELD: u32 = 1;
const ENTRY_STYLE_FIELD: u32 = 2;
const CHILDREN_PARENT_FIELD: u32 = 1;
const CHILDREN_FIELD: u32 = 2;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_NAME_FIELD: u32 = 1;
const STYLE_IDENTIFIER_FIELD: u32 = 2;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_FIELD: u32 = 11;
const STYLE_PARAGRAPH_FIELD: u32 = 12;
const MAX_CHAIN: usize = 64;

const CHARACTER_GROUPS: &[&[u32]] = &[
    &[1],
    &[2],
    &[3],
    &[4, 5],
    &[7, 46],
    &[13],
    &[41],
    &[10],
    &[14],
    &[27],
    &[16],
    &[43, 44],
    &[20, 21],
    &[25, 26],
    &[11],
    &[12],
    &[35],
];
const PARAGRAPH_GROUPS: &[&[u32]] = &[
    &[1],
    &[5, 6],
    &[8],
    &[9],
    &[10],
    &[14],
    &[26],
    &[41, 42],
    &[13],
    &[21],
    &[20],
    &[7],
    &[11],
    &[19],
    &[2, 3],
    &[4],
    &[25],
];
const BORDER_FIELDS: &[u32] = &[15, 16, 17, 18, 31, 32, 45, 46];
/// Borders count as four overrides (one per side), as in the host.
const BORDER_OVERRIDES: u32 = 4;

const fn invalid(reason: &'static str) -> Error {
    Error::InvalidSource { reason }
}

const fn refused(refusal: Refusal) -> Error {
    Error::Refused(refusal)
}

fn common(error: litchi_iwa_common::Error) -> Error {
    native::common(error).into()
}

/// The theme codec's own ceiling on every resource.
const THEME_CODEC_CEILING: usize = 64 * 1024 * 1024;

fn theme_options(limits: WireLimits) -> DecodeOptions {
    let cap = |value: usize| value.min(THEME_CODEC_CEILING);
    DecodeOptions::new(
        cap(limits.max_input_bytes()),
        cap(limits.max_fields()),
        cap(limits.max_rewrite_work()),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(cap(limits.max_output_bytes()))
    .with_max_allocations(cap(limits.max_fields()))
    .with_max_retained_bytes(cap(limits.max_output_bytes()))
    .with_max_scratch_bytes(cap(limits.max_output_bytes()))
}

fn theme_error(error: &DecodeError) -> Error {
    match *error.kind() {
        DecodeErrorKind::Resource {
            limit,
            observed,
            maximum,
        } => {
            let kind = match limit {
                DecodeLimit::InputBytes => LimitKind::InputBytes,
                DecodeLimit::Fields => LimitKind::Fields,
                DecodeLimit::WorkBytes => LimitKind::RewriteWork,
                DecodeLimit::Nesting => LimitKind::Nesting,
                DecodeLimit::Allocations => return Error::Allocation { amount: observed },
                _ => LimitKind::OutputBytes,
            };
            Error::Limit {
                kind,
                observed,
                maximum,
            }
        },
        DecodeErrorKind::Allocation { amount } => Error::Allocation { amount },
        _ => invalid("malformed iWork theme archive"),
    }
}

fn reference_identifier(payload: &[u8], budget: &mut Budget) -> Result<u64> {
    let mut reference = native::Reference::default();
    reference.merge(payload, 1, budget)?;
    Ok(reference.identifier)
}

fn reference_payload(identifier: u64, limits: WireLimits) -> Result<Vec<u8>> {
    let mut writer = Writer::new(limits);
    native::Reference { identifier }.encode(&mut writer)?;
    Ok(writer.into_bytes())
}

fn admit_output(output: &[u8], limits: WireLimits) -> Result<()> {
    Ok(native::admit(
        LimitKind::OutputBytes,
        output.len(),
        limits.max_output_bytes(),
    )?)
}

fn style_id(identifier: u64) -> Result<ParagraphStyleId> {
    from_native_id(identifier).map_err(|_| invalid("iWork paragraph style identifier is zero"))
}

/// The stylesheet that registers paragraph style `style`.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the style is missing, malformed or names
/// no stylesheet; resource errors as configured.
pub fn stylesheet(styles: &(impl Styles + ?Sized), style: u64, limits: WireLimits) -> Result<u64> {
    let mut budget = Budget::new(limits);
    let (decoded, _) = inherit::locate(styles, style, &mut budget)?;
    super::stylesheet_of(&decoded)
}

fn snapshot<'a>(theme: &'a [u8], limits: WireLimits) -> Result<ThemeArchiveSnapshot<'a>> {
    decode_theme_archive(theme, theme_options(limits)).map_err(|error| theme_error(&error))
}

/// The paragraph-style presets of one theme payload, in source order, or
/// `None` when the theme's document stylesheet is not `stylesheet`. A
/// theme without text presets lists none.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed theme or a zero preset;
/// resource errors as configured.
pub fn theme_presets(
    theme: &[u8],
    stylesheet: u64,
    limits: WireLimits,
) -> Result<Option<Vec<u64>>> {
    let snapshot = snapshot(theme, limits)?;
    if snapshot
        .document_stylesheet_reference()
        .map(|reference| reference.identifier)
        != Some(stylesheet)
    {
        return Ok(None);
    }
    let mut presets = Vec::new();
    for reference in snapshot.extension_references(ThemeExtension::Text, PARAGRAPH_PRESETS_FIELD) {
        if reference.identifier == 0 {
            return Err(invalid(
                "iWork stylesheet has a zero paragraph style preset",
            ));
        }
        native::reserve(&mut presets, 1)?;
        presets.push(reference.identifier);
    }
    Ok(Some(presets))
}

/// The named paragraph styles of one stylesheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Catalog {
    stylesheet: u64,
    styles: Vec<NamedParagraphStyle>,
}

impl Catalog {
    /// The stylesheet the styles belong to.
    #[must_use]
    pub const fn stylesheet(&self) -> u64 {
        self.stylesheet
    }

    /// The styles, in theme preset order.
    #[must_use]
    pub fn styles(&self) -> &[NamedParagraphStyle] {
        &self.styles
    }

    /// The named style `id`, if it is one.
    #[must_use]
    pub fn get(&self, id: ParagraphStyleId) -> Option<&NamedParagraphStyle> {
        self.styles.iter().find(|style| style.id() == id)
    }

    fn require(&self, id: ParagraphStyleId) -> Result<&NamedParagraphStyle> {
        self.get(id).ok_or(refused(Refusal::NotNamed))
    }

    /// Check a [create](named_copy) request: `source` is named and `name`
    /// is free.
    ///
    /// # Errors
    ///
    /// [`Refusal::NotNamed`] or [`Refusal::NameTaken`].
    pub fn plan_create(&self, source: ParagraphStyleId, name: &ParagraphStyleName) -> Result<()> {
        self.require(source)?;
        if self
            .styles
            .iter()
            .any(|style| style.name().as_str() == name.as_str())
        {
            return Err(refused(Refusal::NameTaken));
        }
        Ok(())
    }

    /// Plan a rename: `None` when `target` already has `name`, otherwise
    /// the renamed style ([`renamed`] rewrites its payload).
    ///
    /// # Errors
    ///
    /// [`Refusal::NotNamed`] or [`Refusal::NameTaken`].
    pub fn plan_rename(
        &self,
        target: ParagraphStyleId,
        name: &ParagraphStyleName,
    ) -> Result<Option<NamedParagraphStyle>> {
        let current = self.require(target)?;
        if self
            .styles
            .iter()
            .any(|style| style.id() != target && style.name().as_str() == name.as_str())
        {
            return Err(refused(Refusal::NameTaken));
        }
        if current.name().as_str() == name.as_str() {
            return Ok(None);
        }
        Ok(Some(NamedParagraphStyle::new(target, name.clone())))
    }

    /// Check a delete request and return the style it deletes.
    ///
    /// # Errors
    ///
    /// [`Refusal::NotNamed`] or [`Refusal::LastStyle`].
    pub fn plan_delete(&self, target: ParagraphStyleId) -> Result<NamedParagraphStyle> {
        if self.styles.len() <= 1 {
            return Err(refused(Refusal::LastStyle));
        }
        self.require(target).cloned()
    }

    /// Check a delete-with-replacement request against the storage's
    /// applied style: the storage must use `target` and `replacement`
    /// must differ from it.
    ///
    /// # Errors
    ///
    /// [`Refusal::SameReplacement`], [`Refusal::NotApplied`] or
    /// [`Refusal::NotNamed`].
    pub fn plan_replacement(
        &self,
        applied: &Applied,
        target: ParagraphStyleId,
        replacement: ParagraphStyleId,
    ) -> Result<()> {
        if target == replacement {
            return Err(refused(Refusal::SameReplacement));
        }
        if applied.style.id() != target {
            return Err(refused(Refusal::NotApplied));
        }
        self.require(replacement).map(|_| ())
    }
}

/// The styles field (1) of a stylesheet.
fn stylesheet_styles(sheet: &[u8], budget: &mut Budget) -> Result<Vec<u64>> {
    let mut styles = Vec::new();
    native::fields(sheet, 0, budget, |field, budget| {
        if field.number == STYLESHEET_STYLES_FIELD {
            let identifier = reference_identifier(field.bytes()?, budget)
                .map_err(|_| NativeError::Invalid("malformed iWork stylesheet reference"))?;
            native::reserve(&mut styles, 1)?;
            styles.push(identifier);
        }
        Ok(())
    })?;
    Ok(styles)
}

/// Build the catalog of `stylesheet` (payload `sheet`) from the union of
/// its themes' paragraph presets.
///
/// # Errors
///
/// [`Error::InvalidSource`] when there are no presets, or a preset is
/// missing, not registered exactly once in the stylesheet, a variation, of
/// another stylesheet or unnamed; resource errors as configured.
pub fn catalog(
    styles: &(impl Styles + ?Sized),
    stylesheet: u64,
    sheet: &[u8],
    presets: &[u64],
    limits: WireLimits,
) -> Result<Catalog> {
    let mut budget = Budget::new(limits);
    let registered = stylesheet_styles(sheet, &mut budget)?;
    let mut identifiers: Vec<u64> = Vec::new();
    for &preset in presets {
        if preset == 0 {
            return Err(invalid(
                "iWork stylesheet has a zero paragraph style preset",
            ));
        }
        if !identifiers.contains(&preset) {
            native::reserve(&mut identifiers, 1)?;
            identifiers.push(preset);
        }
    }
    if identifiers.is_empty() {
        return Err(invalid(
            "iWork stylesheet has no theme paragraph style presets",
        ));
    }
    let mut named = Vec::new();
    native::reserve(&mut named, identifiers.len())?;
    for identifier in identifiers {
        if registered
            .iter()
            .filter(|&&style| style == identifier)
            .count()
            != 1
        {
            return Err(invalid(
                "iWork paragraph style preset must occur exactly once in its stylesheet",
            ));
        }
        let (style, _) = inherit::locate(styles, identifier, &mut budget)?;
        if style.base.is_variation == Some(true)
            || style.base.stylesheet.map(|reference| reference.identifier) != Some(stylesheet)
        {
            return Err(invalid(
                "iWork paragraph style preset is not a named style in its stylesheet",
            ));
        }
        let name = style
            .base
            .name
            .ok_or(invalid("iWork paragraph style preset has no name"))?;
        named.push(
            NamedParagraphStyle::from_owned(style_id(identifier)?, name)
                .map_err(|_| invalid("iWork paragraph style preset name is not representable"))?,
        );
    }
    Ok(Catalog {
        stylesheet,
        styles: named,
    })
}

/// The named style a storage uses, with its override chain.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Applied {
    style: NamedParagraphStyle,
    chain: Vec<u64>,
}

impl Applied {
    /// The selected named style.
    #[must_use]
    pub const fn style(&self) -> &NamedParagraphStyle {
        &self.style
    }

    /// Whether the storage overrides the named style.
    #[must_use]
    pub fn has_overrides(&self) -> bool {
        !self.chain.is_empty()
    }

    /// The managed variations between the storage and the named style,
    /// from the storage's own style upward.
    #[must_use]
    pub fn chain(&self) -> &[u64] {
        &self.chain
    }
}

fn resolve(
    styles: &(impl Styles + ?Sized),
    catalog: &Catalog,
    first: u64,
    budget: &mut Budget,
) -> Result<Applied> {
    let mut current = first;
    let mut chain = Vec::new();
    for _ in 0..MAX_CHAIN {
        if let Some(style) = catalog
            .styles
            .iter()
            .find(|style| native_id(style.id()) == current)
        {
            return Ok(Applied {
                style: style.clone(),
                chain,
            });
        }
        if chain.contains(&current) {
            return Err(invalid(
                "iWork paragraph-style selection contains an inheritance cycle",
            ));
        }
        let (style, raw) = inherit::locate(styles, current, budget)?;
        if overrides::direct_overrides(&style, raw, budget)?.is_none() {
            return Err(invalid(
                "iWork paragraph style is neither selectable nor a supported variation",
            ));
        }
        native::reserve(&mut chain, 1)?;
        chain.push(current);
        current = super::parent_of(&style)?;
    }
    Err(invalid("iWork paragraph-style override chain is too deep"))
}

/// The named style of a storage styled by `first`.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the chain reaches a style that is neither
/// named nor a managed variation, cycles or is deeper than 64; resource
/// errors as configured.
pub fn applied(
    styles: &(impl Styles + ?Sized),
    catalog: &Catalog,
    first: u64,
    limits: WireLimits,
) -> Result<Applied> {
    resolve(styles, catalog, first, &mut Budget::new(limits))
}

/// One variation an [`ApplyPlan::Replace`] removes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Removal {
    /// The variation.
    pub style: u64,
    /// Its parent.
    pub parent: u64,
    /// The stylesheet that registers it.
    pub stylesheet: u64,
}

/// How to apply a named style to a storage.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ApplyPlan {
    /// The storage already uses the style without overrides.
    Unchanged,
    /// Repoint the storage from `from` to `to`, registering the owner's
    /// reference to `to` and dropping its reference to `from` when unused.
    Reference {
        /// The storage's current style.
        from: u64,
        /// The named style.
        to: u64,
    },
    /// Repoint the storage from `from` to `to` and remove its exclusive
    /// variation chain, in order (each removal's external references are
    /// redirected to `to`).
    Replace {
        /// The storage's current style.
        from: u64,
        /// The named style.
        to: u64,
        /// The variations to remove.
        removals: Vec<Removal>,
    },
}

/// Plan applying named style `target` to a storage styled by `first`.
///
/// `exclusive` says whether `first` is referenced by this storage only and
/// is nobody's parent. `unreferenced(style, removed)` says whether `style`
/// is referenced by no storage paragraph-style table and by no paragraph
/// style as parent, counting the storage as already repointed and ignoring
/// the `removed` variations.
///
/// # Errors
///
/// [`Refusal::NotNamed`] when `target` is not named; as [`applied`].
pub fn plan_apply(
    styles: &(impl Styles + ?Sized),
    catalog: &Catalog,
    first: u64,
    target: ParagraphStyleId,
    exclusive: bool,
    mut unreferenced: impl FnMut(u64, &[u64]) -> bool,
    limits: WireLimits,
) -> Result<ApplyPlan> {
    let mut budget = Budget::new(limits);
    catalog.require(target)?;
    let current = resolve(styles, catalog, first, &mut budget)?;
    if current.style.id() == target && !current.has_overrides() {
        return Ok(ApplyPlan::Unchanged);
    }
    let to = native_id(target);
    let base = native_id(current.style.id());
    if base == first || !exclusive {
        return Ok(ApplyPlan::Reference { from: first, to });
    }
    let mut removals = Vec::new();
    let mut removed = Vec::new();
    let mut style = first;
    for _ in 0..MAX_CHAIN {
        let (decoded, _) = inherit::locate(styles, style, &mut budget)?;
        let parent = super::parent_of(&decoded)?;
        let stylesheet = super::stylesheet_of(&decoded)?;
        native::reserve(&mut removals, 1)?;
        native::reserve(&mut removed, 1)?;
        removals.push(Removal {
            style,
            parent,
            stylesheet,
        });
        removed.push(style);
        if parent == base || !unreferenced(parent, &removed) {
            return Ok(ApplyPlan::Replace {
                from: first,
                to,
                removals,
            });
        }
        style = parent;
        let (decoded, raw) = inherit::locate(styles, style, &mut budget)?;
        if overrides::direct_overrides(&decoded, raw, &mut budget)?.is_none() {
            return Err(invalid(
                "iWork paragraph-style variation chain contains an unsupported style",
            ));
        }
    }
    Err(invalid("iWork paragraph-style override chain is too deep"))
}

/// The stylesheet identifier a created style `identifier` is mapped by.
#[must_use]
pub fn generated_identifier(identifier: u64) -> String {
    format!("{GENERATED_IDENTIFIER_PREFIX}{identifier}")
}

/// The paragraph-style payload of a new named style copied from `source`:
/// renamed to `name`, mapped by `style_identifier` and not a variation.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed source; resource errors as
/// configured.
pub fn named_copy(
    source: &[u8],
    name: &ParagraphStyleName,
    style_identifier: &str,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let decoded = native::ParagraphStyle::decode(source, &mut budget)?;
    budget.admit_bytes(source.len())?;
    let data = patch_nested_length_delimited_field(
        source,
        &[STYLE_SUPER_FIELD, STYLE_NAME_FIELD],
        decoded.base.name.is_some(),
        Some(name.as_str().as_bytes()),
    )
    .map_err(common)?;
    let data = patch_nested_length_delimited_field(
        &data,
        &[STYLE_SUPER_FIELD, STYLE_IDENTIFIER_FIELD],
        decoded.base.style_identifier.is_some(),
        Some(style_identifier.as_bytes()),
    )
    .map_err(common)?;
    let data = patch_nested_varint_field(
        &data,
        &[STYLE_SUPER_FIELD, STYLE_VARIATION_FIELD],
        decoded.base.is_variation.is_some(),
        None,
    )
    .map_err(common)?;
    admit_output(&data, limits)?;
    Ok(data)
}

/// A named style's payload renamed to `name`.
///
/// # Errors
///
/// [`Error::InvalidSource`] for a malformed payload; resource errors as
/// configured.
pub fn renamed(payload: &[u8], name: &ParagraphStyleName, limits: WireLimits) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let decoded = native::ParagraphStyle::decode(payload, &mut budget)?;
    budget.admit_bytes(payload.len())?;
    let data = patch_nested_length_delimited_field(
        payload,
        &[STYLE_SUPER_FIELD, STYLE_NAME_FIELD],
        decoded.base.name.is_some(),
        Some(name.as_str().as_bytes()),
    )
    .map_err(common)?;
    admit_output(&data, limits)?;
    Ok(data)
}

/// `(identifier, style)` of every identifier-map entry of `data`'s field
/// `field_number`.
fn identifier_entries(
    data: &[u8],
    field_number: u32,
    depth: usize,
    budget: &mut Budget,
) -> Result<Vec<(String, u64)>> {
    let mut entries = Vec::new();
    native::fields(data, depth, budget, |field, budget| {
        if field.number == field_number {
            let mut identifier = String::new();
            let mut style = native::Reference::default();
            native::fields(field.bytes()?, depth + 1, budget, |inner, budget| {
                match inner.number {
                    ENTRY_IDENTIFIER_FIELD => identifier = inner.string(budget)?,
                    ENTRY_STYLE_FIELD => style.merge(inner.bytes()?, depth + 2, budget)?,
                    _ => {},
                }
                Ok(())
            })?;
            native::reserve(&mut entries, 1)?;
            entries.push((identifier, style.identifier));
        }
        Ok(())
    })?;
    Ok(entries)
}

/// Register a new named `style` in a `TSS.StylesheetArchive`: appended to
/// its styles (field 1) and mapped by `style_identifier` (field 2).
///
/// # Errors
///
/// [`Error::InvalidSource`] when the stylesheet is malformed or already
/// registers the style or the identifier; resource errors as configured.
pub fn stylesheet_with_named(
    sheet: &[u8],
    style: u64,
    style_identifier: &str,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    let styles = stylesheet_styles(sheet, &mut budget)?;
    let entries = identifier_entries(sheet, STYLESHEET_IDENTIFIER_MAP_FIELD, 0, &mut budget)?;
    if styles.contains(&style)
        || entries
            .iter()
            .any(|(identifier, _)| identifier == style_identifier)
    {
        return Err(invalid(
            "iWork stylesheet already contains the paragraph style",
        ));
    }
    budget.admit_bytes(sheet.len())?;
    let reference = reference_payload(style, limits)?;
    let mut writer = Writer::new(limits);
    writer.bytes(ENTRY_IDENTIFIER_FIELD, style_identifier.as_bytes())?;
    writer.bytes(ENTRY_STYLE_FIELD, &reference)?;
    let entry = writer.into_bytes();
    let data = append_repeated_length_delimited_field(sheet, STYLESHEET_STYLES_FIELD, &reference)
        .map_err(common)?;
    let data =
        append_repeated_length_delimited_field(&data, STYLESHEET_IDENTIFIER_MAP_FIELD, &entry)
            .map_err(common)?;
    admit_output(&data, limits)?;
    Ok(data)
}

/// Whether `children` entries (parent field 1, children field 2) of
/// `data`'s field `field_number` name `style`.
fn children_name(
    data: &[u8],
    field_number: u32,
    depth: usize,
    style: u64,
    budget: &mut Budget,
) -> Result<bool> {
    let mut found = false;
    native::fields(data, depth, budget, |field, budget| {
        if field.number == field_number {
            native::fields(field.bytes()?, depth + 1, budget, |inner, budget| {
                if matches!(inner.number, CHILDREN_PARENT_FIELD | CHILDREN_FIELD) {
                    let mut reference = native::Reference::default();
                    reference.merge(inner.bytes()?, depth + 2, budget)?;
                    found |= reference.identifier == style;
                }
                Ok(())
            })?;
        }
        Ok(())
    })?;
    Ok(found)
}

fn versioned_names(sheet: &[u8], style: u64, budget: &mut Budget) -> Result<bool> {
    let mut versioned = Vec::new();
    native::fields(sheet, 0, budget, |field, _| {
        if STYLESHEET_VERSIONED_FIELDS.contains(&field.number) {
            native::reserve(&mut versioned, 1)?;
            versioned.push(field.bytes()?);
        }
        Ok(())
    })?;
    for payload in versioned {
        let mut listed = false;
        native::fields(payload, 1, budget, |field, budget| {
            if field.number == VERSIONED_STYLES_FIELD {
                let mut reference = native::Reference::default();
                reference.merge(field.bytes()?, 2, budget)?;
                listed |= reference.identifier == style;
            }
            Ok(())
        })?;
        if listed
            || identifier_entries(payload, VERSIONED_IDENTIFIER_MAP_FIELD, 1, budget)?
                .iter()
                .any(|(_, entry)| *entry == style)
            || children_name(payload, VERSIONED_PARENT_CHILDREN_FIELD, 1, style, budget)?
        {
            return Ok(true);
        }
    }
    Ok(false)
}

/// Unregister a named `style` from a `TSS.StylesheetArchive`: its styles
/// entry (field 1) and identifier-map entry (field 2).
///
/// # Errors
///
/// [`Error::InvalidSource`] when the style takes part in the stylesheet's
/// inheritance map or a versioned registry, or is not registered exactly
/// once in each field; resource errors as configured.
pub fn stylesheet_without_named(sheet: &[u8], style: u64, limits: WireLimits) -> Result<Vec<u8>> {
    let mut budget = Budget::new(limits);
    if children_name(
        sheet,
        STYLESHEET_PARENT_CHILDREN_FIELD,
        0,
        style,
        &mut budget,
    )? {
        return Err(invalid(
            "iWork paragraph style participates in stylesheet inheritance",
        ));
    }
    if versioned_names(sheet, style, &mut budget)? {
        return Err(invalid(
            "iWork paragraph style participates in a versioned stylesheet registry",
        ));
    }
    budget.admit_bytes(sheet.len())?;
    let mut kept = Vec::new();
    let mut occurrences = 0usize;
    for payload in
        repeated_length_delimited_payloads(sheet, STYLESHEET_STYLES_FIELD).map_err(common)?
    {
        if reference_identifier(payload, &mut budget)? == style {
            occurrences += 1;
        } else {
            native::reserve(&mut kept, 1)?;
            kept.push(payload.to_vec());
        }
    }
    if occurrences != 1 {
        return Err(invalid(
            "iWork stylesheet must reference the paragraph style exactly once",
        ));
    }
    let data = rewrite_repeated_length_delimited_fields(sheet, STYLESHEET_STYLES_FIELD, &kept)
        .map_err(common)?;
    let mut kept = Vec::new();
    let mut occurrences = 0usize;
    for payload in repeated_length_delimited_payloads(&data, STYLESHEET_IDENTIFIER_MAP_FIELD)
        .map_err(common)?
    {
        let mut entry = native::Reference::default();
        native::fields(payload, 1, &mut budget, |field, budget| {
            if field.number == ENTRY_STYLE_FIELD {
                entry.merge(field.bytes()?, 2, budget)?;
            }
            Ok(())
        })?;
        if entry.identifier == style {
            occurrences += 1;
        } else {
            native::reserve(&mut kept, 1)?;
            kept.push(payload.to_vec());
        }
    }
    if occurrences != 1 {
        return Err(invalid(
            "iWork stylesheet identifier map must reference the paragraph style exactly once",
        ));
    }
    let data =
        rewrite_repeated_length_delimited_fields(&data, STYLESHEET_IDENTIFIER_MAP_FIELD, &kept)
            .map_err(common)?;
    admit_output(&data, limits)?;
    Ok(data)
}

fn text_presets(
    snapshot: ThemeArchiveSnapshot<'_>,
    stylesheet: u64,
) -> Result<Vec<ThemeReferenceSnapshot>> {
    if snapshot
        .document_stylesheet_reference()
        .map(|reference| reference.identifier)
        != Some(stylesheet)
    {
        return Err(invalid(
            "iWork theme does not use the paragraph style's stylesheet",
        ));
    }
    if snapshot.text_presets().is_none() {
        return Err(invalid("iWork theme has no text presets"));
    }
    let mut presets = Vec::new();
    for reference in snapshot.extension_references(ThemeExtension::Text, PARAGRAPH_PRESETS_FIELD) {
        native::reserve(&mut presets, 1)?;
        presets.push(reference);
    }
    Ok(presets)
}

fn rewrite_presets(
    theme: &[u8],
    presets: &[ThemeReferenceSnapshot],
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let write = ThemeArchiveWrite::new().with_reference_list(
        ThemeExtension::Text,
        PARAGRAPH_PRESETS_FIELD,
        presets,
    );
    rewrite_theme_archive(theme, write, theme_options(limits)).map_err(|error| theme_error(&error))
}

/// A theme payload with `style` appended to its paragraph presets, in
/// field-number position (after the last field numbered 7 or lower, as a
/// canonical encoder writes it); every other byte is kept.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the theme does not use `stylesheet`, has
/// no text presets or already lists the style; resource errors as
/// configured.
pub fn theme_with_preset(
    theme: &[u8],
    stylesheet: u64,
    style: u64,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let snapshot = snapshot(theme, limits)?;
    let presets = text_presets(snapshot, stylesheet)?;
    if presets
        .iter()
        .any(|reference| reference.identifier == style)
    {
        return Err(invalid("iWork theme already contains the paragraph style"));
    }
    let text = snapshot
        .text_presets()
        .ok_or(invalid("iWork theme has no text presets"))?;
    let fields = parse_wire_fields(text, limits).map_err(common)?;
    let at = fields
        .iter()
        .filter(|field| field.number() <= PARAGRAPH_PRESETS_FIELD)
        .map(|field| field.end())
        .max()
        .unwrap_or(0);
    let mut writer = Writer::new(limits);
    writer.bytes(PARAGRAPH_PRESETS_FIELD, &reference_payload(style, limits)?)?;
    let field = writer.into_bytes();
    let mut extension = Vec::new();
    native::reserve(&mut extension, text.len() + field.len())?;
    extension.extend_from_slice(&text[..at]);
    extension.extend_from_slice(&field);
    extension.extend_from_slice(&text[at..]);
    let write = ThemeArchiveWrite::new().with_text_presets(&extension);
    rewrite_theme_archive(theme, write, theme_options(limits)).map_err(|error| theme_error(&error))
}

/// A theme payload with `style` removed from its paragraph presets.
///
/// # Errors
///
/// [`Error::InvalidSource`] when the theme does not use `stylesheet`, has
/// no text presets, lists the style in another text preset family, does
/// not list it exactly once or would lose its last paragraph style;
/// resource errors as configured.
pub fn theme_without_preset(
    theme: &[u8],
    stylesheet: u64,
    style: u64,
    limits: WireLimits,
) -> Result<Vec<u8>> {
    let snapshot = snapshot(theme, limits)?;
    let mut presets = text_presets(snapshot, stylesheet)?;
    if TEXT_PRESET_FIELDS
        .filter(|&field| field != PARAGRAPH_PRESETS_FIELD)
        .any(|field| {
            snapshot
                .extension_references(ThemeExtension::Text, field)
                .any(|reference| reference.identifier == style)
        })
    {
        return Err(invalid(
            "iWork text theme uses the paragraph style in another preset family",
        ));
    }
    let before = presets.len();
    presets.retain(|reference| reference.identifier != style);
    if before - presets.len() != 1 {
        return Err(invalid(
            "iWork theme must reference the paragraph style exactly once",
        ));
    }
    if presets.is_empty() {
        return Err(invalid("iWork theme cannot lose its last paragraph style"));
    }
    rewrite_presets(theme, &presets, limits)
}

/// A named style's redefined payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Redefinition {
    style: u64,
    payload: Vec<u8>,
    previous_following: Option<u64>,
    following: Option<u64>,
    managed: ManagedFields,
}

impl Redefinition {
    /// The named style.
    #[must_use]
    pub const fn style(&self) -> u64 {
        self.style
    }

    /// The new payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// The named following style before the redefinition (an object
    /// reference the message metadata drops).
    #[must_use]
    pub const fn previous_following(&self) -> Option<u64> {
        self.previous_following
    }

    /// The named following style after it (an object reference the
    /// message metadata must declare).
    #[must_use]
    pub const fn following(&self) -> Option<u64> {
        self.following
    }

    /// The field paths the message's metadata must declare.
    #[must_use]
    pub const fn managed_fields(&self) -> ManagedFields {
        self.managed
    }
}

/// The host's override count of a paragraph-style payload: one per known
/// character or paragraph property group present, four for borders.
fn known_property_count(data: &[u8], budget: &mut Budget) -> Result<u32> {
    let numbers = |field: u32, budget: &mut Budget| -> Result<Vec<u32>> {
        match repeated_length_delimited_payloads(data, field)
            .map_err(common)?
            .first()
        {
            Some(payload) => Ok(native::field_numbers(payload, budget)?),
            None => Ok(Vec::new()),
        }
    };
    let character = numbers(STYLE_CHARACTER_FIELD, budget)?;
    let paragraph = numbers(STYLE_PARAGRAPH_FIELD, budget)?;
    let count = |fields: &[u32], groups: &[&[u32]]| {
        groups
            .iter()
            .filter(|group| fields.iter().any(|field| group.contains(field)))
            .count()
    };
    let borders = if paragraph.iter().any(|field| BORDER_FIELDS.contains(field)) {
        BORDER_OVERRIDES
    } else {
        0
    };
    u32::try_from(count(&character, CHARACTER_GROUPS) + count(&paragraph, PARAGRAPH_GROUPS))
        .ok()
        .and_then(|count| count.checked_add(borders))
        .ok_or(invalid("paragraph-style override count overflows"))
}

fn overlay(base: &[u8], variation: &[u8]) -> Result<Vec<u8>> {
    let mut output = base.to_vec();
    for field in [STYLE_CHARACTER_FIELD, STYLE_PARAGRAPH_FIELD] {
        let overlays = repeated_length_delimited_payloads(variation, field).map_err(common)?;
        let [overlay] = overlays.as_slice() else {
            return Err(invalid(
                "paragraph-style variation must have exactly one property message",
            ));
        };
        let bases = repeated_length_delimited_payloads(&output, field).map_err(common)?;
        if bases.len() > 1 {
            return Err(invalid("named paragraph style repeats a property message"));
        }
        let merged = match bases.first() {
            Some(base) => overlay_singular_wire_fields(base, overlay).map_err(common)?,
            None => overlay.to_vec(),
        };
        output = patch_length_delimited_field(&output, field, !bases.is_empty(), Some(&merged))
            .map_err(common)?;
    }
    Ok(output)
}

/// Fold the storage's override chain into its applied named style: the
/// named style's payload with every variation's properties overlaid
/// (nearest the named style first) and its override count adjusted.
/// The caller replaces the named style's payload, then applies it to the
/// storage with [`plan_apply`].
///
/// # Errors
///
/// [`Refusal::NoOverrides`] when the storage has none; as [`applied`].
pub fn redefine(
    styles: &(impl Styles + ?Sized),
    catalog: &Catalog,
    first: u64,
    limits: WireLimits,
) -> Result<Redefinition> {
    let mut budget = Budget::new(limits);
    let applied = resolve(styles, catalog, first, &mut budget)?;
    if !applied.has_overrides() {
        return Err(refused(Refusal::NoOverrides));
    }
    let style = native_id(applied.style.id());
    let (decoded, raw) = inherit::locate(styles, style, &mut budget)?;
    if decoded.base.is_variation == Some(true) {
        return Err(invalid("iWork paragraph style is not a named style"));
    }
    let previous_following = decoded
        .paragraph
        .as_ref()
        .and_then(|properties| properties.following_style)
        .map(|reference| reference.identifier);
    let old_known = known_property_count(raw, &mut budget)?;
    let old_overrides = decoded.override_count.unwrap_or(old_known);
    budget.admit_bytes(raw.len())?;
    let mut data = raw.to_vec();
    for &variation in applied.chain.iter().rev() {
        let (variation_style, variation_raw) = inherit::locate(styles, variation, &mut budget)?;
        if overrides::direct_overrides(&variation_style, variation_raw, &mut budget)?.is_none() {
            return Err(invalid(
                "iWork paragraph style is not a supported variation",
            ));
        }
        budget.admit_bytes(data.len())?;
        data = overlay(&data, variation_raw)?;
    }
    let new_known = known_property_count(&data, &mut budget)?;
    let new_overrides = old_overrides
        .saturating_sub(old_known)
        .checked_add(new_known)
        .ok_or(invalid("paragraph-style override count overflows"))?;
    data = patch_varint_field(
        &data,
        STYLE_OVERRIDE_COUNT_FIELD,
        decoded.override_count.is_some(),
        Some(u64::from(new_overrides)),
    )
    .map_err(common)?;
    admit_output(&data, limits)?;
    let redefined = native::ParagraphStyle::decode(&data, &mut budget)?;
    let character = redefined.character.as_ref();
    let managed = ManagedFields {
        text_fill: character.is_some_and(|properties| properties.tsd_fill.is_some()),
        text_stroke: character.is_some_and(|properties| properties.tsd_stroke.is_some()),
        text_shadow: character.is_some_and(|properties| properties.shadow.is_some()),
        text_background: character.is_some_and(|properties| properties.background_color.is_some()),
    };
    Ok(Redefinition {
        style,
        following: redefined
            .paragraph
            .as_ref()
            .and_then(|properties| properties.following_style)
            .map(|reference| reference.identifier),
        payload: data,
        previous_following,
        managed,
    })
}

#[cfg(test)]
mod tests {
    use std::collections::BTreeMap;

    use litchi_iwa_protos::{tsp, tss, tswp};
    use prost::Message;

    use super::*;

    const LIMITS: WireLimits = WireLimits::DEFAULT;

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
    }

    fn style(name: &str, stylesheet: u64) -> Vec<u8> {
        tswp::ParagraphStyleArchive {
            super_: tss::StyleArchive {
                name: Some(name.to_owned()),
                stylesheet: Some(reference(stylesheet)),
                ..Default::default()
            },
            char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
            para_properties: Some(tswp::ParagraphStylePropertiesArchive::default()),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn sheet(styles: &[u64]) -> Vec<u8> {
        tss::StylesheetArchive {
            styles: styles.iter().copied().map(reference).collect(),
            identifier_to_style_map: styles
                .iter()
                .map(|&style| tss::stylesheet_archive::IdentifiedStyleEntry {
                    identifier: format!("s{style}"),
                    style: reference(style),
                })
                .collect(),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn fixture() -> (BTreeMap<u64, Vec<u8>>, Catalog) {
        let styles = BTreeMap::from([(10, style("Body", 1)), (11, style("Title", 1))]);
        let catalog = catalog(&styles, 1, &sheet(&[10, 11]), &[10, 11, 10], LIMITS).unwrap();
        (styles, catalog)
    }

    #[test]
    fn catalog_lists_named_presets_once() {
        let (_, catalog) = fixture();
        let names: Vec<&str> = catalog
            .styles()
            .iter()
            .map(|style| style.name().as_str())
            .collect();
        assert_eq!(names, ["Body", "Title"]);
        let styles = BTreeMap::from([(10, style("Body", 2))]);
        assert!(catalog_of(&styles).is_err());
    }

    fn catalog_of(styles: &BTreeMap<u64, Vec<u8>>) -> Result<Catalog> {
        catalog(styles, 1, &sheet(&[10]), &[10], LIMITS)
    }

    #[test]
    fn catalog_refusals() {
        let (_, catalog) = fixture();
        let body = style_id(10).unwrap();
        let title = ParagraphStyleName::new("Title").unwrap();
        assert_eq!(
            catalog.plan_create(body, &title),
            Err(Error::Refused(Refusal::NameTaken))
        );
        assert_eq!(
            catalog.plan_rename(body, &ParagraphStyleName::new("Body").unwrap()),
            Ok(None)
        );
        assert_eq!(
            catalog.plan_create(style_id(99).unwrap(), &title),
            Err(Error::Refused(Refusal::NotNamed))
        );
        assert!(catalog.plan_delete(body).is_ok());
    }

    #[test]
    fn applied_without_overrides_is_unchanged() {
        let (styles, catalog) = fixture();
        let applied = applied(&styles, &catalog, 10, LIMITS).unwrap();
        assert!(!applied.has_overrides());
        let plan = plan_apply(
            &styles,
            &catalog,
            10,
            style_id(10).unwrap(),
            true,
            |_, _| true,
            LIMITS,
        );
        assert_eq!(plan, Ok(ApplyPlan::Unchanged));
        let plan = plan_apply(
            &styles,
            &catalog,
            10,
            style_id(11).unwrap(),
            true,
            |_, _| true,
            LIMITS,
        );
        assert_eq!(plan, Ok(ApplyPlan::Reference { from: 10, to: 11 }));
        assert_eq!(
            redefine(&styles, &catalog, 10, LIMITS),
            Err(Error::Refused(Refusal::NoOverrides))
        );
    }

    #[test]
    fn stylesheet_named_round_trip() {
        let source = sheet(&[10]);
        let added = stylesheet_with_named(&source, 12, "s12", LIMITS).unwrap();
        let decoded = tss::StylesheetArchive::decode(added.as_slice()).unwrap();
        assert_eq!(decoded.styles.len(), 2);
        assert_eq!(decoded.identifier_to_style_map[1].identifier, "s12");
        assert!(stylesheet_with_named(&added, 12, "other", LIMITS).is_err());
        assert_eq!(
            stylesheet_without_named(&added, 12, LIMITS).unwrap(),
            source
        );
    }

    #[test]
    fn named_copy_clears_variation_and_renames() {
        let mut payload = tswp::ParagraphStyleArchive::decode(style("Body", 1).as_slice()).unwrap();
        payload.super_.is_variation = Some(false);
        let copied = named_copy(
            &payload.encode_to_vec(),
            &ParagraphStyleName::new("Copy").unwrap(),
            "id",
            LIMITS,
        )
        .unwrap();
        let decoded = tswp::ParagraphStyleArchive::decode(copied.as_slice()).unwrap();
        assert_eq!(decoded.super_.name.as_deref(), Some("Copy"));
        assert_eq!(decoded.super_.style_identifier.as_deref(), Some("id"));
        assert_eq!(decoded.super_.is_variation, None);
    }

    #[test]
    fn override_count_groups_properties() {
        let payload = tswp::ParagraphStyleArchive {
            char_properties: Some(tswp::CharacterStylePropertiesArchive {
                bold: Some(true),
                font_name: Some("Helvetica".to_owned()),
                font_name_null: Some(false),
                ..Default::default()
            }),
            para_properties: Some(tswp::ParagraphStylePropertiesArchive {
                alignment: Some(2),
                ..Default::default()
            }),
            ..Default::default()
        }
        .encode_to_vec();
        let mut budget = Budget::new(LIMITS);
        assert_eq!(known_property_count(&payload, &mut budget).unwrap(), 3);
    }
}
