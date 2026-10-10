//! Bounded, source-preserving edits to SpreadsheetML font and cell-format tables.
//!
//! The semantic stylesheet writer is intentionally not used here.  A source
//! stylesheet can contain producer extensions, unknown attributes, and
//! namespace choices that the semantic model does not retain.  This codec
//! scans exact byte spans, clones only the referenced `font` and `xf` records,
//! and changes the requested cell-format properties plus collection counts.

use std::collections::{HashMap, HashSet};
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::color::Rgb;
use crate::error::{Error, Result, allocation, invalid};
use crate::raw::namespace::{STRICT_SPREADSHEETML_NAMESPACE, is_spreadsheetml_name};
use crate::style::cell_format::{Fill, GradientKind};
use crate::style::stylesheet::border::Tint;
use crate::style::stylesheet::{FIRST_CUSTOM_ID, NumberFormat, builtin_format_code};
use crate::style::{
    AlignmentChange, BorderChange, BorderColor, BorderEdge, FillChange, FillColor, FillPattern,
    FontChange, FontColor, FontColorKind, Horizontal, Indent, Line, ProtectionChange, Reading,
    Rotation, Scheme, Script, Side, Underline, Vertical,
};

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const DEFAULT_MAX_CHANGES: usize = 4_096;
const DEFAULT_MAX_FONTS: usize = 65_534;
const DEFAULT_MAX_FILLS: usize = 65_534;
const DEFAULT_MAX_BORDERS: usize = 65_534;
const DEFAULT_MAX_NUMBER_FORMATS: usize = 65_534;
const DEFAULT_MAX_CELL_XFS: usize = 65_430;
const DEFAULT_MAX_DEPTH: usize = 128;
const DEFAULT_MAX_PART_BYTES: usize = 512 * 1024 * 1024;
const DEFAULT_MAX_STAGED_BYTES: usize = crate::style::MAX_FONT_CHANGE_NAME_BYTES;
const VIRTUAL_NONE_FILL: usize = usize::MAX;
const VIRTUAL_GRAY125_FILL: usize = usize::MAX - 1;

/// Caller-selected bounds for one lossless stylesheet rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct StyleEditLimits {
    pub(crate) max_changes: usize,
    pub(crate) max_fonts: usize,
    pub(crate) max_fills: usize,
    pub(crate) max_borders: usize,
    pub(crate) max_number_formats: usize,
    pub(crate) max_cell_xfs: usize,
    pub(crate) max_depth: usize,
    pub(crate) max_input_bytes: usize,
    pub(crate) max_output_bytes: usize,
    pub(crate) max_staged_bytes: usize,
}

impl Default for StyleEditLimits {
    fn default() -> Self {
        Self {
            max_changes: DEFAULT_MAX_CHANGES,
            max_fonts: DEFAULT_MAX_FONTS,
            max_fills: DEFAULT_MAX_FILLS,
            max_borders: DEFAULT_MAX_BORDERS,
            max_number_formats: DEFAULT_MAX_NUMBER_FORMATS,
            max_cell_xfs: DEFAULT_MAX_CELL_XFS,
            max_depth: DEFAULT_MAX_DEPTH,
            max_input_bytes: DEFAULT_MAX_PART_BYTES,
            max_output_bytes: DEFAULT_MAX_PART_BYTES,
            max_staged_bytes: DEFAULT_MAX_STAGED_BYTES,
        }
    }
}

/// One style request against the source cell-format table.
#[derive(Debug, Clone)]
pub(crate) struct StyleRequest {
    pub(crate) style: u32,
    pub(crate) change: FontChange,
    pub(crate) fill: FillChange,
    pub(crate) border: BorderChange,
    pub(crate) alignment: AlignmentChange,
    pub(crate) protection: ProtectionChange,
    pub(crate) quote_prefix: Option<Option<bool>>,
    pub(crate) number_format: Option<Option<Arc<str>>>,
}

/// Result of applying a batch of style requests.
///
/// `styles` is in the same order as the requests.  `bytes` remains byte
/// identical when every requested result deduplicates against an existing
/// resource, but the mapped style indexes may still differ from the source
/// indexes.
#[derive(Debug)]
pub(crate) struct StyleRewrite {
    pub(crate) bytes: Vec<u8>,
    pub(crate) styles: Vec<u32>,
}

/// Rewrite using the native bounded profile.
#[cfg(test)]
pub(crate) fn rewrite(input: &[u8], requests: &[StyleRequest]) -> Result<StyleRewrite> {
    rewrite_with_limits(input, requests, StyleEditLimits::default())
}

/// Rewrite a stylesheet while preserving all bytes outside the requested
/// resource clones and collection counters.
pub(crate) fn rewrite_with_limits(
    input: &[u8],
    requests: &[StyleRequest],
    limits: StyleEditLimits,
) -> Result<StyleRewrite> {
    enforce_limit(
        "styles changes",
        requests.len(),
        limits.max_changes,
        litchi_core::Resource::Objects,
    )?;
    enforce_limit(
        "styles input bytes",
        input.len(),
        limits.max_input_bytes,
        litchi_core::Resource::InputBytes,
    )?;
    // Existing records remain intact, so even an empty/no-op request retains
    // at least the complete input. Check before copying that early result.
    enforce_limit(
        "styles output bytes",
        input.len(),
        limits.max_output_bytes,
        litchi_core::Resource::OutputBytes,
    )?;

    let mut styles = Vec::new();
    styles
        .try_reserve_exact(requests.len())
        .map_err(|source| allocation("font style results", source))?;
    if requests.is_empty() {
        return Ok(StyleRewrite {
            bytes: copy_bytes(input, "styles source")?,
            styles,
        });
    }
    for request in requests {
        if let Some(fill) = request.fill.replacement() {
            fill.validate()?;
        }
    }

    prove_source_projection(input, limits)?;
    let mut document = scan(input, limits.max_depth)?;
    let fonts = document.fonts.ok_or(Error::Unsupported {
        feature: "style edits require a direct styles/fonts collection",
    })?;
    let cell_xfs = document.cell_xfs.ok_or(Error::Unsupported {
        feature: "style edits require a direct styles/cellXfs collection",
    })?;
    let font_items = direct_items(&document.elements, fonts, b"font")?;
    let xf_items = direct_items(&document.elements, cell_xfs, b"xf")?;
    if font_items.is_empty() || xf_items.is_empty() {
        return Err(Error::Unsupported {
            feature: "style edits require existing font and cell-XF records",
        });
    }
    enforce_limit(
        "styles fonts",
        font_items.len(),
        limits.max_fonts,
        litchi_core::Resource::Objects,
    )?;
    enforce_limit(
        "styles cell XFs",
        xf_items.len(),
        limits.max_cell_xfs,
        litchi_core::Resource::Objects,
    )?;
    let needs_fills = requests.iter().any(|request| !request.fill.is_empty());
    let fills = if needs_fills { document.fills } else { None };
    if let Some(fills) = fills {
        let fill_items = direct_items(&document.elements, fills, b"fill")?;
        enforce_limit(
            "styles fills",
            fill_items.len(),
            limits.max_fills,
            litchi_core::Resource::Objects,
        )?;
        validate_count(input, &document.elements[fills], fill_items.len(), "fills")?;
        document.fill_items = fill_items;
        document.virtual_fills = document.fill_items.is_empty();
    } else if needs_fills {
        document.virtual_fills = true;
    }
    if needs_fills {
        validate_fill_references(&document, &xf_items)?;
        let mut source_fill_bytes = 0usize;
        for &fill in &document.fill_items {
            source_fill_bytes = source_fill_bytes
                .checked_add(document.elements[fill].end - document.elements[fill].start)
                .ok_or_else(|| invalid("styles fill bytes overflow"))?;
            enforce_limit(
                "styles source fill bytes",
                source_fill_bytes,
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
        }
    }
    let needs_borders = requests.iter().any(|request| !request.border.is_empty());
    let borders = if needs_borders {
        document.borders
    } else {
        None
    };
    if let Some(borders) = borders {
        let border_items = direct_items(&document.elements, borders, b"border")?;
        enforce_limit(
            "styles borders",
            border_items.len(),
            limits.max_borders,
            litchi_core::Resource::Objects,
        )?;
        validate_count(
            input,
            &document.elements[borders],
            border_items.len(),
            "borders",
        )?;
        document.border_items = border_items;
        document.virtual_borders = document.border_items.is_empty();
    } else if needs_borders {
        document.virtual_borders = true;
    }
    if needs_borders {
        validate_border_references(&document, &xf_items)?;
        let mut source_border_bytes = 0usize;
        for &border in &document.border_items {
            source_border_bytes = source_border_bytes
                .checked_add(document.elements[border].end - document.elements[border].start)
                .ok_or_else(|| invalid("styles border bytes overflow"))?;
            enforce_limit(
                "styles source border bytes",
                source_border_bytes,
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
        }
    }
    let needs_number_formats = requests
        .iter()
        .any(|request| request.number_format.is_some());
    let mut number_formats = if needs_number_formats {
        index_number_formats(
            &document.elements,
            document.num_fmts,
            &xf_items,
            document.cell_style_xfs,
            limits.max_number_formats,
            limits.max_staged_bytes,
        )?
    } else {
        NumberFormatIndex::empty()
    };
    document.font_properties = index_font_properties(&document.elements, &font_items)?;
    document.alignment_properties = index_alignment_properties(&document.elements, &xf_items)?;
    document.fill_properties = if needs_fills {
        index_fill_properties(&document.elements, &document.fill_items)?
    } else {
        Vec::new()
    };
    document.border_properties = if needs_borders {
        index_border_properties(
            &document.elements,
            &document.border_items,
            limits.max_staged_bytes,
        )?
    } else {
        HashMap::new()
    };
    let needs_fill_replacements = requests
        .iter()
        .any(|request| request.fill.replacement().is_some());
    if needs_fill_replacements {
        let mut replacement_fills = Vec::new();
        replacement_fills
            .try_reserve(requests.len())
            .map_err(|source| allocation("styles replacement fill references", source))?;
        for request in requests {
            if request.fill.replacement().is_none() || document.virtual_fills {
                continue;
            }
            let style = usize::try_from(request.style)
                .map_err(|_| invalid("styles cell-XF index does not fit usize"))?;
            let xf = *xf_items
                .get(style)
                .ok_or_else(|| invalid("styles cell-XF index is out of range"))?;
            let id =
                usize::try_from(attribute_u32(&document.elements[xf], b"fillId")?.unwrap_or(0))
                    .map_err(|_| invalid("styles fill index does not fit usize"))?;
            let fill = *document
                .fill_items
                .get(id)
                .ok_or_else(|| invalid("styles fill reference is out of range"))?;
            replacement_fills.push(fill);
        }
        replacement_fills.sort_unstable();
        replacement_fills.dedup();
        index_fill_replacement_metadata(
            input,
            &document.elements,
            &mut document.fill_properties,
            &replacement_fills,
            crate::style::cell_format::MAX_RECORDS,
            limits.max_staged_bytes,
        )?;
    }
    validate_count(input, &document.elements[fonts], font_items.len(), "fonts")?;
    validate_count(
        input,
        &document.elements[cell_xfs],
        xf_items.len(),
        "cellXfs",
    )?;

    // Reject an oversized record before constructing the exact-fragment lookup
    // tables.  Otherwise a request against one large source font could first
    // clone that font (and every other table record) even though the caller's
    // staged-resource budget already makes the edit impossible.
    for request in requests {
        let xf_index = usize::try_from(request.style)
            .map_err(|_| invalid("styles cell-XF index does not fit usize"))?;
        let xf = *xf_items
            .get(xf_index)
            .ok_or_else(|| invalid("styles cell-XF reference is out of range"))?;
        let font_id = attribute_u32(&document.elements[xf], b"fontId")?.unwrap_or(0);
        let font_index = usize::try_from(font_id)
            .map_err(|_| invalid("styles font index does not fit usize"))?;
        let font = *font_items
            .get(font_index)
            .ok_or_else(|| invalid("styles font reference is out of range"))?;
        let font_effective = changes_effective(input, &document, font, xf, &request.change)?;
        let alignment_effective = alignment_changes_effective(&document, xf, &request.alignment)?;
        let protection_effective =
            protection_changes_effective(&document, xf, &request.protection)?;
        let fill_effective = fill_changes_effective(&document, xf, &request.fill)?;
        let border_effective = border_changes_effective(&document, xf, &request.border)?;
        let quote_prefix_effective =
            quote_prefix_changes_effective(&document, xf, request.quote_prefix)?;
        let number_format_effective = number_format_changes_effective(
            &document,
            xf,
            request.number_format.as_ref().map(|value| value.as_deref()),
            &number_formats,
        )?;
        if !font_effective
            && !fill_effective
            && !border_effective
            && !alignment_effective
            && !protection_effective
            && !quote_prefix_effective
            && !number_format_effective
        {
            continue;
        }
        if font_effective {
            enforce_limit(
                "styles candidate font bytes",
                document.elements[font].end - document.elements[font].start,
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
        }
        if fill_effective {
            let fill = direct_fill(&document, xf)?;
            // A missing or empty fills collection is represented by the two
            // virtual default records until the first effective write.  Those
            // records have no source span to charge here; `changed_fill`
            // accounts for their generated bytes below.
            if !document.virtual_fills {
                enforce_limit(
                    "styles candidate fill bytes",
                    document.elements[fill].end - document.elements[fill].start,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
            }
        }
        if border_effective {
            let border = direct_border(&document, xf)?;
            if !document.virtual_borders {
                enforce_limit(
                    "styles candidate border bytes",
                    document.elements[border].end - document.elements[border].start,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
            }
        }
        enforce_limit(
            "styles candidate cell-XF bytes",
            document.elements[xf].end - document.elements[xf].start,
            limits.max_staged_bytes,
            litchi_core::Resource::Memory,
        )?;
    }

    let mut font_lookup = HashMap::<Vec<u8>, u32>::new();
    font_lookup
        .try_reserve(font_items.len())
        .map_err(|source| allocation("styles font lookup", source))?;
    for (index, element) in font_items.iter().copied().enumerate() {
        font_lookup.insert(
            copy_bytes(
                &input[document.elements[element].start..document.elements[element].end],
                "styles font lookup key",
            )?,
            u32::try_from(index).map_err(|_| invalid("styles font index overflows"))?,
        );
    }
    let mut xf_lookup = HashMap::<Vec<u8>, u32>::new();
    xf_lookup
        .try_reserve(xf_items.len())
        .map_err(|source| allocation("styles cell-XF lookup", source))?;
    for (index, element) in xf_items.iter().copied().enumerate() {
        xf_lookup.insert(
            copy_bytes(
                &input[document.elements[element].start..document.elements[element].end],
                "styles cell-XF lookup key",
            )?,
            u32::try_from(index).map_err(|_| invalid("styles cell-XF index overflows"))?,
        );
    }
    font_lookup
        .try_reserve(requests.len())
        .map_err(|source| allocation("styles font lookup", source))?;
    xf_lookup
        .try_reserve(requests.len())
        .map_err(|source| allocation("styles cell-XF lookup", source))?;

    let mut new_fonts = Vec::<Vec<u8>>::new();
    let mut fill_lookup = HashMap::<Vec<u8>, u32>::new();
    if needs_fills {
        fill_lookup
            .try_reserve(document.fill_items.len())
            .map_err(|source| allocation("styles fill lookup", source))?;
        for (index, element) in document.fill_items.iter().copied().enumerate() {
            fill_lookup.insert(
                copy_bytes(
                    &input[document.elements[element].start..document.elements[element].end],
                    "styles fill lookup key",
                )?,
                u32::try_from(index).map_err(|_| invalid("styles fill index overflows"))?,
            );
        }
        fill_lookup
            .try_reserve(requests.len())
            .map_err(|source| allocation("styles fill lookup", source))?;
    }
    let mut new_fills = Vec::<Vec<u8>>::new();
    let mut border_lookup = HashMap::<Vec<u8>, u32>::new();
    if needs_borders {
        border_lookup
            .try_reserve(document.border_items.len().saturating_add(requests.len()))
            .map_err(|source| allocation("styles border lookup", source))?;
        for (index, element) in document.border_items.iter().copied().enumerate() {
            border_lookup.insert(
                copy_bytes(
                    &input[document.elements[element].start..document.elements[element].end],
                    "styles border lookup key",
                )?,
                u32::try_from(index).map_err(|_| invalid("styles border index overflows"))?,
            );
        }
    }
    let mut new_borders = Vec::<Vec<u8>>::new();
    let mut new_num_fmts = Vec::<Vec<u8>>::new();
    let mut new_xfs = Vec::<Vec<u8>>::new();
    let mut staged_bytes = 0usize;
    new_fonts
        .try_reserve(requests.len())
        .map_err(|source| allocation("new styles fonts", source))?;
    new_fills
        .try_reserve(requests.len())
        .map_err(|source| allocation("new styles fills", source))?;
    new_borders
        .try_reserve(requests.len())
        .map_err(|source| allocation("new styles borders", source))?;
    new_num_fmts
        .try_reserve(requests.len())
        .map_err(|source| allocation("new styles number formats", source))?;
    new_xfs
        .try_reserve(requests.len())
        .map_err(|source| allocation("new styles cell XFs", source))?;
    let mut virtual_defaults = None;
    let mut virtual_defaults_initialized = !document.virtual_fills;
    let mut virtual_border_initialized = !document.virtual_borders;

    for request in requests {
        let xf_index = usize::try_from(request.style)
            .map_err(|_| invalid("styles cell-XF index does not fit usize"))?;
        let xf = *xf_items
            .get(xf_index)
            .ok_or_else(|| invalid("styles cell-XF reference is out of range"))?;
        let font_id = attribute_u32(&document.elements[xf], b"fontId")?.unwrap_or(0);
        let font_index = usize::try_from(font_id)
            .map_err(|_| invalid("styles font index does not fit usize"))?;
        let font = *font_items
            .get(font_index)
            .ok_or_else(|| invalid("styles font reference is out of range"))?;

        let font_effective = changes_effective(input, &document, font, xf, &request.change)?;
        let alignment_effective = alignment_changes_effective(&document, xf, &request.alignment)?;
        let protection_effective =
            protection_changes_effective(&document, xf, &request.protection)?;
        let fill_effective = fill_changes_effective(&document, xf, &request.fill)?;
        let border_effective = border_changes_effective(&document, xf, &request.border)?;
        let quote_prefix_effective =
            quote_prefix_changes_effective(&document, xf, request.quote_prefix)?;
        let number_format_effective = number_format_changes_effective(
            &document,
            xf,
            request.number_format.as_ref().map(|value| value.as_deref()),
            &number_formats,
        )?;
        if !font_effective
            && !fill_effective
            && !border_effective
            && !alignment_effective
            && !protection_effective
            && !quote_prefix_effective
            && !number_format_effective
        {
            styles.push(request.style);
            continue;
        }

        let resolved_font = if font_effective {
            enforce_limit(
                "styles candidate font bytes",
                document.elements[font].end - document.elements[font].start,
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            let font_bytes = changed_font(
                input,
                &document,
                font,
                &request.change,
                limits.max_staged_bytes,
            )?;
            enforce_limit(
                "styles candidate font bytes",
                font_bytes.len(),
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            if let Some(index) = font_lookup.get(&font_bytes).copied() {
                index
            } else {
                let index = font_items
                    .len()
                    .checked_add(new_fonts.len())
                    .ok_or_else(|| invalid("styles font count overflows"))?;
                enforce_limit(
                    "styles generated fonts",
                    index.saturating_add(1),
                    limits.max_fonts,
                    litchi_core::Resource::Objects,
                )?;
                let index =
                    u32::try_from(index).map_err(|_| invalid("styles font index overflows"))?;
                staged_bytes = staged_bytes
                    .checked_add(font_bytes.len())
                    .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                enforce_limit(
                    "styles staged resource bytes",
                    staged_bytes,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
                font_lookup.insert(font_bytes.clone(), index);
                new_fonts.push(font_bytes);
                index
            }
        } else {
            font_id
        };

        let resolved_fill = if fill_effective {
            if document.virtual_fills && !virtual_defaults_initialized {
                enforce_limit(
                    "styles default fills",
                    2,
                    limits.max_fills,
                    litchi_core::Resource::Objects,
                )?;
                fill_lookup
                    .try_reserve(2)
                    .map_err(|source| allocation("styles fill lookup", source))?;
                new_fills
                    .try_reserve(2)
                    .map_err(|source| allocation("new styles fills", source))?;
                let none =
                    virtual_fill_record(&document, FillPattern::None, limits.max_staged_bytes)?;
                let gray125 =
                    virtual_fill_record(&document, FillPattern::Gray125, limits.max_staged_bytes)?;
                fill_lookup.insert(none.clone(), 0);
                fill_lookup.insert(gray125.clone(), 1);
                virtual_defaults = Some((none, gray125));
                virtual_defaults_initialized = true;
            }
            if let Some((none, gray125)) = virtual_defaults.take() {
                for record in [none, gray125] {
                    staged_bytes = staged_bytes
                        .checked_add(record.len())
                        .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                    enforce_limit(
                        "styles staged resource bytes",
                        staged_bytes,
                        limits.max_staged_bytes,
                        litchi_core::Resource::Memory,
                    )?;
                    new_fills.push(record);
                }
            }
            let fill = direct_fill(&document, xf)?;
            if !document.virtual_fills {
                enforce_limit(
                    "styles candidate fill bytes",
                    document.elements[fill].end - document.elements[fill].start,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
            }
            let fill_bytes = changed_fill(
                input,
                &document,
                fill,
                &request.fill,
                limits.max_output_bytes,
                limits.max_staged_bytes,
            )?;
            enforce_limit(
                "styles candidate fill bytes",
                fill_bytes.len(),
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            if let Some(index) = fill_lookup.get(&fill_bytes).copied() {
                index
            } else {
                let index = document
                    .fill_items
                    .len()
                    .checked_add(new_fills.len())
                    .ok_or_else(|| invalid("styles fill count overflows"))?;
                enforce_limit(
                    "styles generated fills",
                    index.saturating_add(1),
                    limits.max_fills,
                    litchi_core::Resource::Objects,
                )?;
                let index =
                    u32::try_from(index).map_err(|_| invalid("styles fill index overflows"))?;
                staged_bytes = staged_bytes
                    .checked_add(fill_bytes.len())
                    .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                enforce_limit(
                    "styles staged resource bytes",
                    staged_bytes,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
                fill_lookup.insert(fill_bytes.clone(), index);
                new_fills.push(fill_bytes);
                index
            }
        } else {
            attribute_u32(&document.elements[xf], b"fillId")?.unwrap_or(0)
        };

        let resolved_border = if border_effective {
            if document.virtual_borders && !virtual_border_initialized {
                enforce_limit(
                    "styles default borders",
                    1,
                    limits.max_borders,
                    litchi_core::Resource::Objects,
                )?;
                let default_border = virtual_border_record(&document, limits.max_staged_bytes)?;
                border_lookup.insert(default_border.clone(), 0);
                staged_bytes = staged_bytes
                    .checked_add(default_border.len())
                    .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                enforce_limit(
                    "styles staged resource bytes",
                    staged_bytes,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
                new_borders.push(default_border);
                virtual_border_initialized = true;
            }
            let border = direct_border(&document, xf)?;
            if !document.virtual_borders {
                enforce_limit(
                    "styles candidate border bytes",
                    document.elements[border].end - document.elements[border].start,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
            }
            let border_bytes = changed_border(
                input,
                &document,
                border,
                &request.border,
                limits.max_output_bytes,
            )?;
            enforce_limit(
                "styles candidate border bytes",
                border_bytes.len(),
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            if let Some(index) = border_lookup.get(&border_bytes).copied() {
                index
            } else {
                let index = document
                    .border_items
                    .len()
                    .checked_add(new_borders.len())
                    .ok_or_else(|| invalid("styles border count overflows"))?;
                enforce_limit(
                    "styles generated borders",
                    index.saturating_add(1),
                    limits.max_borders,
                    litchi_core::Resource::Objects,
                )?;
                let index =
                    u32::try_from(index).map_err(|_| invalid("styles border index overflows"))?;
                staged_bytes = staged_bytes
                    .checked_add(border_bytes.len())
                    .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                enforce_limit(
                    "styles staged resource bytes",
                    staged_bytes,
                    limits.max_staged_bytes,
                    litchi_core::Resource::Memory,
                )?;
                border_lookup.insert(border_bytes.clone(), index);
                new_borders.push(border_bytes);
                index
            }
        } else {
            attribute_u32(&document.elements[xf], b"borderId")?.unwrap_or(0)
        };

        let resolved_number_format = if number_format_effective {
            let requested = request
                .number_format
                .as_ref()
                .ok_or_else(|| invalid("number-format request disappeared during planning"))?;
            Some(match requested {
                Some(code) => Some({
                    let (id, record) = resolve_number_format(
                        &document,
                        &mut number_formats,
                        code,
                        limits.max_number_formats,
                        limits.max_staged_bytes,
                    )?;
                    if let Some(record) = record {
                        staged_bytes = staged_bytes
                            .checked_add(record.len())
                            .ok_or_else(|| invalid("styles staged bytes overflow"))?;
                        enforce_limit(
                            "styles staged resource bytes",
                            staged_bytes,
                            limits.max_staged_bytes,
                            litchi_core::Resource::Memory,
                        )?;
                        new_num_fmts.push(record);
                    }
                    id
                }),
                None => None,
            })
        } else {
            None
        };

        enforce_limit(
            "styles candidate cell-XF bytes",
            document.elements[xf].end - document.elements[xf].start,
            limits.max_staged_bytes,
            litchi_core::Resource::Memory,
        )?;
        let xf_bytes = changed_xf(
            input,
            &document,
            xf,
            XfChange {
                font: font_effective.then_some(resolved_font),
                fill: fill_effective.then_some(resolved_fill),
                border: border_effective.then_some(resolved_border),
                alignment: alignment_effective.then_some(&request.alignment),
                protection: protection_effective.then_some(&request.protection),
                quote_prefix: request.quote_prefix,
                number_format: resolved_number_format,
            },
            limits.max_staged_bytes,
        )?;
        enforce_limit(
            "styles candidate cell-XF bytes",
            xf_bytes.len(),
            limits.max_staged_bytes,
            litchi_core::Resource::Memory,
        )?;
        let resolved_xf = if let Some(index) = xf_lookup.get(&xf_bytes).copied() {
            index
        } else {
            let index = xf_items
                .len()
                .checked_add(new_xfs.len())
                .ok_or_else(|| invalid("styles cell-XF count overflows"))?;
            enforce_limit(
                "styles generated cell XFs",
                index.saturating_add(1),
                limits.max_cell_xfs,
                litchi_core::Resource::Objects,
            )?;
            let index =
                u32::try_from(index).map_err(|_| invalid("styles cell-XF index overflows"))?;
            staged_bytes = staged_bytes
                .checked_add(xf_bytes.len())
                .ok_or_else(|| invalid("styles staged bytes overflow"))?;
            enforce_limit(
                "styles staged resource bytes",
                staged_bytes,
                limits.max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            xf_lookup.insert(xf_bytes.clone(), index);
            new_xfs.push(xf_bytes);
            index
        };
        styles.push(resolved_xf);
    }

    if new_fonts.is_empty()
        && new_fills.is_empty()
        && new_borders.is_empty()
        && new_num_fmts.is_empty()
        && new_xfs.is_empty()
    {
        return Ok(StyleRewrite {
            bytes: copy_bytes(input, "styles source")?,
            styles,
        });
    }

    let font_count = font_items
        .len()
        .checked_add(new_fonts.len())
        .ok_or_else(|| invalid("styles font count overflows"))?;
    let fill_count = document
        .fill_items
        .len()
        .checked_add(new_fills.len())
        .ok_or_else(|| invalid("styles fill count overflows"))?;
    let border_count = document
        .border_items
        .len()
        .checked_add(new_borders.len())
        .ok_or_else(|| invalid("styles border count overflows"))?;
    let number_format_count = number_formats
        .items
        .len()
        .checked_add(new_num_fmts.len())
        .ok_or_else(|| invalid("styles number-format count overflows"))?;
    let xf_count = xf_items
        .len()
        .checked_add(new_xfs.len())
        .ok_or_else(|| invalid("styles cell-XF count overflows"))?;
    let mut edits = Vec::new();
    edits
        .try_reserve(11)
        .map_err(|source| allocation("styles rewrite edits", source))?;
    if !new_num_fmts.is_empty() {
        if let Some(num_fmts) = document.num_fmts {
            if !document.elements[num_fmts].empty {
                edits.push(rewrite_count(
                    input,
                    &document.elements[num_fmts],
                    number_format_count,
                    limits.max_output_bytes,
                )?);
            }
            edits.push(append_records_to_collection(
                input,
                &document.elements[num_fmts],
                &new_num_fmts,
                "styles number-format append",
                number_format_count,
                limits.max_output_bytes,
            )?);
        } else {
            let insert_at = document
                .num_fmts_insert_at
                .ok_or_else(|| invalid("styles root has no number-format insertion point"))?;
            edits.push(ByteEdit {
                start: insert_at,
                end: insert_at,
                replacement: new_num_fmts_collection(
                    &document,
                    &new_num_fmts,
                    limits.max_output_bytes,
                )?,
            });
        }
    }
    if !new_fonts.is_empty() {
        edits.push(rewrite_count(
            input,
            &document.elements[fonts],
            font_count,
            limits.max_output_bytes,
        )?);
        edits.push(append_records(
            &document.elements[fonts],
            &new_fonts,
            "styles font append",
        )?);
    }
    if !new_fills.is_empty() {
        if let Some(fills) = fills {
            if !document.elements[fills].empty {
                edits.push(rewrite_count(
                    input,
                    &document.elements[fills],
                    fill_count,
                    limits.max_output_bytes,
                )?);
            }
            edits.push(append_records_to_collection(
                input,
                &document.elements[fills],
                &new_fills,
                "styles fill append",
                fill_count,
                limits.max_output_bytes,
            )?);
        } else {
            let insert_at = document
                .fills_insert_at
                .ok_or_else(|| invalid("styles root has no fill insertion point"))?;
            edits.push(ByteEdit {
                start: insert_at,
                end: insert_at,
                replacement: new_fills_collection(
                    &document,
                    &new_fills,
                    fill_count,
                    limits.max_output_bytes,
                )?,
            });
        }
    }
    if !new_borders.is_empty() {
        if let Some(borders) = borders {
            if !document.elements[borders].empty {
                edits.push(rewrite_count(
                    input,
                    &document.elements[borders],
                    border_count,
                    limits.max_output_bytes,
                )?);
            }
            edits.push(append_border_records_to_collection(
                input,
                &document,
                borders,
                &new_borders,
                border_count,
                limits.max_output_bytes,
            )?);
        } else {
            let insert_at = document
                .borders_insert_at
                .ok_or_else(|| invalid("styles root has no border insertion point"))?;
            edits.push(ByteEdit {
                start: insert_at,
                end: insert_at,
                replacement: new_borders_collection(
                    &document,
                    &new_borders,
                    border_count,
                    limits.max_output_bytes,
                )?,
            });
        }
    }
    if !new_xfs.is_empty() {
        edits.push(rewrite_count(
            input,
            &document.elements[cell_xfs],
            xf_count,
            limits.max_output_bytes,
        )?);
        edits.push(append_records(
            &document.elements[cell_xfs],
            &new_xfs,
            "styles cell-XF append",
        )?);
    }
    let bytes = apply_edits(input, edits, limits.max_output_bytes)?;
    prove_source_projection(
        &bytes,
        StyleEditLimits {
            max_input_bytes: limits.max_output_bytes,
            ..limits
        },
    )?;
    Ok(StyleRewrite { bytes, styles })
}

/// The read model preprocesses MCE. Edit original spans only when the native
/// proof establishes agreement on every understood element and attribute.
/// Ignorable extension markup remains in the original bytes and in clones.
fn prove_source_projection(input: &[u8], limits: StyleEditLimits) -> Result<()> {
    let defaults = mce::Limits::default();
    let processing = mce::Limits {
        max_input_bytes: defaults.max_input_bytes.min(limits.max_input_bytes),
        max_output_bytes: defaults.max_output_bytes.min(limits.max_output_bytes),
        max_depth: defaults.max_depth.min(limits.max_depth),
        ..defaults
    };
    mce::admission::admit(input, &mce::Capabilities::ooxml_baseline(), &processing)
        .map(|_| ())
        .map_err(|_| Error::Unsupported {
            feature: "style edits require equivalent source and processed stylesheet content",
        })
}

#[derive(Debug)]
struct Document {
    elements: Vec<Element>,
    fonts: Option<usize>,
    fills: Option<usize>,
    fills_insert_at: Option<usize>,
    borders: Option<usize>,
    borders_insert_at: Option<usize>,
    num_fmts: Option<usize>,
    num_fmts_insert_at: Option<usize>,
    cell_style_xfs: Option<usize>,
    cell_xfs: Option<usize>,
    fill_items: Vec<usize>,
    virtual_fills: bool,
    border_items: Vec<usize>,
    virtual_borders: bool,
    font_properties: Vec<FontProperties>,
    alignment_properties: Vec<AlignmentProperties>,
    fill_properties: Vec<FillProperties>,
    border_properties: HashMap<usize, BorderProperties>,
}

#[derive(Debug)]
struct NumberFormatIndex {
    items: Vec<usize>,
    by_id: HashMap<u32, usize>,
    by_code: HashMap<Vec<u8>, u32>,
    used_ids: HashSet<u32>,
    next_custom_id: Option<u32>,
    generated_count: usize,
}

impl NumberFormatIndex {
    fn empty() -> Self {
        Self {
            items: Vec::new(),
            by_id: HashMap::new(),
            by_code: HashMap::new(),
            used_ids: HashSet::new(),
            next_custom_id: Some(FIRST_CUSTOM_ID),
            generated_count: 0,
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct AlignmentProperties {
    alignment: Option<usize>,
    /// The insertion point before the first `protection` child.  This keeps
    /// the usual SpreadsheetML order while retaining every existing child.
    alignment_insert_at: Option<usize>,
    protection: Option<usize>,
    /// The insertion point after alignment and before `extLst`.
    protection_insert_at: Option<usize>,
}

#[derive(Debug, Clone, Default)]
struct FillProperties {
    pattern_fill: Option<usize>,
    gradient_fill: Option<usize>,
    foreground: Option<usize>,
    background: Option<usize>,
    pattern_insert_at: Option<usize>,
    foreground_insert_at: Option<usize>,
    background_insert_at: Option<usize>,
    gradient: Option<GradientProjection>,
    pattern_opaque: bool,
}

#[derive(Debug, Clone, Copy, Default)]
struct BorderSideProperties {
    side: Option<usize>,
    color: Option<usize>,
    insert_at: Option<usize>,
}

#[derive(Debug, Clone, Copy, Default)]
struct BorderProperties {
    sides: [BorderSideProperties; 9],
}

#[derive(Debug, Clone, PartialEq)]
struct GradientProjection {
    kind: GradientKind,
    degree: f64,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
    stops: Vec<GradientStopProjection>,
    opaque: bool,
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct GradientStopProjection {
    position: f64,
    color: FillColor,
}

#[derive(Debug, Clone, Copy, Default)]
struct FontProperties {
    bold: Option<usize>,
    italic: Option<usize>,
    strike: Option<usize>,
    condense: Option<usize>,
    extend: Option<usize>,
    outline: Option<usize>,
    shadow: Option<usize>,
    underline: Option<usize>,
    color: Option<usize>,
    name: Option<usize>,
    size: Option<usize>,
    script: Option<usize>,
    family: Option<usize>,
    charset: Option<usize>,
    scheme: Option<usize>,
    /// The byte offset before the first later-schema child, or the closing
    /// tag when no later known child exists.  These are computed while the
    /// direct-child index is built so each request remains O(1).
    bold_insert_at: Option<usize>,
    italic_insert_at: Option<usize>,
    strike_insert_at: Option<usize>,
    condense_insert_at: Option<usize>,
    extend_insert_at: Option<usize>,
    outline_insert_at: Option<usize>,
    shadow_insert_at: Option<usize>,
    underline_insert_at: Option<usize>,
    color_insert_at: Option<usize>,
    name_insert_at: Option<usize>,
    size_insert_at: Option<usize>,
    script_insert_at: Option<usize>,
    family_insert_at: Option<usize>,
    charset_insert_at: Option<usize>,
    scheme_insert_at: Option<usize>,
}

#[derive(Debug)]
struct Element {
    qname: Box<[u8]>,
    local: Box<[u8]>,
    start: usize,
    open_end: usize,
    end: usize,
    close_start: Option<usize>,
    parent: Option<usize>,
    spreadsheetml: bool,
    strict_spreadsheetml: bool,
    empty: bool,
    attributes: Vec<Attribute>,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Box<[u8]>,
    value: Box<[u8]>,
}

#[derive(Debug, Clone)]
struct ByteEdit {
    start: usize,
    end: usize,
    replacement: Vec<u8>,
}

#[derive(Debug, Clone)]
struct RawAttribute {
    name: Vec<u8>,
    start: usize,
    end: usize,
    value_start: usize,
    value_end: usize,
}

fn scan(input: &[u8], max_depth: usize) -> Result<Document> {
    if mce::binds_namespace_by_reference(input) {
        return Err(Error::Unsupported {
            feature: "style edits cannot choose styles markup-compatibility bindings",
        });
    }
    let mut reader = NsReader::from_reader(input);
    let origin = ReaderOrigin::of(input);
    let mut elements = Vec::new();
    let mut stack = Vec::<usize>::new();
    let mut root = None;
    let mut closed_root = false;
    let mut declaration_seen = false;
    let mut event_seen = false;

    loop {
        let start = position(&reader, origin)?;
        let event = reader
            .read_event()
            .map_err(|_| invalid("malformed styles XML"))?
            .into_owned();
        let end = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let preceded_by_event = event_seen;
        let is_eof = matches!(&event, Event::Eof);
        match event {
            Event::Start(element) => {
                if stack.len() >= max_depth {
                    return Err(limit_error(
                        "styles XML depth",
                        stack.len().saturating_add(1),
                        max_depth,
                        litchi_core::Resource::Depth,
                    ));
                }
                let attributes = attributes(&element, reader.decoder())?;
                let qname = copy_bytes(element.name().as_ref(), "styles XML element QName")?;
                let local = copy_bytes(
                    element.name().local_name().as_ref(),
                    "styles XML element local name",
                )?;
                reject_mce(&namespace)?;
                let spreadsheetml = is_spreadsheetml_name(&namespace, element.name(), &local);
                let strict_spreadsheetml = is_strict_spreadsheetml(&namespace);
                if stack.is_empty() {
                    if closed_root || local.as_slice() != b"styleSheet" || !spreadsheetml {
                        return Err(invalid(
                            "styles XML must have one SpreadsheetML styleSheet root",
                        ));
                    }
                }
                elements
                    .try_reserve(1)
                    .map_err(|source| allocation("styles XML elements", source))?;
                let index = elements.len();
                elements.push(Element {
                    qname: qname.into_boxed_slice(),
                    local: local.into_boxed_slice(),
                    start,
                    open_end: end,
                    end,
                    close_start: None,
                    parent: stack.last().copied(),
                    spreadsheetml,
                    strict_spreadsheetml,
                    empty: false,
                    attributes,
                });
                if root.is_none() {
                    root = Some(index);
                }
                stack
                    .try_reserve(1)
                    .map_err(|source| allocation("styles XML element stack", source))?;
                stack.push(index);
            },
            Event::Empty(element) => {
                if stack.len() >= max_depth {
                    return Err(limit_error(
                        "styles XML depth",
                        stack.len().saturating_add(1),
                        max_depth,
                        litchi_core::Resource::Depth,
                    ));
                }
                let attributes = attributes(&element, reader.decoder())?;
                let qname = copy_bytes(element.name().as_ref(), "styles XML element QName")?;
                let local = copy_bytes(
                    element.name().local_name().as_ref(),
                    "styles XML element local name",
                )?;
                reject_mce(&namespace)?;
                let spreadsheetml = is_spreadsheetml_name(&namespace, element.name(), &local);
                let strict_spreadsheetml = is_strict_spreadsheetml(&namespace);
                if stack.is_empty()
                    && (closed_root || local.as_slice() != b"styleSheet" || !spreadsheetml)
                {
                    return Err(invalid(
                        "styles XML must have one SpreadsheetML styleSheet root",
                    ));
                }
                elements
                    .try_reserve(1)
                    .map_err(|source| allocation("styles XML elements", source))?;
                let index = elements.len();
                elements.push(Element {
                    qname: qname.into_boxed_slice(),
                    local: local.into_boxed_slice(),
                    start,
                    open_end: end,
                    end,
                    close_start: None,
                    parent: stack.last().copied(),
                    spreadsheetml,
                    strict_spreadsheetml,
                    empty: true,
                    attributes,
                });
                if root.is_none() {
                    root = Some(index);
                }
                if stack.is_empty() {
                    closed_root = true;
                }
            },
            Event::End(element) => {
                let index = stack
                    .pop()
                    .ok_or_else(|| invalid("styles XML has a closing element outside its root"))?;
                if elements[index].qname.as_ref() != element.name().as_ref() {
                    return Err(invalid("styles XML has mismatched element tags"));
                }
                elements[index].end = end;
                elements[index].close_start = Some(start);
                if stack.is_empty() {
                    closed_root = true;
                }
            },
            Event::Text(text) if stack.is_empty() && !is_xml_whitespace(text.as_ref()) => {
                return Err(invalid("styles XML text is outside root"));
            },
            Event::CData(_) if stack.is_empty() => {
                return Err(invalid("styles XML CDATA is outside root"));
            },
            Event::GeneralRef(_) if stack.is_empty() => {
                return Err(invalid("styles XML entity is outside root"));
            },
            Event::Decl(_) => {
                if root.is_some() || declaration_seen || preceded_by_event {
                    return Err(invalid("invalid styles XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) => {
                return Err(invalid("styles XML DTD is unsupported"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        if !is_eof {
            event_seen = true;
        }
    }

    let root = root.ok_or_else(|| invalid("styles XML has no styleSheet root"))?;
    if !closed_root || !stack.is_empty() || elements[root].local.as_ref() != b"styleSheet" {
        return Err(invalid("styles XML has an incomplete styleSheet root"));
    }
    let fonts = direct_collection(&elements, root, b"fonts")?;
    let fills = direct_collection(&elements, root, b"fills")?;
    let fills_insert_at = fills
        .is_none()
        .then(|| fonts.and_then(|fonts| insert_after_root_child(&elements, root, fonts)))
        .flatten();
    let borders = direct_collection(&elements, root, b"borders")?;
    let borders_insert_at = borders
        .is_none()
        .then(|| {
            fills
                .and_then(|fills| insert_after_root_child(&elements, root, fills))
                .or_else(|| fonts.and_then(|fonts| insert_after_root_child(&elements, root, fonts)))
        })
        .flatten();
    let num_fmts = direct_collection(&elements, root, b"numFmts")?;
    let num_fmts_insert_at = num_fmts
        .is_none()
        .then(|| {
            elements
                .iter()
                .filter(|element| element.parent == Some(root))
                .map(|element| element.start)
                .min()
                .or(elements[root].close_start)
        })
        .flatten();
    let cell_style_xfs = direct_collection(&elements, root, b"cellStyleXfs")?;
    let cell_xfs = direct_collection(&elements, root, b"cellXfs")?;
    Ok(Document {
        elements,
        fonts,
        fills,
        fills_insert_at,
        borders,
        borders_insert_at,
        num_fmts,
        num_fmts_insert_at,
        cell_style_xfs,
        cell_xfs,
        fill_items: Vec::new(),
        virtual_fills: false,
        border_items: Vec::new(),
        virtual_borders: false,
        font_properties: Vec::new(),
        alignment_properties: Vec::new(),
        fill_properties: Vec::new(),
        border_properties: HashMap::new(),
    })
}

fn is_xml_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(*byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn reject_mce(namespace: &ResolveResult<'_>) -> Result<()> {
    if matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE_NAMESPACE) {
        return Err(Error::Unsupported {
            feature: "style edits cannot rewrite styles markup-compatibility content",
        });
    }
    Ok(())
}

fn is_strict_spreadsheetml(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_SPREADSHEETML_NAMESPACE
    )
}

fn direct_collection(elements: &[Element], root: usize, name: &[u8]) -> Result<Option<usize>> {
    let mut found = None;
    for (index, element) in elements.iter().enumerate() {
        if element.parent == Some(root) && element.spreadsheetml && element.local.as_ref() == name {
            if found.replace(index).is_some() {
                return Err(invalid(format!(
                    "styles XML has duplicate {} collections",
                    String::from_utf8_lossy(name)
                )));
            }
        }
    }
    Ok(found)
}

fn insert_after_root_child(elements: &[Element], root: usize, anchor: usize) -> Option<usize> {
    let anchor_end = elements.get(anchor)?.end;
    elements
        .iter()
        .filter(|element| element.parent == Some(root) && element.start >= anchor_end)
        .map(|element| element.start)
        .min()
        .or_else(|| elements.get(root)?.close_start)
}

fn index_number_formats(
    elements: &[Element],
    collection: Option<usize>,
    xfs: &[usize],
    cell_style_xfs: Option<usize>,
    max_number_formats: usize,
    max_staged_bytes: usize,
) -> Result<NumberFormatIndex> {
    let items = collection
        .map(|collection| direct_items(elements, collection, b"numFmt"))
        .transpose()?
        .unwrap_or_default();
    if let Some(collection) = collection {
        validate_count(&[], &elements[collection], items.len(), "number formats")?;
    }

    enforce_limit(
        "styles number formats",
        items.len(),
        max_number_formats,
        litchi_core::Resource::Objects,
    )?;
    let mut source_code_bytes = 0usize;
    for &item in &items {
        let code = attribute_str(&elements[item], b"formatCode", "number-format code")?
            .ok_or_else(|| invalid("styles number-format record is missing formatCode"))?;
        source_code_bytes = source_code_bytes
            .checked_add(code.len())
            .ok_or_else(|| invalid("styles number-format code bytes overflow"))?;
        enforce_limit(
            "styles source number-format code bytes",
            source_code_bytes,
            max_staged_bytes,
            litchi_core::Resource::Memory,
        )?;
    }

    let mut by_id = HashMap::new();
    by_id
        .try_reserve(items.len())
        .map_err(|source| allocation("styles number-format ID index", source))?;
    let mut by_code = HashMap::new();
    by_code
        .try_reserve(items.len())
        .map_err(|source| allocation("styles number-format code index", source))?;
    let mut used_ids = HashSet::new();
    let style_xf_items = cell_style_xfs
        .map(|collection| direct_items(elements, collection, b"xf"))
        .transpose()?
        .unwrap_or_default();
    used_ids
        .try_reserve(
            items
                .len()
                .saturating_add(xfs.len())
                .saturating_add(style_xf_items.len()),
        )
        .map_err(|source| allocation("styles number-format IDs", source))?;
    for &item in &items {
        let id = attribute_u32(&elements[item], b"numFmtId")?
            .ok_or_else(|| invalid("styles number-format record is missing numFmtId"))?;
        let code = attribute_str(&elements[item], b"formatCode", "number-format code")?
            .ok_or_else(|| invalid("styles number-format record is missing formatCode"))?;
        enforce_limit(
            "styles source number-format code",
            code.len(),
            max_staged_bytes,
            litchi_core::Resource::Memory,
        )?;
        if by_id.insert(id, item).is_some() {
            return Err(invalid("styles number-format IDs are duplicated"));
        }
        let code = copy_bytes(code.as_bytes(), "styles number-format code index")?;
        by_code.entry(code).or_insert(id);
        used_ids.insert(id);
    }
    for &xf in xfs {
        if let Some(id) = attribute_u32(&elements[xf], b"numFmtId")? {
            used_ids.insert(id);
        }
    }
    for &xf in &style_xf_items {
        if let Some(id) = attribute_u32(&elements[xf], b"numFmtId")? {
            used_ids.insert(id);
        }
    }
    Ok(NumberFormatIndex {
        items,
        by_id,
        by_code,
        used_ids,
        next_custom_id: Some(FIRST_CUSTOM_ID),
        generated_count: 0,
    })
}

fn index_font_properties(elements: &[Element], fonts: &[usize]) -> Result<Vec<FontProperties>> {
    let mut is_font = Vec::new();
    is_font
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles font membership", source))?;
    is_font.resize(elements.len(), false);
    for &font in fonts {
        if let Some(slot) = is_font.get_mut(font) {
            *slot = true;
        }
    }

    let mut properties = Vec::new();
    properties
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles font property index", source))?;
    properties.resize(elements.len(), FontProperties::default());
    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        if !is_font.get(parent).copied().unwrap_or(false) || !element.spreadsheetml {
            continue;
        }
        let Some(rank) = font_child_rank(element.local.as_ref()) else {
            continue;
        };
        let properties = &mut properties[parent];
        if rank > 0 && properties.bold_insert_at.is_none() {
            properties.bold_insert_at = Some(element.start);
        }
        if rank > 1 && properties.italic_insert_at.is_none() {
            properties.italic_insert_at = Some(element.start);
        }
        if rank > 2 && properties.strike_insert_at.is_none() {
            properties.strike_insert_at = Some(element.start);
        }
        if rank > 3 && properties.condense_insert_at.is_none() {
            properties.condense_insert_at = Some(element.start);
        }
        if rank > 4 && properties.extend_insert_at.is_none() {
            properties.extend_insert_at = Some(element.start);
        }
        if rank > 5 && properties.outline_insert_at.is_none() {
            properties.outline_insert_at = Some(element.start);
        }
        if rank > 6 && properties.shadow_insert_at.is_none() {
            properties.shadow_insert_at = Some(element.start);
        }
        if rank > 7 && properties.underline_insert_at.is_none() {
            properties.underline_insert_at = Some(element.start);
        }
        if rank > 8 && properties.script_insert_at.is_none() {
            properties.script_insert_at = Some(element.start);
        }
        if rank > 9 && properties.size_insert_at.is_none() {
            properties.size_insert_at = Some(element.start);
        }
        if rank > 10 && properties.color_insert_at.is_none() {
            properties.color_insert_at = Some(element.start);
        }
        if rank > 11 && properties.name_insert_at.is_none() {
            properties.name_insert_at = Some(element.start);
        }
        if rank > 12 && properties.family_insert_at.is_none() {
            properties.family_insert_at = Some(element.start);
        }
        if rank > 13 && properties.charset_insert_at.is_none() {
            properties.charset_insert_at = Some(element.start);
        }
        if rank > 14 && properties.scheme_insert_at.is_none() {
            properties.scheme_insert_at = Some(element.start);
        }
        let slot = match element.local.as_ref() {
            b"b" => &mut properties.bold,
            b"i" => &mut properties.italic,
            b"strike" => &mut properties.strike,
            b"condense" => &mut properties.condense,
            b"extend" => &mut properties.extend,
            b"outline" => &mut properties.outline,
            b"shadow" => &mut properties.shadow,
            b"u" => &mut properties.underline,
            b"color" => &mut properties.color,
            b"vertAlign" => &mut properties.script,
            b"sz" => &mut properties.size,
            b"name" => &mut properties.name,
            b"family" => &mut properties.family,
            b"charset" => &mut properties.charset,
            b"scheme" => &mut properties.scheme,
            _ => continue,
        };
        if slot.replace(index).is_some() {
            return Err(invalid("styles font has duplicate properties"));
        }
    }
    for &font in fonts {
        let properties = properties
            .get_mut(font)
            .ok_or_else(|| invalid("styles font index is out of range"))?;
        let close_start = elements.get(font).and_then(|element| element.close_start);
        if properties.bold.is_none() && properties.bold_insert_at.is_none() {
            properties.bold_insert_at = close_start;
        }
        if properties.italic.is_none() && properties.italic_insert_at.is_none() {
            properties.italic_insert_at = close_start;
        }
        if properties.strike.is_none() && properties.strike_insert_at.is_none() {
            properties.strike_insert_at = close_start;
        }
        if properties.condense.is_none() && properties.condense_insert_at.is_none() {
            properties.condense_insert_at = close_start;
        }
        if properties.extend.is_none() && properties.extend_insert_at.is_none() {
            properties.extend_insert_at = close_start;
        }
        if properties.outline.is_none() && properties.outline_insert_at.is_none() {
            properties.outline_insert_at = close_start;
        }
        if properties.shadow.is_none() && properties.shadow_insert_at.is_none() {
            properties.shadow_insert_at = close_start;
        }
        if properties.underline.is_none() && properties.underline_insert_at.is_none() {
            properties.underline_insert_at = close_start;
        }
        if properties.color.is_none() && properties.color_insert_at.is_none() {
            properties.color_insert_at = close_start;
        }
        if properties.script.is_none() && properties.script_insert_at.is_none() {
            properties.script_insert_at = close_start;
        }
        if properties.size.is_none() && properties.size_insert_at.is_none() {
            properties.size_insert_at = close_start;
        }
        if properties.family.is_none() && properties.family_insert_at.is_none() {
            properties.family_insert_at = close_start;
        }
        if properties.charset.is_none() && properties.charset_insert_at.is_none() {
            properties.charset_insert_at = close_start;
        }
        if properties.scheme.is_none() && properties.scheme_insert_at.is_none() {
            properties.scheme_insert_at = close_start;
        }
        if properties.name.is_none() && properties.name_insert_at.is_none() {
            properties.name_insert_at = close_start;
        }
    }
    Ok(properties)
}

/// The direct-child order used by the Microsoft Open XML SDK for
/// SpreadsheetML `CT_Font`.
///
/// Unknown children are deliberately absent.  An insertion anchor is placed
/// before the first known child with a greater rank, leaving every existing
/// child, including opaque producer markup, in its original order.
fn font_child_rank(local: &[u8]) -> Option<usize> {
    match local {
        b"b" => Some(0),
        b"i" => Some(1),
        b"strike" => Some(2),
        b"condense" => Some(3),
        b"extend" => Some(4),
        b"outline" => Some(5),
        b"shadow" => Some(6),
        b"u" => Some(7),
        b"vertAlign" => Some(8),
        b"sz" => Some(9),
        b"color" => Some(10),
        b"name" => Some(11),
        b"family" => Some(12),
        b"charset" => Some(13),
        b"scheme" => Some(14),
        _ => None,
    }
}

fn direct_items(elements: &[Element], parent: usize, name: &[u8]) -> Result<Vec<usize>> {
    let mut result = Vec::new();
    result
        .try_reserve(elements.len())
        .map_err(|source| allocation("styles direct records", source))?;
    result.extend(elements.iter().enumerate().filter_map(|(index, element)| {
        (element.parent == Some(parent) && element.spreadsheetml && element.local.as_ref() == name)
            .then_some(index)
    }));
    Ok(result)
}

fn index_alignment_properties(
    elements: &[Element],
    xfs: &[usize],
) -> Result<Vec<AlignmentProperties>> {
    let mut is_xf = Vec::new();
    is_xf
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles cell-XF membership", source))?;
    is_xf.resize(elements.len(), false);
    for &xf in xfs {
        if let Some(slot) = is_xf.get_mut(xf) {
            *slot = true;
        }
    }

    let mut properties = Vec::new();
    properties
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles cell-XF alignment index", source))?;
    properties.resize(elements.len(), AlignmentProperties::default());
    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        if !is_xf.get(parent).copied().unwrap_or(false) || !element.spreadsheetml {
            continue;
        }
        let slot = &mut properties[parent];
        match element.local.as_ref() {
            b"alignment" => {
                if slot.alignment.replace(index).is_some() {
                    return Err(invalid("styles cell-XF has duplicate alignment properties"));
                }
            },
            b"protection" => {
                if slot.protection.replace(index).is_some() {
                    return Err(invalid(
                        "styles cell-XF has duplicate protection properties",
                    ));
                }
                if slot.alignment_insert_at.is_none() {
                    slot.alignment_insert_at = Some(element.start);
                }
            },
            b"extLst" => {
                if slot.alignment_insert_at.is_none() {
                    slot.alignment_insert_at = Some(element.start);
                }
                if slot.protection_insert_at.is_none() {
                    slot.protection_insert_at = Some(element.start);
                }
            },
            _ => {},
        }
    }
    for &xf in xfs {
        let slot = properties
            .get_mut(xf)
            .ok_or_else(|| invalid("styles cell-XF index is out of range"))?;
        if slot.alignment.is_none() && slot.alignment_insert_at.is_none() {
            slot.alignment_insert_at = elements.get(xf).and_then(|element| element.close_start);
        }
        if slot.protection.is_none() && slot.protection_insert_at.is_none() {
            slot.protection_insert_at = elements.get(xf).and_then(|element| element.close_start);
        }
    }
    Ok(properties)
}

fn index_fill_properties(elements: &[Element], fills: &[usize]) -> Result<Vec<FillProperties>> {
    let mut is_fill = Vec::new();
    is_fill
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles fill membership", source))?;
    is_fill.resize(elements.len(), false);
    for &fill in fills {
        if let Some(slot) = is_fill.get_mut(fill) {
            *slot = true;
        }
    }

    let mut properties = Vec::new();
    properties
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles fill property index", source))?;
    properties.resize(elements.len(), FillProperties::default());

    let mut first_child = Vec::new();
    first_child
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles fill child positions", source))?;
    first_child.resize(elements.len(), None);

    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        if is_fill.get(parent).copied().unwrap_or(false) {
            if first_child[parent].is_none() {
                first_child[parent] = Some(index);
            }
            if !element.spreadsheetml {
                continue;
            }
            let slot = &mut properties[parent];
            match element.local.as_ref() {
                b"patternFill" => {
                    if slot.pattern_fill.replace(index).is_some() {
                        return Err(invalid("styles fill has duplicate patternFill children"));
                    }
                },
                b"gradientFill" => {
                    if slot.gradient_fill.replace(index).is_some() {
                        return Err(invalid("styles fill has duplicate gradientFill children"));
                    }
                },
                _ => {},
            }
        }
    }

    let mut pattern_fill = Vec::new();
    pattern_fill
        .try_reserve_exact(elements.len())
        .map_err(|source| allocation("styles pattern-fill ownership", source))?;
    pattern_fill.resize(elements.len(), None);
    for &fill in fills {
        if let Some(pattern) = properties
            .get(fill)
            .and_then(|properties| properties.pattern_fill)
        {
            let owner = pattern_fill
                .get_mut(pattern)
                .ok_or_else(|| invalid("styles patternFill index is out of range"))?;
            if owner.replace(fill).is_some() {
                return Err(invalid("styles patternFill has multiple owning fills"));
            }
        }
    }

    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        let Some(fill) = pattern_fill.get(parent).copied().flatten() else {
            continue;
        };
        if !element.spreadsheetml {
            continue;
        }
        let slot = &mut properties[fill];
        match element.local.as_ref() {
            b"fgColor" => {
                if slot.foreground.replace(index).is_some() {
                    return Err(invalid("styles patternFill has duplicate fgColor children"));
                }
            },
            b"bgColor" => {
                if slot.background.replace(index).is_some() {
                    return Err(invalid("styles patternFill has duplicate bgColor children"));
                }
            },
            _ => {},
        }
    }

    for &fill in fills {
        let slot = properties
            .get_mut(fill)
            .ok_or_else(|| invalid("styles fill index is out of range"))?;
        let pattern = slot.pattern_fill;
        if slot.pattern_fill.is_none() && slot.gradient_fill.is_none() {
            slot.pattern_insert_at = first_child[fill]
                .map(|child| elements[child].start)
                .or(elements[fill].close_start);
        }
        let Some(pattern) = pattern else {
            continue;
        };
        if slot.foreground.is_none() {
            slot.foreground_insert_at = slot
                .background
                .map(|background| elements[background].start)
                .or(elements[pattern].close_start);
        }
        if slot.background.is_none() {
            slot.background_insert_at = elements[pattern].close_start;
        }
    }
    Ok(properties)
}

fn border_edge_index(local: &[u8]) -> Option<usize> {
    match local {
        b"start" => Some(0),
        b"end" => Some(1),
        b"left" => Some(2),
        b"right" => Some(3),
        b"top" => Some(4),
        b"bottom" => Some(5),
        b"diagonal" => Some(6),
        b"vertical" => Some(7),
        b"horizontal" => Some(8),
        _ => None,
    }
}

fn index_border_properties(
    elements: &[Element],
    borders: &[usize],
    max_staged_bytes: usize,
) -> Result<HashMap<usize, BorderProperties>> {
    if borders.is_empty() {
        return Ok(HashMap::new());
    }
    let per_border = size_of::<usize>()
        .checked_add(size_of::<BorderProperties>())
        .and_then(|size| size.checked_mul(4))
        .ok_or_else(|| invalid("styles border property index size overflows"))?;
    let index_bytes = borders
        .len()
        .checked_mul(per_border)
        .ok_or_else(|| invalid("styles border property index size overflows"))?;
    enforce_limit(
        "styles border property index",
        index_bytes,
        max_staged_bytes,
        litchi_core::Resource::Memory,
    )?;
    let mut properties = HashMap::new();
    properties
        .try_reserve(borders.len())
        .map_err(|source| allocation("styles border property index", source))?;
    for &border in borders {
        if properties
            .insert(border, BorderProperties::default())
            .is_some()
        {
            return Err(invalid("styles border collection repeats a border record"));
        }
    }

    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        if !properties.contains_key(&parent) || !element.spreadsheetml {
            continue;
        }
        let Some(edge) = border_edge_index(element.local.as_ref()) else {
            continue;
        };
        let slot = properties
            .get_mut(&parent)
            .ok_or_else(|| invalid("styles border index is out of range"))?
            .sides
            .get_mut(edge)
            .ok_or_else(|| invalid("styles border edge index is out of range"))?;
        if slot.side.replace(index).is_some() {
            return Err(invalid("styles border has duplicate side children"));
        }
    }
    for (index, element) in elements.iter().enumerate() {
        let Some(side) = element.parent else {
            continue;
        };
        if !element.spreadsheetml || element.local.as_ref() != b"color" {
            continue;
        }
        let Some(side_element) = elements.get(side) else {
            continue;
        };
        let Some(border) = side_element.parent else {
            continue;
        };
        if !properties.contains_key(&border) || !side_element.spreadsheetml {
            continue;
        }
        let Some(edge) = border_edge_index(side_element.local.as_ref()) else {
            continue;
        };
        let slot = properties
            .get_mut(&border)
            .ok_or_else(|| invalid("styles border index is out of range"))?
            .sides
            .get_mut(edge)
            .ok_or_else(|| invalid("styles border edge index is out of range"))?;
        if slot.color.replace(index).is_some() {
            return Err(invalid("styles border side has duplicate color children"));
        }
    }

    for element in elements.iter() {
        let Some(border) = element.parent else {
            continue;
        };
        if !properties.contains_key(&border) {
            continue;
        }
        let rank = element
            .spreadsheetml
            .then(|| border_edge_index(element.local.as_ref()))
            .flatten();
        let anchor_rank = rank.unwrap_or(9);
        for edge in 0..anchor_rank {
            let slot = properties
                .get_mut(&border)
                .ok_or_else(|| invalid("styles border index is out of range"))?
                .sides
                .get_mut(edge)
                .ok_or_else(|| invalid("styles border edge index is out of range"))?;
            if slot.side.is_none() && slot.insert_at.is_none() {
                slot.insert_at = Some(element.start);
            }
        }
    }
    for &border in borders {
        for edge in 0..9 {
            let slot = properties
                .get_mut(&border)
                .ok_or_else(|| invalid("styles border index is out of range"))?
                .sides
                .get_mut(edge)
                .ok_or_else(|| invalid("styles border edge index is out of range"))?;
            if slot.insert_at.is_none() {
                slot.insert_at = elements.get(border).and_then(|element| element.close_start);
            }
        }
    }
    Ok(properties)
}

fn validate_border_references(document: &Document, xfs: &[usize]) -> Result<()> {
    let border_count = if document.virtual_borders {
        1
    } else {
        document.border_items.len()
    };
    let style_xfs = document
        .cell_style_xfs
        .map(|collection| direct_items(&document.elements, collection, b"xf"))
        .transpose()?
        .unwrap_or_default();
    for &xf in xfs.iter().chain(style_xfs.iter()) {
        let Some(id) = attribute_u32(&document.elements[xf], b"borderId")? else {
            continue;
        };
        let id =
            usize::try_from(id).map_err(|_| invalid("styles border index does not fit usize"))?;
        if document.virtual_borders && id != 0 {
            return Err(invalid("styles border reference is out of range"));
        }
        if id >= border_count {
            return Err(invalid("styles border reference is out of range"));
        }
    }
    Ok(())
}

fn index_fill_replacement_metadata(
    input: &[u8],
    elements: &[Element],
    properties: &mut [FillProperties],
    fills: &[usize],
    max_gradient_stops: usize,
    max_staged_bytes: usize,
) -> Result<()> {
    let mut parents = Vec::new();
    parents
        .try_reserve(fills.len().saturating_mul(3))
        .map_err(|source| allocation("styles fill child index", source))?;
    for &fill in fills {
        parents.push(fill);
        if let Some(properties) = properties.get(fill) {
            if let Some(pattern) = properties.pattern_fill {
                parents.push(pattern);
            }
            if let Some(gradient) = properties.gradient_fill {
                parents.push(gradient);
            }
        }
    }
    let mut relevant = HashSet::new();
    relevant
        .try_reserve(parents.len())
        .map_err(|source| allocation("styles fill child index", source))?;
    for &parent in &parents {
        relevant.insert(parent);
    }
    for (index, element) in elements.iter().enumerate() {
        if !element.spreadsheetml || element.local.as_ref() != b"stop" {
            continue;
        }
        if element
            .parent
            .is_some_and(|parent| relevant.contains(&parent))
        {
            relevant
                .try_reserve(1)
                .map_err(|source| allocation("styles fill child index", source))?;
            parents
                .try_reserve(1)
                .map_err(|source| allocation("styles fill child index", source))?;
            relevant.insert(index);
            parents.push(index);
        }
    }
    for (index, element) in elements.iter().enumerate() {
        if !element.spreadsheetml || element.local.as_ref() != b"color" {
            continue;
        }
        if element
            .parent
            .is_some_and(|parent| relevant.contains(&parent))
        {
            relevant
                .try_reserve(1)
                .map_err(|source| allocation("styles fill child index", source))?;
            parents
                .try_reserve(1)
                .map_err(|source| allocation("styles fill child index", source))?;
            relevant.insert(index);
            parents.push(index);
        }
    }
    parents.sort_unstable();
    parents.dedup();
    let children = direct_children_index(elements, &parents)?;
    let mut source_gradient_bytes = 0usize;
    for &fill in fills {
        let slot = properties
            .get_mut(fill)
            .ok_or_else(|| invalid("styles fill index is out of range"))?;
        if let Some(pattern) = slot.pattern_fill {
            slot.pattern_opaque = pattern_has_opaque_markup(input, elements, pattern, &children);
        }
        if let Some(gradient) = slot.gradient_fill {
            let (gradient, gradient_bytes) = parse_gradient_projection(
                input,
                elements,
                gradient,
                &children,
                max_gradient_stops,
                source_gradient_bytes,
                max_staged_bytes,
            )?;
            source_gradient_bytes = source_gradient_bytes
                .checked_add(gradient_bytes)
                .ok_or_else(|| invalid("styles gradient bytes overflow"))?;
            enforce_limit(
                "styles source gradient bytes",
                source_gradient_bytes,
                max_staged_bytes,
                litchi_core::Resource::Memory,
            )?;
            slot.gradient = Some(gradient);
        }
    }
    Ok(())
}

fn direct_children_index(
    elements: &[Element],
    parents: &[usize],
) -> Result<HashMap<usize, Vec<usize>>> {
    let mut result = HashMap::new();
    result
        .try_reserve(parents.len())
        .map_err(|source| allocation("styles fill child index", source))?;
    for &parent in parents {
        result.insert(parent, Vec::new());
    }
    for (index, element) in elements.iter().enumerate() {
        let Some(parent) = element.parent else {
            continue;
        };
        let Some(children) = result.get_mut(&parent) else {
            continue;
        };
        children
            .try_reserve(1)
            .map_err(|source| allocation("styles fill child index", source))?;
        children.push(index);
    }
    Ok(result)
}

fn pattern_has_opaque_markup(
    input: &[u8],
    elements: &[Element],
    pattern: usize,
    children: &HashMap<usize, Vec<usize>>,
) -> bool {
    if element_has_opaque_markup(
        input,
        elements,
        pattern,
        children,
        &[b"patternType"],
        &[b"fgColor", b"bgColor"],
    ) {
        return true;
    }
    children.get(&pattern).into_iter().flatten().any(|&index| {
        matches!(elements[index].local.as_ref(), b"fgColor" | b"bgColor")
            && element_has_opaque_markup(
                input,
                elements,
                index,
                children,
                &[b"rgb", b"theme", b"indexed", b"auto", b"tint"],
                &[],
            )
    })
}

fn element_has_opaque_markup(
    input: &[u8],
    elements: &[Element],
    element: usize,
    children: &HashMap<usize, Vec<usize>>,
    known_attributes: &[&[u8]],
    known_children: &[&[u8]],
) -> bool {
    let Some(parent) = elements.get(element) else {
        return true;
    };
    if parent.attributes.iter().any(|attribute| {
        !known_attributes
            .iter()
            .any(|name| *name == attribute.name.as_ref())
    }) {
        return true;
    }
    let mut cursor = parent.open_end;
    for &index in children.get(&element).into_iter().flatten() {
        let child = &elements[index];
        if !is_xml_whitespace(&input[cursor..child.start])
            || !child.spreadsheetml
            || !known_children
                .iter()
                .any(|name| *name == child.local.as_ref())
        {
            return true;
        }
        cursor = child.end;
    }
    let end = parent.close_start.unwrap_or(parent.end);
    !is_xml_whitespace(&input[cursor..end])
}

fn parse_gradient_projection(
    input: &[u8],
    elements: &[Element],
    gradient: usize,
    children: &HashMap<usize, Vec<usize>>,
    max_gradient_stops: usize,
    staged_bytes: usize,
    max_staged_bytes: usize,
) -> Result<(GradientProjection, usize)> {
    let element = elements
        .get(gradient)
        .ok_or_else(|| invalid("styles gradientFill index is out of range"))?;
    let mut kind = GradientKind::Linear;
    let mut degree = 0.0;
    let mut left = 0.0;
    let mut right = 0.0;
    let mut top = 0.0;
    let mut bottom = 0.0;
    let mut opaque = false;
    for attribute in &element.attributes {
        let target = match attribute.name.as_ref() {
            b"type" => {
                let value = std::str::from_utf8(attribute.value.as_ref())
                    .map_err(|_| invalid("styles gradient type is not UTF-8"))?;
                kind = GradientKind::from_token(value.trim())
                    .ok_or_else(|| invalid("styles gradient type is invalid"))?;
                continue;
            },
            b"degree" => &mut degree,
            b"left" => &mut left,
            b"right" => &mut right,
            b"top" => &mut top,
            b"bottom" => &mut bottom,
            _ => {
                opaque = true;
                continue;
            },
        };
        let value = std::str::from_utf8(attribute.value.as_ref())
            .map_err(|_| invalid("styles gradient geometry is not UTF-8"))?
            .trim()
            .parse::<f64>()
            .ok()
            .filter(|value| value.is_finite())
            .ok_or_else(|| invalid("styles gradient geometry is invalid"))?;
        *target = value;
    }

    let stop_count = children
        .get(&gradient)
        .into_iter()
        .flatten()
        .filter(|&&child| {
            elements[child].spreadsheetml && elements[child].local.as_ref() == b"stop"
        })
        .count();
    enforce_limit(
        "styles gradient stops",
        stop_count,
        max_gradient_stops,
        litchi_core::Resource::Objects,
    )?;
    let mut stop_bytes = 0usize;
    for &index in children.get(&gradient).into_iter().flatten() {
        let child = &elements[index];
        if child.spreadsheetml && child.local.as_ref() == b"stop" {
            stop_bytes = stop_bytes
                .checked_add(child.end - child.start)
                .ok_or_else(|| invalid("styles gradient stop bytes overflow"))?;
        }
    }
    let projected_bytes = size_of::<GradientProjection>()
        .checked_add(
            stop_count
                .checked_mul(size_of::<GradientStopProjection>())
                .ok_or_else(|| invalid("styles gradient projection size overflows"))?,
        )
        .and_then(|size| size.checked_add(stop_bytes))
        .ok_or_else(|| invalid("styles gradient projection size overflows"))?;
    let aggregate = staged_bytes
        .checked_add(projected_bytes)
        .ok_or_else(|| invalid("styles gradient bytes overflow"))?;
    enforce_limit(
        "styles source gradient bytes",
        aggregate,
        max_staged_bytes,
        litchi_core::Resource::Memory,
    )?;
    let mut stops = Vec::new();
    stops
        .try_reserve_exact(stop_count)
        .map_err(|source| allocation("styles gradient stops", source))?;
    let mut cursor = element.open_end;
    for &index in children.get(&gradient).into_iter().flatten() {
        let child = &elements[index];
        if !is_xml_whitespace(&input[cursor..child.start]) {
            opaque = true;
        }
        if !child.spreadsheetml || child.local.as_ref() != b"stop" {
            opaque = true;
            cursor = child.end;
            continue;
        }
        let (stop, stop_opaque) = parse_gradient_stop_projection(input, elements, index, children)?;
        opaque |= stop_opaque;
        stops.push(stop);
        cursor = child.end;
    }
    let end = element.close_start.unwrap_or(element.end);
    if !is_xml_whitespace(&input[cursor..end]) {
        opaque = true;
    }
    Ok((
        GradientProjection {
            kind,
            degree,
            left,
            right,
            top,
            bottom,
            stops,
            opaque,
        },
        projected_bytes,
    ))
}

fn parse_gradient_stop_projection(
    input: &[u8],
    elements: &[Element],
    stop: usize,
    children: &HashMap<usize, Vec<usize>>,
) -> Result<(GradientStopProjection, bool)> {
    let element = elements
        .get(stop)
        .ok_or_else(|| invalid("styles gradient stop index is out of range"))?;
    let position = attribute_str(element, b"position", "gradient stop position")?
        .ok_or_else(|| invalid("styles gradient stop is missing position"))?
        .trim()
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .ok_or_else(|| invalid("styles gradient stop position is invalid"))?;
    let mut opaque = element
        .attributes
        .iter()
        .any(|attribute| attribute.name.as_ref() != b"position");
    let mut color = None;
    let mut cursor = element.open_end;
    for &index in children.get(&stop).into_iter().flatten() {
        let child = &elements[index];
        if !is_xml_whitespace(&input[cursor..child.start]) {
            opaque = true;
        }
        if !child.spreadsheetml || child.local.as_ref() != b"color" || color.is_some() {
            opaque = true;
        } else {
            color = Some(fill_color_value_in_elements(elements, index)?);
            opaque |= element_has_opaque_markup(
                input,
                elements,
                index,
                children,
                &[b"rgb", b"theme", b"indexed", b"auto", b"tint"],
                &[],
            );
        }
        cursor = child.end;
    }
    let end = element.close_start.unwrap_or(element.end);
    if !is_xml_whitespace(&input[cursor..end]) {
        opaque = true;
    }
    let color = color.ok_or_else(|| invalid("styles gradient stop is missing color"))?;
    Ok((GradientStopProjection { position, color }, opaque))
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<Attribute>> {
    let mut result = Vec::new();
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|_| invalid("malformed styles XML attribute"))?;
        // NsReader resolves literal namespace values, while MCE decodes
        // references. Do not let the span scanner assign different indexes.
        if (attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:"))
            && attribute.value.contains(&b'&')
        {
            return Err(Error::Unsupported {
                feature: "style edits require literal namespace bindings",
            });
        }
        let name = copy_bytes(attribute.key.as_ref(), "styles XML attribute name")?;
        if result
            .iter()
            .any(|item: &Attribute| item.name.as_ref() == name.as_slice())
        {
            return Err(invalid("styles XML has a duplicate attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|_| invalid("malformed styles XML attribute value"))?
            .into_owned();
        let value = copy_bytes(value.as_bytes(), "styles XML attribute value")?;
        result
            .try_reserve(1)
            .map_err(|source| allocation("styles XML attributes", source))?;
        result.push(Attribute {
            name: name.into_boxed_slice(),
            value: value.into_boxed_slice(),
        });
    }
    Ok(result)
}

fn attribute_u32(element: &Element, name: &[u8]) -> Result<Option<u32>> {
    let Some(value) = element
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| attribute.value.as_ref())
    else {
        return Ok(None);
    };
    let text = std::str::from_utf8(value)
        .map_err(|_| invalid("styles numeric attribute is not UTF-8"))?
        .trim();
    text.parse::<u32>()
        .map(Some)
        .map_err(|_| invalid("styles numeric attribute is invalid"))
}

fn attribute_bool(element: &Element, name: &[u8]) -> Result<Option<bool>> {
    let Some(value) = element
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| attribute.value.as_ref())
    else {
        return Ok(None);
    };
    let value = std::str::from_utf8(value)
        .map_err(|_| invalid("styles boolean attribute is not UTF-8"))?
        .trim();
    match value {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(invalid("styles boolean attribute is invalid")),
    }
}

fn validate_count(input: &[u8], element: &Element, actual: usize, label: &str) -> Result<()> {
    if let Some(declared) = attribute_u32(element, b"count")? {
        let declared = usize::try_from(declared)
            .map_err(|_| invalid("styles collection count does not fit usize"))?;
        if declared != actual {
            return Err(invalid(format!(
                "styles {label} count {declared} disagrees with {actual} records"
            )));
        }
    }
    let _ = input;
    Ok(())
}

fn property_value(document: &Document, font: usize, property: &[u8]) -> Result<Option<bool>> {
    let Some(child) = direct_property(document, font, property)? else {
        return Ok(None);
    };
    let Some(value) = document.elements[child]
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == b"val")
        .map(|attribute| attribute.value.as_ref())
    else {
        return Ok(Some(true));
    };
    let value = std::str::from_utf8(value)
        .map_err(|_| invalid("styles font boolean value is not UTF-8"))?
        .trim();
    match value {
        "1" | "true" => Ok(Some(true)),
        "0" | "false" => Ok(Some(false)),
        _ => Err(invalid("styles font boolean property has an invalid value")),
    }
}

fn string_value<'a>(
    document: &'a Document,
    font: usize,
    property: &[u8],
) -> Result<Option<&'a str>> {
    let Some(child) = direct_property(document, font, property)? else {
        return Ok(None);
    };
    let value = document.elements[child]
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == b"val")
        .map(|attribute| attribute.value.as_ref())
        .ok_or_else(|| invalid("styles font string property is missing val"))?;
    std::str::from_utf8(value)
        .map(Some)
        .map_err(|_| invalid("styles font string property is not UTF-8"))
}

fn u32_property_value(document: &Document, font: usize, property: &[u8]) -> Result<Option<u32>> {
    let Some(value) = string_value(document, font, property)? else {
        return Ok(None);
    };
    value
        .trim()
        .parse::<u32>()
        .map(Some)
        .map_err(|_| invalid("styles font integer property has an invalid value"))
}

fn scheme_value(document: &Document, font: usize) -> Result<Option<Scheme>> {
    let Some(value) = string_value(document, font, b"scheme")? else {
        return Ok(None);
    };
    value
        .trim()
        .parse::<Scheme>()
        .map(Some)
        .map_err(|_| invalid("styles font scheme property has an invalid value"))
}

fn size_value(document: &Document, font: usize) -> Result<Option<f64>> {
    let Some(value) = string_value(document, font, b"sz")? else {
        return Ok(None);
    };
    let value = value
        .trim()
        .parse::<f64>()
        .map_err(|_| invalid("styles font size property has an invalid value"))?;
    if !(value > 0.0 && value.is_finite()) {
        return Err(invalid(
            "styles font size property must be finite and greater than 0",
        ));
    }
    Ok(Some(value))
}

fn script_value(document: &Document, font: usize) -> Result<Option<Script>> {
    let Some(value) = string_value(document, font, b"vertAlign")? else {
        return Ok(None);
    };
    value
        .trim()
        .parse::<Script>()
        .map(Some)
        .map_err(|_| invalid("styles font script property has an invalid value"))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ColorState {
    kind: FontColorKind,
    tint: Option<Tint>,
}

fn color_state(color: &FontColor) -> ColorState {
    ColorState {
        kind: color.kind(),
        tint: color.tint(),
    }
}

fn color_value(document: &Document, font: usize) -> Result<Option<ColorState>> {
    let Some(child) = direct_property(document, font, b"color")? else {
        return Ok(None);
    };
    let element = &document.elements[child];
    let rgb = attribute_str(element, b"rgb", "font color RGB")?;
    let theme = attribute_u32(element, b"theme")?;
    let indexed = attribute_u32(element, b"indexed")?;
    let automatic = attribute_bool(element, b"auto")?;
    let base_count = usize::from(rgb.is_some())
        + usize::from(theme.is_some())
        + usize::from(indexed.is_some())
        + usize::from(automatic.is_some());
    if base_count > 1 {
        return Err(invalid(
            "styles font color has multiple mutually exclusive base values",
        ));
    }
    let kind = if let Some(rgb) = rgb {
        FontColorKind::Rgb(
            rgb.parse()
                .map_err(|_| invalid("styles font color RGB value is invalid"))?,
        )
    } else if let Some(theme) = theme {
        FontColorKind::Theme(theme)
    } else if let Some(indexed) = indexed {
        FontColorKind::Indexed(indexed)
    } else if let Some(automatic) = automatic {
        FontColorKind::Auto(automatic)
    } else {
        FontColorKind::Default
    };
    let tint = attribute_str(element, b"tint", "font color tint")?
        .map(|value| {
            value
                .trim()
                .parse::<f64>()
                .map_err(|_| invalid("styles font color tint is invalid"))
                .and_then(|value| {
                    Tint::new(value).map_err(|_| invalid("styles font color tint is invalid"))
                })
        })
        .transpose()?;
    Ok(Some(ColorState { kind, tint }))
}

fn attribute_str<'a>(
    element: &'a Element,
    name: &[u8],
    label: &'static str,
) -> Result<Option<&'a str>> {
    let Some(value) = element
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == name)
        .map(|attribute| attribute.value.as_ref())
    else {
        return Ok(None);
    };
    std::str::from_utf8(value)
        .map(Some)
        .map_err(|_| invalid(format!("styles {label} value is not UTF-8")))
}

fn changes_effective(
    _input: &[u8],
    document: &Document,
    font: usize,
    xf: usize,
    change: &FontChange,
) -> Result<bool> {
    let apply_font_disabled = attribute_bool(&document.elements[xf], b"applyFont")? == Some(false);
    let bold = property_value(document, font, b"b")?;
    let italic = property_value(document, font, b"i")?;
    let strike = property_value(document, font, b"strike")?;
    let condense = property_value(document, font, b"condense")?;
    let extend = property_value(document, font, b"extend")?;
    let outline = property_value(document, font, b"outline")?;
    let shadow = property_value(document, font, b"shadow")?;
    let underline = underline_value(document, font)?;
    let color = if change.color_change().is_some() {
        color_value(document, font)?
    } else {
        None
    };
    let name = if change.name_change().is_some() {
        string_value(document, font, b"name")?
    } else {
        None
    };
    let size = if change.size_change().is_some() {
        size_value(document, font)?
    } else {
        None
    };
    let script = if change.script_change().is_some() {
        script_value(document, font)?
    } else {
        None
    };
    let family = if change.family_change().is_some() {
        u32_property_value(document, font, b"family")?
    } else {
        None
    };
    let charset = if change.charset_change().is_some() {
        u32_property_value(document, font, b"charset")?
    } else {
        None
    };
    let scheme = if change.scheme_change().is_some() {
        scheme_value(document, font)?
    } else {
        None
    };
    let property_change = change.bold_change().is_some_and(|value| value != bold)
        || change.italic_change().is_some_and(|value| value != italic)
        || change.strike_change().is_some_and(|value| value != strike)
        || change
            .condense_change()
            .is_some_and(|value| value != condense)
        || change.extend_change().is_some_and(|value| value != extend)
        || change
            .outline_change()
            .is_some_and(|value| value != outline)
        || change.shadow_change().is_some_and(|value| value != shadow)
        || change
            .underline_change()
            .is_some_and(|value| value != underline)
        || change
            .color_change()
            .is_some_and(|value| value.map(color_state) != color)
        || change.name_change().is_some_and(|value| value != name)
        || change.size_change().is_some_and(|value| value != size)
        || change.script_change().is_some_and(|value| value != script)
        || change.family_change().is_some_and(|value| value != family)
        || change
            .charset_change()
            .is_some_and(|value| value != charset)
        || change.scheme_change().is_some_and(|value| value != scheme);
    if property_change {
        return Ok(true);
    }
    let activates_font = change
        .bold_change()
        .is_some_and(|value| value.is_some() && value == bold)
        || change
            .italic_change()
            .is_some_and(|value| value.is_some() && value == italic)
        || change
            .strike_change()
            .is_some_and(|value| value.is_some() && value == strike)
        || change
            .condense_change()
            .is_some_and(|value| value.is_some() && value == condense)
        || change
            .extend_change()
            .is_some_and(|value| value.is_some() && value == extend)
        || change
            .outline_change()
            .is_some_and(|value| value.is_some() && value == outline)
        || change
            .shadow_change()
            .is_some_and(|value| value.is_some() && value == shadow)
        || change
            .underline_change()
            .is_some_and(|value| value.is_some() && value == underline)
        || change
            .color_change()
            .is_some_and(|value| value.is_some() && value.map(color_state) == color)
        || change
            .name_change()
            .is_some_and(|value| value.is_some() && value == name)
        || change
            .size_change()
            .is_some_and(|value| value.is_some() && value == size)
        || change
            .script_change()
            .is_some_and(|value| value.is_some() && value == script)
        || change
            .family_change()
            .is_some_and(|value| value.is_some() && value == family)
        || change
            .charset_change()
            .is_some_and(|value| value.is_some() && value == charset)
        || change
            .scheme_change()
            .is_some_and(|value| value.is_some() && value == scheme);
    Ok(apply_font_disabled && activates_font)
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct AlignmentState {
    horizontal: Option<Horizontal>,
    vertical: Option<Vertical>,
    text_rotation: Option<Rotation>,
    wrap_text: Option<bool>,
    indent: Option<Indent>,
    relative_indent: Option<i32>,
    justify_last_line: Option<bool>,
    shrink_to_fit: Option<bool>,
    reading_order: Option<Reading>,
}

fn direct_alignment(document: &Document, xf: usize) -> Result<Option<usize>> {
    document
        .alignment_properties
        .get(xf)
        .map(|properties| properties.alignment)
        .ok_or_else(|| invalid("styles cell-XF index is out of range"))
}

fn direct_protection(document: &Document, xf: usize) -> Result<Option<usize>> {
    document
        .alignment_properties
        .get(xf)
        .map(|properties| properties.protection)
        .ok_or_else(|| invalid("styles cell-XF index is out of range"))
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct ProtectionState {
    locked: Option<bool>,
    hidden: Option<bool>,
}

fn protection_state(document: &Document, xf: usize) -> Result<ProtectionState> {
    let Some(protection) = direct_protection(document, xf)? else {
        return Ok(ProtectionState::default());
    };
    let element = &document.elements[protection];
    Ok(ProtectionState {
        locked: attribute_bool(element, b"locked")?,
        hidden: attribute_bool(element, b"hidden")?,
    })
}

fn protection_changes_effective(
    document: &Document,
    xf: usize,
    change: &ProtectionChange,
) -> Result<bool> {
    if change.is_empty() {
        return Ok(false);
    }
    let apply_protection_disabled =
        attribute_bool(&document.elements[xf], b"applyProtection")? == Some(false);
    let state = protection_state(document, xf)?;
    let property_change = change
        .locked_change()
        .is_some_and(|value| value != state.locked)
        || change
            .hidden_change()
            .is_some_and(|value| value != state.hidden);
    if property_change {
        return Ok(true);
    }
    let activates_protection = change
        .locked_change()
        .is_some_and(|value| value.is_some() && value == state.locked)
        || change
            .hidden_change()
            .is_some_and(|value| value.is_some() && value == state.hidden);
    Ok(apply_protection_disabled && activates_protection)
}

fn quote_prefix_changes_effective(
    document: &Document,
    xf: usize,
    change: Option<Option<bool>>,
) -> Result<bool> {
    let Some(change) = change else {
        return Ok(false);
    };
    let state = attribute_bool(&document.elements[xf], b"quotePrefix")?;
    Ok(change != state)
}

#[derive(Debug, Clone, Default, PartialEq)]
struct FillState<'a> {
    pattern: Option<FillPattern>,
    foreground: Option<FillColor>,
    background: Option<FillColor>,
    gradient: Option<&'a GradientProjection>,
}

fn direct_fill(document: &Document, xf: usize) -> Result<usize> {
    let id = attribute_u32(&document.elements[xf], b"fillId")?.unwrap_or(0);
    let index = usize::try_from(id).map_err(|_| invalid("styles fill index does not fit usize"))?;
    if document.virtual_fills {
        return match index {
            0 => Ok(VIRTUAL_NONE_FILL),
            1 => Ok(VIRTUAL_GRAY125_FILL),
            _ => Err(invalid("styles fill reference is out of range")),
        };
    }
    document
        .fill_items
        .get(index)
        .copied()
        .ok_or_else(|| invalid("styles fill reference is out of range"))
}

const VIRTUAL_NONE_BORDER: usize = usize::MAX - 2;

fn direct_border(document: &Document, xf: usize) -> Result<usize> {
    let id = attribute_u32(&document.elements[xf], b"borderId")?.unwrap_or(0);
    let index =
        usize::try_from(id).map_err(|_| invalid("styles border index does not fit usize"))?;
    if document.virtual_borders {
        return if index == 0 {
            Ok(VIRTUAL_NONE_BORDER)
        } else {
            Err(invalid("styles border reference is out of range"))
        };
    }
    document
        .border_items
        .get(index)
        .copied()
        .ok_or_else(|| invalid("styles border reference is out of range"))
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct BorderSideState {
    present: bool,
    line: Option<Line>,
    color: Option<BorderColor>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct BorderState {
    sides: [BorderSideState; 9],
    diagonal_up: Option<bool>,
    diagonal_down: Option<bool>,
    outline: Option<bool>,
}

fn border_color_value(document: &Document, color: usize) -> Result<BorderColor> {
    let element = document
        .elements
        .get(color)
        .ok_or_else(|| invalid("styles border color index is out of range"))?;
    let rgb = attribute_str(element, b"rgb", "border RGB color")?;
    let theme = attribute_u32(element, b"theme")?;
    let indexed = attribute_u32(element, b"indexed")?;
    let automatic = attribute_bool(element, b"auto")?;
    let base_count = usize::from(rgb.is_some())
        + usize::from(theme.is_some())
        + usize::from(indexed.is_some())
        + usize::from(automatic.is_some());
    if base_count > 1 {
        return Err(invalid(
            "styles border color has multiple mutually exclusive base values",
        ));
    }
    let base = if let Some(rgb) = rgb {
        BorderColor::from_rgb(
            rgb.trim()
                .parse::<Rgb>()
                .map_err(|_| invalid("styles border RGB color is invalid"))?,
        )
    } else if let Some(theme) = theme {
        BorderColor::theme(theme)
    } else if let Some(indexed) = indexed {
        BorderColor::indexed(indexed)
    } else if let Some(automatic) = automatic {
        BorderColor::auto_value(automatic)
    } else {
        BorderColor::default_base()
    };
    let tint = attribute_str(element, b"tint", "border tint")?
        .map(|value| {
            let value = value
                .trim()
                .parse::<f64>()
                .map_err(|_| invalid("styles border tint is invalid"))?;
            Tint::new(value).map_err(|_| invalid("styles border tint is invalid"))
        })
        .transpose()?;
    Ok(tint.map_or(base, |tint| base.with_tint(tint)))
}

fn border_state(document: &Document, border: usize) -> Result<BorderState> {
    if border == VIRTUAL_NONE_BORDER {
        return Ok(BorderState::default());
    }
    let element = document
        .elements
        .get(border)
        .ok_or_else(|| invalid("styles border index is out of range"))?;
    let mut state = BorderState {
        diagonal_up: attribute_bool(element, b"diagonalUp")?,
        diagonal_down: attribute_bool(element, b"diagonalDown")?,
        outline: attribute_bool(element, b"outline")?,
        ..BorderState::default()
    };
    let properties = document
        .border_properties
        .get(&border)
        .ok_or_else(|| invalid("styles border property index is out of range"))?;
    for edge in 0..9 {
        let side = properties.sides[edge];
        let Some(side_element) = side.side else {
            continue;
        };
        let line = attribute_str(
            &document.elements[side_element],
            b"style",
            "border side style",
        )?
        .map(|value| {
            Line::from_xml(value.trim()).map_err(|_| invalid("styles border side style is invalid"))
        })
        .transpose()?
        .flatten();
        let color = side
            .color
            .map(|color| border_color_value(document, color))
            .transpose()?;
        state.sides[edge] = BorderSideState {
            present: true,
            line,
            color,
        };
    }
    Ok(state)
}

fn border_side_change_effective(
    state: BorderSideState,
    change: &crate::style::BorderSideChange,
) -> Result<bool> {
    let (requested_line, requested_color) = border_side_requested(change);
    if requested_color.is_some_and(|value| value.is_some()) {
        let requested_line_value = requested_line.flatten().or(state.line);
        if !state.present && requested_line_value.is_none() {
            return Err(Error::Unsupported {
                feature: "border color edits require an existing typed side",
            });
        }
    }
    let wants_present = match change.whole_change() {
        Some(Some(_)) => true,
        Some(None) => {
            requested_line.is_some_and(|value| value.is_some())
                || requested_color.is_some_and(|value| value.is_some())
        },
        None => {
            state.present
                || requested_line.is_some_and(|value| value.is_some())
                || requested_color.is_some_and(|value| value.is_some())
        },
    };
    if wants_present != state.present {
        return Ok(true);
    }
    if !wants_present {
        return Ok(false);
    }
    Ok(requested_line.is_some_and(|value| value != state.line)
        || requested_color.is_some_and(|value| value != state.color))
}

fn border_side_requested(
    change: &crate::style::BorderSideChange,
) -> (Option<Option<Line>>, Option<Option<BorderColor>>) {
    let mut line = change
        .whole_change()
        .map(|value| value.map(|side| side.line));
    if let Some(value) = change.line_change() {
        line = Some(value);
    }
    let mut color = change
        .whole_change()
        .map(|value| value.and_then(|side| side.color));
    if let Some(value) = change.color_change() {
        color = Some(value);
    }
    (line, color)
}

fn border_changes_effective(document: &Document, xf: usize, change: &BorderChange) -> Result<bool> {
    if change.is_empty() {
        return Ok(false);
    }
    let apply_disabled = attribute_bool(&document.elements[xf], b"applyBorder")? == Some(false);
    let border = direct_border(document, xf)?;
    let strict = if border == VIRTUAL_NONE_BORDER {
        document
            .cell_xfs
            .and_then(|cell_xfs| document.elements.get(cell_xfs))
            .is_some_and(|element| element.strict_spreadsheetml)
    } else {
        document.elements[border].strict_spreadsheetml
    };
    if strict
        && (change.side_change(BorderEdge::Left).is_some()
            || change.side_change(BorderEdge::Right).is_some())
    {
        return Err(Error::Unsupported {
            feature: "physical left/right borders require Transitional SpreadsheetML",
        });
    }
    if !strict
        && (change.side_change(BorderEdge::Start).is_some()
            || change.side_change(BorderEdge::End).is_some())
    {
        return Err(Error::Unsupported {
            feature: "logical start/end borders require Strict SpreadsheetML",
        });
    }
    let state = border_state(document, border)?;
    let mut property_change = false;
    for (edge, index) in [
        (BorderEdge::Start, 0usize),
        (BorderEdge::End, 1),
        (BorderEdge::Left, 2),
        (BorderEdge::Right, 3),
        (BorderEdge::Top, 4),
        (BorderEdge::Bottom, 5),
        (BorderEdge::Diagonal, 6),
        (BorderEdge::Vertical, 7),
        (BorderEdge::Horizontal, 8),
    ] {
        if let Some(side) = change.side_change(edge) {
            property_change |= border_side_change_effective(state.sides[index], side)?;
        }
    }
    property_change |= change
        .diagonal_up_change()
        .is_some_and(|value| value != state.diagonal_up);
    property_change |= change
        .diagonal_down_change()
        .is_some_and(|value| value != state.diagonal_down);
    property_change |= change
        .outline_change()
        .is_some_and(|value| value != state.outline);
    if property_change {
        return Ok(true);
    }
    let activates = change
        .diagonal_up_change()
        .is_some_and(|value| value.is_some() && value == state.diagonal_up)
        || change
            .diagonal_down_change()
            .is_some_and(|value| value.is_some() && value == state.diagonal_down)
        || change
            .outline_change()
            .is_some_and(|value| value.is_some() && value == state.outline)
        || [
            (BorderEdge::Start, 0usize),
            (BorderEdge::End, 1),
            (BorderEdge::Left, 2),
            (BorderEdge::Right, 3),
            (BorderEdge::Top, 4),
            (BorderEdge::Bottom, 5),
            (BorderEdge::Diagonal, 6),
            (BorderEdge::Vertical, 7),
            (BorderEdge::Horizontal, 8),
        ]
        .into_iter()
        .any(|(edge, index)| {
            change
                .side_change(edge)
                .is_some_and(|side| border_side_activates(state.sides[index], side))
        });
    Ok(apply_disabled && activates)
}

fn border_side_activates(state: BorderSideState, change: &crate::style::BorderSideChange) -> bool {
    let (line, color) = border_side_requested(change);
    let concrete = change.whole_change().is_some_and(|value| value.is_some())
        || line.is_some_and(|value| value.is_some())
        || color.is_some_and(|value| value.is_some());
    state.present
        && concrete
        && line.is_none_or(|value| value == state.line)
        && color.is_none_or(|value| value == state.color)
}

fn validate_fill_references(document: &Document, xfs: &[usize]) -> Result<()> {
    let fill_count = if document.virtual_fills {
        2
    } else {
        document.fill_items.len()
    };
    let style_xfs = document
        .cell_style_xfs
        .map(|collection| direct_items(&document.elements, collection, b"xf"))
        .transpose()?
        .unwrap_or_default();
    for &xf in xfs.iter().chain(style_xfs.iter()) {
        let Some(id) = attribute_u32(&document.elements[xf], b"fillId")? else {
            continue;
        };
        let id =
            usize::try_from(id).map_err(|_| invalid("styles fill index does not fit usize"))?;
        if document.virtual_fills && id != 0 {
            return Err(invalid("styles fill reference is out of range"));
        }
        if id >= fill_count {
            return Err(invalid("styles fill reference is out of range"));
        }
    }
    Ok(())
}

fn direct_fill_pattern(document: &Document, fill: usize) -> Result<Option<usize>> {
    document
        .fill_properties
        .get(fill)
        .map(|properties| properties.pattern_fill)
        .ok_or_else(|| invalid("styles fill index is out of range"))
}

fn fill_color_value(document: &Document, color: usize) -> Result<FillColor> {
    fill_color_value_in_elements(&document.elements, color)
}

fn fill_color_value_in_elements(elements: &[Element], color: usize) -> Result<FillColor> {
    let element = elements
        .get(color)
        .ok_or_else(|| invalid("styles fill color index is out of range"))?;
    let rgb = attribute_str(element, b"rgb", "fill RGB color")?;
    let theme = attribute_u32(element, b"theme")?;
    let indexed = attribute_u32(element, b"indexed")?;
    let automatic = attribute_bool(element, b"auto")?;
    let base_count = usize::from(rgb.is_some())
        + usize::from(theme.is_some())
        + usize::from(indexed.is_some())
        + usize::from(automatic.is_some());
    if base_count > 1 {
        return Err(invalid(
            "styles fill color has multiple mutually exclusive base values",
        ));
    }
    let base = if let Some(rgb) = rgb {
        FillColor::from_rgb(
            rgb.trim()
                .parse::<Rgb>()
                .map_err(|_| invalid("styles fill RGB color is invalid"))?,
        )
    } else if let Some(theme) = theme {
        FillColor::theme(theme)
    } else if let Some(indexed) = indexed {
        FillColor::indexed(indexed)
    } else if let Some(automatic) = automatic {
        FillColor::auto_value(automatic)
    } else {
        FillColor::default_base()
    };
    let tint = attribute_str(element, b"tint", "fill tint")?
        .map(|value| {
            let value = value
                .trim()
                .parse::<f64>()
                .map_err(|_| invalid("styles fill tint is invalid"))?;
            Tint::new(value).map_err(|_| invalid("styles fill tint is invalid"))
        })
        .transpose()?;
    Ok(tint.map_or(base, |tint| base.with_tint(tint)))
}

fn fill_state<'a>(document: &'a Document, fill: usize) -> Result<FillState<'a>> {
    if document.virtual_fills {
        return match fill {
            VIRTUAL_NONE_FILL => Ok(FillState {
                // A missing or empty fills collection has no authored
                // patternType on its implicit fillId=0 record.  Keep that
                // distinction so clearing an absent pattern is an exact
                // no-op; an explicit patternType="none" remains writable.
                pattern: None,
                foreground: None,
                background: None,
                gradient: None,
            }),
            VIRTUAL_GRAY125_FILL => Ok(FillState {
                pattern: Some(FillPattern::Gray125),
                foreground: None,
                background: None,
                gradient: None,
            }),
            _ => Err(invalid("styles virtual fill index is out of range")),
        };
    }
    let properties = document
        .fill_properties
        .get(fill)
        .ok_or_else(|| invalid("styles fill index is out of range"))?;
    if properties.pattern_fill.is_some() && properties.gradient_fill.is_some() {
        return Err(invalid(
            "styles fill has both patternFill and gradientFill children",
        ));
    }
    if properties.gradient_fill.is_some() {
        let gradient = properties.gradient.as_ref().ok_or(Error::Unsupported {
            feature: "source gradient fills are only writable by whole-fill replacement",
        })?;
        return Ok(FillState {
            gradient: Some(gradient),
            ..FillState::default()
        });
    }
    let Some(pattern) = properties.pattern_fill else {
        return Ok(FillState::default());
    };
    let pattern_element = &document.elements[pattern];
    let pattern = attribute_str(pattern_element, b"patternType", "fill pattern")?
        .map(|value| {
            FillPattern::from_token(value.trim())
                .ok_or_else(|| invalid("styles fill patternType is invalid"))
        })
        .transpose()?;
    let foreground = properties
        .foreground
        .map(|color| fill_color_value(document, color))
        .transpose()?;
    let background = properties
        .background
        .map(|color| fill_color_value(document, color))
        .transpose()?;
    Ok(FillState {
        pattern,
        foreground,
        background,
        gradient: None,
    })
}

fn fill_changes_effective(document: &Document, xf: usize, change: &FillChange) -> Result<bool> {
    if change.is_empty() {
        return Ok(false);
    }
    let apply_fill_disabled = attribute_bool(&document.elements[xf], b"applyFill")? == Some(false);
    let fill = direct_fill(document, xf)?;
    let state = fill_state(document, fill)?;
    if let Some(replacement) = change.replacement() {
        if replacement.gradient().is_some()
            && (change.pattern_change().is_some()
                || change.foreground_change().is_some()
                || change.background_change().is_some())
        {
            return Err(Error::Unsupported {
                feature: "whole gradient-fill replacements cannot be combined with pattern properties",
            });
        }
        let matches = replacement_matches_state(&state, replacement, change)?;
        return Ok(!matches || apply_fill_disabled);
    }
    if state.gradient.is_some() {
        return Err(Error::Unsupported {
            feature: "source gradient fills are not writable by pattern-fill edits",
        });
    }
    let property_change = change
        .pattern_change()
        .is_some_and(|value| value != state.pattern)
        || change
            .foreground_change()
            .is_some_and(|value| value != state.foreground)
        || change
            .background_change()
            .is_some_and(|value| value != state.background);
    if property_change {
        return Ok(true);
    }
    let activates_fill = change
        .pattern_change()
        .is_some_and(|value| value.is_some() && value == state.pattern)
        || change
            .foreground_change()
            .is_some_and(|value| value.is_some() && value == state.foreground)
        || change
            .background_change()
            .is_some_and(|value| value.is_some() && value == state.background);
    Ok(apply_fill_disabled && activates_fill)
}

fn replacement_matches_state(
    state: &FillState<'_>,
    replacement: &Fill,
    change: &FillChange,
) -> Result<bool> {
    match replacement {
        Fill::Pattern(pattern) => {
            let requested_pattern = change.pattern_change().unwrap_or(pattern.pattern());
            let requested_foreground = change.foreground_change().unwrap_or(pattern.foreground());
            let requested_background = change.background_change().unwrap_or(pattern.background());
            Ok(state.gradient.is_none()
                && state.pattern == requested_pattern
                && state.foreground == requested_foreground
                && state.background == requested_background)
        },
        Fill::Gradient(gradient) => Ok(state
            .gradient
            .as_ref()
            .is_some_and(|source| gradient_matches(source, gradient))),
    }
}

fn gradient_matches(
    source: &GradientProjection,
    requested: &crate::style::cell_format::GradientFill,
) -> bool {
    source.kind == requested.kind()
        && source.degree == requested.degree()
        && source.left == requested.left()
        && source.right == requested.right()
        && source.top == requested.top()
        && source.bottom == requested.bottom()
        && source.stops.len() == requested.stops().len()
        && source
            .stops
            .iter()
            .zip(requested.stops())
            .all(|(source, requested)| {
                source.position == requested.position() && source.color == requested.color()
            })
}

fn has_attribute(element: &Element, name: &[u8]) -> bool {
    element
        .attributes
        .iter()
        .any(|attribute| attribute.name.as_ref() == name)
}

fn current_number_format_code<'a>(
    document: &'a Document,
    xf: usize,
    number_formats: &NumberFormatIndex,
) -> Result<Option<&'a str>> {
    let id = attribute_u32(&document.elements[xf], b"numFmtId")?.unwrap_or(0);
    if let Some(&num_fmt) = number_formats.by_id.get(&id) {
        return attribute_str(
            &document.elements[num_fmt],
            b"formatCode",
            "number-format code",
        );
    }
    Ok(builtin_format_code(id))
}

fn number_format_changes_effective(
    document: &Document,
    xf: usize,
    change: Option<Option<&str>>,
    number_formats: &NumberFormatIndex,
) -> Result<bool> {
    let Some(change) = change else {
        return Ok(false);
    };
    // Validate the activation flag even when the requested code is a semantic
    // no-op.  A malformed source attribute must never be silently repaired by
    // an unrelated-looking style request.
    let apply = attribute_bool(&document.elements[xf], b"applyNumberFormat")?;
    let has_num_fmt_id = has_attribute(&document.elements[xf], b"numFmtId");
    let has_apply = has_attribute(&document.elements[xf], b"applyNumberFormat");
    match change {
        None => Ok(has_num_fmt_id || has_apply),
        Some(code) => {
            NumberFormat::validate_code(code).map_err(|error| invalid(error.to_string()))?;
            let current = current_number_format_code(document, xf, number_formats)?;
            // Match the other native style facets: an explicit false
            // activation flag is repaired, while an absent flag is retained
            // as an exact semantic no-op when the authored code already
            // matches.
            Ok(current != Some(code) || apply == Some(false))
        },
    }
}

fn canonical_builtin_format_id(
    document: &Document,
    number_formats: &NumberFormatIndex,
    code: &str,
) -> Result<Option<u32>> {
    for id in 0..FIRST_CUSTOM_ID {
        if builtin_format_code(id) != Some(code) {
            continue;
        }
        // A source declaration for a reserved ID overrides the built-in
        // meaning in the source document. Reuse that ID only when its
        // declaration agrees; otherwise allocate a safe custom ID.
        if let Some(&num_fmt) = number_formats.by_id.get(&id) {
            let declared = attribute_str(
                &document.elements[num_fmt],
                b"formatCode",
                "number-format code",
            )?;
            if declared != Some(code) {
                continue;
            }
        }
        return Ok(Some(id));
    }
    Ok(None)
}

fn resolve_number_format(
    document: &Document,
    number_formats: &mut NumberFormatIndex,
    code: &str,
    max_number_formats: usize,
    max_output: usize,
) -> Result<(u32, Option<Vec<u8>>)> {
    NumberFormat::validate_code(code).map_err(|error| invalid(error.to_string()))?;
    if let Some(id) = canonical_builtin_format_id(document, number_formats, code)? {
        return Ok((id, None));
    }
    if let Some(&id) = number_formats.by_code.get(code.as_bytes()) {
        return Ok((id, None));
    }

    let generated_count = number_formats
        .items
        .len()
        .checked_add(number_formats.generated_count)
        .and_then(|count| count.checked_add(1))
        .ok_or_else(|| invalid("styles number-format count overflows"))?;
    enforce_limit(
        "styles generated number formats",
        generated_count,
        max_number_formats,
        litchi_core::Resource::Objects,
    )?;

    let mut id = number_formats
        .next_custom_id
        .ok_or_else(|| invalid("styles number-format ID space is exhausted"))?;
    while number_formats.used_ids.contains(&id) {
        id = id
            .checked_add(1)
            .ok_or_else(|| invalid("styles number-format ID space is exhausted"))?;
    }
    let record = new_num_fmt_element(document, id, code, max_output)?;
    let key = copy_bytes(code.as_bytes(), "styles number-format code index")?;
    number_formats.by_code.insert(key, id);
    number_formats.used_ids.insert(id);
    number_formats.next_custom_id = id.checked_add(1);
    number_formats.generated_count = number_formats
        .generated_count
        .checked_add(1)
        .ok_or_else(|| invalid("styles generated number-format count overflows"))?;
    Ok((id, Some(record)))
}

fn alignment_state(document: &Document, xf: usize) -> Result<AlignmentState> {
    let Some(alignment) = direct_alignment(document, xf)? else {
        return Ok(AlignmentState::default());
    };
    let element = &document.elements[alignment];
    let horizontal = attribute_str(element, b"horizontal", "alignment horizontal")?
        .map(str::trim)
        .map(|value| {
            value
                .parse::<Horizontal>()
                .map_err(|_| invalid("styles alignment horizontal value is invalid"))
        })
        .transpose()?;
    let vertical = attribute_str(element, b"vertical", "alignment vertical")?
        .map(str::trim)
        .map(|value| {
            value
                .parse::<Vertical>()
                .map_err(|_| invalid("styles alignment vertical value is invalid"))
        })
        .transpose()?;
    let text_rotation = attribute_u32(element, b"textRotation")?
        .map(|value| {
            Rotation::try_from(value)
                .map_err(|_| invalid("styles alignment textRotation value is invalid"))
        })
        .transpose()?;
    let wrap_text = attribute_bool(element, b"wrapText")?;
    let indent = attribute_u32(element, b"indent")?
        .map(|value| {
            Indent::try_from(value).map_err(|_| invalid("styles alignment indent value is invalid"))
        })
        .transpose()?;
    let shrink_to_fit = attribute_bool(element, b"shrinkToFit")?;
    let reading_order = attribute_u32(element, b"readingOrder")?
        .map(|value| {
            Reading::try_from(value)
                .map_err(|_| invalid("styles alignment readingOrder value is invalid"))
        })
        .transpose()?;
    let relative_indent = attribute_str(element, b"relativeIndent", "relative alignment indent")?
        .map(|value| {
            value
                .trim()
                .parse::<i32>()
                .map_err(|_| invalid("styles alignment relativeIndent value is invalid"))
        })
        .transpose()?;
    let justify_last_line = attribute_bool(element, b"justifyLastLine")?;
    Ok(AlignmentState {
        horizontal,
        vertical,
        text_rotation,
        wrap_text,
        indent,
        relative_indent,
        justify_last_line,
        shrink_to_fit,
        reading_order,
    })
}

fn alignment_changes_effective(
    document: &Document,
    xf: usize,
    change: &AlignmentChange,
) -> Result<bool> {
    if change.is_empty() {
        return Ok(false);
    }
    let strict_spreadsheetml = document
        .elements
        .get(xf)
        .ok_or_else(|| invalid("styles cell-XF index is out of range"))?
        .strict_spreadsheetml;
    let strict_alignment = direct_alignment(document, xf)?
        .is_some_and(|index| document.elements[index].strict_spreadsheetml);
    if (strict_spreadsheetml || strict_alignment)
        && change
            .text_rotation_change()
            .is_some_and(|value| value.is_some_and(|rotation| rotation.is_contextual()))
    {
        return Err(invalid(
            "context-dependent text rotation 254 requires Transitional SpreadsheetML",
        ));
    }
    let apply_alignment_disabled =
        attribute_bool(&document.elements[xf], b"applyAlignment")? == Some(false);
    let state = alignment_state(document, xf)?;
    let property_change = change
        .horizontal_change()
        .is_some_and(|value| value != state.horizontal)
        || change
            .vertical_change()
            .is_some_and(|value| value != state.vertical)
        || change
            .text_rotation_change()
            .is_some_and(|value| value != state.text_rotation)
        || change
            .wrap_text_change()
            .is_some_and(|value| value != state.wrap_text)
        || change
            .indent_change()
            .is_some_and(|value| value != state.indent)
        || change
            .relative_indent_change()
            .is_some_and(|value| value != state.relative_indent)
        || change
            .justify_last_line_change()
            .is_some_and(|value| value != state.justify_last_line)
        || change
            .shrink_to_fit_change()
            .is_some_and(|value| value != state.shrink_to_fit)
        || change
            .reading_order_change()
            .is_some_and(|value| value != state.reading_order);
    if property_change {
        return Ok(true);
    }
    let activates_alignment = change
        .horizontal_change()
        .is_some_and(|value| value.is_some() && value == state.horizontal)
        || change
            .vertical_change()
            .is_some_and(|value| value.is_some() && value == state.vertical)
        || change
            .text_rotation_change()
            .is_some_and(|value| value.is_some() && value == state.text_rotation)
        || change
            .wrap_text_change()
            .is_some_and(|value| value.is_some() && value == state.wrap_text)
        || change
            .indent_change()
            .is_some_and(|value| value.is_some() && value == state.indent)
        || change
            .relative_indent_change()
            .is_some_and(|value| value.is_some() && value == state.relative_indent)
        || change
            .justify_last_line_change()
            .is_some_and(|value| value.is_some() && value == state.justify_last_line)
        || change
            .shrink_to_fit_change()
            .is_some_and(|value| value.is_some() && value == state.shrink_to_fit)
        || change
            .reading_order_change()
            .is_some_and(|value| value.is_some() && value == state.reading_order);
    Ok(apply_alignment_disabled && activates_alignment)
}

fn changed_font(
    input: &[u8],
    document: &Document,
    font: usize,
    change: &FontChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let element = &document.elements[font];
    let source = &input[element.start..element.end];
    if element.empty {
        let mut insertions = Vec::new();
        insertions
            .try_reserve_exact(15)
            .map_err(|source| allocation("styles font property insertions", source))?;
        for (property, requested) in [
            (b"b".as_slice(), change.bold_change()),
            (b"i".as_slice(), change.italic_change()),
            (b"strike".as_slice(), change.strike_change()),
            (b"condense".as_slice(), change.condense_change()),
            (b"extend".as_slice(), change.extend_change()),
            (b"outline".as_slice(), change.outline_change()),
            (b"shadow".as_slice(), change.shadow_change()),
        ] {
            if let Some(Some(value)) = requested
                && direct_property(document, font, property)?.is_none()
            {
                insertions.push(new_property_name(document, font, property, value)?);
            }
        }
        if let Some(Some(value)) = change.underline_change()
            && direct_property(document, font, b"u")?.is_none()
        {
            insertions.push(new_underline_property_name(document, font, value)?);
        }
        if let Some(Some(value)) = change.script_change()
            && direct_property(document, font, b"vertAlign")?.is_none()
        {
            insertions.push(new_script_property_name(document, font, value, max_output)?);
        }
        if let Some(Some(value)) = change.size_change()
            && direct_property(document, font, b"sz")?.is_none()
        {
            insertions.push(new_size_property_name(document, font, value, max_output)?);
        }
        if let Some(Some(value)) = change.color_change()
            && direct_property(document, font, b"color")?.is_none()
        {
            insertions.push(new_color_property_name(document, font, value, max_output)?);
        }
        if let Some(Some(value)) = change.name_change()
            && direct_property(document, font, b"name")?.is_none()
        {
            insertions.push(new_name_property_name(document, font, value, max_output)?);
        }
        if let Some(Some(value)) = change.family_change()
            && direct_property(document, font, b"family")?.is_none()
        {
            insertions.push(new_value_property_name(
                document,
                font,
                b"family",
                &value.to_string(),
                max_output,
            )?);
        }
        if let Some(Some(value)) = change.charset_change()
            && direct_property(document, font, b"charset")?.is_none()
        {
            insertions.push(new_value_property_name(
                document,
                font,
                b"charset",
                &value.to_string(),
                max_output,
            )?);
        }
        if let Some(Some(value)) = change.scheme_change()
            && direct_property(document, font, b"scheme")?.is_none()
        {
            insertions.push(new_value_property_name(
                document,
                font,
                b"scheme",
                value.as_str(),
                max_output,
            )?);
        }
        if insertions.is_empty() {
            return Ok(source.to_vec());
        }
        return expand_empty_font(source, element, &insertions, max_output);
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(15)
        .map_err(|source| allocation("styles font property edits", source))?;
    apply_property_change(
        input,
        document,
        font,
        b"b",
        change.bold_change(),
        &mut edits,
        max_output,
    )?;
    apply_property_change(
        input,
        document,
        font,
        b"i",
        change.italic_change(),
        &mut edits,
        max_output,
    )?;
    apply_property_change(
        input,
        document,
        font,
        b"strike",
        change.strike_change(),
        &mut edits,
        max_output,
    )?;
    for (property, change) in [
        (b"condense".as_slice(), change.condense_change()),
        (b"extend".as_slice(), change.extend_change()),
        (b"outline".as_slice(), change.outline_change()),
        (b"shadow".as_slice(), change.shadow_change()),
    ] {
        apply_property_change(
            input, document, font, property, change, &mut edits, max_output,
        )?;
    }
    apply_underline_change(
        input,
        document,
        font,
        change.underline_change(),
        &mut edits,
        max_output,
    )?;
    apply_script_change(
        input,
        document,
        font,
        change.script_change(),
        &mut edits,
        max_output,
    )?;
    apply_size_change(
        input,
        document,
        font,
        change.size_change(),
        &mut edits,
        max_output,
    )?;
    apply_color_change(
        input,
        document,
        font,
        change.color_change(),
        &mut edits,
        max_output,
    )?;
    apply_name_change(
        input,
        document,
        font,
        change.name_change(),
        &mut edits,
        max_output,
    )?;
    apply_u32_change(
        input,
        document,
        font,
        b"family",
        change.family_change(),
        &mut edits,
        max_output,
    )?;
    apply_u32_change(
        input,
        document,
        font,
        b"charset",
        change.charset_change(),
        &mut edits,
        max_output,
    )?;
    apply_scheme_change(
        input,
        document,
        font,
        change.scheme_change(),
        &mut edits,
        max_output,
    )?;
    let mut relative = Vec::new();
    relative
        .try_reserve_exact(edits.len())
        .map_err(|source| allocation("styles relative font edits", source))?;
    for edit in edits {
        relative.push(ByteEdit {
            start: edit.start - element.start,
            end: edit.end - element.start,
            replacement: edit.replacement,
        });
    }
    apply_edits(source, relative, max_output)
}

fn direct_property(document: &Document, font: usize, property: &[u8]) -> Result<Option<usize>> {
    let properties = document
        .font_properties
        .get(font)
        .ok_or_else(|| invalid("styles font index is out of range"))?;
    Ok(match property {
        b"b" => properties.bold,
        b"i" => properties.italic,
        b"strike" => properties.strike,
        b"condense" => properties.condense,
        b"extend" => properties.extend,
        b"outline" => properties.outline,
        b"shadow" => properties.shadow,
        b"u" => properties.underline,
        b"color" => properties.color,
        b"vertAlign" => properties.script,
        b"sz" => properties.size,
        b"name" => properties.name,
        b"family" => properties.family,
        b"charset" => properties.charset,
        b"scheme" => properties.scheme,
        _ => None,
    })
}

fn apply_property_change(
    input: &[u8],
    document: &Document,
    font: usize,
    property: &[u8],
    change: Option<Option<bool>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, property)?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if property_value(document, font, property)? == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_boolean_opening(
                        &input[element.start..element.open_end],
                        value,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| match property {
            b"b" => properties.bold_insert_at,
            b"i" => properties.italic_insert_at,
            b"strike" => properties.strike_insert_at,
            b"condense" => properties.condense_insert_at,
            b"extend" => properties.extend_insert_at,
            b"outline" => properties.outline_insert_at,
            b"shadow" => properties.shadow_insert_at,
            b"u" => properties.underline_insert_at,
            _ => None,
        })
        .ok_or_else(|| invalid("styles font has no insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_property_name(document, font, property, value)?,
    });
    Ok(())
}

fn apply_u32_change(
    input: &[u8],
    document: &Document,
    font: usize,
    property: &[u8],
    change: Option<Option<u32>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, property)?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if u32_property_value(document, font, property)? == Some(value) {
                    return Ok(());
                }
                let lexical = value.to_string();
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        &lexical,
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| match property {
            b"family" => properties.family_insert_at,
            b"charset" => properties.charset_insert_at,
            _ => None,
        })
        .ok_or_else(|| invalid("styles font has no insertion point"))?;
    let lexical = value.to_string();
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_value_property_name(document, font, property, &lexical, max_output)?,
    });
    Ok(())
}

fn apply_scheme_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<Scheme>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"scheme")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if scheme_value(document, font)? == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        value.as_str(),
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.scheme_insert_at)
        .ok_or_else(|| invalid("styles font has no scheme insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_value_property_name(
            document,
            font,
            b"scheme",
            value.as_str(),
            max_output,
        )?,
    });
    Ok(())
}

fn underline_value(document: &Document, font: usize) -> Result<Option<Underline>> {
    let Some(child) = direct_property(document, font, b"u")? else {
        return Ok(None);
    };
    let Some(value) = document.elements[child]
        .attributes
        .iter()
        .find(|attribute| attribute.name.as_ref() == b"val")
        .map(|attribute| attribute.value.as_ref())
    else {
        return Ok(Some(Underline::Single));
    };
    let value = std::str::from_utf8(value)
        .map_err(|_| invalid("styles font underline value is not UTF-8"))?
        .trim();
    value
        .parse::<Underline>()
        .map(Some)
        .map_err(|_| invalid("styles font underline property has an invalid value"))
}

fn apply_underline_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<Underline>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"u")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if underline_value(document, font)? == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        value.as_str(),
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.underline_insert_at)
        .ok_or_else(|| invalid("styles font has no underline insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_underline_property_name(document, font, value)?,
    });
    Ok(())
}

fn apply_color_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<&FontColor>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"color")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if color_value(document, font)? == Some(color_state(value)) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_color_opening(
                        &input[element.start..element.open_end],
                        color_state(value),
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.color_insert_at)
        .ok_or_else(|| invalid("styles font has no color insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_color_property_name(document, font, value, max_output)?,
    });
    Ok(())
}

fn apply_script_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<Script>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"vertAlign")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if script_value(document, font)? == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        value.as_str(),
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.script_insert_at)
        .ok_or_else(|| invalid("styles font has no script insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_script_property_name(document, font, value, max_output)?,
    });
    Ok(())
}

fn apply_size_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<f64>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"sz")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if size_value(document, font)? == Some(value) {
                    return Ok(());
                }
                let lexical = value.to_string();
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        &lexical,
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.size_insert_at)
        .ok_or_else(|| invalid("styles font has no size insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_size_property_name(document, font, value, max_output)?,
    });
    Ok(())
}

fn apply_name_change(
    input: &[u8],
    document: &Document,
    font: usize,
    change: Option<Option<&str>>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = change else {
        return Ok(());
    };
    let child = direct_property(document, font, b"name")?;
    if let Some(child) = child {
        let element = &document.elements[child];
        match requested {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if string_value(document, font, b"name")? == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_attribute(
                        &input[element.start..element.open_end],
                        b"val",
                        value,
                        true,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = requested else {
        return Ok(());
    };
    let insert_at = document
        .font_properties
        .get(font)
        .and_then(|properties| properties.name_insert_at)
        .ok_or_else(|| invalid("styles font has no name insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_name_property_name(document, font, value, max_output)?,
    });
    Ok(())
}

fn new_property_name(
    document: &Document,
    font: usize,
    property: &[u8],
    value: bool,
) -> Result<Vec<u8>> {
    let font_element = &document.elements[font];
    // The font QName is known to be bound in the containing scope.  A
    // different prefix used by an existing child may have been declared only
    // on that child, so reusing it for a sibling could create an unbound name.
    let prefix_len = font_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname_len = prefix_len
        .checked_add(property.len())
        .ok_or_else(|| invalid("styles font property name size overflows"))?;
    let output_len = 1usize
        .checked_add(qname_len)
        .and_then(|size| size.checked_add(if value { 2 } else { 10 }))
        .ok_or_else(|| invalid("styles font property size overflows"))?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles font property", source))?;
    result.extend_from_slice(b"<");
    result.extend_from_slice(&font_element.qname[..prefix_len]);
    result.extend_from_slice(property);
    if !value {
        result.extend_from_slice(b" val=\"0\"");
    }
    result.extend_from_slice(b"/>");
    Ok(result)
}

fn new_underline_property_name(
    document: &Document,
    font: usize,
    value: Underline,
) -> Result<Vec<u8>> {
    let font_element = &document.elements[font];
    let prefix_len = font_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname_len = prefix_len
        .checked_add(1)
        .ok_or_else(|| invalid("styles font property name size overflows"))?;
    let token = value.as_str();
    let output_len = if value == Underline::Single {
        1usize
            .checked_add(qname_len)
            .and_then(|size| size.checked_add(2))
            .ok_or_else(|| invalid("styles font underline property size overflows"))?
    } else {
        1usize
            .checked_add(qname_len)
            .and_then(|size| size.checked_add(8))
            .and_then(|size| size.checked_add(token.len()))
            .ok_or_else(|| invalid("styles font underline property size overflows"))?
    };
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles font underline property", source))?;
    result.extend_from_slice(b"<");
    result.extend_from_slice(&font_element.qname[..prefix_len]);
    result.extend_from_slice(b"u");
    if value == Underline::Single {
        result.extend_from_slice(b"/>");
    } else {
        result.extend_from_slice(b" val=\"");
        result.extend_from_slice(token.as_bytes());
        result.extend_from_slice(b"\"/>");
    }
    Ok(result)
}

fn new_script_property_name(
    document: &Document,
    font: usize,
    value: Script,
    max_output: usize,
) -> Result<Vec<u8>> {
    new_value_property_name(document, font, b"vertAlign", value.as_str(), max_output)
}

fn new_size_property_name(
    document: &Document,
    font: usize,
    value: f64,
    max_output: usize,
) -> Result<Vec<u8>> {
    let value = value.to_string();
    new_value_property_name(document, font, b"sz", &value, max_output)
}

fn new_name_property_name(
    document: &Document,
    font: usize,
    value: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    new_value_property_name(document, font, b"name", value, max_output)
}

fn color_base_attribute(kind: FontColorKind) -> Option<(&'static [u8], String)> {
    match kind {
        FontColorKind::Default => None,
        FontColorKind::Rgb(value) => Some((b"rgb", value.to_string())),
        FontColorKind::Theme(index) => Some((b"theme", index.to_string())),
        FontColorKind::Indexed(index) => Some((b"indexed", index.to_string())),
        FontColorKind::Auto(enabled) => Some((b"auto", if enabled { "1" } else { "0" }.into())),
    }
}

fn new_color_property_name(
    document: &Document,
    font: usize,
    value: &FontColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let state = color_state(value);
    let font_element = &document.elements[font];
    let prefix_len = font_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname_len = prefix_len
        .checked_add(b"color".len())
        .ok_or_else(|| invalid("styles font color name size overflows"))?;
    let base = color_base_attribute(state.kind);
    let tint = state.tint.map(|value| value.get().to_string());
    let mut output_len = qname_len
        .checked_add(3)
        .ok_or_else(|| invalid("styles font color size overflows"))?;
    for (name, value) in base.iter().map(|(name, value)| (*name, value)) {
        let value_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(name.len())
            .and_then(|size| size.checked_add(4))
            .and_then(|size| size.checked_add(value_len))
            .ok_or_else(|| invalid("styles font color size overflows"))?;
    }
    if let Some(value) = tint.as_deref() {
        let value_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(b"tint".len())
            .and_then(|size| size.checked_add(4))
            .and_then(|size| size.checked_add(value_len))
            .ok_or_else(|| invalid("styles font color size overflows"))?;
    }
    enforce_limit(
        "styles font color property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles font color property", source))?;
    result.push(b'<');
    result.extend_from_slice(&font_element.qname[..prefix_len]);
    result.extend_from_slice(b"color");
    if let Some((name, value)) = base {
        append_color_attribute(&mut result, name, &value)?;
    }
    if let Some(value) = tint.as_deref() {
        append_color_attribute(&mut result, b"tint", value)?;
    }
    result.extend_from_slice(b"/>");
    if result.len() != output_len {
        return Err(invalid("styles font color size accounting mismatch"));
    }
    Ok(result)
}

fn rewrite_color_opening(input: &[u8], state: ColorState, max_output: usize) -> Result<Vec<u8>> {
    let attributes = raw_attributes(input)?;
    enforce_limit(
        "styles font color rewrite",
        input.len(),
        max_output,
        litchi_core::Resource::Memory,
    )?;

    let desired_base = color_base_attribute(state.kind);
    let desired_base_name = desired_base.as_ref().map(|(name, _)| *name);
    let has_desired_base = desired_base_name.is_some_and(|name| {
        attributes
            .iter()
            .any(|attribute| attribute.name.as_slice() == name)
    });
    let has_tint = attributes
        .iter()
        .any(|attribute| attribute.name.as_slice() == b"tint");

    let mut output = Vec::new();
    output
        .try_reserve_exact(input.len())
        .map_err(|source| allocation("styles font color rewrite", source))?;
    let mut cursor = 0;
    for attribute in &attributes {
        let remove = if is_color_base_attribute(&attribute.name) {
            desired_base_name.is_none_or(|name| name != attribute.name.as_slice())
        } else {
            attribute.name.as_slice() == b"tint" && state.tint.is_none()
        };
        if remove {
            let mut start = attribute.start;
            while start > cursor && input[start - 1].is_ascii_whitespace() {
                start -= 1;
            }
            output.extend_from_slice(&input[cursor..start]);
            cursor = attribute.end;
        }
    }
    output.extend_from_slice(&input[cursor..]);

    if let Some((name, value)) = desired_base {
        if has_desired_base {
            output = rewrite_attribute(&output, name, &value, false, max_output)?;
        } else if has_tint {
            output = insert_attribute_before(&output, b"tint", name, &value, max_output)?;
        } else {
            output = rewrite_attribute(&output, name, &value, true, max_output)?;
        }
    }
    if let Some(value) = state.tint.map(|value| value.get().to_string()) {
        output = rewrite_attribute(&output, b"tint", &value, true, max_output)?;
    }
    Ok(output)
}

fn is_color_base_attribute(name: &[u8]) -> bool {
    matches!(name, b"rgb" | b"theme" | b"indexed" | b"auto")
}

fn insert_attribute_before(
    input: &[u8],
    anchor: &[u8],
    name: &[u8],
    value: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let attributes = raw_attributes(input)?;
    let Some(anchor) = attributes
        .iter()
        .find(|attribute| attribute.name.as_slice() == anchor)
    else {
        return rewrite_attribute(input, name, value, true, max_output);
    };
    let escaped_len = escaped_attribute_len(value)?;
    let insertion_len = name
        .len()
        .checked_add(escaped_len)
        .and_then(|size| size.checked_add(4))
        .ok_or_else(|| invalid("styles color attribute size overflows"))?;
    let output_len = input
        .len()
        .checked_add(insertion_len)
        .ok_or_else(|| invalid("styles color attribute insertion size overflows"))?;
    enforce_limit(
        "styles color attribute insertion",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles color attribute insertion", source))?;
    output.extend_from_slice(&input[..anchor.start]);
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(&mut output, value, escaped_len)?;
    output.extend_from_slice(b"\" ");
    output.extend_from_slice(&input[anchor.start..]);
    Ok(output)
}

fn append_color_attribute(output: &mut Vec<u8>, name: &[u8], value: &str) -> Result<()> {
    let escaped_len = escaped_attribute_len(value)?;
    output.extend_from_slice(b" ");
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(output, value, escaped_len)?;
    output.push(b'"');
    Ok(())
}

fn new_value_property_name(
    document: &Document,
    font: usize,
    property: &[u8],
    value: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let font_element = &document.elements[font];
    let prefix_len = font_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname_len = prefix_len
        .checked_add(property.len())
        .ok_or_else(|| invalid("styles font property name size overflows"))?;
    let escaped_len = escaped_attribute_len(value)?;
    let output_len = qname_len
        .checked_add(9)
        .and_then(|size| size.checked_add(escaped_len))
        .ok_or_else(|| invalid("styles font property size overflows"))?;
    enforce_limit(
        "styles font property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles font property", source))?;
    result.push(b'<');
    result.extend_from_slice(&font_element.qname[..prefix_len]);
    result.extend_from_slice(property);
    result.extend_from_slice(b" val=\"");
    append_escaped_attribute(&mut result, value, escaped_len)?;
    result.extend_from_slice(b"\"/>");
    Ok(result)
}

fn border_prefix(document: &Document) -> &[u8] {
    document
        .borders
        .or(document.fills)
        .or(document.fonts)
        .or(document.cell_xfs)
        .and_then(|element| document.elements.get(element))
        .map_or(&[][..], |element| qname_prefix(&element.qname))
}

fn border_edge_name(edge: BorderEdge) -> &'static [u8] {
    match edge {
        BorderEdge::Start => b"start",
        BorderEdge::End => b"end",
        BorderEdge::Left => b"left",
        BorderEdge::Right => b"right",
        BorderEdge::Top => b"top",
        BorderEdge::Bottom => b"bottom",
        BorderEdge::Diagonal => b"diagonal",
        BorderEdge::Vertical => b"vertical",
        BorderEdge::Horizontal => b"horizontal",
    }
}

fn border_color_attributes(color: BorderColor) -> Vec<(&'static [u8], String)> {
    let mut attributes: Vec<(&'static [u8], String)> = Vec::new();
    match color {
        BorderColor::Default { .. } => {},
        BorderColor::Rgb { value, .. } => attributes.push((b"rgb", value.to_string())),
        BorderColor::Theme { index, .. } => attributes.push((b"theme", index.to_string())),
        BorderColor::Indexed { index, .. } => attributes.push((b"indexed", index.to_string())),
        BorderColor::Auto { enabled, .. } => {
            attributes.push((b"auto", if enabled { "1" } else { "0" }.to_owned()));
        },
    }
    if let Some(tint) = color.tint() {
        attributes.push((b"tint", tint.get().to_string()));
    }
    attributes
}

fn new_border_color(
    document: &Document,
    border: usize,
    color: BorderColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = if border == VIRTUAL_NONE_BORDER {
        border_prefix(document)
    } else {
        qname_prefix(&document.elements[border].qname)
    };
    let attributes = border_color_attributes(color);
    let attributes_len = attributes.iter().try_fold(0usize, |size, (name, value)| {
        let value_len = escaped_attribute_len(value)?;
        size.checked_add(name.len())
            .and_then(|size| size.checked_add(value_len))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| invalid("styles border color size overflows"))
    })?;
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"color".len()))
        .and_then(|size| size.checked_add(attributes_len))
        .and_then(|size| size.checked_add(2))
        .ok_or_else(|| invalid("styles border color size overflows"))?;
    enforce_limit(
        "styles border color",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles border color", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"color");
    for (name, value) in attributes {
        append_color_attribute(&mut output, name, &value)?;
    }
    output.extend_from_slice(b"/>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn new_border_side(
    document: &Document,
    border: usize,
    edge: BorderEdge,
    side: Side,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = if border == VIRTUAL_NONE_BORDER {
        border_prefix(document)
    } else {
        qname_prefix(&document.elements[border].qname)
    };
    let style = side.line.as_str();
    let color = side
        .color
        .map(|color| new_border_color(document, border, color, max_output))
        .transpose()?;
    let color_len = color.as_ref().map_or(0, Vec::len);
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(border_edge_name(edge).len()))
        .and_then(|size| size.checked_add(b" style=\"".len()))
        .and_then(|size| size.checked_add(style.len()))
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(color_len))
        .and_then(|size| size.checked_add(3))
        .and_then(|size| size.checked_add(prefix.len()))
        .and_then(|size| size.checked_add(border_edge_name(edge).len()))
        .ok_or_else(|| invalid("styles border side size overflows"))?;
    enforce_limit(
        "styles border side",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles border side", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(border_edge_name(edge));
    output.extend_from_slice(b" style=\"");
    output.extend_from_slice(style.as_bytes());
    output.extend_from_slice(b"\">");
    if let Some(color) = color {
        output.extend_from_slice(&color);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(border_edge_name(edge));
    output.push(b'>');
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn virtual_border_record(document: &Document, max_output: usize) -> Result<Vec<u8>> {
    let prefix = border_prefix(document);
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"border".len()))
        .and_then(|size| size.checked_add(2))
        .ok_or_else(|| invalid("styles default border size overflows"))?;
    enforce_limit(
        "styles default border",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles default border", source))?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"border/>");
    Ok(output)
}

fn new_borders_collection(
    document: &Document,
    records: &[Vec<u8>],
    count: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = border_prefix(document);
    let count = count.to_string();
    let records_len = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles border collection size overflows"))?;
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"borders count=\"".len()))
        .and_then(|size| size.checked_add(count.len()))
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(records_len))
        .and_then(|size| size.checked_add(2 + prefix.len() + b"borders".len() + 1))
        .ok_or_else(|| invalid("styles border collection size overflows"))?;
    enforce_limit(
        "styles border collection",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles border collection", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"borders count=\"");
    output.extend_from_slice(count.as_bytes());
    output.extend_from_slice(b"\">");
    for record in records {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"borders>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn border_side_opaque(input: &[u8], document: &Document, side: usize) -> bool {
    let element = &document.elements[side];
    if element
        .attributes
        .iter()
        .any(|attribute| attribute.name.as_ref() != b"style")
    {
        return true;
    }
    let mut cursor = element.open_end;
    for (child_index, child) in document.elements.iter().enumerate().skip(side + 1) {
        if child.start >= element.end {
            break;
        }
        if child.parent != Some(side) {
            continue;
        }
        if !is_xml_whitespace(&input[cursor..child.start])
            || !child.spreadsheetml
            || child.local.as_ref() != b"color"
        {
            return true;
        }
        if border_color_opaque(input, document, child_index) {
            return true;
        }
        cursor = child.end;
    }
    let end = element.close_start.unwrap_or(element.end);
    !is_xml_whitespace(&input[cursor..end])
}

fn border_color_opaque(input: &[u8], document: &Document, color: usize) -> bool {
    let element = &document.elements[color];
    if element.attributes.iter().any(|attribute| {
        !matches!(
            attribute.name.as_ref(),
            b"rgb" | b"theme" | b"indexed" | b"auto" | b"tint"
        )
    }) {
        return true;
    }
    let end = element.close_start.unwrap_or(element.end);
    !is_xml_whitespace(&input[element.open_end..end])
}

fn rewrite_border_color_opening(
    input: &[u8],
    color: BorderColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let desired = border_color_attributes(color);
    let mut output = copy_bytes(input, "styles border color opening")?;
    for name in [b"rgb".as_slice(), b"theme", b"indexed", b"auto", b"tint"] {
        if !desired
            .iter()
            .any(|(desired_name, _)| *desired_name == name)
        {
            output = remove_attribute(&output, name, max_output)?;
        }
    }
    for (name, value) in desired {
        output = rewrite_attribute(&output, name, &value, true, max_output)?;
    }
    Ok(output)
}

fn expand_empty_border_side(
    element: &Element,
    opening: &[u8],
    color: Option<Vec<u8>>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = opening
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty border side has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && opening[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || opening[slash - 1] != b'/' {
        return Err(invalid(
            "styles empty border side has no self-closing slash",
        ));
    }
    slash -= 1;
    let color_len = color.as_ref().map_or(0, Vec::len);
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(color_len))
        .and_then(|size| size.checked_add(3 + element.qname.len()))
        .ok_or_else(|| invalid("styles expanded border side size overflows"))?;
    enforce_limit(
        "styles expanded border side",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded border side", source))?;
    output.extend_from_slice(&opening[..slash]);
    output.push(b'>');
    if let Some(color) = color {
        output.extend_from_slice(&color);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(&element.qname);
    output.push(b'>');
    Ok(output)
}

fn rewrite_border_side(
    input: &[u8],
    document: &Document,
    side: usize,
    line: Option<Option<Line>>,
    color: Option<Option<BorderColor>>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let element = &document.elements[side];
    let state = border_state_for_side(document, side)?;
    let mut opening = copy_bytes(
        &input[element.start..element.open_end],
        "styles border side opening",
    )?;
    if let Some(requested) = line {
        if requested != state.line {
            opening = match requested {
                Some(value) => {
                    rewrite_attribute(&opening, b"style", value.as_str(), true, max_output)?
                },
                None => remove_attribute(&opening, b"style", max_output)?,
            };
        }
    }
    if element.empty {
        let color_markup = match color {
            Some(Some(value)) if Some(value) != state.color => {
                Some(new_border_color(document, side, value, max_output)?)
            },
            _ => None,
        };
        return if color_markup.is_some() {
            expand_empty_border_side(element, &opening, color_markup, max_output)
        } else {
            Ok(opening)
        };
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(2)
        .map_err(|source| allocation("styles border side edits", source))?;
    edits.push(ByteEdit {
        start: 0,
        end: element.open_end - element.start,
        replacement: opening,
    });
    if let Some(requested) = color {
        if requested != state.color {
            let parent = element
                .parent
                .ok_or_else(|| invalid("styles border side has no parent"))?;
            let edge = border_edge_index(element.local.as_ref())
                .ok_or_else(|| invalid("styles border side is not recognized"))?;
            if let Some(color_element) = document
                .border_properties
                .get(&parent)
                .and_then(|properties| properties.sides[edge].color)
            {
                let color_element_ref = &document.elements[color_element];
                if let Some(value) = requested {
                    let replacement = rewrite_border_color_opening(
                        &input[color_element_ref.start..color_element_ref.open_end],
                        value,
                        max_output,
                    )?;
                    edits.push(ByteEdit {
                        start: color_element_ref.start - element.start,
                        end: color_element_ref.open_end - element.start,
                        replacement,
                    });
                } else {
                    if border_color_opaque(input, document, color_element) {
                        return Err(Error::Unsupported {
                            feature: "border color clear would discard opaque markup",
                        });
                    }
                    edits.push(ByteEdit {
                        start: color_element_ref.start - element.start,
                        end: color_element_ref.end - element.start,
                        replacement: Vec::new(),
                    });
                }
            } else if let Some(value) = requested {
                let color_markup = new_border_color(document, side, value, max_output)?;
                // A modeled color precedes any retained future child content.
                let insert_at = element.open_end;
                edits.push(ByteEdit {
                    start: insert_at - element.start,
                    end: insert_at - element.start,
                    replacement: color_markup,
                });
            }
        }
    }
    let source = &input[element.start..element.end];
    apply_edits(source, edits, max_output)
}

fn border_state_for_side(document: &Document, side: usize) -> Result<BorderSideState> {
    let parent = document.elements[side]
        .parent
        .ok_or_else(|| invalid("styles border side has no parent"))?;
    let edge = border_edge_index(document.elements[side].local.as_ref())
        .ok_or_else(|| invalid("styles border side is not recognized"))?;
    Ok(border_state(document, parent)?.sides[edge])
}

fn changed_border(
    input: &[u8],
    document: &Document,
    border: usize,
    change: &BorderChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    if border == VIRTUAL_NONE_BORDER {
        let prefix = border_prefix(document);
        let side_edges = [
            BorderEdge::Start,
            BorderEdge::End,
            BorderEdge::Left,
            BorderEdge::Right,
            BorderEdge::Top,
            BorderEdge::Bottom,
            BorderEdge::Diagonal,
            BorderEdge::Vertical,
            BorderEdge::Horizontal,
        ];
        let mut sides = Vec::new();
        sides
            .try_reserve_exact(9)
            .map_err(|source| allocation("styles generated border sides", source))?;
        for edge in side_edges {
            let Some(side_change) = change.side_change(edge) else {
                continue;
            };
            let (line, color) = border_side_requested(side_change);
            let line = line.flatten();
            let color = color.flatten();
            if let Some(line) = line {
                sides.push(new_border_side(
                    document,
                    border,
                    edge,
                    Side { line, color },
                    max_output,
                )?);
            } else if color.is_some() {
                return Err(Error::Unsupported {
                    feature: "border color edits require an authored side line",
                });
            }
        }
        let flags_len = [
            (b"diagonalUp".as_slice(), change.diagonal_up_change()),
            (b"diagonalDown", change.diagonal_down_change()),
            (b"outline", change.outline_change()),
        ]
        .into_iter()
        .filter_map(|(name, value)| value.and_then(|value| value.map(|_| name.len() + 5)))
        .try_fold(0usize, |size, length| size.checked_add(length))
        .ok_or_else(|| invalid("styles generated border size overflows"))?;
        let side_len = sides
            .iter()
            .try_fold(0usize, |size, side| size.checked_add(side.len()))
            .ok_or_else(|| invalid("styles generated border size overflows"))?;
        let has_children = !sides.is_empty();
        let output_len = 1usize
            .checked_add(prefix.len())
            .and_then(|size| size.checked_add(b"border".len()))
            .and_then(|size| size.checked_add(flags_len))
            .and_then(|size| size.checked_add(if has_children { 1 } else { 2 }))
            .and_then(|size| size.checked_add(side_len))
            .and_then(|size| {
                has_children
                    .then_some(3 + prefix.len() + b"border".len())
                    .map_or(Some(size), |tail| size.checked_add(tail))
            })
            .ok_or_else(|| invalid("styles generated border size overflows"))?;
        enforce_limit(
            "styles generated border",
            output_len,
            max_output,
            litchi_core::Resource::Memory,
        )?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| allocation("styles generated border", source))?;
        output.extend_from_slice(b"<");
        output.extend_from_slice(prefix);
        output.extend_from_slice(b"border");
        for (name, value) in [
            (b"diagonalUp".as_slice(), change.diagonal_up_change()),
            (b"diagonalDown", change.diagonal_down_change()),
            (b"outline", change.outline_change()),
        ] {
            if let Some(Some(value)) = value {
                output.push(b' ');
                output.extend_from_slice(name);
                output.extend_from_slice(b"=\"");
                output.push(if value { b'1' } else { b'0' });
                output.push(b'"');
            }
        }
        if has_children {
            output.push(b'>');
            for side in sides {
                output.extend_from_slice(&side);
            }
            output.extend_from_slice(b"</");
            output.extend_from_slice(prefix);
            output.extend_from_slice(b"border>");
        } else {
            output.extend_from_slice(b"/>");
        }
        enforce_limit(
            "styles generated border",
            output.len(),
            max_output,
            litchi_core::Resource::Memory,
        )?;
        return Ok(output);
    }
    let element = &document.elements[border];
    let state = border_state(document, border)?;
    let mut opening = copy_bytes(
        &input[element.start..element.open_end],
        "styles border opening",
    )?;
    for (name, requested, current) in [
        (
            b"diagonalUp".as_slice(),
            change.diagonal_up_change(),
            state.diagonal_up,
        ),
        (
            b"diagonalDown",
            change.diagonal_down_change(),
            state.diagonal_down,
        ),
        (b"outline", change.outline_change(), state.outline),
    ] {
        if let Some(requested) = requested {
            if requested != current {
                opening = match requested {
                    Some(value) => rewrite_attribute(
                        &opening,
                        name,
                        if value { "1" } else { "0" },
                        true,
                        max_output,
                    )?,
                    None => remove_attribute(&opening, name, max_output)?,
                };
            }
        }
    }
    let mut edits = Vec::new();
    let mut empty_children = Vec::new();
    for (edge, index) in [
        (BorderEdge::Start, 0usize),
        (BorderEdge::End, 1),
        (BorderEdge::Left, 2),
        (BorderEdge::Right, 3),
        (BorderEdge::Top, 4),
        (BorderEdge::Bottom, 5),
        (BorderEdge::Diagonal, 6),
        (BorderEdge::Vertical, 7),
        (BorderEdge::Horizontal, 8),
    ] {
        let Some(side_change) = change.side_change(edge) else {
            continue;
        };
        let properties = document
            .border_properties
            .get(&border)
            .ok_or_else(|| invalid("styles border property index is out of range"))?
            .sides[index];
        let state = state.sides[index];
        let (line_change, color_change) = border_side_requested(side_change);
        let wants_present = match side_change.whole_change() {
            Some(Some(_)) => true,
            Some(None) => {
                line_change.is_some_and(|value| value.is_some())
                    || color_change.is_some_and(|value| value.is_some())
            },
            None => {
                state.present
                    || line_change.is_some_and(|value| value.is_some())
                    || color_change.is_some_and(|value| value.is_some())
            },
        };
        if !wants_present {
            if let Some(side) = properties.side {
                if border_side_opaque(input, document, side) {
                    return Err(Error::Unsupported {
                        feature: "border side clear would discard opaque markup",
                    });
                }
                edits.push(ByteEdit {
                    start: document.elements[side].start - element.start,
                    end: document.elements[side].end - element.start,
                    replacement: Vec::new(),
                });
            }
            continue;
        }
        if let Some(side) = properties.side {
            let replacement =
                rewrite_border_side(input, document, side, line_change, color_change, max_output)?;
            if replacement != input[document.elements[side].start..document.elements[side].end] {
                edits.push(ByteEdit {
                    start: document.elements[side].start - element.start,
                    end: document.elements[side].end - element.start,
                    replacement,
                });
            }
        } else {
            let line = line_change.and_then(|value| value);
            let color = color_change.and_then(|value| value);
            if let Some(line) = line {
                let markup =
                    new_border_side(document, border, edge, Side { line, color }, max_output)?;
                if element.empty {
                    empty_children.push(markup);
                } else {
                    let insert_at = properties
                        .insert_at
                        .ok_or_else(|| invalid("styles border has no side insertion point"))?;
                    edits.push(ByteEdit {
                        start: insert_at - element.start,
                        end: insert_at - element.start,
                        replacement: markup,
                    });
                }
            } else if color.is_some() {
                return Err(Error::Unsupported {
                    feature: "border color edits require an authored side line",
                });
            } else if state.present {
                return Err(invalid("styles border side state is inconsistent"));
            }
        }
    }
    let source = &input[element.start..element.end];
    if element.empty {
        if empty_children.is_empty() {
            return Ok(opening);
        }
        return expand_empty_border(element, &opening, &empty_children, max_output);
    }
    edits.insert(
        0,
        ByteEdit {
            start: 0,
            end: element.open_end - element.start,
            replacement: opening,
        },
    );
    apply_edits(source, edits, max_output)
}

fn expand_empty_border(
    element: &Element,
    opening: &[u8],
    children: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = opening
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty border has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && opening[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || opening[slash - 1] != b'/' {
        return Err(invalid("styles empty border has no self-closing slash"));
    }
    slash -= 1;
    let children_len = children
        .iter()
        .try_fold(0usize, |size, child| size.checked_add(child.len()))
        .ok_or_else(|| invalid("styles expanded border size overflows"))?;
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(children_len))
        .and_then(|size| size.checked_add(3 + element.qname.len()))
        .ok_or_else(|| invalid("styles expanded border size overflows"))?;
    enforce_limit(
        "styles expanded border",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded border", source))?;
    output.extend_from_slice(&opening[..slash]);
    output.push(b'>');
    for child in children {
        output.extend_from_slice(child);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(&element.qname);
    output.push(b'>');
    Ok(output)
}

fn expand_empty_font(
    source: &[u8],
    element: &Element,
    insertions: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = source
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty font has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && source[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || source[slash - 1] != b'/' {
        return Err(invalid("styles empty font has no self-closing slash"));
    }
    slash -= 1;
    let insertion_len = insertions
        .iter()
        .try_fold(0usize, |size, insertion| size.checked_add(insertion.len()))
        .ok_or_else(|| invalid("styles expanded font size overflows"))?;
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(insertion_len))
        .and_then(|size| size.checked_add(element.qname.len()))
        .and_then(|size| size.checked_add(3))
        .ok_or_else(|| invalid("styles expanded font size overflows"))?;
    enforce_limit(
        "styles candidate font bytes",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded font", source))?;
    result.extend_from_slice(&source[..slash]);
    result.push(b'>');
    for insertion in insertions {
        result.extend_from_slice(insertion);
    }
    result.extend_from_slice(b"</");
    result.extend_from_slice(&element.qname);
    result.push(b'>');
    Ok(result)
}

fn changed_fill(
    input: &[u8],
    document: &Document,
    fill: usize,
    change: &FillChange,
    max_output: usize,
    max_staged: usize,
) -> Result<Vec<u8>> {
    if change.replacement().is_some() {
        return changed_fill_replacement(input, document, fill, change, max_output, max_staged);
    }
    if document.virtual_fills {
        return changed_virtual_fill(document, fill, change, max_output);
    }
    let element = &document.elements[fill];
    let source = &input[element.start..element.end];
    let properties = document
        .fill_properties
        .get(fill)
        .ok_or_else(|| invalid("styles fill index is out of range"))?;
    if properties.gradient_fill.is_some() {
        return Err(Error::Unsupported {
            feature: "source gradient fills are not writable by pattern-fill edits",
        });
    }
    let Some(pattern) = properties.pattern_fill else {
        let pattern_markup = new_pattern_fill(document, fill, change, max_output)?;
        if element.empty {
            return expand_empty_fill(source, element, &pattern_markup, max_output);
        }
        let insert_at = properties
            .pattern_insert_at
            .ok_or_else(|| invalid("styles fill has no patternFill insertion point"))?;
        let offset = insert_at - element.start;
        let mut edits = Vec::new();
        edits
            .try_reserve_exact(1)
            .map_err(|source| allocation("styles fill insertion", source))?;
        edits.push(ByteEdit {
            start: offset,
            end: offset,
            replacement: pattern_markup,
        });
        return apply_edits(source, edits, max_output);
    };

    let pattern_element = &document.elements[pattern];
    let pattern_source = &input[pattern_element.start..pattern_element.end];
    let opening = rewrite_pattern_opening(
        &input[pattern_element.start..pattern_element.open_end],
        change.pattern_change(),
        max_output,
    )?;
    let state = fill_state(document, fill)?;
    let mut color_edits = Vec::new();
    color_edits
        .try_reserve_exact(3)
        .map_err(|source| allocation("styles fill property edits", source))?;
    if pattern_element.empty {
        let mut children = Vec::new();
        children
            .try_reserve_exact(2)
            .map_err(|source| allocation("styles empty pattern children", source))?;
        if let Some(Some(value)) = change.foreground_change() {
            children.push(new_fill_color_markup(
                fill_pattern_prefix(document, pattern),
                b"fgColor",
                value,
                max_output,
            )?);
        }
        if let Some(Some(value)) = change.background_change() {
            children.push(new_fill_color_markup(
                fill_pattern_prefix(document, pattern),
                b"bgColor",
                value,
                max_output,
            )?);
        }
        let replacement = if children.is_empty() {
            opening
        } else {
            expand_empty_pattern(&opening, &children, max_output)?
        };
        return replace_pattern_in_fill(
            source,
            element.start,
            pattern_element,
            replacement,
            max_output,
        );
    }

    apply_fill_color_change(
        input,
        document,
        fill,
        true,
        change.foreground_change(),
        state.foreground,
        &mut color_edits,
        max_output,
    )?;
    apply_fill_color_change(
        input,
        document,
        fill,
        false,
        change.background_change(),
        state.background,
        &mut color_edits,
        max_output,
    )?;

    let mut edits = Vec::new();
    edits
        .try_reserve_exact(color_edits.len().saturating_add(1))
        .map_err(|source| allocation("styles pattern edits", source))?;
    edits.push(ByteEdit {
        start: 0,
        end: pattern_element.open_end - pattern_element.start,
        replacement: opening,
    });
    for edit in color_edits {
        edits.push(ByteEdit {
            start: edit.start - pattern_element.start,
            end: edit.end - pattern_element.start,
            replacement: edit.replacement,
        });
    }
    let replacement = apply_edits(pattern_source, edits, max_output)?;
    replace_pattern_in_fill(
        source,
        element.start,
        pattern_element,
        replacement,
        max_output,
    )
}

fn changed_fill_replacement(
    input: &[u8],
    document: &Document,
    fill: usize,
    change: &FillChange,
    max_output: usize,
    max_staged: usize,
) -> Result<Vec<u8>> {
    let replacement = change
        .replacement()
        .ok_or_else(|| invalid("styles fill replacement disappeared during rewrite"))?;
    let state = fill_state(document, fill)?;
    let matches = replacement_matches_state(&state, replacement, change)?;
    let existing_child = if document.virtual_fills {
        None
    } else {
        document
            .fill_properties
            .get(fill)
            .and_then(|properties| properties.pattern_fill.or(properties.gradient_fill))
    };
    if matches && !document.virtual_fills {
        let element = document
            .elements
            .get(fill)
            .ok_or_else(|| invalid("styles fill index is out of range"))?;
        return copy_bytes(
            &input[element.start..element.end],
            "styles unchanged fill replacement",
        );
    }

    let (element, properties) = if document.virtual_fills {
        (None, None)
    } else {
        let element = document
            .elements
            .get(fill)
            .ok_or_else(|| invalid("styles fill index is out of range"))?;
        let properties = document
            .fill_properties
            .get(fill)
            .ok_or_else(|| invalid("styles fill index is out of range"))?;
        let opaque = properties
            .gradient
            .as_ref()
            .is_some_and(|gradient| gradient.opaque)
            || (properties.gradient_fill.is_none() && properties.pattern_opaque);
        if opaque {
            return Err(Error::Unsupported {
                feature: "whole fill replacement would discard opaque child markup",
            });
        }
        (Some(element), Some(properties))
    };

    let replacement_markup = match replacement {
        Fill::Pattern(pattern_fill) => {
            let pattern = change.pattern_change().unwrap_or(pattern_fill.pattern());
            let foreground = change
                .foreground_change()
                .unwrap_or(pattern_fill.foreground());
            let background = change
                .background_change()
                .unwrap_or(pattern_fill.background());
            let prefix = existing_child
                .and_then(|child| document.elements.get(child))
                .map(|element| qname_prefix(&element.qname))
                .unwrap_or_else(|| fill_qname_prefix(document, fill));
            new_pattern_fill_values_with_prefix(
                prefix, pattern, foreground, background, max_output,
            )?
        },
        Fill::Gradient(gradient) => new_gradient_fill(
            document,
            fill,
            existing_child,
            gradient,
            crate::style::cell_format::MAX_RECORDS,
            max_output,
            max_staged,
        )?,
    };
    if document.virtual_fills {
        return new_fill_record(document, &replacement_markup, max_output);
    }

    let element = element.ok_or_else(|| invalid("styles fill index is out of range"))?;
    let properties = properties.ok_or_else(|| invalid("styles fill index is out of range"))?;
    let child = properties.pattern_fill.or(properties.gradient_fill);
    if let Some(child) = child {
        return replace_pattern_in_fill(
            &input[element.start..element.end],
            element.start,
            &document.elements[child],
            replacement_markup,
            max_output,
        );
    }
    let insert_at = properties
        .pattern_insert_at
        .or(element.close_start)
        .ok_or_else(|| invalid("styles fill has no child insertion point"))?;
    let offset = insert_at
        .checked_sub(element.start)
        .ok_or_else(|| invalid("styles fill child insertion precedes fill"))?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(1)
        .map_err(|source| allocation("styles fill replacement insertion", source))?;
    edits.push(ByteEdit {
        start: offset,
        end: offset,
        replacement: replacement_markup,
    });
    apply_edits(&input[element.start..element.end], edits, max_output)
}

fn replace_pattern_in_fill(
    source: &[u8],
    fill_start: usize,
    pattern: &Element,
    replacement: Vec<u8>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let start = pattern
        .start
        .checked_sub(fill_start)
        .ok_or_else(|| invalid("styles patternFill precedes its fill record"))?;
    let end = pattern
        .end
        .checked_sub(fill_start)
        .ok_or_else(|| invalid("styles patternFill precedes its fill record"))?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(1)
        .map_err(|source| allocation("styles fill pattern replacement", source))?;
    edits.push(ByteEdit {
        start,
        end,
        replacement,
    });
    apply_edits(source, edits, max_output)
}

fn rewrite_pattern_opening(
    input: &[u8],
    change: Option<Option<FillPattern>>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let Some(change) = change else {
        return copy_bytes(input, "styles pattern opening");
    };
    match change {
        Some(value) => rewrite_attribute(input, b"patternType", value.as_str(), true, max_output),
        None => remove_attribute(input, b"patternType", max_output),
    }
}

fn apply_fill_color_change(
    input: &[u8],
    document: &Document,
    fill: usize,
    foreground: bool,
    change: Option<Option<FillColor>>,
    current: Option<FillColor>,
    edits: &mut Vec<ByteEdit>,
    max_output: usize,
) -> Result<()> {
    let Some(change) = change else {
        return Ok(());
    };
    let properties = document
        .fill_properties
        .get(fill)
        .ok_or_else(|| invalid("styles fill index is out of range"))?;
    let child = if foreground {
        properties.foreground
    } else {
        properties.background
    };
    if let Some(child) = child {
        let element = &document.elements[child];
        match change {
            None => edits.push(ByteEdit {
                start: element.start,
                end: element.end,
                replacement: Vec::new(),
            }),
            Some(value) => {
                if current == Some(value) {
                    return Ok(());
                }
                edits.push(ByteEdit {
                    start: element.start,
                    end: element.open_end,
                    replacement: rewrite_fill_color_opening(
                        &input[element.start..element.open_end],
                        value,
                        max_output,
                    )?,
                });
            },
        }
        return Ok(());
    }
    let Some(value) = change else {
        return Ok(());
    };
    let insert_at = if foreground {
        properties.foreground_insert_at
    } else {
        properties.background_insert_at
    }
    .ok_or_else(|| invalid("styles patternFill has no color insertion point"))?;
    edits.push(ByteEdit {
        start: insert_at,
        end: insert_at,
        replacement: new_fill_color(document, fill, foreground, value, max_output)?,
    });
    Ok(())
}

fn fill_pattern_prefix(document: &Document, pattern: usize) -> &[u8] {
    let qname = &document.elements[pattern].qname;
    let prefix_len = qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    &qname[..prefix_len]
}

fn fill_qname_prefix(document: &Document, fill: usize) -> &[u8] {
    let qname = document
        .elements
        .get(fill)
        .or_else(|| {
            document
                .fills
                .and_then(|fills| document.elements.get(fills))
        })
        .or_else(|| {
            document
                .elements
                .iter()
                .find(|element| element.parent.is_none())
        })
        .map(|element| element.qname.as_ref())
        .unwrap_or_default();
    let prefix_len = qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    &qname[..prefix_len]
}

fn new_fill_color(
    document: &Document,
    fill: usize,
    foreground: bool,
    value: FillColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let pattern = direct_fill_pattern(document, fill)?
        .ok_or_else(|| invalid("styles fill has no patternFill child"))?;
    new_fill_color_markup(
        fill_pattern_prefix(document, pattern),
        if foreground { b"fgColor" } else { b"bgColor" },
        value,
        max_output,
    )
}

fn new_pattern_fill(
    document: &Document,
    fill: usize,
    change: &FillChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let pattern = change.pattern_change().and_then(|value| value);
    let foreground = change.foreground_change().and_then(|value| value);
    let background = change.background_change().and_then(|value| value);
    new_pattern_fill_values(document, fill, pattern, foreground, background, max_output)
}

fn new_pattern_fill_values(
    document: &Document,
    fill: usize,
    pattern: Option<FillPattern>,
    foreground: Option<FillColor>,
    background: Option<FillColor>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = fill_qname_prefix(document, fill);
    new_pattern_fill_values_with_prefix(prefix, pattern, foreground, background, max_output)
}

fn new_pattern_fill_values_with_prefix(
    prefix: &[u8],
    pattern: Option<FillPattern>,
    foreground: Option<FillColor>,
    background: Option<FillColor>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let pattern_token = pattern.map(|value| value.as_str());
    let mut children = Vec::new();
    children
        .try_reserve_exact(2)
        .map_err(|source| allocation("styles new pattern colors", source))?;
    if let Some(value) = foreground {
        children.push(new_fill_color_markup(
            prefix, b"fgColor", value, max_output,
        )?);
    }
    if let Some(value) = background {
        children.push(new_fill_color_markup(
            prefix, b"bgColor", value, max_output,
        )?);
    }
    let pattern_name_len = prefix
        .len()
        .checked_add(b"patternFill".len())
        .ok_or_else(|| invalid("styles patternFill QName size overflows"))?;
    let mut output_len = 1usize
        .checked_add(pattern_name_len)
        .and_then(|size| {
            size.checked_add(pattern_token.map_or(0, |token| {
                b" patternType=\""
                    .len()
                    .saturating_add(token.len())
                    .saturating_add(1)
            }))
        })
        .ok_or_else(|| invalid("styles patternFill size overflows"))?;
    if children.is_empty() {
        output_len = output_len
            .checked_add(2)
            .ok_or_else(|| invalid("styles patternFill size overflows"))?;
    } else {
        output_len = output_len
            .checked_add(1)
            .and_then(|size| {
                children
                    .iter()
                    .try_fold(size, |size, child| size.checked_add(child.len()))
            })
            .and_then(|size| size.checked_add(3usize.saturating_add(pattern_name_len)))
            .ok_or_else(|| invalid("styles patternFill size overflows"))?;
    }
    enforce_limit(
        "styles patternFill property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles patternFill property", source))?;
    output.extend_from_slice(b"<");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"patternFill");
    if let Some(token) = pattern_token {
        output.extend_from_slice(b" patternType=\"");
        output.extend_from_slice(token.as_bytes());
        output.push(b'"');
    }
    if children.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        for child in children {
            output.extend_from_slice(&child);
        }
        output.extend_from_slice(b"</");
        output.extend_from_slice(prefix);
        output.extend_from_slice(b"patternFill>");
    }
    if output.len() != output_len {
        return Err(invalid("styles patternFill size accounting mismatch"));
    }
    Ok(output)
}

fn new_gradient_fill(
    document: &Document,
    fill: usize,
    child: Option<usize>,
    gradient: &crate::style::cell_format::GradientFill,
    max_gradient_stops: usize,
    max_output: usize,
    max_staged: usize,
) -> Result<Vec<u8>> {
    gradient.validate()?;
    let stops = gradient.stops();
    enforce_limit(
        "styles gradient stops",
        stops.len(),
        max_gradient_stops,
        litchi_core::Resource::Objects,
    )?;
    let prefix = child
        .and_then(|child| document.elements.get(child))
        .map(|element| qname_prefix(&element.qname))
        .unwrap_or_else(|| fill_qname_prefix(document, fill));
    let generated_limit = max_output.min(max_staged);
    let mut attributes: Vec<(&[u8], String)> = Vec::new();
    attributes
        .try_reserve_exact(6)
        .map_err(|source| allocation("styles gradient attributes", source))?;
    attributes.push((b"type", gradient.kind().as_str().to_owned()));
    if gradient.degree() != 0.0 {
        attributes.push((b"degree", gradient.degree().to_string()));
    }
    if gradient.left() != 0.0 {
        attributes.push((b"left", gradient.left().to_string()));
    }
    if gradient.right() != 0.0 {
        attributes.push((b"right", gradient.right().to_string()));
    }
    if gradient.top() != 0.0 {
        attributes.push((b"top", gradient.top().to_string()));
    }
    if gradient.bottom() != 0.0 {
        attributes.push((b"bottom", gradient.bottom().to_string()));
    }
    let qname_len = prefix
        .len()
        .checked_add(b"gradientFill".len())
        .ok_or_else(|| invalid("styles gradientFill QName size overflows"))?;
    let mut outer_len = 1usize
        .checked_add(qname_len)
        .ok_or_else(|| invalid("styles gradientFill size overflows"))?;
    for (name, value) in &attributes {
        let value_len = escaped_attribute_len(value)?;
        outer_len = outer_len
            .checked_add(name.len())
            .and_then(|size| size.checked_add(value_len))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| invalid("styles gradientFill size overflows"))?;
    }
    outer_len = outer_len
        .checked_add(1)
        .and_then(|size| size.checked_add(3usize.saturating_add(qname_len)))
        .ok_or_else(|| invalid("styles gradientFill size overflows"))?;
    let mut projected_stop_bytes = 0usize;
    for stop in stops {
        let projected = gradient_stop_markup_len(prefix, stop)?;
        projected_stop_bytes = projected_stop_bytes
            .checked_add(projected)
            .ok_or_else(|| invalid("styles generated gradient bytes overflow"))?;
        enforce_limit(
            "styles generated gradient bytes",
            projected_stop_bytes,
            generated_limit,
            litchi_core::Resource::Memory,
        )?;
    }
    let output_len = outer_len
        .checked_add(projected_stop_bytes)
        .ok_or_else(|| invalid("styles gradientFill size overflows"))?;
    enforce_limit(
        "styles generated gradient bytes",
        output_len,
        generated_limit,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles gradientFill", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"gradientFill");
    for (name, value) in &attributes {
        append_fill_color_attribute(&mut output, name, value)?;
    }
    output.push(b'>');
    for stop in stops {
        let remaining = generated_limit
            .checked_sub(output.len())
            .ok_or_else(|| invalid("styles generated gradient bytes underflow"))?;
        let markup = new_gradient_stop_markup(prefix, stop, remaining)?;
        output.extend_from_slice(&markup);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"gradientFill>");
    if output.len() != output_len {
        return Err(invalid("styles gradientFill size accounting mismatch"));
    }
    Ok(output)
}

fn new_gradient_stop_markup(
    prefix: &[u8],
    stop: &crate::style::cell_format::GradientStop,
    max_output: usize,
) -> Result<Vec<u8>> {
    let position = stop.position().to_string();
    let color = new_fill_color_markup(prefix, b"color", stop.color(), max_output)?;
    let qname_len = prefix
        .len()
        .checked_add(b"stop".len())
        .ok_or_else(|| invalid("styles gradient stop QName size overflows"))?;
    let position_len = escaped_attribute_len(&position)?;
    let output_len = 1usize
        .checked_add(qname_len)
        .and_then(|size| size.checked_add(b" position=\"".len()))
        .and_then(|size| size.checked_add(position_len))
        .and_then(|size| size.checked_add(b"\">".len()))
        .and_then(|size| size.checked_add(color.len()))
        .and_then(|size| size.checked_add(3usize.saturating_add(qname_len)))
        .ok_or_else(|| invalid("styles gradient stop size overflows"))?;
    enforce_limit(
        "styles generated gradient bytes",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles gradient stop", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"stop position=\"");
    append_escaped_attribute(&mut output, &position, position_len)?;
    output.extend_from_slice(b"\">");
    output.extend_from_slice(&color);
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"stop>");
    if output.len() != output_len {
        return Err(invalid("styles gradient stop size accounting mismatch"));
    }
    Ok(output)
}

fn gradient_stop_markup_len(
    prefix: &[u8],
    stop: &crate::style::cell_format::GradientStop,
) -> Result<usize> {
    let position = stop.position().to_string();
    let qname_len = prefix
        .len()
        .checked_add(b"stop".len())
        .ok_or_else(|| invalid("styles gradient stop QName size overflows"))?;
    let color_len = fill_color_markup_len(prefix, b"color", stop.color())?;
    let position_len = escaped_attribute_len(&position)?;
    1usize
        .checked_add(qname_len)
        .and_then(|size| size.checked_add(b" position=\"".len()))
        .and_then(|size| size.checked_add(position_len))
        .and_then(|size| size.checked_add(b"\">".len()))
        .and_then(|size| size.checked_add(color_len))
        .and_then(|size| size.checked_add(3usize.saturating_add(qname_len)))
        .ok_or_else(|| invalid("styles gradient stop size overflows"))
}

fn new_fill_record(document: &Document, pattern: &[u8], max_output: usize) -> Result<Vec<u8>> {
    let prefix = fill_qname_prefix(document, VIRTUAL_NONE_FILL);
    let qname_len = prefix
        .len()
        .checked_add(b"fill".len())
        .ok_or_else(|| invalid("styles fill QName size overflows"))?;
    let output_len = 1usize
        .checked_add(qname_len)
        .and_then(|size| size.checked_add(1))
        .and_then(|size| size.checked_add(pattern.len()))
        .and_then(|size| size.checked_add(3usize.saturating_add(qname_len)))
        .ok_or_else(|| invalid("styles fill size overflows"))?;
    enforce_limit(
        "styles fill record",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles fill record", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"fill>");
    output.extend_from_slice(pattern);
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"fill>");
    if output.len() != output_len {
        return Err(invalid("styles fill size accounting mismatch"));
    }
    Ok(output)
}

fn virtual_fill_record(
    document: &Document,
    pattern: FillPattern,
    max_output: usize,
) -> Result<Vec<u8>> {
    let pattern = new_pattern_fill_values(
        document,
        VIRTUAL_NONE_FILL,
        Some(pattern),
        None,
        None,
        max_output,
    )?;
    new_fill_record(document, &pattern, max_output)
}

fn changed_virtual_fill(
    document: &Document,
    fill: usize,
    change: &FillChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let state = fill_state(document, fill)?;
    let pattern = change.pattern_change().unwrap_or(state.pattern);
    let foreground = change.foreground_change().unwrap_or(state.foreground);
    let background = change.background_change().unwrap_or(state.background);
    let pattern =
        new_pattern_fill_values(document, fill, pattern, foreground, background, max_output)?;
    new_fill_record(document, &pattern, max_output)
}

fn fill_color_base_attribute(value: FillColor) -> Option<(&'static [u8], String)> {
    match value {
        FillColor::Default { .. } => None,
        FillColor::Rgb { value, .. } => Some((b"rgb", value.to_string())),
        FillColor::Theme { index, .. } => Some((b"theme", index.to_string())),
        FillColor::Indexed { index, .. } => Some((b"indexed", index.to_string())),
        FillColor::Auto { enabled, .. } => Some((b"auto", if enabled { "1" } else { "0" }.into())),
    }
}

fn fill_color_tint(value: FillColor) -> Option<String> {
    value.tint().map(|value| value.get().to_string())
}

fn new_fill_color_markup(
    prefix: &[u8],
    name: &[u8],
    value: FillColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let base = fill_color_base_attribute(value);
    let tint = fill_color_tint(value);
    let output_len = fill_color_markup_len(prefix, name, value)?;
    enforce_limit(
        "styles fill color property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles fill color property", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(name);
    if let Some((attribute, value)) = base {
        append_fill_color_attribute(&mut output, attribute, &value)?;
    }
    if let Some(value) = tint.as_deref() {
        append_fill_color_attribute(&mut output, b"tint", value)?;
    }
    output.extend_from_slice(b"/>");
    if output.len() != output_len {
        return Err(invalid("styles fill color size accounting mismatch"));
    }
    Ok(output)
}

fn fill_color_markup_len(prefix: &[u8], name: &[u8], value: FillColor) -> Result<usize> {
    let base = fill_color_base_attribute(value);
    let tint = fill_color_tint(value);
    let qname_len = prefix
        .len()
        .checked_add(name.len())
        .ok_or_else(|| invalid("styles fill color QName size overflows"))?;
    let mut output_len = 1usize
        .checked_add(qname_len)
        .and_then(|size| size.checked_add(2))
        .ok_or_else(|| invalid("styles fill color size overflows"))?;
    for (attribute, value) in base.iter() {
        let value_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(1)
            .and_then(|size| size.checked_add(attribute.len()))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(value_len))
            .and_then(|size| size.checked_add(1))
            .ok_or_else(|| invalid("styles fill color size overflows"))?;
    }
    if let Some(value) = tint.as_deref() {
        let value_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(1)
            .and_then(|size| size.checked_add(b"tint".len()))
            .and_then(|size| size.checked_add(2))
            .and_then(|size| size.checked_add(value_len))
            .and_then(|size| size.checked_add(1))
            .ok_or_else(|| invalid("styles fill color size overflows"))?;
    }
    Ok(output_len)
}

fn append_fill_color_attribute(output: &mut Vec<u8>, name: &[u8], value: &str) -> Result<()> {
    let escaped_len = escaped_attribute_len(value)?;
    output.push(b' ');
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(output, value, escaped_len)?;
    output.push(b'"');
    Ok(())
}

fn rewrite_fill_color_opening(
    input: &[u8],
    value: FillColor,
    max_output: usize,
) -> Result<Vec<u8>> {
    let attributes = raw_attributes(input)?;
    enforce_limit(
        "styles fill color rewrite",
        input.len(),
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let desired_base = fill_color_base_attribute(value);
    let desired_base_name = desired_base.as_ref().map(|(name, _)| *name);
    let has_desired_base = desired_base_name.is_some_and(|name| {
        attributes
            .iter()
            .any(|attribute| attribute.name.as_slice() == name)
    });
    let has_tint = attributes
        .iter()
        .any(|attribute| attribute.name.as_slice() == b"tint");
    let mut output = Vec::new();
    output
        .try_reserve_exact(input.len())
        .map_err(|source| allocation("styles fill color rewrite", source))?;
    let mut cursor = 0;
    for attribute in &attributes {
        let remove = if is_color_base_attribute(&attribute.name) {
            desired_base_name.is_none_or(|name| name != attribute.name.as_slice())
        } else {
            attribute.name.as_slice() == b"tint" && value.tint().is_none()
        };
        if remove {
            let mut start = attribute.start;
            while start > cursor && input[start - 1].is_ascii_whitespace() {
                start -= 1;
            }
            output.extend_from_slice(&input[cursor..start]);
            cursor = attribute.end;
        }
    }
    output.extend_from_slice(&input[cursor..]);

    if let Some((name, replacement)) = desired_base {
        if has_desired_base {
            output = rewrite_attribute(&output, name, &replacement, false, max_output)?;
        } else if has_tint {
            output = insert_attribute_before(&output, b"tint", name, &replacement, max_output)?;
        } else {
            output = rewrite_attribute(&output, name, &replacement, true, max_output)?;
        }
    }
    if let Some(tint) = fill_color_tint(value) {
        output = rewrite_attribute(&output, b"tint", &tint, true, max_output)?;
    }
    Ok(output)
}

fn expand_empty_fill(
    source: &[u8],
    element: &Element,
    pattern: &[u8],
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = source
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty fill has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && source[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || source[slash - 1] != b'/' {
        return Err(invalid("styles empty fill has no self-closing slash"));
    }
    slash -= 1;
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(pattern.len()))
        .and_then(|size| size.checked_add(3usize.saturating_add(element.qname.len())))
        .ok_or_else(|| invalid("styles expanded fill size overflows"))?;
    enforce_limit(
        "styles candidate fill bytes",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded fill", source))?;
    output.extend_from_slice(&source[..slash]);
    output.push(b'>');
    output.extend_from_slice(pattern);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&element.qname);
    output.push(b'>');
    Ok(output)
}

fn expand_empty_pattern(
    opening: &[u8],
    children: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = opening
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty patternFill has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && opening[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || opening[slash - 1] != b'/' {
        return Err(invalid(
            "styles empty patternFill has no self-closing slash",
        ));
    }
    slash -= 1;
    let children_len = children
        .iter()
        .try_fold(0usize, |size, child| size.checked_add(child.len()))
        .ok_or_else(|| invalid("styles expanded patternFill size overflows"))?;
    let qname_start = opening
        .iter()
        .position(|byte| *byte == b'<')
        .and_then(|start| {
            let start = start.checked_add(1)?;
            let end = opening[start..]
                .iter()
                .position(|byte| matches!(*byte, b' ' | b'/' | b'>'))?;
            Some((start, start + end))
        })
        .ok_or_else(|| invalid("styles empty patternFill has no QName"))?;
    let qname = &opening[qname_start.0..qname_start.1];
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(children_len))
        .and_then(|size| size.checked_add(3usize.saturating_add(qname.len())))
        .ok_or_else(|| invalid("styles expanded patternFill size overflows"))?;
    enforce_limit(
        "styles candidate patternFill bytes",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded patternFill", source))?;
    output.extend_from_slice(&opening[..slash]);
    output.push(b'>');
    for child in children {
        output.extend_from_slice(child);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(qname);
    output.push(b'>');
    Ok(output)
}

#[derive(Debug, Clone, Copy)]
struct XfChange<'a> {
    font: Option<u32>,
    fill: Option<u32>,
    border: Option<u32>,
    alignment: Option<&'a AlignmentChange>,
    protection: Option<&'a ProtectionChange>,
    quote_prefix: Option<Option<bool>>,
    number_format: Option<Option<u32>>,
}

fn changed_xf(
    input: &[u8],
    document: &Document,
    xf: usize,
    change: XfChange<'_>,
    max_output: usize,
) -> Result<Vec<u8>> {
    let XfChange {
        font,
        fill,
        border,
        alignment,
        protection,
        quote_prefix,
        number_format,
    } = change;
    let element = &document.elements[xf];
    let source = &input[element.start..element.end];
    enforce_limit(
        "styles cell-XF opening",
        element.open_end - element.start,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut opening = copy_bytes(
        &input[element.start..element.open_end],
        "styles cell-XF opening",
    )?;
    if let Some(font) = font {
        opening = rewrite_attribute(&opening, b"fontId", &font.to_string(), true, max_output)?;
        opening = rewrite_attribute(&opening, b"applyFont", "1", true, max_output)?;
    }
    if let Some(fill) = fill {
        opening = rewrite_attribute(&opening, b"fillId", &fill.to_string(), true, max_output)?;
        opening = rewrite_attribute(&opening, b"applyFill", "1", true, max_output)?;
    }
    if let Some(border) = border {
        opening = rewrite_attribute(&opening, b"borderId", &border.to_string(), true, max_output)?;
        opening = rewrite_attribute(&opening, b"applyBorder", "1", true, max_output)?;
    }
    if alignment.is_some() {
        opening = rewrite_attribute(&opening, b"applyAlignment", "1", true, max_output)?;
    }
    if protection.is_some() {
        opening = rewrite_attribute(&opening, b"applyProtection", "1", true, max_output)?;
    }
    if let Some(value) = quote_prefix {
        opening = match value {
            Some(value) => rewrite_attribute(
                &opening,
                b"quotePrefix",
                if value { "1" } else { "0" },
                true,
                max_output,
            )?,
            None => remove_attribute(&opening, b"quotePrefix", max_output)?,
        };
    }
    if let Some(value) = number_format {
        opening = match value {
            Some(value) => {
                rewrite_attribute(&opening, b"numFmtId", &value.to_string(), true, max_output)?
            },
            None => remove_attribute(&opening, b"numFmtId", max_output)?,
        };
        opening = match value {
            Some(_) => rewrite_attribute(&opening, b"applyNumberFormat", "1", true, max_output)?,
            None => remove_attribute(&opening, b"applyNumberFormat", max_output)?,
        };
    }

    if element.empty {
        let mut children = Vec::new();
        children
            .try_reserve_exact(2)
            .map_err(|source| allocation("styles cell-XF child insertions", source))?;
        if let Some(alignment) = alignment {
            children.push(new_alignment_element(document, xf, alignment, max_output)?);
        }
        if let Some(protection) = protection {
            children.push(new_protection_element(
                document, xf, protection, max_output,
            )?);
        }
        if children.is_empty() {
            return Ok(opening);
        }
        return expand_empty_xf(element, &opening, &children, max_output);
    }

    let mut edits = Vec::new();
    edits
        .try_reserve(5)
        .map_err(|source| allocation("styles cell-XF edits", source))?;
    edits.push(ByteEdit {
        start: 0,
        end: element.open_end - element.start,
        replacement: opening,
    });
    if let Some(alignment) = alignment {
        if let Some(alignment_element) = direct_alignment(document, xf)? {
            let replacement = rewrite_alignment_element(
                input,
                document,
                xf,
                alignment_element,
                alignment,
                max_output,
            )?;
            edits.push(ByteEdit {
                start: document.elements[alignment_element].start - element.start,
                end: document.elements[alignment_element].end - element.start,
                replacement,
            });
        } else {
            let alignment_markup = new_alignment_element(document, xf, alignment, max_output)?;
            let insert_at = document
                .alignment_properties
                .get(xf)
                .and_then(|properties| properties.alignment_insert_at)
                .ok_or_else(|| invalid("styles cell-XF has no alignment insertion point"))?;
            edits.push(ByteEdit {
                start: insert_at - element.start,
                end: insert_at - element.start,
                replacement: alignment_markup,
            });
        }
    }
    if let Some(protection) = protection {
        if let Some(protection_element) = direct_protection(document, xf)? {
            let replacement = rewrite_protection_element(
                input,
                document,
                xf,
                protection_element,
                protection,
                max_output,
            )?;
            edits.push(ByteEdit {
                start: document.elements[protection_element].start - element.start,
                end: document.elements[protection_element].end - element.start,
                replacement,
            });
        } else {
            let protection_markup = new_protection_element(document, xf, protection, max_output)?;
            let insert_at = document
                .alignment_properties
                .get(xf)
                .and_then(|properties| properties.protection_insert_at)
                .ok_or_else(|| invalid("styles cell-XF has no protection insertion point"))?;
            edits.push(ByteEdit {
                start: insert_at - element.start,
                end: insert_at - element.start,
                replacement: protection_markup,
            });
        }
    }
    apply_edits(source, edits, max_output)
}

fn expand_empty_xf(
    element: &Element,
    opening: &[u8],
    children: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>> {
    let close = opening
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty cell XF has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && opening[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || opening[slash - 1] != b'/' {
        return Err(invalid("styles empty cell XF has no self-closing slash"));
    }
    slash -= 1;
    let children_len = children
        .iter()
        .try_fold(0usize, |size, child| size.checked_add(child.len()))
        .ok_or_else(|| invalid("styles expanded cell-XF size overflows"))?;
    let output_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(children_len))
        .and_then(|size| size.checked_add(element.qname.len()))
        .and_then(|size| size.checked_add(3))
        .ok_or_else(|| invalid("styles expanded cell-XF size overflows"))?;
    enforce_limit(
        "styles candidate cell-XF bytes",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded cell XF", source))?;
    result.extend_from_slice(&opening[..slash]);
    result.push(b'>');
    for child in children {
        result.extend_from_slice(child);
    }
    result.extend_from_slice(b"</");
    result.extend_from_slice(&element.qname);
    result.push(b'>');
    Ok(result)
}

fn new_alignment_element(
    document: &Document,
    xf: usize,
    change: &AlignmentChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let xf_element = &document.elements[xf];
    let prefix_len = xf_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname = &xf_element.qname[..prefix_len];
    let text_rotation = change
        .text_rotation_change()
        .and_then(|value| value.map(|value| value.get().to_string()));
    let indent = change
        .indent_change()
        .and_then(|value| value.map(|value| value.get().to_string()));
    let relative_indent = change
        .relative_indent_change()
        .and_then(|value| value.map(|value| value.to_string()));
    let reading_order = change
        .reading_order_change()
        .and_then(|value| value.map(|value| value.get().to_string()));
    let attributes: [(&[u8], Option<&str>); 9] = [
        (
            b"horizontal",
            change.horizontal_change().flatten().map(Horizontal::as_str),
        ),
        (
            b"vertical",
            change.vertical_change().flatten().map(Vertical::as_str),
        ),
        (b"textRotation", text_rotation.as_deref()),
        (
            b"wrapText",
            change
                .wrap_text_change()
                .flatten()
                .map(|value| if value { "1" } else { "0" }),
        ),
        (b"indent", indent.as_deref()),
        (b"relativeIndent", relative_indent.as_deref()),
        (
            b"justifyLastLine",
            change
                .justify_last_line_change()
                .flatten()
                .map(|value| if value { "1" } else { "0" }),
        ),
        (
            b"shrinkToFit",
            change
                .shrink_to_fit_change()
                .flatten()
                .map(|value| if value { "1" } else { "0" }),
        ),
        (b"readingOrder", reading_order.as_deref()),
    ];
    let mut output_len = prefix_len
        .checked_add(1 + b"alignment".len() + 2)
        .ok_or_else(|| invalid("styles alignment property size overflows"))?;
    for (name, value) in attributes
        .iter()
        .filter_map(|(name, value)| value.map(|value| (*name, value)))
    {
        let escaped_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(name.len())
            .and_then(|size| size.checked_add(escaped_len))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| invalid("styles alignment property size overflows"))?;
    }
    enforce_limit(
        "styles alignment property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles alignment property", source))?;
    output.push(b'<');
    output.extend_from_slice(qname);
    output.extend_from_slice(b"alignment");
    for (name, value) in attributes
        .iter()
        .filter_map(|(name, value)| value.map(|value| (*name, value)))
    {
        append_alignment_attribute(&mut output, name, value)?;
    }
    output.extend_from_slice(b"/>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn append_alignment_attribute(output: &mut Vec<u8>, name: &[u8], value: &str) -> Result<()> {
    let escaped_len = escaped_attribute_len(value)?;
    output.push(b' ');
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(output, value, escaped_len)?;
    output.push(b'\"');
    Ok(())
}

fn rewrite_alignment_element(
    input: &[u8],
    document: &Document,
    xf: usize,
    alignment: usize,
    change: &AlignmentChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let element = &document.elements[alignment];
    let state = alignment_state(document, xf)?;
    enforce_limit(
        "styles alignment opening",
        element.open_end - element.start,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut opening = copy_bytes(
        &input[element.start..element.open_end],
        "styles alignment opening",
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"horizontal",
        change.horizontal_change(),
        state.horizontal,
        |value| value.as_str().to_owned(),
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"vertical",
        change.vertical_change(),
        state.vertical,
        |value| value.as_str().to_owned(),
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"textRotation",
        change.text_rotation_change(),
        state.text_rotation,
        |value| value.get().to_string(),
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"wrapText",
        change.wrap_text_change(),
        state.wrap_text,
        |value| {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        },
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"indent",
        change.indent_change(),
        state.indent,
        |value| value.get().to_string(),
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"relativeIndent",
        change.relative_indent_change(),
        state.relative_indent,
        |value| value.to_string(),
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"justifyLastLine",
        change.justify_last_line_change(),
        state.justify_last_line,
        |value| {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        },
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"shrinkToFit",
        change.shrink_to_fit_change(),
        state.shrink_to_fit,
        |value| {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        },
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"readingOrder",
        change.reading_order_change(),
        state.reading_order,
        |value| value.get().to_string(),
        max_output,
    )?;
    let output_len = opening
        .len()
        .checked_add(element.end - element.open_end)
        .ok_or_else(|| invalid("styles alignment element size overflows"))?;
    enforce_limit(
        "styles alignment element",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles alignment element", source))?;
    result.extend_from_slice(&opening);
    result.extend_from_slice(&input[element.open_end..element.end]);
    Ok(result)
}

fn apply_alignment_attribute<T: Copy + PartialEq>(
    opening: &mut Vec<u8>,
    name: &[u8],
    requested: Option<Option<T>>,
    current: Option<T>,
    format: impl Fn(T) -> String,
    max_output: usize,
) -> Result<()> {
    let Some(requested) = requested else {
        return Ok(());
    };
    if requested == current {
        return Ok(());
    }
    match requested {
        Some(value) => {
            let lexical = format(value);
            *opening = rewrite_attribute(opening, name, &lexical, true, max_output)?;
        },
        None => {
            *opening = remove_attribute(opening, name, max_output)?;
        },
    }
    Ok(())
}

fn new_protection_element(
    document: &Document,
    xf: usize,
    change: &ProtectionChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let xf_element = &document.elements[xf];
    let prefix_len = xf_element
        .qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    let qname = &xf_element.qname[..prefix_len];
    let attributes: [(&[u8], Option<&str>); 2] = [
        (
            b"locked",
            change
                .locked_change()
                .flatten()
                .map(|value| if value { "1" } else { "0" }),
        ),
        (
            b"hidden",
            change
                .hidden_change()
                .flatten()
                .map(|value| if value { "1" } else { "0" }),
        ),
    ];
    let mut output_len = prefix_len
        .checked_add(1 + b"protection".len() + 2)
        .ok_or_else(|| invalid("styles protection property size overflows"))?;
    for (name, value) in attributes
        .iter()
        .filter_map(|(name, value)| value.map(|value| (*name, value)))
    {
        let escaped_len = escaped_attribute_len(value)?;
        output_len = output_len
            .checked_add(name.len())
            .and_then(|size| size.checked_add(escaped_len))
            .and_then(|size| size.checked_add(4))
            .ok_or_else(|| invalid("styles protection property size overflows"))?;
    }
    enforce_limit(
        "styles protection property",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles protection property", source))?;
    output.push(b'<');
    output.extend_from_slice(qname);
    output.extend_from_slice(b"protection");
    for (name, value) in attributes
        .iter()
        .filter_map(|(name, value)| value.map(|value| (*name, value)))
    {
        append_alignment_attribute(&mut output, name, value)?;
    }
    output.extend_from_slice(b"/>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn qname_prefix(qname: &[u8]) -> &[u8] {
    let prefix_len = qname
        .iter()
        .position(|byte| *byte == b':')
        .map_or(0, |colon| colon + 1);
    &qname[..prefix_len]
}

fn number_format_prefix(document: &Document) -> Result<&[u8]> {
    let element = document
        .num_fmts
        .or_else(|| {
            document
                .elements
                .iter()
                .position(|element| element.parent.is_none())
        })
        .ok_or_else(|| invalid("styles root has no number-format namespace context"))?;
    Ok(qname_prefix(&document.elements[element].qname))
}

fn new_fills_collection(
    document: &Document,
    records: &[Vec<u8>],
    count: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = fill_qname_prefix(document, VIRTUAL_NONE_FILL);
    let count = count.to_string();
    let records_len = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles fill collection size overflows"))?;
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"fills".len()))
        .and_then(|size| size.checked_add(1 + b"count".len() + 2 + count.len()))
        .and_then(|size| size.checked_add(2))
        .and_then(|size| size.checked_add(records_len))
        .and_then(|size| size.checked_add(2 + prefix.len() + b"fills".len() + 1))
        .ok_or_else(|| invalid("styles fill collection size overflows"))?;
    enforce_limit(
        "styles fill collection",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles fill collection", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"fills count=\"");
    output.extend_from_slice(count.as_bytes());
    output.extend_from_slice(b"\">");
    for record in records {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"fills>");
    if output.len() != output_len {
        return Err(invalid("styles fill collection size accounting mismatch"));
    }
    Ok(output)
}

fn new_num_fmt_element(
    document: &Document,
    id: u32,
    code: &str,
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = number_format_prefix(document)?;
    let id = id.to_string();
    let escaped_len = escaped_attribute_len(code)?;
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"numFmt".len()))
        .and_then(|size| size.checked_add(1 + b"numFmtId".len() + 3 + id.len()))
        .and_then(|size| size.checked_add(1 + b"formatCode".len() + 3 + escaped_len))
        .and_then(|size| size.checked_add(2))
        .ok_or_else(|| invalid("styles number-format element size overflows"))?;
    enforce_limit(
        "styles number-format element",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles number-format element", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"numFmt numFmtId=\"");
    output.extend_from_slice(id.as_bytes());
    output.extend_from_slice(b"\" formatCode=\"");
    append_escaped_attribute(&mut output, code, escaped_len)?;
    output.extend_from_slice(b"\"/>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn new_num_fmts_collection(
    document: &Document,
    records: &[Vec<u8>],
    max_output: usize,
) -> Result<Vec<u8>> {
    let prefix = number_format_prefix(document)?;
    let count = records.len().to_string();
    let records_len = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles number-format collection size overflows"))?;
    let output_len = 1usize
        .checked_add(prefix.len())
        .and_then(|size| size.checked_add(b"numFmts".len()))
        .and_then(|size| size.checked_add(1 + b"count".len() + 4 + count.len()))
        .and_then(|size| size.checked_add(1 + prefix.len()))
        .and_then(|size| size.checked_add(b"numFmts".len() + 2))
        .and_then(|size| size.checked_add(records_len))
        .ok_or_else(|| invalid("styles number-format collection size overflows"))?;
    enforce_limit(
        "styles number-format collection",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles number-format collection", source))?;
    output.push(b'<');
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"numFmts count=\"");
    output.extend_from_slice(count.as_bytes());
    output.extend_from_slice(b"\">");
    for record in records {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"numFmts>");
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn rewrite_protection_element(
    input: &[u8],
    document: &Document,
    xf: usize,
    protection: usize,
    change: &ProtectionChange,
    max_output: usize,
) -> Result<Vec<u8>> {
    let element = &document.elements[protection];
    let state = protection_state(document, xf)?;
    enforce_limit(
        "styles protection opening",
        element.open_end - element.start,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut opening = copy_bytes(
        &input[element.start..element.open_end],
        "styles protection opening",
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"locked",
        change.locked_change(),
        state.locked,
        |value| {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        },
        max_output,
    )?;
    apply_alignment_attribute(
        &mut opening,
        b"hidden",
        change.hidden_change(),
        state.hidden,
        |value| {
            if value {
                "1".to_owned()
            } else {
                "0".to_owned()
            }
        },
        max_output,
    )?;
    let output_len = opening
        .len()
        .checked_add(element.end - element.open_end)
        .ok_or_else(|| invalid("styles protection element size overflows"))?;
    enforce_limit(
        "styles protection element",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut result = Vec::new();
    result
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles protection element", source))?;
    result.extend_from_slice(&opening);
    result.extend_from_slice(&input[element.open_end..element.end]);
    Ok(result)
}

fn rewrite_boolean_opening(input: &[u8], value: bool, max_output: usize) -> Result<Vec<u8>> {
    rewrite_attribute(
        input,
        b"val",
        if value { "1" } else { "0" },
        true,
        max_output,
    )
}

fn rewrite_count(
    input: &[u8],
    element: &Element,
    count: usize,
    max_output: usize,
) -> Result<ByteEdit> {
    Ok(ByteEdit {
        start: element.start,
        end: element.open_end,
        replacement: rewrite_attribute(
            &input[element.start..element.open_end],
            b"count",
            &count.to_string(),
            true,
            max_output,
        )?,
    })
}

fn append_records_to_collection(
    input: &[u8],
    element: &Element,
    records: &[Vec<u8>],
    resource: &'static str,
    count: usize,
    max_output: usize,
) -> Result<ByteEdit> {
    if element.empty {
        return Ok(ByteEdit {
            start: element.start,
            end: element.end,
            replacement: expand_empty_collection(input, element, records, count, max_output)?,
        });
    }
    let mut replacement = Vec::new();
    let length = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles collection append size overflows"))?;
    enforce_limit(resource, length, max_output, litchi_core::Resource::Memory)?;
    replacement
        .try_reserve_exact(length)
        .map_err(|source| allocation(resource, source))?;
    for record in records {
        replacement.extend_from_slice(record);
    }
    let close_start = element
        .close_start
        .ok_or_else(|| invalid("styles collection has no closing tag"))?;
    Ok(ByteEdit {
        start: close_start,
        end: close_start,
        replacement,
    })
}

fn append_border_records_to_collection(
    input: &[u8],
    document: &Document,
    collection: usize,
    records: &[Vec<u8>],
    count: usize,
    max_output: usize,
) -> Result<ByteEdit> {
    let element = &document.elements[collection];
    if element.empty {
        return Ok(ByteEdit {
            start: element.start,
            end: element.end,
            replacement: expand_empty_collection(input, element, records, count, max_output)?,
        });
    }
    let mut insertion = element
        .close_start
        .ok_or_else(|| invalid("styles border collection has no closing tag"))?;
    for child in document.elements.iter() {
        if child.parent == Some(collection)
            && (!child.spreadsheetml || child.local.as_ref() == b"extLst")
        {
            insertion = child.start;
            break;
        }
    }
    let length = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles border append size overflows"))?;
    enforce_limit(
        "styles border append",
        length,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut replacement = Vec::new();
    replacement
        .try_reserve_exact(length)
        .map_err(|source| allocation("styles border append", source))?;
    for record in records {
        replacement.extend_from_slice(record);
    }
    Ok(ByteEdit {
        start: insertion,
        end: insertion,
        replacement,
    })
}

fn expand_empty_collection(
    input: &[u8],
    element: &Element,
    records: &[Vec<u8>],
    count: usize,
    max_output: usize,
) -> Result<Vec<u8>> {
    let source = &input[element.start..element.end];
    let close = source
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles empty collection has no closing bracket"))?;
    let mut slash = close;
    while slash > 0 && source[slash - 1].is_ascii_whitespace() {
        slash -= 1;
    }
    if slash == 0 || source[slash - 1] != b'/' {
        return Err(invalid("styles empty collection has no self-closing slash"));
    }
    slash -= 1;
    let records_len = records
        .iter()
        .try_fold(0usize, |size, record| size.checked_add(record.len()))
        .ok_or_else(|| invalid("styles expanded collection size overflows"))?;
    let mut expanded = Vec::new();
    let expanded_len = slash
        .checked_add(1)
        .and_then(|size| size.checked_add(records_len))
        .and_then(|size| size.checked_add(element.qname.len()))
        .and_then(|size| size.checked_add(3))
        .ok_or_else(|| invalid("styles expanded collection size overflows"))?;
    enforce_limit(
        "styles expanded collection",
        expanded_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    expanded
        .try_reserve_exact(expanded_len)
        .map_err(|source| allocation("styles expanded collection", source))?;
    expanded.extend_from_slice(&source[..slash]);
    expanded.push(b'>');
    for record in records {
        expanded.extend_from_slice(record);
    }
    expanded.extend_from_slice(b"</");
    expanded.extend_from_slice(&element.qname);
    expanded.push(b'>');

    let open_end = expanded
        .iter()
        .position(|byte| *byte == b'>')
        .and_then(|position| position.checked_add(1))
        .ok_or_else(|| invalid("styles expanded collection has no opening bracket"))?;
    let opening = rewrite_attribute(
        &expanded[..open_end],
        b"count",
        &count.to_string(),
        true,
        max_output,
    )?;
    let output_len = opening
        .len()
        .checked_add(expanded.len() - open_end)
        .ok_or_else(|| invalid("styles expanded collection size overflows"))?;
    enforce_limit(
        "styles expanded collection",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles expanded collection", source))?;
    output.extend_from_slice(&opening);
    output.extend_from_slice(&expanded[open_end..]);
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn append_records(
    element: &Element,
    records: &[Vec<u8>],
    resource: &'static str,
) -> Result<ByteEdit> {
    let mut replacement = Vec::new();
    for record in records {
        replacement
            .try_reserve(record.len())
            .map_err(|source| allocation(resource, source))?;
        replacement.extend_from_slice(record);
    }
    let close_start = element
        .close_start
        .ok_or_else(|| invalid("styles collection has no closing tag"))?;
    Ok(ByteEdit {
        start: close_start,
        end: close_start,
        replacement,
    })
}

fn rewrite_attribute(
    input: &[u8],
    name: &[u8],
    value: &str,
    add_if_missing: bool,
    max_output: usize,
) -> Result<Vec<u8>> {
    let escaped_len = escaped_attribute_len(value)?;
    enforce_limit(
        "styles escaped attribute",
        escaped_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let attributes = raw_attributes(input)?;
    let mut matched = None;
    for attribute in &attributes {
        if attribute.name == name {
            if matched.is_some() {
                return Err(invalid(format!(
                    "styles XML start tag repeats {}",
                    String::from_utf8_lossy(name)
                )));
            }
            matched = Some(attribute);
        }
    }
    if matched.is_none() && !add_if_missing {
        return Ok(input.to_vec());
    }
    if let Some(attribute) = matched {
        let output_len = input
            .len()
            .checked_sub(attribute.value_end - attribute.value_start)
            .and_then(|size| size.checked_add(escaped_len))
            .ok_or_else(|| invalid("styles attribute rewrite size overflows"))?;
        enforce_limit(
            "styles attribute rewrite",
            output_len,
            max_output,
            litchi_core::Resource::Memory,
        )?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(output_len)
            .map_err(|source| allocation("styles attribute rewrite", source))?;
        output.extend_from_slice(&input[..attribute.value_start]);
        append_escaped_attribute(&mut output, value, escaped_len)?;
        output.extend_from_slice(&input[attribute.value_end..]);
        return Ok(output);
    }
    let close = input
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid("styles XML start tag is malformed"))?;
    let mut insertion = close;
    while insertion > 0 && input[insertion - 1].is_ascii_whitespace() {
        insertion -= 1;
    }
    if insertion == 0 || input[insertion - 1] != b'/' {
        insertion = close;
    } else {
        insertion -= 1;
    }
    let output_len = input
        .len()
        .checked_add(name.len())
        .and_then(|size| size.checked_add(escaped_len + 4))
        .ok_or_else(|| invalid("styles attribute size overflows"))?;
    enforce_limit(
        "styles attribute insertion",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles attribute insertion", source))?;
    output.extend_from_slice(&input[..insertion]);
    output.push(b' ');
    output.extend_from_slice(name);
    output.extend_from_slice(b"=\"");
    append_escaped_attribute(&mut output, value, escaped_len)?;
    output.push(b'"');
    output.extend_from_slice(&input[insertion..]);
    Ok(output)
}

fn remove_attribute(input: &[u8], name: &[u8], max_output: usize) -> Result<Vec<u8>> {
    let attributes = raw_attributes(input)?;
    let mut matches = attributes
        .iter()
        .filter(|attribute| attribute.name.as_slice() == name);
    let Some(first) = matches.next() else {
        return copy_bytes(input, "styles unchanged attribute");
    };
    if matches.next().is_some() {
        return Err(invalid(format!(
            "styles XML start tag repeats {}",
            String::from_utf8_lossy(name)
        )));
    }
    let mut start = first.start;
    while start > 0 && input[start - 1].is_ascii_whitespace() {
        start -= 1;
    }
    let output_len = input
        .len()
        .checked_sub(first.end - start)
        .ok_or_else(|| invalid("styles attribute removal size underflows"))?;
    enforce_limit(
        "styles attribute removal",
        output_len,
        max_output,
        litchi_core::Resource::Memory,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles attribute removal", source))?;
    output.extend_from_slice(&input[..start]);
    output.extend_from_slice(&input[first.end..]);
    Ok(output)
}

fn raw_attributes(input: &[u8]) -> Result<Vec<RawAttribute>> {
    let mut result = Vec::new();
    let mut index = 1usize;
    while index < input.len()
        && !input[index].is_ascii_whitespace()
        && !matches!(input[index], b'>' | b'/')
    {
        index += 1;
    }
    while index < input.len() {
        while index < input.len() && input[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= input.len() || input[index] == b'>' {
            break;
        }
        if input[index] == b'/' {
            index += 1;
            while index < input.len() && input[index].is_ascii_whitespace() {
                index += 1;
            }
            if input.get(index) == Some(&b'>') {
                break;
            }
            return Err(invalid(
                "styles XML start tag has malformed self-closing syntax",
            ));
        }
        let attribute_start = index;
        let name_start = index;
        while index < input.len()
            && !input[index].is_ascii_whitespace()
            && !matches!(input[index], b'=' | b'/' | b'>')
        {
            index += 1;
        }
        if name_start == index {
            return Err(invalid("styles XML attribute name is empty"));
        }
        let name = copy_bytes(&input[name_start..index], "styles raw attribute name")?;
        while index < input.len() && input[index].is_ascii_whitespace() {
            index += 1;
        }
        if input.get(index) != Some(&b'=') {
            return Err(invalid("styles XML attribute has no value"));
        }
        index += 1;
        while index < input.len() && input[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *input
            .get(index)
            .ok_or_else(|| invalid("styles XML attribute value is missing"))?;
        if quote != b'\'' && quote != b'"' {
            return Err(invalid("styles XML attribute value is not quoted"));
        }
        index += 1;
        let value_start = index;
        while input.get(index).is_some_and(|byte| *byte != quote) {
            index += 1;
        }
        if index >= input.len() {
            return Err(invalid("styles XML attribute value is unterminated"));
        }
        let value_end = index;
        index += 1;
        result
            .try_reserve(1)
            .map_err(|source| allocation("styles raw attributes", source))?;
        result.push(RawAttribute {
            name,
            start: attribute_start,
            end: index,
            value_start,
            value_end,
        });
    }
    Ok(result)
}

fn escaped_attribute_len(value: &str) -> Result<usize> {
    value.bytes().try_fold(0usize, |size, byte| {
        let addition = match byte {
            b'&' => 5,
            b'<' => 4,
            b'"' | b'\'' => 6,
            b'\r' | b'\n' | b'\t' => 5,
            _ => 1,
        };
        size.checked_add(addition)
            .ok_or_else(|| invalid("styles escaped attribute size overflows"))
    })
}

fn append_escaped_attribute(output: &mut Vec<u8>, value: &str, expected: usize) -> Result<()> {
    let start = output.len();
    for byte in value.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'"' => output.extend_from_slice(b"&quot;"),
            b'\'' => output.extend_from_slice(b"&apos;"),
            b'\r' => output.extend_from_slice(b"&#xD;"),
            b'\n' => output.extend_from_slice(b"&#xA;"),
            b'\t' => output.extend_from_slice(b"&#x9;"),
            byte => output.push(byte),
        }
    }
    if output.len().saturating_sub(start) != expected {
        return Err(invalid("styles escaped attribute size underflow"));
    }
    Ok(())
}

fn apply_edits(input: &[u8], mut edits: Vec<ByteEdit>, max_output: usize) -> Result<Vec<u8>> {
    edits.sort_by(|left, right| {
        right
            .start
            .cmp(&left.start)
            .then_with(|| right.end.cmp(&left.end))
    });
    // `into_iter().rev()` below emits edits from left to right.  Keep two
    // insertions at one offset together first; otherwise stable sorting keeps
    // their input order and the reverse traversal would emit the later one
    // first (for example `<i/><b/>`).
    let mut merged = Vec::new();
    merged
        .try_reserve_exact(edits.len())
        .map_err(|source| allocation("styles merged edits", source))?;
    for edit in edits {
        if edit.start == edit.end
            && merged.last().is_some_and(|previous: &ByteEdit| {
                previous.start == edit.start && previous.end == edit.end
            })
        {
            let Some(previous) = merged.last_mut() else {
                return Err(invalid("styles merged insertion has no predecessor"));
            };
            previous
                .replacement
                .try_reserve(edit.replacement.len())
                .map_err(|source| allocation("styles merged insertion", source))?;
            previous.replacement.extend(edit.replacement);
        } else {
            merged.push(edit);
        }
    }
    let mut output_len = input.len();
    let mut next_start = input.len();
    for edit in &merged {
        if edit.start > edit.end || edit.end > next_start || edit.end > input.len() {
            return Err(invalid("styles XML edits overlap"));
        }
        output_len = output_len
            .checked_sub(edit.end - edit.start)
            .and_then(|size| size.checked_add(edit.replacement.len()))
            .ok_or_else(|| invalid("styles output size overflows"))?;
        next_start = edit.start;
    }
    enforce_limit(
        "styles output bytes",
        output_len,
        max_output,
        litchi_core::Resource::OutputBytes,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| allocation("styles output", source))?;
    let mut cursor = 0;
    for edit in merged.into_iter().rev() {
        output.extend_from_slice(&input[cursor..edit.start]);
        output.extend_from_slice(&edit.replacement);
        cursor = edit.end;
    }
    output.extend_from_slice(&input[cursor..]);
    Ok(output)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("styles XML position does not fit usize"))
}

fn enforce_limit(
    scope: &'static str,
    observed: usize,
    maximum: usize,
    resource: litchi_core::Resource,
) -> Result<()> {
    if observed > maximum {
        return Err(limit_error(scope, observed, maximum, resource));
    }
    Ok(())
}

fn limit_error(
    scope: &'static str,
    observed: usize,
    maximum: usize,
    resource: litchi_core::Resource,
) -> Error {
    Error::ResourceLimit(litchi_core::ResourceLimit {
        resource,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
        scope: Arc::from(scope),
    })
}

fn copy_bytes(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut copy = Vec::new();
    copy.try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    copy.extend_from_slice(bytes);
    Ok(copy)
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use super::{
        DEFAULT_MAX_DEPTH, StyleEditLimits, StyleRequest, changed_font, direct_collection,
        direct_items, index_fill_properties, index_font_properties, rewrite, rewrite_with_limits,
        scan,
    };
    use crate::style::cell_format::{Fill, Format, FormatTable, GradientFill, GradientStop};
    use crate::style::stylesheet::border::Tint;
    use crate::style::{
        AlignmentChange, BorderChange, BorderColor, FillChange, FillColor, FillPattern, FontChange,
        FontColor, FontColorKind, Horizontal, Indent, Line, ProtectionChange, Reading, Rotation,
        Scheme, Script, Side, Underline, Vertical,
    };

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_NS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    fn source() -> Vec<u8> {
        format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font f:keep="yes"><x:name val="Calibri"/><f:future keep="1"/><x:sz val="11"/></x:font></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" f:xf="yes"/></x:cellXfs></x:styleSheet>"#
        )
        .into_bytes()
    }

    fn request(change: FontChange) -> StyleRequest {
        StyleRequest {
            style: 0,
            change,
            fill: FillChange::new(),
            border: BorderChange::new(),
            alignment: AlignmentChange::new(),
            protection: ProtectionChange::new(),
            quote_prefix: None,
            number_format: None,
        }
    }

    fn alignment_request(alignment: AlignmentChange) -> StyleRequest {
        StyleRequest {
            style: 0,
            change: FontChange::new(),
            fill: FillChange::new(),
            border: BorderChange::new(),
            alignment,
            protection: ProtectionChange::new(),
            quote_prefix: None,
            number_format: None,
        }
    }

    fn protection_request(protection: ProtectionChange) -> StyleRequest {
        StyleRequest {
            style: 0,
            change: FontChange::new(),
            fill: FillChange::new(),
            border: BorderChange::new(),
            alignment: AlignmentChange::new(),
            protection,
            quote_prefix: None,
            number_format: None,
        }
    }

    fn number_format_request(value: Option<Option<&str>>) -> StyleRequest {
        StyleRequest {
            style: 0,
            change: FontChange::new(),
            fill: FillChange::new(),
            border: BorderChange::new(),
            alignment: AlignmentChange::new(),
            protection: ProtectionChange::new(),
            quote_prefix: None,
            number_format: value.map(|value| value.map(Arc::<str>::from)),
        }
    }

    fn fill_request(change: FillChange) -> StyleRequest {
        StyleRequest {
            style: 0,
            change: FontChange::new(),
            fill: change,
            border: BorderChange::new(),
            alignment: AlignmentChange::new(),
            protection: ProtectionChange::new(),
            quote_prefix: None,
            number_format: None,
        }
    }

    fn border_request(change: BorderChange) -> StyleRequest {
        StyleRequest {
            style: 0,
            change: FontChange::new(),
            fill: FillChange::new(),
            border: change,
            alignment: AlignmentChange::new(),
            protection: ProtectionChange::new(),
            quote_prefix: None,
            number_format: None,
        }
    }

    fn border_source() -> Vec<u8> {
        format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill></x:fills><x:borders count="1"><x:border f:keep="yes"><x:left style='thin' f:left="yes"><x:color rgb='FF112233'/></x:left><f:future marker="keep"/></x:border></x:borders><x:cellXfs count="1"><x:xf fontId="0" fillId="0" borderId="0" applyBorder='0'/></x:cellXfs></x:styleSheet>"#
        )
        .into_bytes()
    }

    fn fill_source() -> Vec<u8> {
        format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill f:keep="yes"><x:patternFill patternType='solid' f:pattern="yes"><x:fgColor rgb='FFFF0000' f:keep="fg"/><x:bgColor indexed="64"/><f:future/></x:patternFill></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        )
        .into_bytes()
    }

    fn cloned_xf<'a>(output: &'a str, qname: &str) -> &'a str {
        output.rsplit_once(qname).map_or("", |(_, tail)| tail)
    }

    #[test]
    fn number_format_builtin_writes_no_custom_collection_and_preserves_noop() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[number_format_request(Some(Some("0.00")))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes.clone()).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        assert_eq!(result.styles, vec![1]);
        assert!(!output.contains("numFmts"));
        assert!(cloned.contains(r#"numFmtId="2""#));
        assert!(cloned.contains(r#"applyNumberFormat="1""#));

        let repeated = rewrite(
            &result.bytes,
            &[StyleRequest {
                style: result.styles[0],
                ..number_format_request(Some(Some("0.00")))
            }],
        )
        .unwrap();
        assert_eq!(repeated.bytes, result.bytes);
        assert_eq!(repeated.styles, result.styles);
    }

    #[test]
    fn number_format_custom_collection_is_lossless_deduplicated_and_bounded() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:numFmts count="2"><x:numFmt numFmtId="164" formatCode="0.00&amp; units" f:keep="yes"/><x:numFmt numFmtId="166" formatCode="0.000"/></x:numFmts><x:fonts count="1"><x:font/></x:fonts><x:cellStyleXfs count="1"><x:xf numFmtId="165" fontId="0"/></x:cellStyleXfs><x:cellXfs count="1"><x:xf numFmtId="164" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[number_format_request(Some(Some("0.0000\" units\"")))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes.clone()).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        assert_eq!(result.styles, vec![1]);
        assert!(output.contains(r#"<x:numFmts count="3">"#));
        assert!(output.contains(r#"numFmtId="164" formatCode="0.00&amp; units" f:keep="yes""#));
        assert!(
            output.contains(r#"<x:numFmt numFmtId="167" formatCode="0.0000&quot; units&quot;"/>"#)
        );
        assert!(cloned.contains(r#"numFmtId="167""#));
        assert!(cloned.contains(r#"applyNumberFormat="1""#));

        let repeated = rewrite(
            &result.bytes,
            &[StyleRequest {
                style: result.styles[0],
                ..number_format_request(Some(Some("0.0000\" units\"")))
            }],
        )
        .unwrap();
        assert_eq!(repeated.bytes, result.bytes);
        assert_eq!(repeated.styles, result.styles);
    }

    #[test]
    fn number_format_creates_or_expands_collection_at_schema_position() {
        let absent = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let created = rewrite(
            absent.as_bytes(),
            &[number_format_request(Some(Some("0.00\" units\"")))],
        )
        .unwrap();
        let created = String::from_utf8(created.bytes).unwrap();
        let num_fmts = created.find("<x:numFmts").unwrap();
        let fonts = created.find("<x:fonts").unwrap();
        assert!(num_fmts < fonts);
        assert!(created.contains(r#"<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.00&quot; units&quot;"/></x:numFmts>"#));

        let empty = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:numFmts count="0"/><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let expanded = rewrite(
            empty.as_bytes(),
            &[number_format_request(Some(Some("0.00\" units\"")))],
        )
        .unwrap();
        let expanded = String::from_utf8(expanded.bytes).unwrap();
        assert!(expanded.contains(r#"<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.00&quot; units&quot;"/></x:numFmts>"#));
    }

    #[test]
    fn new_number_format_element_checks_exact_budget() {
        let input = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let document = scan(input.as_bytes(), DEFAULT_MAX_DEPTH).unwrap();
        let expected =
            super::new_num_fmt_element(&document, 164, "0.00\" units\"", usize::MAX).unwrap();
        assert_eq!(
            super::new_num_fmt_element(&document, 164, "0.00\" units\"", expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            super::new_num_fmt_element(
                &document,
                164,
                "0.00\" units\"",
                expected.len() - 1,
            ),
            Err(crate::Error::ResourceLimit(limit))
                if limit.resource == litchi_core::Resource::Memory
        ));
    }

    #[test]
    fn number_format_index_limits_source_records_before_hash_map_growth() {
        let input = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.00"/></x:numFmts><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="0" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let limits = StyleEditLimits {
            max_number_formats: 0,
            ..StyleEditLimits::default()
        };
        let error = rewrite_with_limits(
            input.as_bytes(),
            &[number_format_request(Some(Some("0.000")))],
            limits,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            crate::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::Objects
        ));
    }

    #[test]
    fn number_format_clear_removes_authored_id_and_activation_only_on_clone() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.00"/></x:numFmts><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf numFmtId="164" applyNumberFormat="1" fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(source.as_bytes(), &[number_format_request(Some(None))]).unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        assert_eq!(result.styles, vec![1]);
        assert!(cloned.contains(r#"fontId="0""#));
        assert!(!cloned.contains("numFmtId"));
        assert!(!cloned.contains("applyNumberFormat"));
        assert_eq!(output.matches(r#"numFmtId="164""#).count(), 2);
        assert_eq!(output.matches(r#"applyNumberFormat="1""#).count(), 1);
    }

    #[test]
    fn number_format_rejects_malformed_authored_codes_before_mutation() {
        for code in ["", "0;0;0;0;0", "[Red", "\"unterminated"] {
            let error = rewrite(&source(), &[number_format_request(Some(Some(code)))]).unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{code}: {error}");
        }
    }

    #[test]
    fn number_format_preserves_declared_builtin_overrides() {
        let source = format!(
            r#"<styleSheet xmlns="{NS}"><numFmts count="1"><numFmt numFmtId="2" formatCode="0.000"/></numFmts><fonts count="1"><font/></fonts><cellXfs count="1"><xf fontId="0" numFmtId="2"/></cellXfs></styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[number_format_request(Some(Some("0.00")))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert!(output.contains(r#"<numFmt numFmtId="2" formatCode="0.000"/>"#));
        assert!(output.contains(r#"<numFmt numFmtId="164" formatCode="0.00"/>"#));
        assert!(cloned_xf(&output, "<xf").contains(r#"numFmtId="164""#));
    }

    #[test]
    fn number_format_matching_source_code_without_apply_flag_is_exact_noop() {
        let source = format!(
            r#"<styleSheet xmlns="{NS}"><fonts count="1"><font/></fonts><cellXfs count="1"><xf fontId="0" numFmtId="2"/></cellXfs></styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[number_format_request(Some(Some("0.00")))],
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, [0]);
    }

    #[test]
    fn number_format_generated_records_obey_count_limit() {
        let error = rewrite_with_limits(
            &source(),
            &[number_format_request(Some(Some("0.00000")))],
            StyleEditLimits {
                max_number_formats: 0,
                ..StyleEditLimits::default()
            },
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::ResourceLimit(ref limit)
            if limit.resource == litchi_core::Resource::Objects
                && limit.scope.as_ref() == "styles generated number formats"));
    }

    #[test]
    fn pattern_fill_rewrites_colors_and_pattern_losslessly() {
        let source = fill_source();
        let tint = Tint::new(0.5).unwrap();
        let change = FillChange::new()
            .pattern(Some(FillPattern::DarkGray))
            .foreground(Some(FillColor::theme(2).with_tint(tint)))
            .background(Some(FillColor::auto_value(true)));
        let result = rewrite(&source, &[fill_request(change.clone())]).unwrap();
        assert_eq!(result.styles, vec![1]);
        let output = String::from_utf8(result.bytes.clone()).unwrap();
        let fill_records = output
            .match_indices(r#"<x:fill"#)
            .filter(|(offset, _)| {
                matches!(
                    output.as_bytes().get(offset + b"<x:fill".len()),
                    Some(b'>' | b'/' | b' ' | b'\t' | b'\r' | b'\n')
                )
            })
            .count();
        assert_eq!(fill_records, 3);
        assert_eq!(output.matches(r#"f:keep="yes""#).count(), 2);
        assert_eq!(output.matches(r#"<f:future/>"#).count(), 2);
        assert!(output.contains(r#"<x:patternFill patternType='solid' f:pattern="yes">"#));
        assert!(output.contains(r#"theme="2""#));
        assert!(output.contains(r#"tint="0.5""#));
        assert!(output.contains(r#"auto="1""#));
        assert_eq!(output.matches(r#"<x:fgColor"#).count(), 2);
        assert_eq!(output.matches(r#"<x:bgColor"#).count(), 2);
        let parsed = FormatTable::parse(&result.bytes).expect("rewritten fills parse");
        assert_eq!(parsed.len(), 2);
        let reopened = FormatTable::parse(&result.bytes).expect("reopened fills parse");
        assert_eq!(reopened, parsed);

        let repeated = rewrite(
            &result.bytes,
            &[StyleRequest {
                style: result.styles[0],
                ..fill_request(change.clone())
            }],
        )
        .unwrap();
        assert_eq!(repeated.bytes, result.bytes);
        assert_eq!(repeated.styles, result.styles);
    }

    #[test]
    fn pattern_fill_clear_removes_authored_properties_only_on_clone() {
        let source = fill_source();
        let result = rewrite(
            &source,
            &[fill_request(
                FillChange::new()
                    .pattern(None)
                    .foreground(None)
                    .background(None),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert_eq!(output.matches(r#"patternType='solid'"#).count(), 1);
        assert_eq!(output.matches(r#"<x:fgColor"#).count(), 1);
        assert_eq!(output.matches(r#"<x:bgColor"#).count(), 1);
        assert!(output.contains(r#"<x:patternFill f:pattern="yes"><f:future/></x:patternFill>"#));
        let cloned = output
            .rsplit_once(r#"<x:fill"#)
            .map_or("", |(_, tail)| tail);
        assert!(cloned.contains(r#"applyFill="1""#));
    }

    fn linear_gradient() -> GradientFill {
        GradientFill::linear(
            45.0,
            vec![
                GradientStop::new(0.0, FillColor::rgb(0x10, 0x20, 0x30)).unwrap(),
                GradientStop::new(1.0, FillColor::theme(4).with_tint(Tint::new(0.25).unwrap()))
                    .unwrap(),
            ]
            .into_boxed_slice(),
        )
        .unwrap()
    }

    #[test]
    fn whole_gradient_replacement_preserves_outer_markup_and_reopens() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill f:keep="yes"><x:gradientFill type="linear" degree="0"><x:stop position="0"><x:color rgb="FFFF0000"/></x:stop><x:stop position="1"><x:color theme="2"/></x:stop></x:gradientFill><f:fillFuture/></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[fill_request(FillChange::replace(Fill::Gradient(
                linear_gradient(),
            )))],
        )
        .unwrap();
        assert_eq!(result.styles, [1]);
        let output = String::from_utf8(result.bytes.clone()).unwrap();
        assert_eq!(output.matches(r#"f:keep="yes""#).count(), 2);
        assert_eq!(output.matches(r#"<f:fillFuture/>"#).count(), 2);
        assert_eq!(output.matches(r#"<x:gradientFill"#).count(), 2);
        assert!(output.contains(r#"degree="45""#));
        assert!(output.contains(r#"rgb="FF102030""#));
        assert!(output.contains(r#"theme="4" tint="0.25""#));
        let parsed = FormatTable::parse(&result.bytes).expect("rewritten gradient parses");
        let reopened = FormatTable::parse(&result.bytes).expect("reopened gradient parses");
        assert_eq!(reopened, parsed);
    }

    #[test]
    fn read_created_gradient_replacement_retains_geometry_for_both_kinds() {
        for kind in ["linear", "path"] {
            let donor = format!(
                r#"<styleSheet xmlns="{NS}"><fonts count="1"><font/></fonts><fills count="1"><fill><gradientFill type="{kind}" degree="45.5" left="0.1" right="0.2" top="0.3" bottom="0.4"><stop position="0.25"><color indexed="7" tint="-0.5"/></stop></gradientFill></fill></fills><cellXfs count="1"><xf fontId="0" fillId="0"/></cellXfs></styleSheet>"#
            );
            let donor_format =
                Format::new(Arc::new(FormatTable::parse(donor.as_bytes()).unwrap()), 0).unwrap();
            let requested = donor_format.fill().unwrap().clone();
            let result = rewrite(
                &source(),
                &[fill_request(FillChange::replace(requested.clone()))],
            )
            .unwrap();
            let reopened = Format::new(
                Arc::new(FormatTable::parse(&result.bytes).unwrap()),
                result.styles[0],
            )
            .unwrap();
            assert_eq!(reopened.fill(), Some(&requested), "{kind}");
        }
    }

    #[test]
    fn whole_gradient_semantic_noop_retains_opaque_source_and_bytes() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill f:keep="yes"><x:gradientFill type="linear" degree="45" f:future="yes"><x:stop position="0"><x:color rgb="FF102030"/></x:stop><x:stop position="1"><x:color theme="4" tint="0.25"/></x:stop></x:gradientFill><f:fillFuture/></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[fill_request(FillChange::replace(Fill::Gradient(
                linear_gradient(),
            )))],
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, [0]);
    }

    #[test]
    fn whole_gradient_match_only_activates_disabled_apply_without_touching_opaque_fill() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:gradientFill type="linear" degree="45" f:future="yes"><x:stop position="0"><x:color rgb="FF102030"/></x:stop><x:stop position="1"><x:color theme="4" tint="0.25"/></x:stop></x:gradientFill></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1" applyFill="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[fill_request(FillChange::replace(Fill::Gradient(
                linear_gradient(),
            )))],
        )
        .unwrap();
        assert_eq!(result.styles, [1]);
        let output = String::from_utf8(result.bytes).unwrap();
        assert_eq!(output.matches(r#"f:future="yes""#).count(), 1);
        assert!(output.contains(r#"applyFill="1""#));
    }

    #[test]
    fn whole_gradient_change_refuses_opaque_source_before_generation() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:gradientFill type="linear" f:future="yes"><x:stop position="0"><x:color rgb="FFFF0000"/></x:stop></x:gradientFill></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite(
            source.as_bytes(),
            &[fill_request(FillChange::replace(Fill::Gradient(
                linear_gradient(),
            )))],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Unsupported { .. }), "{error}");
    }

    #[test]
    fn whole_gradient_generation_obeys_staged_bytes_before_child_allocations() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:fills count="1"><x:fill/></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite_with_limits(
            source.as_bytes(),
            &[fill_request(FillChange::replace(Fill::Gradient(
                linear_gradient(),
            )))],
            StyleEditLimits {
                max_staged_bytes: 96,
                ..StyleEditLimits::default()
            },
        )
        .unwrap_err();
        assert!(
            matches!(error, crate::Error::ResourceLimit(ref limit)
            if limit.resource == litchi_core::Resource::Memory),
            "{error}"
        );
    }

    #[test]
    fn gradient_fill_edits_refuse_before_mutation() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:gradientFill type="linear"><x:stop position="0"/></x:gradientFill></x:fill></x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite(
            source.as_bytes(),
            &[fill_request(
                FillChange::new().pattern(Some(FillPattern::Solid)),
            )],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Unsupported { .. }), "{error}");
    }

    #[test]
    fn fill_count_limit_is_checked_before_fill_lookup_growth() {
        let limits = StyleEditLimits {
            max_fills: 1,
            ..StyleEditLimits::default()
        };
        let error = rewrite_with_limits(
            &fill_source(),
            &[fill_request(
                FillChange::new().pattern(Some(FillPattern::Solid)),
            )],
            limits,
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::ResourceLimit(_)), "{error}");
    }

    #[test]
    fn missing_or_empty_fill_collection_gets_safe_defaults_before_new_fill() {
        for collection in ["", "<x:fills count=\"0\"></x:fills>"] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts>{collection}<x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let result = rewrite(
                source.as_bytes(),
                &[fill_request(
                    FillChange::new()
                        .pattern(Some(FillPattern::Solid))
                        .foreground(Some(FillColor::rgb(0x10, 0x20, 0x30))),
                )],
            )
            .unwrap();
            let output = String::from_utf8(result.bytes).unwrap();
            assert!(output.contains(r#"<x:fills count="3">"#));
            assert!(output.contains(r#"<x:fill><x:patternFill patternType="none"/></x:fill>"#));
            assert!(output.contains(r#"<x:fill><x:patternFill patternType="gray125"/></x:fill>"#));
            assert!(output.contains(r#"fillId="2""#));
            assert!(output.contains(r#"<x:fgColor rgb="FF102030"/>"#));
        }
    }

    #[test]
    fn missing_fills_are_inserted_after_fonts_before_later_collections() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:numFmts count="0"/><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[fill_request(
                FillChange::new().pattern(Some(FillPattern::Solid)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let num_fmts = output.find("<x:numFmts").unwrap();
        let fonts = output.find("<x:fonts").unwrap();
        let fills = output.find("<x:fills").unwrap();
        let cell_xfs = output.find("<x:cellXfs").unwrap();
        assert!(num_fmts < fonts && fonts < fills && fills < cell_xfs);
    }

    #[test]
    fn clearing_absent_virtual_pattern_is_exact_noop_without_default_budget() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite_with_limits(
            source.as_bytes(),
            &[fill_request(FillChange::new().pattern(None))],
            StyleEditLimits {
                max_fills: 0,
                max_staged_bytes: 0,
                ..StyleEditLimits::default()
            },
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, [0]);
    }

    #[test]
    fn fill_property_index_handles_large_record_sets_without_nested_lookup() {
        const COUNT: usize = 1_024;
        let fills = "<x:fill><x:patternFill patternType=\"solid\"/></x:fill>".repeat(COUNT);
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:fills count="{COUNT}">{fills}</x:fills><x:cellXfs count="1"><x:xf fontId="0" fillId="{last}"/></x:cellXfs></x:styleSheet>"#,
            last = COUNT - 1,
        );
        let document = scan(source.as_bytes(), DEFAULT_MAX_DEPTH).unwrap();
        let root = document
            .elements
            .iter()
            .position(|element| element.parent.is_none())
            .unwrap();
        let fills = direct_collection(&document.elements, root, b"fills");
        let fills = fills.unwrap().unwrap();
        let items = direct_items(&document.elements, fills, b"fill").unwrap();
        let properties = index_fill_properties(&document.elements, &items).unwrap();
        assert_eq!(items.len(), COUNT);
        assert!(
            items
                .iter()
                .all(|&fill| properties[fill].pattern_fill.is_some())
        );
    }

    #[test]
    fn missing_fill_collection_rejects_nonzero_dangling_references() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" fillId="1"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite(
            source.as_bytes(),
            &[fill_request(
                FillChange::new().pattern(Some(FillPattern::Solid)),
            )],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
    }

    #[test]
    fn stylesheet_envelope_rejects_content_outside_root_and_misplaced_declarations() {
        let source = String::from_utf8(source()).unwrap();
        let declaration = r#"<?xml version="1.0"?>"#;
        let invalid = [
            ("leading text", format!("text{source}")),
            ("trailing text", format!("{source}tail")),
            ("leading cdata", format!("<![CDATA[leading]]>{source}")),
            ("trailing cdata", format!("{source}<![CDATA[trailing]]>")),
            ("leading entity reference", format!("&amp;{source}")),
            ("trailing unmatched end", format!("{source}</x:styleSheet>")),
            ("post-root declaration", format!("{source}{declaration}")),
            (
                "declaration inside root",
                source.replacen("<x:fonts", &format!("{declaration}<x:fonts"), 1),
            ),
            (
                "duplicate declaration",
                format!("{declaration}{declaration}{source}"),
            ),
            (
                "declaration after whitespace",
                format!(" \n{declaration}{source}"),
            ),
            (
                "declaration after comment",
                format!("<!-- before declaration -->{declaration}{source}"),
            ),
            (
                "declaration after processing instruction",
                format!("<?before declaration?>{declaration}{source}"),
            ),
        ];
        for (label, input) in invalid {
            assert!(
                rewrite(
                    input.as_bytes(),
                    &[request(FontChange::new().bold(Some(true)))]
                )
                .is_err(),
                "malformed stylesheet envelope was accepted: {label}"
            );
        }

        let valid = format!(
            " \n<!-- before root --><?before root?>{source}<?after root?><!-- after root -->\n"
        );
        assert!(
            rewrite(
                valid.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
            )
            .is_ok()
        );

        let mut bom_valid = vec![0xef, 0xbb, 0xbf];
        bom_valid.extend_from_slice(format!("{declaration}{source}").as_bytes());
        assert!(rewrite(&bom_valid, &[request(FontChange::new().bold(Some(true)))],).is_ok());

        let non_xml_whitespace = format!("\u{000b}{source}");
        assert!(
            rewrite(
                non_xml_whitespace.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
            )
            .is_err()
        );
    }

    #[test]
    fn output_limit_applies_to_empty_and_noop_requests() {
        let input = source();
        let limits = StyleEditLimits {
            max_output_bytes: input.len() - 1,
            ..StyleEditLimits::default()
        };
        for requests in [Vec::new(), vec![request(FontChange::new().bold(None))]] {
            assert!(matches!(rewrite_with_limits(&input, &requests, limits),
                Err(crate::Error::ResourceLimit(limit))
                    if limit.resource == litchi_core::Resource::OutputBytes));
        }
    }

    #[test]
    fn expanded_empty_font_uses_exact_candidate_budget() {
        let input = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let mut document = scan(input.as_bytes(), DEFAULT_MAX_DEPTH).unwrap();
        let fonts = document.fonts.unwrap();
        let font_items = direct_items(&document.elements, fonts, b"font").unwrap();
        document.font_properties = index_font_properties(&document.elements, &font_items).unwrap();
        let font = font_items[0];
        let change = FontChange::new().bold(Some(true));
        let expected =
            changed_font(input.as_bytes(), &document, font, &change, usize::MAX).unwrap();
        assert_eq!(
            changed_font(input.as_bytes(), &document, font, &change, expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            changed_font(
                input.as_bytes(),
                &document,
                font,
                &change,
                expected.len() - 1,
            ),
            Err(crate::Error::ResourceLimit(limit))
                if limit.resource == litchi_core::Resource::Memory
        ));
    }

    fn ignorable_source() -> String {
        String::from_utf8(source()).unwrap().replace(
            "xmlns:f=\"urn:future\"",
            "xmlns:f=\"urn:future\" xmlns:compat=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" compat:Ignorable=\"f\"",
        )
    }

    #[test]
    fn ignorable_extensions_are_preserved_and_repeated_edits_are_exact_noops() {
        for namespace in [NS, "http://purl.oclc.org/ooxml/spreadsheetml/main"] {
            let input = ignorable_source().replace(NS, namespace);
            let result = rewrite(
                input.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
            )
            .unwrap();
            let text = std::str::from_utf8(&result.bytes).unwrap();
            assert!(text.contains("compat:Ignorable=\"f\""));
            assert_eq!(text.matches("<f:future keep=\"1\"/>").count(), 2);
            assert_eq!(text.matches("f:xf=\"yes\"").count(), 2);
            let repeat = rewrite(
                &result.bytes,
                &[StyleRequest {
                    style: result.styles[0],
                    change: FontChange::new().bold(Some(true)),
                    fill: FillChange::new(),
                    border: BorderChange::new(),
                    alignment: AlignmentChange::new(),
                    protection: ProtectionChange::new(),
                    quote_prefix: None,
                    number_format: None,
                }],
            )
            .unwrap();
            assert_eq!(repeat.bytes, result.bytes);
            assert_eq!(repeat.styles, result.styles);
        }
    }

    #[test]
    fn unproven_mce_projections_refuse_before_resource_rewriting() {
        let source = ignorable_source();
        for input in [
            source.replace("<f:future keep=\"1\"/>", "<f:future><x:b/></f:future>"),
            source.replace("compat:Ignorable=\"f\"", "compat:Ignorable=\"f\" compat:ProcessContent=\"f:future\""),
            source.replace("compat:Ignorable=\"f\"", "compat:MustUnderstand=\"f\""),
            source.replace("compat:Ignorable=\"f\"", "compat:Ignorable=\"missing\""),
            source.replace("<x:font f:keep", "<x:font compat:Ignorable=\"f\" f:keep"),
            source.replace("<f:future keep=\"1\"/>", "<compat:AlternateContent><compat:Fallback><x:b/></compat:Fallback></compat:AlternateContent>"),
            source.replace("spreadsheetml/2006/main", "spreadsheetml/2006/m&#97;in"),
        ] {
            assert!(matches!(rewrite(input.as_bytes(), &[request(FontChange::new().bold(Some(true)))]),
                Err(crate::Error::Unsupported { .. })), "{input}");
        }
    }

    #[test]
    fn nested_ignorable_understood_namespaces_keep_their_clone_context() {
        let input =
            ignorable_source().replace("<x:font f:keep", "<x:font compat:Ignorable=\"x\" f:keep");
        let result = rewrite(
            input.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        assert_eq!(
            std::str::from_utf8(&result.bytes)
                .unwrap()
                .matches("compat:Ignorable=\"x\"")
                .count(),
            2
        );
    }

    #[test]
    fn cloned_font_retains_locally_rebound_prefixes() {
        let input = ignorable_source()
            .replace(
                "<x:font f:keep",
                &format!("<p:font xmlns:p=\"{NS}\" xmlns:x=\"urn:shadow\" f:keep"),
            )
            .replace("</x:font>", "</p:font>")
            .replace("<x:name", "<p:name")
            .replace("<x:sz", "<p:sz");
        let result = rewrite(
            input.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        let output = std::str::from_utf8(&result.bytes).unwrap();
        assert!(output.contains("<p:b/>"));
        assert_eq!(output.matches("xmlns:x=\"urn:shadow\"").count(), 2);
    }

    #[test]
    fn new_font_properties_use_sdk_order_without_moving_opaque_children() {
        let input = String::from_utf8(source()).unwrap().replace(
            "<x:name val=\"Calibri\"/>",
            "<f:before keep=\"1\"/><x:name val=\"Calibri\"/>",
        );
        let result = rewrite(
            input.as_bytes(),
            &[request(
                FontChange::new().bold(Some(true)).italic(Some(true)),
            )],
        )
        .unwrap();
        let output = std::str::from_utf8(&result.bytes).unwrap();
        let before = output.rfind("<f:before keep=\"1\"/>").unwrap();
        let future = output.rfind("<f:future keep=\"1\"/>").unwrap();
        let bold = output.find("<x:b/>").unwrap();
        let italic = output.find("<x:i/>").unwrap();
        let name = output.rfind("<x:name val=\"Calibri\"/>").unwrap();
        let size = output.rfind("<x:sz val=\"11\"/>").unwrap();
        assert!(before < bold, "opaque child before the anchor moved");
        assert!(bold < italic, "same-point properties were reversed");
        assert!(
            italic < name && name < future && future < size,
            "known font order was not kept"
        );
        assert_eq!(result.styles, vec![1]);
    }

    #[test]
    fn empty_elements_obey_the_same_depth_bound_as_start_elements() {
        let input = format!(
            "<styleSheet xmlns=\"{NS}\"><fonts><font/></fonts><cellXfs><xf/></cellXfs></styleSheet>"
        );
        let limits = StyleEditLimits {
            max_depth: 2,
            ..StyleEditLimits::default()
        };
        assert!(
            matches!(rewrite_with_limits(input.as_bytes(), &[request(FontChange::new().bold(Some(true)))], limits),
            Err(crate::Error::ResourceLimit(limit)) if limit.resource == litchi_core::Resource::Depth)
        );
    }

    #[test]
    fn creates_lossless_font_and_xf_clones() {
        let result = rewrite(&source(), &[request(FontChange::new().bold(Some(true)))]).unwrap();
        let after = String::from_utf8(result.bytes).unwrap();
        assert!(after.contains(r#"f:keep="yes""#));
        assert!(after.contains(r#"<f:future keep="1"/>"#));
        assert!(after.contains(r#"f:xf="yes""#));
        assert!(after.contains(r#"<x:b/>"#));
        assert!(after.contains(r#"fontId="1""#));
        assert!(after.contains(r#"applyFont="1""#));
        assert_eq!(result.styles, vec![1]);
    }

    #[test]
    fn clear_and_explicit_false_are_distinct() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:b/><x:i/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let clear = rewrite(source.as_bytes(), &[request(FontChange::new().bold(None))]).unwrap();
        let false_value = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(false)))],
        )
        .unwrap();
        let clear = String::from_utf8(clear.bytes).unwrap();
        let false_value = String::from_utf8(false_value.bytes).unwrap();
        // Existing records remain available to every other cell. Only the
        // appended font loses bold, while italic remains authored.
        assert!(clear.contains(r#"<x:font><x:b/><x:i/></x:font>"#));
        assert!(clear.contains(r#"<x:font><x:i/></x:font>"#));
        assert!(false_value.contains(r#"<x:b val="0"/>"#));
    }

    #[test]
    fn strike_and_underline_preserve_tri_state_values() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:u val="double"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .strike(Some(true))
                    .underline(Some(Underline::SingleAccounting)),
            )],
        )
        .unwrap();
        let changed_text = String::from_utf8(changed.bytes).unwrap();
        assert!(
            changed_text.contains(r#"<x:strike/>"#)
                || changed_text.contains(r#"<x:strike val="1"/>"#)
        );
        assert!(changed_text.contains(r#"<x:u val="singleAccounting"/>"#));

        let cleared = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().strike(None).underline(None))],
        )
        .unwrap();
        let cleared_text = String::from_utf8(cleared.bytes).unwrap();
        assert!(cleared_text.contains(r#"<x:font><x:u val="double"/></x:font>"#));
        assert!(cleared_text.contains(r#"<x:font></x:font>"#));

        let explicit_none = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().underline(Some(Underline::None)))],
        )
        .unwrap();
        let explicit_none_text = String::from_utf8(explicit_none.bytes).unwrap();
        assert!(explicit_none_text.contains(r#"<x:u val="none"/>"#));
    }

    #[test]
    fn strike_and_underline_noop_values_keep_exact_source() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:strike val="true"/><x:u/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .strike(Some(true))
                    .underline(Some(Underline::Single)),
            )],
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, vec![0]);
    }

    #[test]
    fn name_size_and_script_keep_lexical_noops_and_escape_attribute_controls() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font f:keep="yes"><x:vertAlign val="baseline"/><x:sz val="11.0"/><x:name val="A&amp;B"/><f:future keep="1"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" f:xf="yes"/></x:cellXfs></x:styleSheet>"#
        );
        let noop = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .name(Some("A&B"))
                    .unwrap()
                    .size(Some(11.0))
                    .unwrap()
                    .script(Some(Script::Baseline)),
            )],
        )
        .unwrap();
        assert_eq!(noop.bytes, source.as_bytes());
        assert_eq!(noop.styles, vec![0]);

        let changed = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .name(Some("A&B\nC\rD\tE"))
                    .unwrap()
                    .size(Some(14.5))
                    .unwrap()
                    .script(Some(Script::Superscript)),
            )],
        )
        .unwrap();
        let changed = String::from_utf8(changed.bytes).unwrap();
        assert!(changed.contains(r#"val="A&amp;B&#xA;C&#xD;D&#x9;E""#));
        assert!(changed.contains(r#"<x:sz val="14.5"/>"#));
        assert!(changed.contains(r#"<x:vertAlign val="superscript"/>"#));
        assert!(changed.contains(r#"f:keep="yes""#));
        assert!(changed.contains(r#"f:future keep="1""#));
    }

    #[test]
    fn font_name_rewrite_preserves_single_quote_attribute_delimiters() {
        let input = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:name val='Old' keep="yes"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let change = FontChange::new().name(Some("O'Brien & \"Face\"")).unwrap();
        let result = rewrite(input.as_bytes(), &[request(change.clone())]).unwrap();
        let output = std::str::from_utf8(&result.bytes).unwrap();
        assert!(output.contains("val='O&apos;Brien &amp; &quot;Face&quot;' keep=\"yes\""));
        let repeated = rewrite(
            &result.bytes,
            &[StyleRequest {
                style: result.styles[0],
                change,
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new(),
                protection: ProtectionChange::new(),
                quote_prefix: None,
                number_format: None,
            }],
        )
        .unwrap();
        assert_eq!(repeated.bytes, result.bytes);
    }

    #[test]
    fn name_size_and_script_clear_only_the_cloned_font_properties() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:vertAlign val="subscript"/><x:sz val="11"/><x:name val="Calibri"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .name(None)
                    .unwrap()
                    .size(None)
                    .unwrap()
                    .script(None),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert_eq!(
            output.matches(r#"<x:vertAlign val="subscript"/>"#).count(),
            1
        );
        assert_eq!(output.matches(r#"<x:sz val="11"/>"#).count(), 1);
        assert_eq!(output.matches(r#"<x:name val="Calibri"/>"#).count(), 1);
        assert_eq!(result.styles, vec![1]);
    }

    #[test]
    fn unrelated_edits_preserve_source_sizes_outside_the_authoring_domain() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:sz val="500"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert!(output.contains(r#"<x:sz val="500"/>"#));
        assert!(FontChange::new().size(Some(500.0)).is_err());
    }

    #[test]
    fn unrelated_edits_preserve_source_family_and_charset_outside_authoring_domains() {
        assert!(FontChange::new().family(Some(6)).is_err());
        assert!(FontChange::new().charset(Some(256)).is_err());
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:family val="6"/><x:charset val="256"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert!(output.contains(r#"<x:family val="6"/>"#));
        assert!(output.contains(r#"<x:charset val="256"/>"#));
    }

    #[test]
    fn authored_name_and_size_inputs_are_validated_and_bounded() {
        assert!(FontChange::new().name(Some("")).is_err());
        assert!(FontChange::new().name(Some("bad\u{0001}")).is_err());
        assert!(FontChange::new().size(Some(0.1)).is_err());
        assert!(FontChange::new().size(Some(f64::NAN)).is_err());
        assert!(FontChange::new().size(Some(f64::INFINITY)).is_err());
        assert!(FontChange::new().size(Some(409.1)).is_err());

        let input = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let limits = StyleEditLimits {
            max_staged_bytes: 32,
            ..StyleEditLimits::default()
        };
        let error = rewrite_with_limits(
            input.as_bytes(),
            &[request(FontChange::new().name(Some("&&&&&&&&")).unwrap())],
            limits,
        )
        .unwrap_err();
        assert!(matches!(
            error,
            crate::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::Memory
        ));
    }

    #[test]
    fn color_switches_bases_preserves_opaque_markup_and_keeps_semantic_noops_exact() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font f:keep="yes"><x:color f:opaque='yes' theme='2' tint='0.25'><f:unknown marker="keep"/></x:color><f:after/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" f:xf="yes"/></x:cellXfs></x:styleSheet>"#
        );
        let tint = Tint::new(0.25).unwrap();
        let noop = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new().color(Some(FontColor::theme(2).with_tint(tint))),
            )],
        )
        .unwrap();
        assert_eq!(noop.bytes, source.as_bytes());
        assert_eq!(noop.styles, vec![0]);

        let changed = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new().color(Some(
                    FontColor::rgb(crate::Rgb::argb(0x80, 0x11, 0x22, 0x33))
                        .with_tint(Tint::new(-0.5).unwrap()),
                )),
            )],
        )
        .unwrap();
        let output = String::from_utf8(changed.bytes).unwrap();
        assert!(output.contains(
            r#"<x:color f:opaque='yes' rgb="80112233" tint='-0.5'><f:unknown marker="keep"/></x:color>"#
        ));
        assert_eq!(output.matches(r#"<f:unknown marker="keep"/>"#).count(), 2);
        assert_eq!(output.matches(r#"f:after/>"#).count(), 2);
        assert_eq!(changed.styles, vec![1]);
    }

    #[test]
    fn color_rewrites_preserve_same_base_quotes_and_manage_tint_presence() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:color rgb='FF112233'/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let changed = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().color(Some(FontColor::rgb(
                crate::Rgb::argb(0xff, 0x44, 0x55, 0x66),
            ))))],
        )
        .unwrap();
        let output = String::from_utf8(changed.bytes).unwrap();
        assert!(output.contains(r#"<x:color rgb='FF445566'/>"#));

        let with_tint = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new().color(Some(
                    FontColor::rgb(crate::Rgb::argb(0xff, 0x11, 0x22, 0x33))
                        .with_tint(Tint::new(0.25).unwrap()),
                )),
            )],
        )
        .unwrap();
        let output = String::from_utf8(with_tint.bytes).unwrap();
        assert!(output.contains(r#"<x:color rgb='FF112233' tint="0.25"/>"#));

        let tinted = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:color rgb='FF112233' tint='0.25'/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let switched = rewrite(
            tinted.as_bytes(),
            &[request(FontChange::new().color(Some(FontColor::theme(2))))],
        )
        .unwrap();
        let output = String::from_utf8(switched.bytes).unwrap();
        assert!(output.contains(r#"<x:color theme="2"/>"#));
        assert_eq!(output.matches("tint=").count(), 1);
    }

    #[test]
    fn color_clear_removes_only_the_cloned_color_and_default_or_auto_are_typed() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:color f:opaque="yes" rgb="FF112233" tint="0.5" xmlns:f="urn:opaque"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let cleared =
            rewrite(source.as_bytes(), &[request(FontChange::new().color(None))]).unwrap();
        let output = String::from_utf8(cleared.bytes).unwrap();
        assert_eq!(output.matches(r#"<x:color f:opaque="yes""#).count(), 1);
        assert_eq!(cleared.styles, vec![1]);

        let default_color = FontColor::default_base();
        let auto_false = FontColor::parsed(
            FontColorKind::Auto(false),
            None,
            "<x:color auto='0'/>".into(),
        );
        for (color, expected) in [
            (default_color, r#"<x:color/>"#),
            (auto_false, r#"<x:color auto="0"/>"#),
        ] {
            let input = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let result = rewrite(
                input.as_bytes(),
                &[request(FontChange::new().color(Some(color)))],
            )
            .unwrap();
            assert!(String::from_utf8(result.bytes).unwrap().contains(expected));
        }
    }

    #[test]
    fn malformed_source_colors_refuse_before_mutation() {
        for color in [
            r#"<x:color rgb="FF112233" theme="2"/>"#,
            r#"<x:color rgb="112233"/>"#,
            r#"<x:color tint="2"/>"#,
            r#"<x:color auto="maybe"/>"#,
            r#"<x:color auto="TRUE"/>"#,
        ] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font>{color}</x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let error = rewrite(
                source.as_bytes(),
                &[request(FontChange::new().color(Some(FontColor::theme(1))))],
            )
            .unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
        }
    }

    #[test]
    fn explicit_apply_font_false_is_activated_for_a_semantically_equal_color() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:color theme="2" tint="0.25"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyFont="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().color(Some(
                FontColor::theme(2).with_tint(Tint::new(0.25).unwrap()),
            )))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert!(output.contains(r#"applyFont="1""#));
        assert_eq!(result.styles, vec![1]);
    }

    #[test]
    fn uppercase_boolean_tokens_are_rejected_even_for_semantic_noops() {
        let cases = [
            (r#"<x:b val="TRUE"/>"#, FontChange::new().bold(Some(true))),
            (
                r#"<x:color auto="TRUE"/>"#,
                FontChange::new().color(Some(FontColor::automatic())),
            ),
        ];
        for (property, change) in cases {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font>{property}</x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let error = rewrite(source.as_bytes(), &[request(change)]).unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
        }

        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:b/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyFont="FALSE"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
    }

    #[test]
    fn empty_font_expands_strike_and_underline_in_schema_order() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .strike(Some(false))
                    .underline(Some(Underline::Double)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let strike = output.find(r#"<x:strike val="0"/>"#).unwrap();
        let underline = output.find(r#"<x:u val="double"/>"#).unwrap();
        assert!(strike < underline);
    }

    #[test]
    fn empty_font_expands_name_size_and_script_in_schema_order() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .name(Some("A&B"))
                    .unwrap()
                    .size(Some(12.0))
                    .unwrap()
                    .script(Some(Script::Subscript)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let script = output.find(r#"<x:vertAlign val="subscript"/>"#).unwrap();
        let size = output.find(r#"<x:sz val="12"/>"#).unwrap();
        let name = output.find(r#"<x:name val="A&amp;B"/>"#).unwrap();
        assert!(script < size && size < name);
    }

    #[test]
    fn empty_fonts_insert_color_after_size_and_repeat_exactly() {
        for font_markup in ["<x:font/>", "<x:font></x:font>"] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1">{font_markup}</x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let change = FontChange::new()
                .name(Some("Arial"))
                .unwrap()
                .size(Some(12.0))
                .unwrap()
                .color(Some(FontColor::theme(2)))
                .script(Some(Script::Superscript));
            let result = rewrite(source.as_bytes(), &[request(change.clone())]).unwrap();
            let output = String::from_utf8(result.bytes.clone()).unwrap();
            let start = output.rfind("<x:font>").unwrap();
            let end = start + output[start..].find("</x:font>").unwrap() + "</x:font>".len();
            let cloned_font = &output[start..end];
            let script = cloned_font
                .find(r#"<x:vertAlign val="superscript"/>"#)
                .unwrap();
            let size = cloned_font.find(r#"<x:sz val="12"/>"#).unwrap();
            let color = cloned_font.find(r#"<x:color theme="2"/>"#).unwrap();
            let name = cloned_font.find(r#"<x:name val="Arial"/>"#).unwrap();
            assert!(script < size && size < color && color < name);

            let repeated = rewrite(
                &result.bytes,
                &[StyleRequest {
                    style: result.styles[0],
                    change,
                    fill: FillChange::new(),
                    border: BorderChange::new(),
                    alignment: AlignmentChange::new(),
                    protection: ProtectionChange::new(),
                    quote_prefix: None,
                    number_format: None,
                }],
            )
            .unwrap();
            assert_eq!(repeated.bytes, result.bytes);
        }
    }

    #[test]
    fn remaining_font_fields_insert_in_schema_order_and_repeat_exactly() {
        for font_markup in ["<x:font/>", "<x:font></x:font>"] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1">{font_markup}</x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let change = FontChange::new()
                .condense(Some(true))
                .extend(Some(false))
                .outline(Some(true))
                .shadow(Some(false))
                .family(Some(2))
                .unwrap()
                .charset(Some(134))
                .unwrap()
                .scheme(Some(Scheme::Minor));
            let result = rewrite(source.as_bytes(), &[request(change.clone())]).unwrap();
            let output = String::from_utf8(result.bytes.clone()).unwrap();
            let start = output.rfind("<x:font").unwrap();
            let end = start + output[start..].find("</x:font>").unwrap() + "</x:font>".len();
            let cloned_font = &output[start..end];
            let condense = cloned_font.find(r#"<x:condense/>"#).unwrap();
            let extend = cloned_font.find(r#"<x:extend val="0"/>"#).unwrap();
            let outline = cloned_font.find(r#"<x:outline/>"#).unwrap();
            let shadow = cloned_font.find(r#"<x:shadow val="0"/>"#).unwrap();
            let family = cloned_font.find(r#"<x:family val="2"/>"#).unwrap();
            let charset = cloned_font.find(r#"<x:charset val="134"/>"#).unwrap();
            let scheme = cloned_font.find(r#"<x:scheme val="minor"/>"#).unwrap();
            assert!(
                condense < extend
                    && extend < outline
                    && outline < shadow
                    && shadow < family
                    && family < charset
                    && charset < scheme
            );

            let repeated = rewrite(
                &result.bytes,
                &[StyleRequest {
                    style: result.styles[0],
                    change,
                    fill: FillChange::new(),
                    border: BorderChange::new(),
                    alignment: AlignmentChange::new(),
                    protection: ProtectionChange::new(),
                    quote_prefix: None,
                    number_format: None,
                }],
            )
            .unwrap();
            assert_eq!(repeated.bytes, result.bytes);
        }
    }

    #[test]
    fn remaining_font_fields_preserve_lexical_noops_and_support_clear() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:condense extra='yes' val='true'/><x:family val='2'/><x:scheme val='minor'/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let noop = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .condense(Some(true))
                    .family(Some(2))
                    .unwrap()
                    .scheme(Some(Scheme::Minor)),
            )],
        )
        .unwrap();
        assert_eq!(noop.bytes, source.as_bytes());
        assert_eq!(noop.styles, vec![0]);

        let changed = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .condense(Some(false))
                    .family(Some(3))
                    .unwrap()
                    .scheme(Some(Scheme::Major)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(changed.bytes).unwrap();
        assert!(output.contains(r#"<x:condense extra='yes' val='0'/>"#));
        assert!(output.contains(r#"<x:family val='3'/>"#));
        assert!(output.contains(r#"<x:scheme val='major'/>"#));
        assert_eq!(changed.styles, vec![1]);

        let cleared = rewrite(
            source.as_bytes(),
            &[request(
                FontChange::new()
                    .condense(None)
                    .family(None)
                    .unwrap()
                    .scheme(None),
            )],
        )
        .unwrap();
        let output = String::from_utf8(cleared.bytes).unwrap();
        assert_eq!(output.matches("<x:condense").count(), 1);
        assert_eq!(output.matches("<x:family").count(), 1);
        assert_eq!(output.matches("<x:scheme").count(), 1);
    }

    #[test]
    fn malformed_remaining_font_fields_refuse_before_mutation() {
        for (property, change) in [
            (
                r#"<x:condense val="maybe"/>"#,
                FontChange::new().condense(Some(false)),
            ),
            (
                r#"<x:family val="-1"/>"#,
                FontChange::new().family(Some(2)).unwrap(),
            ),
            (
                r#"<x:scheme val="body"/>"#,
                FontChange::new().scheme(Some(Scheme::Minor)),
            ),
        ] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font>{property}</x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let error = rewrite(source.as_bytes(), &[request(change)]).unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
        }
    }

    #[test]
    fn empty_font_is_expanded_inside_the_font_element() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().italic(Some(false)))],
        )
        .unwrap();
        let after = String::from_utf8(result.bytes).unwrap();
        assert!(after.contains(r#"<x:font><x:i val="0"/></x:font>"#));
        assert!(!after.contains(r#"<x:font/><x:i"#));
    }

    #[test]
    fn semantic_noop_keeps_exact_source() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:b val="true"/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, vec![0]);
    }

    #[test]
    fn explicit_apply_font_false_is_activated_for_matching_property() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:b/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyFont="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap();
        let after = String::from_utf8(result.bytes).unwrap();
        assert!(after.contains(r#"applyFont="1""#));
        assert_eq!(result.styles, vec![1]);
    }

    #[test]
    fn alignment_only_clones_xf_without_cloning_font_or_activating_font() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font><x:b/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyFont="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new()
                    .horizontal(Some(Horizontal::Center))
                    .vertical(Some(Vertical::Bottom))
                    .text_rotation(Some(Rotation::degrees(45).unwrap()))
                    .wrap_text(Some(true))
                    .indent(Some(Indent::new(2)))
                    .shrink_to_fit(Some(true))
                    .reading_order(Some(Reading::RightToLeft)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert_eq!(result.styles, vec![1]);
        assert_eq!(output.matches("<x:font>").count(), 1);
        assert!(output.contains(r#"applyFont="0" applyAlignment="1""#));
        assert!(output.contains(
            r#"<x:alignment horizontal="center" vertical="bottom" textRotation="45" wrapText="1" indent="2" shrinkToFit="1" readingOrder="2"/>"#
        ));
    }

    #[test]
    fn new_alignment_checks_exact_budget_with_long_source_prefix() {
        let prefix = "p".repeat(256);
        let input = format!(
            r#"<{prefix}:styleSheet xmlns:{prefix}="{NS}"><{prefix}:fonts count="1"><{prefix}:font/></{prefix}:fonts><{prefix}:cellXfs count="1"><{prefix}:xf fontId="0"/></{prefix}:cellXfs></{prefix}:styleSheet>"#
        );
        let document = scan(input.as_bytes(), DEFAULT_MAX_DEPTH).unwrap();
        let xfs = direct_items(&document.elements, document.cell_xfs.unwrap(), b"xf").unwrap();
        let change = AlignmentChange::new()
            .horizontal(Some(Horizontal::CenterContinuous))
            .vertical(Some(Vertical::Distributed))
            .text_rotation(Some(Rotation::stacked()))
            .wrap_text(Some(true))
            .indent(Some(Indent::new(255)))
            .shrink_to_fit(Some(false))
            .reading_order(Some(Reading::RightToLeft))
            .relative_indent(Some(i32::MIN))
            .justify_last_line(Some(false));
        let expected =
            super::new_alignment_element(&document, xfs[0], &change, usize::MAX).unwrap();
        assert_eq!(
            super::new_alignment_element(&document, xfs[0], &change, expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            super::new_alignment_element(&document, xfs[0], &change, expected.len() - 1),
            Err(crate::Error::ResourceLimit(limit))
                if limit.resource == litchi_core::Resource::Memory
        ));
    }

    #[test]
    fn new_protection_checks_exact_budget_with_long_source_prefix() {
        let prefix = "p".repeat(256);
        let input = format!(
            r#"<{prefix}:styleSheet xmlns:{prefix}="{NS}"><{prefix}:fonts count="1"><{prefix}:font/></{prefix}:fonts><{prefix}:cellXfs count="1"><{prefix}:xf fontId="0"/></{prefix}:cellXfs></{prefix}:styleSheet>"#
        );
        let document = scan(input.as_bytes(), DEFAULT_MAX_DEPTH).unwrap();
        let xfs = direct_items(&document.elements, document.cell_xfs.unwrap(), b"xf").unwrap();
        let change = ProtectionChange::new()
            .locked(Some(true))
            .hidden(Some(false));
        let expected =
            super::new_protection_element(&document, xfs[0], &change, usize::MAX).unwrap();
        assert_eq!(
            super::new_protection_element(&document, xfs[0], &change, expected.len()).unwrap(),
            expected
        );
        assert!(matches!(
            super::new_protection_element(&document, xfs[0], &change, expected.len() - 1),
            Err(crate::Error::ResourceLimit(limit))
                if limit.resource == litchi_core::Resource::Memory
        ));
    }

    #[test]
    fn new_alignment_precedes_ext_lst_when_protection_is_absent() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"><x:extLst/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new().horizontal(Some(Horizontal::Right)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = output.rsplit_once("<x:xf").map_or("", |(_, tail)| tail);
        assert!(cloned.find("<x:alignment").unwrap() < cloned.find("<x:extLst").unwrap());
    }

    #[test]
    fn new_alignment_uses_the_xf_namespace_prefix_binding() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:p="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><p:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new().horizontal(Some(Horizontal::Left)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = output.rsplit_once("<p:xf").map_or("", |(_, tail)| tail);
        assert!(cloned.contains(r#"<p:alignment horizontal="left"/>"#));
    }

    #[test]
    fn alignment_rewrite_preserves_opaque_markup_and_attribute_lexical_forms() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="f"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyAlignment='0' f:xf='yes'><x:alignment horizontal='left' vertical="top" f:opaque="yes"><f:future/></x:alignment></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new()
                    .horizontal(None)
                    .vertical(Some(Vertical::Center))
                    .wrap_text(Some(true)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = output.rsplit_once("<x:xf").map_or("", |(_, tail)| tail);
        assert_eq!(result.styles, vec![1]);
        assert!(cloned.contains(r#"applyAlignment='1'"#));
        assert!(cloned.contains(r#"vertical="center""#));
        assert!(cloned.contains(r#"wrapText="1""#));
        assert!(!cloned.contains(r#"horizontal='left'"#));
        assert!(cloned.contains(r#"f:opaque="yes""#));
        assert!(cloned.contains(r#"<f:future/>"#));
        assert!(!cloned.contains(r#"applyFont="1""#));
    }

    #[test]
    fn alignment_matching_value_activates_apply_alignment_without_rewriting_markup() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyAlignment="0"><x:alignment horizontal='center'/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new().horizontal(Some(Horizontal::Center)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert_eq!(result.styles, vec![1]);
        assert!(output.contains(r#"applyAlignment="1"><x:alignment horizontal='center'/>"#));
    }

    #[test]
    fn alignment_semantic_noop_keeps_exact_source() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyAlignment="1"><x:alignment horizontal='center' wrapText="0" relativeIndent='-2' justifyLastLine="0"/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new()
                    .horizontal(Some(Horizontal::Center))
                    .wrap_text(Some(false))
                    .relative_indent(Some(-2))
                    .justify_last_line(Some(false)),
            )],
        )
        .unwrap();
        assert_eq!(result.bytes, source.as_bytes());
        assert_eq!(result.styles, vec![0]);
    }

    #[test]
    fn alignment_clear_removes_only_requested_attributes_and_keeps_wrapper_markup() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="f"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"><x:alignment horizontal="left" f:opaque="yes"><f:future/></x:alignment></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(AlignmentChange::new().horizontal(None))],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = output.rsplit_once("<x:xf").map_or("", |(_, tail)| tail);
        assert!(cloned.contains(r#"<x:alignment f:opaque="yes"><f:future/></x:alignment>"#));
        assert!(!cloned.contains(r#"horizontal="left""#));
    }

    #[test]
    fn malformed_alignment_flags_are_rejected_before_property_repair() {
        for flag in ["TRUE", "maybe"] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyAlignment="{flag}"><x:alignment horizontal="left"/></x:xf></x:cellXfs></x:styleSheet>"#
            );
            let error = rewrite(
                source.as_bytes(),
                &[alignment_request(
                    AlignmentChange::new().horizontal(Some(Horizontal::Center)),
                )],
            )
            .unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
        }
    }

    #[test]
    fn protection_only_clone_writes_both_flags_and_keeps_font_untouched() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[protection_request(
                ProtectionChange::new()
                    .locked(Some(true))
                    .hidden(Some(false)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        assert_eq!(result.styles, vec![1]);
        assert_eq!(output.matches("<x:font/>").count(), 1);
        assert!(cloned.contains(r#"applyProtection="1""#));
        assert!(cloned.contains(r#"<x:protection locked="1" hidden="0"/>"#));
        assert!(!cloned.contains("applyFont"));
        assert!(!cloned.contains("applyAlignment"));
    }

    #[test]
    fn protection_inserts_after_alignment_before_ext_lst_and_preserves_opaque_markup() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="f"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" f:xf="yes"><x:alignment horizontal="left"/><x:extLst><f:future/></x:extLst></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[protection_request(
                ProtectionChange::new().locked(Some(true)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        let alignment = cloned.find("<x:alignment").unwrap();
        let protection = cloned.find("<x:protection").unwrap();
        let ext_lst = cloned.find("<x:extLst").unwrap();
        assert!(alignment < protection && protection < ext_lst);
        assert!(cloned.contains(r#"f:xf="yes""#));
        assert!(cloned.contains(r#"<f:future/>"#));
    }

    #[test]
    fn protection_activation_keeps_lexical_values_and_clear_keeps_opaque_wrapper() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="f"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyProtection='0'><x:protection locked='1' f:opaque="yes"><f:future/></x:protection></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let activated = rewrite(
            source.as_bytes(),
            &[protection_request(
                ProtectionChange::new().locked(Some(true)),
            )],
        )
        .unwrap();
        let activated = String::from_utf8(activated.bytes).unwrap();
        let activated = cloned_xf(&activated, "<x:xf");
        assert!(activated.contains(r#"applyProtection='1'"#));
        assert!(activated.contains(r#"locked='1'"#));
        assert!(activated.contains(r#"f:opaque="yes""#));
        assert!(activated.contains(r#"<f:future/>"#));

        let cleared = rewrite(
            source.as_bytes(),
            &[protection_request(ProtectionChange::new().locked(None))],
        )
        .unwrap();
        let cleared = String::from_utf8(cleared.bytes).unwrap();
        let cleared = cloned_xf(&cleared, "<x:xf");
        assert!(cleared.contains(r#"<x:protection f:opaque="yes"><f:future/></x:protection>"#));
        assert!(!cleared.contains("locked="));
    }

    #[test]
    fn protection_false_is_authored_and_repeated_false_is_an_exact_noop() {
        let absent = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let authored = rewrite(
            absent.as_bytes(),
            &[protection_request(
                ProtectionChange::new().hidden(Some(false)),
            )],
        )
        .unwrap();
        let authored_text = String::from_utf8(authored.bytes.clone()).unwrap();
        assert!(authored_text.contains(r#"<x:protection hidden="0"/>"#));
        let repeated = rewrite(
            &authored.bytes,
            &[StyleRequest {
                style: authored.styles[0],
                change: FontChange::new(),
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new(),
                protection: ProtectionChange::new().hidden(Some(false)),
                quote_prefix: None,
                number_format: None,
            }],
        )
        .unwrap();
        assert_eq!(repeated.bytes, authored.bytes);
        assert_eq!(repeated.styles, authored.styles);
    }

    #[test]
    fn malformed_protection_flags_are_rejected_before_mutation() {
        for (xf_flag, child_flag) in [("TRUE", "1"), ("0", "TRUE")] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" applyProtection="{xf_flag}"><x:protection locked="{child_flag}"/></x:xf></x:cellXfs></x:styleSheet>"#
            );
            let error = rewrite(
                source.as_bytes(),
                &[protection_request(
                    ProtectionChange::new().locked(Some(true)),
                )],
            )
            .unwrap_err();
            assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
        }
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" quotePrefix="TRUE"/></x:cellXfs></x:styleSheet>"#
        );
        let error = rewrite(
            source.as_bytes(),
            &[StyleRequest {
                style: 0,
                change: FontChange::new(),
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new(),
                protection: ProtectionChange::new(),
                quote_prefix: Some(Some(true)),
                number_format: None,
            }],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Invalid(_)), "{error}");
    }

    #[test]
    fn quote_prefix_is_independent_and_preserves_delimiters_on_set_and_clear() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0" quotePrefix='0'/></x:cellXfs></x:styleSheet>"#
        );
        let set = rewrite(
            source.as_bytes(),
            &[StyleRequest {
                style: 0,
                change: FontChange::new(),
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new(),
                protection: ProtectionChange::new(),
                quote_prefix: Some(Some(true)),
                number_format: None,
            }],
        )
        .unwrap();
        let set = String::from_utf8(set.bytes).unwrap();
        let set = cloned_xf(&set, "<x:xf");
        assert!(set.contains(r#"quotePrefix='1'"#));
        assert!(!set.contains("applyFont"));
        assert!(!set.contains("applyAlignment"));
        assert!(!set.contains("applyProtection"));

        let cleared = rewrite(
            source.as_bytes(),
            &[StyleRequest {
                style: 0,
                change: FontChange::new(),
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new(),
                protection: ProtectionChange::new(),
                quote_prefix: Some(None),
                number_format: None,
            }],
        )
        .unwrap();
        let cleared = String::from_utf8(cleared.bytes).unwrap();
        let cleared = cloned_xf(&cleared, "<x:xf");
        assert!(!cleared.contains("quotePrefix"));
    }

    #[test]
    fn combined_style_facets_share_one_xf_clone_and_follow_child_order() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"><x:extLst/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[StyleRequest {
                style: 0,
                change: FontChange::new(),
                fill: FillChange::new(),
                border: BorderChange::new(),
                alignment: AlignmentChange::new()
                    .horizontal(Some(Horizontal::Center))
                    .relative_indent(Some(-2))
                    .justify_last_line(Some(false)),
                protection: ProtectionChange::new().locked(Some(false)),
                quote_prefix: Some(Some(true)),
                number_format: None,
            }],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = cloned_xf(&output, "<x:xf");
        assert_eq!(result.styles, vec![1]);
        assert_eq!(output.matches(r#"<x:xf fontId="0""#).count(), 2);
        assert!(cloned.contains(r#"applyAlignment="1""#));
        assert!(cloned.contains(r#"applyProtection="1""#));
        assert!(cloned.contains(r#"quotePrefix="1""#));
        let alignment = cloned.find("<x:alignment").unwrap();
        let protection = cloned.find("<x:protection").unwrap();
        let ext_lst = cloned.find("<x:extLst").unwrap();
        assert!(alignment < protection && protection < ext_lst);
        assert!(cloned.contains(r#"relativeIndent="-2""#));
        assert!(cloned.contains(r#"justifyLastLine="0""#));
    }

    #[test]
    fn contextual_rotation_follows_stylesheet_conformance_namespace() {
        for (namespace, accepted) in [(NS, true), (STRICT_NS, false)] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{namespace}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let result = rewrite(
                source.as_bytes(),
                &[alignment_request(
                    AlignmentChange::new().text_rotation(Some(Rotation::contextual())),
                )],
            );
            assert_eq!(result.is_ok(), accepted, "namespace {namespace}");
            if let Ok(result) = result {
                let output = String::from_utf8(result.bytes).unwrap();
                assert!(output.contains(r#"textRotation="254""#));
            }
        }

        let source = format!(
            r#"<x:styleSheet xmlns:x="{STRICT_NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"><x:alignment textRotation="254" horizontal="left"/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[alignment_request(
                AlignmentChange::new().horizontal(Some(Horizontal::Center)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        let cloned = output.rsplit_once("<x:xf").map_or("", |(_, tail)| tail);
        assert!(cloned.contains(r#"textRotation="254""#));
        assert!(cloned.contains(r#"horizontal="center""#));

        let mixed = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:s="{STRICT_NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"><s:alignment horizontal="left"/></x:xf></x:cellXfs></x:styleSheet>"#
        );
        assert!(
            rewrite(
                mixed.as_bytes(),
                &[alignment_request(
                    AlignmentChange::new().text_rotation(Some(Rotation::contextual())),
                )],
            )
            .is_err()
        );
    }

    #[test]
    fn duplicate_requests_deduplicate_records_and_keep_positions() {
        let result = rewrite(
            &source(),
            &[
                request(FontChange::new().bold(Some(true))),
                request(FontChange::new().bold(Some(true))),
            ],
        )
        .unwrap();
        let text = String::from_utf8(result.bytes).unwrap();
        assert_eq!(text.matches(r#"<x:b/>"#).count(), 1);
        assert_eq!(result.styles, vec![1, 1]);
    }

    #[test]
    fn missing_property_uses_the_font_prefix_binding() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:p="{NS}"><x:fonts count="1"><x:font><p:b/></x:font></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().italic(Some(true)))],
        )
        .unwrap();
        let after = String::from_utf8(result.bytes).unwrap();
        assert!(after.contains(r#"<x:i/>"#));
    }

    #[test]
    fn missing_collections_are_typed_unsupported() {
        let source = format!(r#"<x:styleSheet xmlns:x="{NS}"/>"#);
        let error = rewrite(
            source.as_bytes(),
            &[request(FontChange::new().bold(Some(true)))],
        )
        .unwrap_err();
        assert!(matches!(error, crate::error::Error::Unsupported { .. }));
    }

    #[test]
    fn namespace_and_markup_compatibility_are_checked() {
        let foreign = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let limits = StyleEditLimits {
            max_changes: 0,
            ..StyleEditLimits::default()
        };
        assert!(
            rewrite_with_limits(
                foreign.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
                limits
            )
            .is_err()
        );

        let limits = StyleEditLimits {
            max_output_bytes: foreign.len(),
            ..StyleEditLimits::default()
        };
        assert!(
            rewrite_with_limits(
                foreign.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
                limits,
            )
            .is_err()
        );

        let mce = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent/></x:styleSheet>"#
        );
        assert!(
            rewrite(
                mce.as_bytes(),
                &[request(FontChange::new().bold(Some(true)))],
            )
            .is_err()
        );
    }

    #[test]
    fn border_changes_clone_sides_flags_and_preserve_opaque_markup() {
        let result = rewrite(
            &border_source(),
            &[border_request(
                BorderChange::new()
                    .left_line(Some(Line::Double))
                    .right(Some(
                        Side::new(Line::Dotted).with_color(BorderColor::theme(4)),
                    ))
                    .outline(Some(false)),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes.clone()).unwrap();
        assert_eq!(result.styles, vec![1]);
        assert!(output.contains(r#"borders count="2""#));
        assert!(output.contains(r#"f:keep="yes""#));
        assert!(output.contains(r#"f:left="yes""#));
        assert!(output.contains(r#"<x:right style="dotted"><x:color theme="4"/></x:right>"#));
        assert!(output.contains(r#"outline="0""#));
        assert!(output.contains("applyBorder='1'"));
        assert!(output.contains(r#"<f:future marker="keep"/>"#));
        FormatTable::parse(&result.bytes).expect("rewritten border parses");
    }

    #[test]
    fn border_semantic_noop_keeps_source_bytes_and_color_only_without_side_refuses() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:borders count="1"><x:border><x:left style='thin'><x:color rgb='FF112233'/></x:left></x:border></x:borders><x:cellXfs count="1"><x:xf fontId="0" borderId="0" applyBorder="1"/></x:cellXfs></x:styleSheet>"#
        );
        let noop = rewrite(
            source.as_bytes(),
            &[border_request(
                BorderChange::new().left_line(Some(Line::Thin)),
            )],
        )
        .unwrap();
        assert_eq!(noop.bytes, source.as_bytes());
        let error = rewrite(
            source.as_bytes(),
            &[border_request(
                BorderChange::new().right_color(Some(BorderColor::rgb(1, 2, 3))),
            )],
        )
        .unwrap_err();
        assert!(matches!(error, crate::Error::Unsupported { .. }));

        let layered = rewrite(
            source.as_bytes(),
            &[border_request(
                BorderChange::new()
                    .left(Some(Side::new(Line::Thin)))
                    .left_line(Some(Line::Double))
                    .left_color(Some(BorderColor::theme(2))),
            )],
        )
        .unwrap();
        let layered = String::from_utf8(layered.bytes).unwrap();
        assert!(layered.contains(r#"<x:left style='double'><x:color theme="2"/></x:left>"#));

        let recreated = rewrite(
            source.as_bytes(),
            &[border_request(
                BorderChange::new().left(None).left_line(Some(Line::Double)),
            )],
        )
        .unwrap();
        let recreated = String::from_utf8(recreated.bytes).unwrap();
        assert!(recreated.contains(r#"<x:left style='double'></x:left>"#));
    }

    #[test]
    fn border_clear_retains_unrelated_side_markup_and_empty_collection_is_authorable() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:borders count="1"><x:border><x:left style="thin"><x:color rgb="FF112233"/></x:left><f:future/></x:border></x:borders><x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let cleared = rewrite(
            source.as_bytes(),
            &[border_request(BorderChange::new().left(None))],
        )
        .unwrap();
        let text = String::from_utf8(cleared.bytes).unwrap();
        assert!(text.contains("<x:border><f:future/></x:border>"));

        let empty = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:borders count="0"/><x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            empty.as_bytes(),
            &[border_request(
                BorderChange::new().top(Some(Side::new(Line::Medium))),
            )],
        )
        .unwrap();
        let output = String::from_utf8(result.bytes).unwrap();
        assert!(output.contains(r#"<x:borders count="2"><x:border/><x:border><x:top style="medium"></x:top></x:border></x:borders>"#));
        FormatTable::parse(output.as_bytes()).expect("empty border collection parses");
    }

    #[test]
    fn border_empty_and_absent_collections_round_trip_all_native_colors() {
        for namespace in [NS, "http://purl.oclc.org/ooxml/spreadsheetml/main"] {
            for collection in [
                "",
                "<x:borders count=\"0\"/>",
                "<x:borders count=\"1\"><x:border/></x:borders>",
            ] {
                for color in [
                    BorderColor::argb(0x80, 0x11, 0x22, 0x33),
                    BorderColor::theme(4),
                    BorderColor::indexed(12),
                    BorderColor::auto_value(false),
                    BorderColor::default_base(),
                ] {
                    let color = color.with_tint(Tint::new(-0.25).unwrap());
                    let source = format!(
                        r#"<x:styleSheet xmlns:x="{namespace}"><x:fonts count="1"><x:font/></x:fonts>{collection}<x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
                    );
                    let change = BorderChange::new()
                        .top_line(Some(Line::Thin))
                        .top_color(Some(color));
                    let result =
                        rewrite(source.as_bytes(), &[border_request(change.clone())]).unwrap();
                    let format = Format::new(
                        Arc::new(FormatTable::parse(&result.bytes).unwrap()),
                        result.styles[0],
                    )
                    .unwrap();
                    assert_eq!(
                        format.border().unwrap().top.as_ref(),
                        Some(&Side::new(Line::Thin).with_color(color))
                    );
                    let mut again = border_request(change);
                    again.style = result.styles[0];
                    assert_eq!(
                        rewrite(&result.bytes, &[again]).unwrap().bytes,
                        result.bytes
                    );
                }
            }
        }
    }

    #[test]
    fn border_color_removal_preserves_line_and_refuses_opaque_child_content() {
        for extra in [r#" f:keep="yes""#, ""] {
            let source = format!(
                r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:borders count="1"><x:border><x:left style='thin'><x:color rgb='FF112233'{extra}/></x:left></x:border></x:borders><x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
            );
            let result = rewrite(
                source.as_bytes(),
                &[border_request(BorderChange::new().left_color(None))],
            );
            if extra.is_empty() {
                let result = result.unwrap();
                assert!(
                    String::from_utf8_lossy(&result.bytes)
                        .contains("<x:left style='thin'></x:left>")
                );
                let format = Format::new(
                    Arc::new(FormatTable::parse(&result.bytes).unwrap()),
                    result.styles[0],
                )
                .unwrap();
                assert_eq!(format.border().unwrap().left, Some(Side::new(Line::Thin)));
            } else {
                assert!(matches!(result, Err(crate::Error::Unsupported { .. })));
            }
        }
    }

    #[test]
    fn border_additions_precede_future_children_and_collection_extensions() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}" xmlns:f="urn:future"><x:fonts count="1"><x:font/></x:fonts><x:borders count="1"><x:border><x:left style="thin"><f:future/></x:left><f:future/></x:border><x:extLst><x:ext uri="urn:test"/></x:extLst></x:borders><x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let result = rewrite(
            source.as_bytes(),
            &[border_request(
                BorderChange::new()
                    .left_color(Some(BorderColor::theme(4)))
                    .top_line(Some(Line::Thin)),
            )],
        )
        .unwrap();
        let text = String::from_utf8_lossy(&result.bytes);
        assert!(text.contains(r#"<x:left style="thin"><x:color theme="4"/><f:future/></x:left><x:top style="thin"></x:top><f:future/></x:border><x:extLst>"#));
        FormatTable::parse(&result.bytes).unwrap();
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                &[border_request(BorderChange::new().left(None))]
            ),
            Err(crate::Error::Unsupported { .. })
        ));
    }

    #[test]
    fn border_conformance_and_staged_index_limits_are_checked() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:borders count="1"><x:border/></x:borders><x:cellXfs count="1"><x:xf fontId="0" borderId="0"/></x:cellXfs></x:styleSheet>"#
        );
        assert!(
            rewrite(
                source.as_bytes(),
                &[border_request(
                    BorderChange::new().start_line(Some(Line::Thin))
                )]
            )
            .is_err()
        );
        let strict = source.replace(NS, "http://purl.oclc.org/ooxml/spreadsheetml/main");
        assert!(
            rewrite(
                strict.as_bytes(),
                &[border_request(
                    BorderChange::new().left_line(Some(Line::Thin))
                )]
            )
            .is_err()
        );
        let result = rewrite(
            strict.as_bytes(),
            &[border_request(
                BorderChange::new()
                    .start_line(Some(Line::Thin))
                    .end_line(Some(Line::Double)),
            )],
        )
        .unwrap();
        let format = Format::new(
            Arc::new(FormatTable::parse(&result.bytes).unwrap()),
            result.styles[0],
        )
        .unwrap();
        assert_eq!(format.border().unwrap().start, Some(Side::new(Line::Thin)));
        assert_eq!(format.border().unwrap().end, Some(Side::new(Line::Double)));
        let limited = rewrite_with_limits(
            source.as_bytes(),
            &[border_request(
                BorderChange::new().top_line(Some(Line::Thin)),
            )],
            StyleEditLimits {
                max_staged_bytes: 64,
                ..StyleEditLimits::default()
            },
        );
        assert!(matches!(limited, Err(crate::Error::ResourceLimit(_))));
    }

    #[test]
    fn fill_and_border_missing_collections_are_inserted_in_schema_order() {
        let source = format!(
            r#"<x:styleSheet xmlns:x="{NS}"><x:fonts count="1"><x:font/></x:fonts><x:cellXfs count="1"><x:xf fontId="0"/></x:cellXfs></x:styleSheet>"#
        );
        let mut request = border_request(BorderChange::new().top_line(Some(Line::Thin)));
        request.fill = FillChange::new()
            .pattern(Some(FillPattern::Solid))
            .foreground(Some(FillColor::rgb(1, 2, 3)));
        let result = rewrite(source.as_bytes(), &[request]).unwrap();
        let text = String::from_utf8_lossy(&result.bytes);
        assert!(text.find("<x:fonts").unwrap() < text.find("<x:fills").unwrap());
        assert!(text.find("<x:fills").unwrap() < text.find("<x:borders").unwrap());
        assert!(text.find("<x:borders").unwrap() < text.find("<x:cellXfs").unwrap());
        FormatTable::parse(&result.bytes).unwrap();
    }
    #[test]
    fn border_apply_activation_reuses_opaque_record_and_repeat_is_exact() {
        let source = border_source();
        let change = BorderChange::new().left_line(Some(Line::Thin));
        let result = rewrite(&source, &[border_request(change.clone())]).unwrap();
        let text = String::from_utf8_lossy(&result.bytes);
        assert!(text.contains(r#"borders count="1""#));
        assert!(text.contains("applyBorder='1'"));
        assert!(
            text.contains("<x:left style='thin' f:left=\"yes\"><x:color rgb='FF112233'/></x:left>")
        );
        let mut again = border_request(change);
        again.style = result.styles[0];
        assert_eq!(
            rewrite(&result.bytes, &[again]).unwrap().bytes,
            result.bytes
        );
    }
}
