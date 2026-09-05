//! Bounded catalog reads for Keynote table appearances.
//!
//! The migration host owns the physical object catalog, but it must not own
//! Keynote's style semantics.  This module keeps the semantic traversal in
//! the focused package and accepts only a short-lived source callback from the
//! host.  A callback borrows one decompressed payload for the duration of the
//! Buffa decode; no archive or payload is retained here.

use litchi_iwa_common::WireLimits;
use litchi_iwa_protos::table_appearance_codec;

use crate::slide::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

/// Source callbacks used by the migration host's catalog appearance adapter.
///
/// This trait is intentionally hidden behind the `internal-iwork-source`
/// feature.  It is a temporary bridge for the legacy host and is not part of
/// the supported Keynote facade.  Implementations must execute `read` while
/// the selected payload is borrowed and must not retain the payload.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub trait __CatalogTableAppearanceSource {
    fn message_type_count(&self, identifier: u64, message_type: u32) -> Result<usize, String>;

    fn with_message_data_type<T>(
        &mut self,
        identifier: u64,
        message_type: u32,
        type_name: &str,
        read: impl FnOnce(&[u8]) -> Result<T, String>,
    ) -> Result<T, String>;
}

/// Aggregate wire budget for one catalog appearance operation.
///
/// The codec protects each borrowed payload with `DecodeOptions`; this ledger
/// additionally bounds the complete style/preset traversal so a long chain
/// cannot multiply that per-payload allowance without limit.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
#[derive(Debug, Clone, Copy)]
struct AppearanceBudget {
    max_input_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: usize,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
}

#[cfg(feature = "internal-iwork-source")]
impl AppearanceBudget {
    /// Build an aggregate budget from the caller's wire profile.
    #[must_use]
    const fn from_wire_limits(limits: WireLimits) -> Self {
        Self {
            max_input_bytes: limits.max_input_bytes(),
            max_fields: limits.max_fields(),
            max_work_bytes: limits.max_rewrite_work(),
            max_nesting: limits.max_nesting(),
            input_bytes: 0,
            fields: 0,
            work_bytes: 0,
        }
    }

    fn options(&self, source: &[u8]) -> table_appearance_codec::DecodeOptions {
        let base = table_appearance_codec::DecodeOptions::for_source(source);
        let recursion_limit = u32::try_from(self.max_nesting).unwrap_or(u32::MAX);
        base.with_max_input_bytes(
            base.max_input_bytes()
                .min(self.max_input_bytes.saturating_sub(self.input_bytes)),
        )
        .with_max_fields(
            base.max_fields()
                .min(self.max_fields.saturating_sub(self.fields)),
        )
        .with_max_work_bytes(
            base.max_work_bytes()
                .min(self.max_work_bytes.saturating_sub(self.work_bytes)),
        )
        .with_recursion_limit(base.recursion_limit().min(recursion_limit))
    }

    fn consume(&mut self, report: table_appearance_codec::DecodeReport) -> Result<(), String> {
        self.input_bytes = checked_budget_add(
            self.input_bytes,
            report.input_bytes(),
            self.max_input_bytes,
            "input bytes",
        )?;
        self.fields = checked_budget_add(self.fields, report.fields(), self.max_fields, "fields")?;
        self.work_bytes = checked_budget_add(
            self.work_bytes,
            report.work_bytes(),
            self.max_work_bytes,
            "rewrite work",
        )?;
        if report.max_depth() as usize > self.max_nesting {
            return Err(format!(
                "Keynote table appearance nesting exceeds {} levels",
                self.max_nesting
            ));
        }
        Ok(())
    }

    fn check_style_depth(&self, depth: usize) -> Result<(), String> {
        if depth > self.max_nesting {
            return Err(format!(
                "Keynote table appearance style depth exceeds {} levels",
                self.max_nesting
            ));
        }
        Ok(())
    }
}

#[cfg(feature = "internal-iwork-source")]
fn checked_budget_add(
    current: usize,
    amount: usize,
    maximum: usize,
    axis: &str,
) -> Result<usize, String> {
    let observed = current
        .checked_add(amount)
        .ok_or_else(|| format!("Keynote table appearance {axis} budget overflow"))?;
    if observed > maximum {
        return Err(format!(
            "Keynote table appearance {axis} budget exceeded: observed {observed}, maximum {maximum}"
        ));
    }
    Ok(observed)
}

/// Decode only the style edges needed to enter the catalog appearance graph.
///
/// The host may still discover table names and dimensions through its own
/// catalog projection, but style ownership remains in this focused codec
/// owner.  The returned values are scalar facts; the generated or borrowed
/// model snapshot never crosses the seam.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __catalog_table_style_edges(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(u64, Option<u64>), String> {
    let mut budget = AppearanceBudget::from_wire_limits(limits);
    let (model, report) =
        table_appearance_codec::decode_table_model_with_report(payload, budget.options(payload))
            .map_err(|error| error.to_string())?;
    budget.consume(report)?;
    Ok((model.style_identifier(), model.style_preset_identifier()))
}

/// Resolve one table's effective appearance through a borrowed catalog.
///
/// The concrete style, preset, and style-network object IDs are accepted only
/// by this feature-gated migration seam.  The supported API resolves tables by
/// semantic selectors and returns the value types from `slide::table`.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __catalog_table_appearance<S: __CatalogTableAppearanceSource>(
    source: &mut S,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
    limits: WireLimits,
) -> Result<Appearance, String> {
    let mut budget = AppearanceBudget::from_wire_limits(limits);
    let Some(first_style_identifier) = effective_style_identifier(
        source,
        style_identifier,
        style_preset_identifier,
        &mut budget,
    )?
    else {
        return Ok(Appearance::default());
    };

    let mut visited = [0_u64; MAX_STYLE_INHERITANCE_DEPTH];
    let mut current = Some(first_style_identifier);
    let mut row_banding = None;
    let mut row_sizing = None;
    let mut body_horizontal = None;
    let mut body_vertical = None;
    let mut header_columns_horizontal = None;
    let mut header_rows_vertical = None;
    let mut footer_rows_vertical = None;

    for visited_len in 0..=MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = current else {
            return Ok(appearance_from_overrides(
                row_banding,
                row_sizing,
                body_horizontal,
                body_vertical,
                header_columns_horizontal,
                header_rows_vertical,
                footer_rows_vertical,
            ));
        };
        if visited[..visited_len].contains(&identifier) {
            return Err(format!(
                "iWork table style inheritance cycles at {identifier}"
            ));
        }
        if visited_len == visited.len() {
            return Err(format!(
                "iWork table style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
            ));
        }
        visited[visited_len] = identifier;
        budget.check_style_depth(visited_len + 1)?;

        validate_role(source, identifier, TABLE_STYLE_MESSAGE_TYPE, "table style")?;
        let (parent_identifier, overrides) = source.with_message_data_type(
            identifier,
            TABLE_STYLE_MESSAGE_TYPE,
            "TableStyleArchive",
            |payload| {
                let (style, report) = table_appearance_codec::decode_table_style_with_report(
                    payload,
                    budget.options(payload),
                )
                .map_err(|error| error.to_string())?;
                budget.consume(report)?;
                Ok((style.parent_identifier(), style.overrides()))
            },
        )?;

        row_banding = row_banding.or(overrides.row_banding);
        row_sizing = row_sizing.or(overrides.row_sizing);
        body_horizontal = body_horizontal.or(overrides.body_horizontal);
        body_vertical = body_vertical.or(overrides.body_vertical);
        header_columns_horizontal =
            header_columns_horizontal.or(overrides.header_columns_horizontal);
        header_rows_vertical = header_rows_vertical.or(overrides.header_rows_vertical);
        footer_rows_vertical = footer_rows_vertical.or(overrides.footer_rows_vertical);
        current = parent_identifier.filter(|parent| *parent != 0);
    }

    Err(format!(
        "iWork table style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    ))
}

#[cfg(feature = "internal-iwork-source")]
fn effective_style_identifier<S: __CatalogTableAppearanceSource>(
    source: &mut S,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
    budget: &mut AppearanceBudget,
) -> Result<Option<u64>, String> {
    if style_identifier != 0 {
        // A concrete model style wins over its optional preset.  This keeps a
        // malformed, unused preset from poisoning a valid direct style.
        return Ok(Some(style_identifier));
    }
    let Some(preset_identifier) = style_preset_identifier else {
        return Ok(None);
    };
    validate_role(
        source,
        preset_identifier,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        "table style preset",
    )?;
    let network_identifier = source
        .with_message_data_type(
            preset_identifier,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            "TableStylePresetArchive",
            |payload| {
                let (preset, report) =
                    table_appearance_codec::decode_table_style_preset_with_report(
                        payload,
                        budget.options(payload),
                    )
                    .map_err(|error| error.to_string())?;
                budget.consume(report)?;
                Ok(preset.style_network_identifier())
            },
        )?
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            format!("iWork table style preset {preset_identifier} has no style network")
        })?;
    validate_role(
        source,
        network_identifier,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        "table style network",
    )?;
    let style_identifier = source.with_message_data_type(
        network_identifier,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        "TableStyleNetworkArchive",
        |payload| {
            let (network, report) = table_appearance_codec::decode_table_style_network_with_report(
                payload,
                budget.options(payload),
            )
            .map_err(|error| error.to_string())?;
            budget.consume(report)?;
            Ok(network.table_style_identifier())
        },
    )?;
    if style_identifier == 0 {
        return Err(format!(
            "iWork table style network {network_identifier} has no table style"
        ));
    }
    Ok(Some(style_identifier))
}

#[cfg(feature = "internal-iwork-source")]
fn validate_role<S: __CatalogTableAppearanceSource>(
    source: &S,
    identifier: u64,
    expected_type: u32,
    expected_name: &str,
) -> Result<(), String> {
    if source.message_type_count(identifier, expected_type)? != 1 {
        return Err(format!(
            "iWork {expected_name} {identifier} must contain exactly one role payload"
        ));
    }
    for role in [
        TABLE_INFO_MESSAGE_TYPE,
        TABLE_STYLE_MESSAGE_TYPE,
        TABLE_STYLE_PRESET_MESSAGE_TYPE,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        TABLE_MODEL_MESSAGE_TYPE,
    ] {
        if role != expected_type && source.message_type_count(identifier, role)? != 0 {
            return Err(format!(
                "iWork {expected_name} {identifier} contains an appearance role alias"
            ));
        }
    }
    Ok(())
}

fn appearance_from_overrides(
    row_banding: Option<bool>,
    row_sizing: Option<bool>,
    body_horizontal: Option<bool>,
    body_vertical: Option<bool>,
    header_columns_horizontal: Option<bool>,
    header_rows_vertical: Option<bool>,
    footer_rows_vertical: Option<bool>,
) -> Appearance {
    Appearance {
        row_banding: if row_banding.unwrap_or(false) {
            Banding::Enabled
        } else {
            Banding::Disabled
        },
        row_sizing: if row_sizing.unwrap_or(false) {
            RowSizing::FitCellContents
        } else {
            RowSizing::Fixed
        },
        gridlines: Gridlines {
            body_horizontal: visibility(body_horizontal.unwrap_or(true)),
            header_columns_horizontal: visibility(header_columns_horizontal.unwrap_or(true)),
            body_vertical: visibility(body_vertical.unwrap_or(true)),
            header_rows_vertical: visibility(header_rows_vertical.unwrap_or(true)),
            footer_rows_vertical: visibility(footer_rows_vertical.unwrap_or(true)),
        },
    }
}

const fn visibility(value: bool) -> GridlineVisibility {
    if value {
        GridlineVisibility::Visible
    } else {
        GridlineVisibility::Hidden
    }
}
