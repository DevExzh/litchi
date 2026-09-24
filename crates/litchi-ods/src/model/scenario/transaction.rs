//! Source-bound, inert transactions for ODF `table:scenario` metadata.

use super::{
    Error as ScenarioError, Limits, MAX_INPUT_BYTES, MAX_TEXT_BYTES, OptionalSetting, Scenario,
    Snapshot, xml_text_is_valid,
};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::Event,
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{collections::HashMap, sync::Arc};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_SPANS: usize = 1_048_576;

/// A staged, source-checked scenario edit.
#[derive(Clone, Debug)]
pub struct Edit {
    before: Snapshot,
    draft: Vec<Scenario>,
    draft_aggregate_bytes: usize,
}

impl Edit {
    pub(crate) fn new(before: Snapshot) -> Self {
        Self {
            draft: before.scenarios.clone(),
            draft_aggregate_bytes: before.aggregate_bytes,
            before,
        }
    }

    /// Borrow the staged scenario declarations in source order.
    #[must_use]
    pub fn scenarios(&self) -> &[Scenario] {
        &self.draft
    }

    /// Borrow the immutable source snapshot from which this edit was staged.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Find one staged scenario by its exact worksheet name.
    #[must_use]
    pub fn for_sheet(&self, sheet_name: &str) -> Option<&Scenario> {
        let mut matches = self
            .draft
            .iter()
            .filter(|scenario| scenario.sheet() == sheet_name);
        let first = matches.next();
        first.filter(|_| matches.next().is_none())
    }

    /// Replace the complete ordered scenario catalog.
    pub fn replace(&mut self, scenarios: Vec<Scenario>) -> Result<()> {
        let aggregate_bytes = validate_scenarios(&scenarios, self.before.limits)?;
        self.draft = scenarios;
        self.draft_aggregate_bytes = aggregate_bytes;
        Ok(())
    }

    /// Add one scenario to the staged catalog.
    pub fn add(&mut self, scenario: Scenario) -> Result<()> {
        let limits = self.before.limits;
        if self.draft.len() >= limits.scenarios {
            return Err(resource_limit(
                "scenarios",
                self.draft.len() + 1,
                limits.scenarios,
            ));
        }
        scenario.validate().map_err(model_error)?;
        let scenario_bytes = scenario.aggregate_bytes().map_err(model_error)?;
        let aggregate_bytes = self
            .draft_aggregate_bytes
            .checked_add(scenario_bytes)
            .ok_or_else(|| invalid("ODS scenario aggregate size overflows"))?;
        if aggregate_bytes > limits.aggregate_bytes {
            return Err(resource_limit(
                "aggregate scenario bytes",
                aggregate_bytes,
                limits.aggregate_bytes,
            ));
        }
        if self
            .draft
            .iter()
            .any(|candidate| candidate.sheet() == scenario.sheet())
        {
            return Err(invalid(format!(
                "duplicate ODS scenario worksheet '{}', exact worksheet selectors are required",
                scenario.sheet()
            )));
        }
        self.draft
            .try_reserve(1)
            .map_err(|_| invalid("ODS scenario catalog allocation failed"))?;
        self.draft.push(scenario);
        self.draft_aggregate_bytes = aggregate_bytes;
        Ok(())
    }

    /// Replace a scenario selected by its source-order index.
    pub fn replace_at(&mut self, index: usize, scenario: Scenario) -> Result<()> {
        let limits = self.before.limits;
        let previous = self
            .draft
            .get(index)
            .ok_or_else(|| invalid(format!("ODS scenario index {index} did not match")))?;
        scenario.validate().map_err(model_error)?;
        if self
            .draft
            .iter()
            .enumerate()
            .any(|(candidate_index, candidate)| {
                candidate_index != index && candidate.sheet() == scenario.sheet()
            })
        {
            return Err(invalid(format!(
                "duplicate ODS scenario worksheet '{}', exact worksheet selectors are required",
                scenario.sheet()
            )));
        }
        let previous_bytes = previous.aggregate_bytes().map_err(model_error)?;
        let replacement_bytes = scenario.aggregate_bytes().map_err(model_error)?;
        let aggregate_bytes = self
            .draft_aggregate_bytes
            .checked_sub(previous_bytes)
            .and_then(|value| value.checked_add(replacement_bytes))
            .ok_or_else(|| invalid("ODS scenario aggregate size overflows"))?;
        if aggregate_bytes > limits.aggregate_bytes {
            return Err(resource_limit(
                "aggregate scenario bytes",
                aggregate_bytes,
                limits.aggregate_bytes,
            ));
        }
        let slot = self
            .draft
            .get_mut(index)
            .ok_or_else(|| invalid(format!("ODS scenario index {index} did not match")))?;
        *slot = scenario;
        self.draft_aggregate_bytes = aggregate_bytes;
        Ok(())
    }

    /// Replace a scenario selected by its exact worksheet name.
    pub fn replace_sheet(&mut self, sheet_name: &str, scenario: Scenario) -> Result<()> {
        let mut matches = self
            .draft
            .iter()
            .enumerate()
            .filter(|(_, candidate)| candidate.sheet() == sheet_name);
        let Some((index, _)) = matches.next() else {
            return Err(invalid(format!(
                "ODS scenario for worksheet '{sheet_name}' was not found"
            )));
        };
        if matches.next().is_some() {
            return Err(invalid(format!(
                "ODS scenario worksheet '{sheet_name}' is ambiguous"
            )));
        }
        self.replace_at(index, scenario)
    }

    /// Alias for [`Self::replace_sheet`] using the source vocabulary's
    /// name-oriented selector terminology.
    pub fn replace_named(&mut self, sheet_name: &str, scenario: Scenario) -> Result<()> {
        self.replace_sheet(sheet_name, scenario)
    }

    /// Remove one scenario selected by its source-order index.
    pub fn remove_at(&mut self, index: usize) -> Result<Scenario> {
        let existing = self
            .draft
            .get(index)
            .ok_or_else(|| invalid(format!("ODS scenario index {index} did not match")))?;
        let removed_bytes = existing.aggregate_bytes().map_err(model_error)?;
        let removed = self.draft.remove(index);
        self.draft_aggregate_bytes = self
            .draft_aggregate_bytes
            .checked_sub(removed_bytes)
            .ok_or_else(|| invalid("ODS scenario aggregate size underflows"))?;
        Ok(removed)
    }

    /// Remove one scenario selected by its exact worksheet name.
    pub fn remove_sheet(&mut self, sheet_name: &str) -> Result<Scenario> {
        let mut matches = self
            .draft
            .iter()
            .enumerate()
            .filter(|(_, candidate)| candidate.sheet() == sheet_name);
        let Some((index, _)) = matches.next() else {
            return Err(invalid(format!(
                "ODS scenario for worksheet '{sheet_name}' was not found"
            )));
        };
        if matches.next().is_some() {
            return Err(invalid(format!(
                "ODS scenario worksheet '{sheet_name}' is ambiguous"
            )));
        }
        self.remove_at(index)
    }

    /// Alias for [`Self::remove_sheet`] using the source vocabulary's
    /// name-oriented selector terminology.
    pub fn remove_named(&mut self, sheet_name: &str) -> Result<Scenario> {
        self.remove_sheet(sheet_name)
    }

    /// Remove all scenario declarations while retaining the source document.
    pub fn clear(&mut self) -> Result<()> {
        self.replace(Vec::new())
    }

    /// Validate and materialize this edit as a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        if self.draft == self.before.scenarios {
            let source = Arc::clone(&self.before.content);
            return Ok(Commit {
                snapshot: self.before,
                patch: Patch {
                    source: Arc::clone(&source),
                    target: source,
                },
                changed: false,
            });
        }
        validate_scenarios(&self.draft, self.before.limits)?;

        let target_xml = render_source(
            &self.before.content,
            &self.before.scenarios,
            &self.draft,
            self.before.limits.input_bytes,
        )?;
        let target = Snapshot::parse_with(&target_xml, self.before.limits).map_err(model_error)?;
        if target.scenarios != self.draft {
            return Err(invalid(
                "ODS scenario commit failed typed readback or changed document order",
            ));
        }
        let source = Arc::clone(&self.before.content);
        let target_source: Arc<str> = Arc::from(target_xml);
        Ok(Commit {
            snapshot: target,
            patch: Patch {
                source,
                target: Arc::clone(&target_source),
            },
            changed: true,
        })
    }
}

/// An exact-source, reversible `content.xml` scenario patch.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source: Arc<str>,
    target: Arc<str>,
}

impl Patch {
    /// Whether this patch leaves `content.xml` byte-for-byte unchanged.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.source.as_ref() == self.target.as_ref()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
        }
    }

    /// Apply this patch only to the exact source snapshot from which it came.
    pub fn apply(&self, snapshot: &Snapshot) -> Result<Commit> {
        if snapshot.content.as_ref() != self.source.as_ref() {
            return Err(invalid("ODS scenario patch source snapshot does not match"));
        }
        let target =
            Snapshot::parse_with(self.target.as_ref(), snapshot.limits).map_err(model_error)?;
        Ok(Commit {
            snapshot: target,
            patch: self.clone(),
            changed: !self.is_empty(),
        })
    }
}

/// A fully rehydrated scenario commit.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the package content changed.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the resulting scenario snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible content patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }
}

fn validate_scenarios(scenarios: &[Scenario], limits: Limits) -> Result<usize> {
    if scenarios.len() > limits.scenarios {
        return Err(resource_limit(
            "scenarios",
            scenarios.len(),
            limits.scenarios,
        ));
    }
    let mut names = HashMap::with_capacity(scenarios.len());
    let mut aggregate_bytes = 0usize;
    for scenario in scenarios {
        scenario.validate().map_err(model_error)?;
        aggregate_bytes = aggregate_bytes
            .checked_add(scenario.aggregate_bytes().map_err(model_error)?)
            .ok_or_else(|| invalid("ODS scenario aggregate size overflows"))?;
        if aggregate_bytes > limits.aggregate_bytes {
            return Err(resource_limit(
                "aggregate scenario bytes",
                aggregate_bytes,
                limits.aggregate_bytes,
            ));
        }
        if names.insert(scenario.sheet(), ()).is_some() {
            return Err(invalid(format!(
                "duplicate ODS scenario worksheet '{}', exact worksheet selectors are required",
                scenario.sheet()
            )));
        }
    }
    Ok(aggregate_bytes)
}

#[derive(Clone, Debug)]
struct Span {
    namespace: Option<String>,
    local: String,
    qname: String,
    start: usize,
    tag_end: usize,
    close_start: usize,
    end: usize,
    parent: Option<usize>,
    empty: bool,
    opaque_attributes: bool,
    opaque_content: bool,
    table_name: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct EditRange {
    start: usize,
    end: usize,
}

#[derive(Clone, Debug)]
struct TableSite {
    span: usize,
    name: Option<String>,
    scenario: Option<usize>,
}

#[derive(Clone, Copy, Debug)]
enum ReplacementKind {
    Remove,
    Scenario { draft: usize, table: usize },
    Add { draft: usize, table: usize },
}

#[derive(Clone, Copy, Debug)]
struct Replacement {
    range: EditRange,
    kind: ReplacementKind,
}

fn render_source(
    source: &str,
    original: &[Scenario],
    draft: &[Scenario],
    max_output_bytes: usize,
) -> Result<String> {
    let output_limit = max_output_bytes.min(MAX_INPUT_BYTES);
    if source.len() > output_limit {
        return Err(invalid("ODS content.xml exceeds the scenario input limit"));
    }
    let spans = scan(source)?;
    let spreadsheet = spans
        .iter()
        .position(|span| span.namespace.as_deref() == Some(OFFICE) && span.local == "spreadsheet")
        .ok_or_else(|| invalid("ODS content.xml has no office:spreadsheet element"))?;
    let mut tables = spans
        .iter()
        .enumerate()
        .filter(|(_, span)| {
            span.namespace.as_deref() == Some(TABLE)
                && span.local == "table"
                && span.parent == Some(spreadsheet)
        })
        .map(|(span, value)| TableSite {
            span,
            name: value.table_name.clone(),
            scenario: None,
        })
        .collect::<Vec<_>>();

    let mut table_by_span = HashMap::with_capacity(tables.len());
    for (index, table) in tables.iter().enumerate() {
        table_by_span.insert(table.span, index);
    }
    let mut scenario_spans = Vec::new();
    for (index, span) in spans.iter().enumerate() {
        if span.namespace.as_deref() != Some(TABLE) || span.local != "scenario" {
            continue;
        }
        let Some(parent) = span.parent.and_then(|parent| table_by_span.get(&parent)) else {
            continue;
        };
        if tables[*parent].scenario.replace(index).is_some() {
            return Err(invalid("ODS table contains more than one table:scenario"));
        }
        scenario_spans.push(index);
    }
    if scenario_spans.len() != original.len() {
        return Err(invalid(
            "ODS scenario source catalog changed before the transaction committed",
        ));
    }
    for (index, scenario_span) in scenario_spans.iter().copied().enumerate() {
        let parent = spans[scenario_span]
            .parent
            .and_then(|parent| table_by_span.get(&parent))
            .copied()
            .ok_or_else(|| invalid("ODS scenario table parent disappeared"))?;
        if tables[parent].name.as_deref() != Some(original[index].sheet()) {
            return Err(invalid(
                "ODS scenario source catalog no longer matches its worksheet selectors",
            ));
        }
    }

    let mut replacements = Vec::<Replacement>::new();
    let mut original_names = HashMap::with_capacity(original.len());
    for (index, scenario) in original.iter().enumerate() {
        if original_names.insert(scenario.sheet(), index).is_some() {
            return Err(invalid("ODS scenario source has duplicate worksheet names"));
        }
        let target = draft
            .iter()
            .enumerate()
            .find(|(_, candidate)| candidate.sheet() == scenario.sheet());
        match target {
            Some((draft_index, candidate)) if candidate != scenario => {
                let span = scenario_spans[index];
                reject_opaque(&spans, span)?;
                let table = spans[span]
                    .parent
                    .and_then(|parent| table_by_span.get(&parent))
                    .copied()
                    .ok_or_else(|| invalid("ODS scenario has no table parent"))?;
                replacements.push(Replacement {
                    range: EditRange {
                        start: spans[span].start,
                        end: spans[span].end,
                    },
                    kind: ReplacementKind::Scenario {
                        draft: draft_index,
                        table: tables[table].span,
                    },
                });
            },
            Some(_) => {},
            None => {
                let span = scenario_spans[index];
                reject_opaque(&spans, span)?;
                replacements.push(Replacement {
                    range: EditRange {
                        start: spans[span].start,
                        end: spans[span].end,
                    },
                    kind: ReplacementKind::Remove,
                });
            },
        }
    }

    let mut draft_names = HashMap::with_capacity(draft.len());
    for (draft_index, scenario) in draft.iter().enumerate() {
        if draft_names.insert(scenario.sheet(), ()).is_some() {
            return Err(invalid(
                "ODS scenario worksheet selector catalog is not unique",
            ));
        }
        if original_names.contains_key(scenario.sheet()) {
            continue;
        }
        let matching_tables = tables
            .iter()
            .filter(|table| table.name.as_deref() == Some(scenario.sheet()))
            .collect::<Vec<_>>();
        let table = match matching_tables.as_slice() {
            [table] => *table,
            [] => {
                return Err(invalid(format!(
                    "ODS scenario worksheet '{}' was not found",
                    scenario.sheet()
                )));
            },
            _ => {
                return Err(invalid(format!(
                    "ODS scenario worksheet '{}' is ambiguous",
                    scenario.sheet()
                )));
            },
        };
        if table.scenario.is_some() {
            return Err(invalid(format!(
                "ODS worksheet '{}' already has a scenario",
                scenario.sheet()
            )));
        }
        let table_span = &spans[table.span];
        let range = if table_span.empty {
            EditRange {
                start: table_span.start,
                end: table_span.end,
            }
        } else {
            let insertion = direct_structural_child(&spans, table.span)
                .map(|child| spans[child].start)
                .unwrap_or(table_span.close_start);
            EditRange {
                start: insertion,
                end: insertion,
            }
        };
        replacements.push(Replacement {
            range,
            kind: ReplacementKind::Add {
                draft: draft_index,
                table: table.span,
            },
        });
    }

    if replacements.is_empty() {
        return copy_source(source, output_limit);
    }
    replacements.sort_unstable_by(|left, right| {
        left.range
            .start
            .cmp(&right.range.start)
            .then(left.range.end.cmp(&right.range.end))
    });
    let mut previous: Option<EditRange> = None;
    for replacement in &replacements {
        let range = replacement.range;
        if range.end < range.start || range.end > source.len() {
            return Err(invalid("ODS scenario replacement span is invalid"));
        }
        if let Some(previous) = previous {
            if range.start < previous.end || range.start == previous.start {
                return Err(invalid(
                    "ODS scenario replacement spans overlap or share an insertion point",
                ));
            }
        }
        previous = Some(range);
    }

    let mut output_len = source.len();
    for replacement in &replacements {
        let replacement_len = replacement_length(source, &spans, draft, *replacement)?;
        output_len = output_len
            .checked_sub(replacement.range.end - replacement.range.start)
            .and_then(|value| value.checked_add(replacement_len))
            .ok_or_else(|| invalid("patched ODS content.xml size overflows"))?;
        if output_len > output_limit {
            return Err(invalid(
                "patched ODS content.xml exceeds the scenario input limit",
            ));
        }
    }

    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| invalid("patched ODS content.xml allocation failed"))?;
    let mut cursor = 0usize;
    for replacement in &replacements {
        output.push_str(
            source
                .get(cursor..replacement.range.start)
                .ok_or_else(|| invalid("ODS scenario replacement span is invalid"))?,
        );
        append_replacement(&mut output, source, &spans, draft, *replacement)?;
        cursor = replacement.range.end;
    }
    output.push_str(
        source
            .get(cursor..)
            .ok_or_else(|| invalid("ODS scenario replacement span is invalid"))?,
    );
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn copy_source(source: &str, output_limit: usize) -> Result<String> {
    if source.len() > output_limit {
        return Err(invalid("ODS content.xml exceeds the scenario input limit"));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| invalid("ODS content.xml allocation failed"))?;
    output.push_str(source);
    Ok(output)
}

fn replacement_length(
    source: &str,
    spans: &[Span],
    draft: &[Scenario],
    replacement: Replacement,
) -> Result<usize> {
    match replacement.kind {
        ReplacementKind::Remove => Ok(0),
        ReplacementKind::Scenario {
            draft: draft_index,
            table,
        } => scenario_xml_length(&draft[draft_index], element_prefix(&spans[table].qname)),
        ReplacementKind::Add {
            draft: draft_index,
            table,
        } => {
            let span = &spans[table];
            let fragment = scenario_xml_length(&draft[draft_index], element_prefix(&span.qname))?;
            if span.empty {
                let opening = table_opening(source, span)?;
                checked_sum(&[opening.len(), 1, fragment, 2, span.qname.len(), 1])
            } else {
                Ok(fragment)
            }
        },
    }
}

fn append_replacement(
    output: &mut String,
    source: &str,
    spans: &[Span],
    draft: &[Scenario],
    replacement: Replacement,
) -> Result<()> {
    match replacement.kind {
        ReplacementKind::Remove => Ok(()),
        ReplacementKind::Scenario {
            draft: draft_index,
            table,
        } => {
            append_scenario_xml(
                output,
                &draft[draft_index],
                element_prefix(&spans[table].qname),
            );
            Ok(())
        },
        ReplacementKind::Add {
            draft: draft_index,
            table,
        } => {
            let span = &spans[table];
            let prefix = element_prefix(&span.qname);
            if span.empty {
                output.push_str(table_opening(source, span)?);
                output.push('>');
                append_scenario_xml(output, &draft[draft_index], prefix);
                output.push_str("</");
                output.push_str(&span.qname);
                output.push('>');
            } else {
                append_scenario_xml(output, &draft[draft_index], prefix);
            }
            Ok(())
        },
    }
}

fn table_opening<'a>(source: &'a str, span: &Span) -> Result<&'a str> {
    source
        .get(span.start..span.tag_end)
        .ok_or_else(|| invalid("ODS table opening span is invalid"))?
        .trim_end()
        .strip_suffix("/>")
        .ok_or_else(|| invalid("ODS self-closing table span is invalid"))
}

fn scenario_xml_length(scenario: &Scenario, source_prefix: &str) -> Result<usize> {
    let prefix = if source_prefix.is_empty() {
        "table"
    } else {
        source_prefix
    };
    let mut length = 0usize;
    add_len(&mut length, 1 + prefix.len() + ":scenario".len())?;
    if source_prefix.is_empty() {
        add_len(&mut length, " xmlns:table=\"".len() + TABLE.len() + 1)?;
    }
    add_len(&mut length, 1 + prefix.len() + ":scenario-ranges=\"".len())?;
    for (index, range) in scenario.ranges().iter().enumerate() {
        add_len(&mut length, usize::from(index != 0))?;
        add_len(&mut length, escaped_xml_length(range.as_str())?)?;
    }
    add_len(&mut length, 1 + 1 + prefix.len() + ":is-active=\"".len())?;
    let active_length = if scenario.is_active() { 4 } else { 5 };
    add_len(&mut length, active_length + 1)?;
    add_optional_bool_length(
        &mut length,
        prefix,
        "display-border",
        scenario.display_border(),
    )?;
    if scenario.border_color().is_some() {
        add_len(
            &mut length,
            1 + prefix.len() + ":border-color=\"".len() + 7 + 1,
        )?;
    }
    add_optional_bool_length(&mut length, prefix, "copy-back", scenario.copy_back())?;
    add_optional_bool_length(&mut length, prefix, "copy-styles", scenario.copy_styles())?;
    add_optional_bool_length(
        &mut length,
        prefix,
        "copy-formulas",
        scenario.copy_formulas(),
    )?;
    if let Some(comment) = scenario.comment() {
        add_len(
            &mut length,
            1 + prefix.len() + ":comment=\"".len() + escaped_xml_length(comment)? + 1,
        )?;
    }
    add_optional_bool_length(&mut length, prefix, "protected", scenario.protected())?;
    add_len(&mut length, 2)?;
    Ok(length)
}

fn add_optional_bool_length(
    length: &mut usize,
    prefix: &str,
    name: &str,
    value: OptionalSetting,
) -> Result<()> {
    let value_length = match value {
        OptionalSetting::Unspecified => return Ok(()),
        OptionalSetting::Enabled => 4,
        OptionalSetting::Disabled => 5,
    };
    add_len(
        length,
        1 + prefix.len() + 1 + name.len() + 2 + value_length + 1,
    )
}

fn escaped_xml_length(value: &str) -> Result<usize> {
    let mut length = 0usize;
    for byte in value.bytes() {
        add_len(
            &mut length,
            match byte {
                b'&' => 5,
                b'<' | b'>' => 4,
                b'"' | b'\'' => 6,
                _ => 1,
            },
        )?;
    }
    Ok(length)
}

fn append_scenario_xml(output: &mut String, scenario: &Scenario, source_prefix: &str) {
    let prefix = if source_prefix.is_empty() {
        "table"
    } else {
        source_prefix
    };
    output.push('<');
    output.push_str(prefix);
    output.push_str(":scenario");
    if source_prefix.is_empty() {
        output.push_str(" xmlns:table=\"");
        output.push_str(TABLE);
        output.push('"');
    }
    output.push(' ');
    output.push_str(prefix);
    output.push_str(":scenario-ranges=\"");
    for (index, range) in scenario.ranges().iter().enumerate() {
        if index != 0 {
            output.push(' ');
        }
        append_escaped_xml(output, range.as_str());
    }
    output.push_str("\" ");
    output.push_str(prefix);
    output.push_str(":is-active=\"");
    output.push_str(if scenario.is_active() {
        "true"
    } else {
        "false"
    });
    output.push('"');
    append_optional_bool_attr(output, prefix, "display-border", scenario.display_border());
    if let Some(color) = scenario.border_color() {
        output.push(' ');
        output.push_str(prefix);
        output.push_str(":border-color=\"");
        append_color(output, color);
        output.push('"');
    }
    append_optional_bool_attr(output, prefix, "copy-back", scenario.copy_back());
    append_optional_bool_attr(output, prefix, "copy-styles", scenario.copy_styles());
    append_optional_bool_attr(output, prefix, "copy-formulas", scenario.copy_formulas());
    if let Some(comment) = scenario.comment() {
        output.push(' ');
        output.push_str(prefix);
        output.push_str(":comment=\"");
        append_escaped_xml(output, comment);
        output.push('"');
    }
    append_optional_bool_attr(output, prefix, "protected", scenario.protected());
    output.push_str("/>");
}

fn append_optional_bool_attr(
    output: &mut String,
    prefix: &str,
    name: &str,
    value: OptionalSetting,
) {
    let value = match value {
        OptionalSetting::Unspecified => return,
        OptionalSetting::Enabled => "true",
        OptionalSetting::Disabled => "false",
    };
    output.push(' ');
    output.push_str(prefix);
    output.push(':');
    output.push_str(name);
    output.push_str("=\"");
    output.push_str(value);
    output.push('"');
}

fn append_color(output: &mut String, color: super::RgbColor) {
    const HEX: &[u8; 16] = b"0123456789ABCDEF";
    output.push('#');
    for component in [color.red(), color.green(), color.blue()] {
        output.push(HEX[usize::from(component >> 4)] as char);
        output.push(HEX[usize::from(component & 0x0F)] as char);
    }
}

fn append_escaped_xml(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            '\'' => output.push_str("&apos;"),
            character => output.push(character),
        }
    }
}

fn checked_sum(values: &[usize]) -> Result<usize> {
    let mut total = 0usize;
    for value in values {
        add_len(&mut total, *value)?;
    }
    Ok(total)
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| invalid("ODS scenario output size overflows"))?;
    Ok(())
}

fn direct_structural_child(spans: &[Span], table: usize) -> Option<usize> {
    spans.iter().enumerate().find_map(|(index, span)| {
        (span.parent == Some(table) && !is_table_preamble(span)).then_some(index)
    })
}

fn is_table_preamble(span: &Span) -> bool {
    matches!(
        (span.namespace.as_deref(), span.local.as_str()),
        (Some(TABLE), "title" | "desc" | "table-source") | (Some(OFFICE), "dde-source")
    )
}

fn reject_opaque(spans: &[Span], target: usize) -> Result<()> {
    if spans
        .get(target)
        .is_some_and(|span| span.opaque_attributes || span.opaque_content)
    {
        return Err(invalid(
            "changed ODS scenario contains opaque extension markup",
        ));
    }
    for (index, span) in spans.iter().enumerate() {
        if index == target || !is_descendant(spans, index, target) {
            continue;
        }
        if span.namespace.as_deref() != Some(TABLE) || span.opaque_attributes || span.opaque_content
        {
            return Err(invalid(
                "changed ODS scenario contains opaque extension markup",
            ));
        }
    }
    Ok(())
}

fn is_descendant(spans: &[Span], candidate: usize, ancestor: usize) -> bool {
    let mut parent = spans.get(candidate).and_then(|span| span.parent);
    while let Some(index) = parent {
        if index == ancestor {
            return true;
        }
        parent = spans.get(index).and_then(|span| span.parent);
    }
    false
}

fn element_prefix(qname: &str) -> &str {
    qname.split_once(':').map_or("", |(prefix, _)| prefix)
}

fn scan(xml: &str) -> Result<Vec<Span>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut spans = Vec::<Span>::new();
    let mut open = Vec::<usize>::new();
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("ODS scenario XML position overflows usize"))?;
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid ODS scenario XML: {error}")))?;
        let namespace = resolve_namespace(&resolved)?;
        let event = event.into_owned();
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("ODS scenario XML position overflows usize"))?;
        match event {
            Event::Start(element) => {
                if spans.len() >= MAX_SPANS {
                    return Err(invalid("ODS scenario XML exceeds the element limit"));
                }
                let index = spans.len();
                spans.push(Span {
                    namespace: namespace.clone(),
                    local: decode(element.local_name().as_ref(), "scenario element local name")?,
                    qname: decode(element.name().as_ref(), "scenario element qualified name")?,
                    start,
                    tag_end: end,
                    close_start: end,
                    end,
                    parent: open.last().copied(),
                    empty: false,
                    opaque_attributes: has_opaque_attributes(&reader, &element)?,
                    opaque_content: false,
                    table_name: if namespace.as_deref() == Some(TABLE)
                        && element.local_name().as_ref() == b"table"
                    {
                        table_name(&reader, &element)?
                    } else {
                        None
                    },
                });
                open.push(index);
            },
            Event::Empty(element) => {
                if spans.len() >= MAX_SPANS {
                    return Err(invalid("ODS scenario XML exceeds the element limit"));
                }
                spans.push(Span {
                    namespace: namespace.clone(),
                    local: decode(element.local_name().as_ref(), "scenario element local name")?,
                    qname: decode(element.name().as_ref(), "scenario element qualified name")?,
                    start,
                    tag_end: end,
                    close_start: end,
                    end,
                    parent: open.last().copied(),
                    empty: true,
                    opaque_attributes: has_opaque_attributes(&reader, &element)?,
                    opaque_content: false,
                    table_name: if namespace.as_deref() == Some(TABLE)
                        && element.local_name().as_ref() == b"table"
                    {
                        table_name(&reader, &element)?
                    } else {
                        None
                    },
                });
            },
            Event::End(_) => {
                let index = open
                    .pop()
                    .ok_or_else(|| invalid("ODS scenario XML span underflow"))?;
                spans[index].close_start = start;
                spans[index].end = end;
            },
            Event::Comment(_) => {
                if let Some(parent) = open.last().copied() {
                    spans[parent].opaque_content = true;
                }
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {
                if let Some(parent) = open.last().copied()
                    && spans[parent].namespace.as_deref() == Some(TABLE)
                    && spans[parent].local == "scenario"
                {
                    spans[parent].opaque_content = true;
                }
            },
            Event::DocType(_) => return Err(invalid("DTD content is not accepted")),
            Event::Eof => break,
            Event::PI(_) => {
                if let Some(parent) = open.last().copied()
                    && spans[parent].namespace.as_deref() == Some(TABLE)
                    && spans[parent].local == "scenario"
                {
                    spans[parent].opaque_content = true;
                }
            },
            Event::Decl(_) => {},
        }
        buffer.clear();
    }
    if !open.is_empty() {
        return Err(invalid("ODS scenario XML ended with open elements"));
    }
    Ok(spans)
}

fn has_opaque_attributes(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<bool> {
    let is_scenario = element.local_name().as_ref() == b"scenario";
    let mut opaque = false;
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute
            .map_err(|error| invalid(format!("invalid ODS scenario attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        match resolved {
            ResolveResult::Bound(Namespace(uri)) => {
                let uri = std::str::from_utf8(uri)
                    .map_err(|_| invalid("ODS scenario attribute namespace is not UTF-8"))?;
                let allowed = uri == TABLE
                    && (!is_scenario
                        || matches!(
                            name.as_ref(),
                            b"scenario-ranges"
                                | b"is-active"
                                | b"display-border"
                                | b"border-color"
                                | b"copy-back"
                                | b"copy-styles"
                                | b"copy-formulas"
                                | b"comment"
                                | b"protected"
                        ));
                if !allowed {
                    opaque = true;
                }
            },
            ResolveResult::Unbound => opaque = true,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "unbound ODS scenario attribute namespace prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        }
    }
    Ok(opaque)
}

fn table_name(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<Option<String>> {
    for raw_attribute in element.attributes().with_checks(true) {
        let attribute = raw_attribute
            .map_err(|error| invalid(format!("invalid ODS table attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == TABLE.as_bytes())
            && name.as_ref() == b"name"
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("invalid ODS table name: {error}")))?
                .into_owned();
            if value.is_empty() || value.len() > MAX_TEXT_BYTES || !xml_text_is_valid(&value) {
                return Err(invalid("invalid or oversized ODS table name"));
            }
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn resolve_namespace(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Ok(Some(
            std::str::from_utf8(uri)
                .map_err(|_| invalid("ODS scenario namespace is not UTF-8"))?
                .to_owned(),
        )),
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound ODS scenario namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn decode(value: &[u8], label: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_| invalid(format!("{label} is not UTF-8")))
}

fn model_error(error: ScenarioError) -> Error {
    Error::InvalidFormat(format!("ODS scenario metadata: {error}"))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn resource_limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::InvalidFormat(format!(
        "ODS {resource} limit exceeded: observed {actual}, maximum {maximum}"
    ))
}
