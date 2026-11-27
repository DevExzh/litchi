//! Logical retained-size accounting for decoded XLSB worksheets.
//!
//! The worksheet cache uses this estimate as an admission weight.  It is a
//! deterministic accounting of the worksheet value and its owned children,
//! rather than a measurement of allocator or resident memory.  All arithmetic
//! saturates so a hostile model cannot wrap the cache's accounting.

use std::mem::size_of;

use crate::conditional_formatting::{Formatting, Rule, Scale, Value};
use crate::data_validation::Validation;
use crate::formula::ParsedFormula;
use crate::package::cell::Cell;
use crate::package::scenarios::Manager;
use crate::package::web_extension_bindings::Binding;
use crate::sheet::Worksheet;
use crate::{slicer, timeline};

/// Approximate allocation overhead retained by an `Arc<Worksheet>`.
const ARC_HEADER_BYTES: usize = 2 * size_of::<usize>();

/// The standard library's B-tree node reserves eleven key/value slots.  The
/// worksheet map is sparse, so each occupied entry is charged a complete
/// node-width slot budget rather than only its live pair.
const BTREE_NODE_SLOT_CAPACITY: usize = 11;
/// Fixed pointer/length bookkeeping words charged with each conservatively
/// accounted node-width budget above.
const BTREE_NODE_METADATA_WORDS: usize = 16;

/// The saturating accumulator used by the worksheet cache.
#[derive(Clone, Copy)]
struct Weight(usize);

impl Weight {
    const fn new(initial: usize) -> Self {
        Self(initial)
    }

    fn add(&mut self, bytes: usize) {
        self.0 = self.0.saturating_add(bytes);
    }

    fn mul_add(&mut self, count: usize, bytes: usize) {
        self.add(count.saturating_mul(bytes));
    }

    fn vector<T>(&mut self, values: &Vec<T>) {
        self.mul_add(values.capacity(), size_of::<T>());
    }

    fn string(&mut self, text: &String) {
        self.add(text.capacity());
    }

    fn optional_string(&mut self, text: &Option<String>) {
        if let Some(text) = text {
            self.string(text);
        }
    }

    fn formula(&mut self, formula: &ParsedFormula) {
        self.add(formula.rgce.capacity());
        self.add(formula.rgcb.capacity());
    }

    const fn total(self) -> usize {
        self.0
    }
}

impl Worksheet {
    /// Return the logical retained size of this decoded worksheet.
    ///
    /// The estimate includes the worksheet allocation, sparse-map entries,
    /// cell-owned formula/value buffers, every worksheet record vector, and
    /// the dynamically sized children of views, validations, formatting,
    /// scenarios, slicers, timelines, sparklines, comments, and repairs.
    /// This is a saturated logical weight, not an allocator measurement.
    pub(crate) fn retained_bytes(&self) -> usize {
        self.retained_bytes_within(usize::MAX).unwrap_or(usize::MAX)
    }

    /// Compute the retained size while allowing a cache to stop once a
    /// configured admission ceiling is crossed.
    pub(crate) fn retained_bytes_within(&self, limit: usize) -> Option<usize> {
        let mut weight = Weight::new(ARC_HEADER_BYTES.saturating_add(size_of::<Self>()));
        weight.string(&self.name);

        // A BTreeMap node stores keys and values in partially occupied arrays;
        // charge every reserved key/value slot and the node's fixed
        // bookkeeping words for each occupied entry.  This intentionally
        // overcounts a multi-entry node, while covering a one-cell sparse
        // node's eleven reserved slots.
        let map_entry = size_of::<((u32, u32), Cell)>()
            .saturating_mul(BTREE_NODE_SLOT_CAPACITY)
            .saturating_add(size_of::<usize>().saturating_mul(BTREE_NODE_METADATA_WORDS));
        weight.mul_add(self.cells.len(), map_entry);
        if weight.total() > limit {
            return None;
        }
        for cell in self.cells.values() {
            weight.add(cell.heap_weight());
            if weight.total() > limit {
                return None;
            }
        }

        weight.vector(&self.merged_cells);
        weight.vector(&self.hyperlinks);
        for hyperlink in &self.hyperlinks {
            weight.string(&hyperlink.r_id);
            for text in [
                &hyperlink.location,
                &hyperlink.tooltip,
                &hyperlink.display,
                &hyperlink.target,
            ]
            .into_iter()
            .flatten()
            {
                weight.string(text);
            }
        }

        weight.vector(&self.comments);
        for comment in &self.comments {
            weight.string(&comment.author);
            weight.string(&comment.text);
            weight.vector(&comment.runs);
        }

        weight.vector(&self.column_infos);
        weight.vector(&self.row_infos);
        for row in &self.row_infos {
            weight.vector(&row.column_spans);
        }

        if let Some(protection) = &self.strong_sheet_protection {
            weight.vector(&protection.hash);
            weight.vector(&protection.salt);
            weight.string(&protection.algorithm);
        }

        weight.vector(&self.data_validations);
        for validation in &self.data_validations {
            validation_weight(&mut weight, validation);
        }

        weight.vector(&self.conditional_formattings);
        for formatting in &self.conditional_formattings {
            formatting_weight(&mut weight, formatting);
        }

        weight.vector(&self.web_extension_bindings);
        for binding in &self.web_extension_bindings {
            binding_weight(&mut weight, binding);
        }

        weight.vector(&self.views);
        for view in &self.views {
            // `selections` is public; each selection lends its range slice,
            // whose vector capacity is private to the format-neutral model.
            weight.vector(&view.selections);
            for selection in &view.selections {
                weight.add(selection.retained_bytes());
            }
        }

        if let Some(scenarios) = &self.scenarios {
            scenarios_weight(&mut weight, scenarios);
        }
        if let Some(slicers) = &self.slicers {
            slicers_weight(&mut weight, slicers);
        }
        if let Some(timelines) = &self.timelines {
            timelines_weight(&mut weight, timelines);
        }
        if let Some(groups) = &self.sparkline_groups {
            sparklines_weight(&mut weight, groups);
        }

        weight.add(self.tolerance.retained_bytes());

        (weight.total() <= limit).then_some(weight.total())
    }
}

fn validation_weight(weight: &mut Weight, validation: &Validation) {
    for text in [
        &validation.formula1,
        &validation.formula2,
        &validation.list_formula,
        &validation.input_title,
        &validation.input_text,
        &validation.error_title,
        &validation.error_text,
    ] {
        weight.optional_string(text);
    }
    for formula in [&validation.formula1_binary, &validation.formula2_binary]
        .into_iter()
        .flatten()
    {
        weight.formula(formula);
    }
    weight.string(&validation.cell_ranges);
}

fn binding_weight(weight: &mut Weight, binding: &Binding) {
    weight.string(&binding.application_reference);
    weight.formula(&binding.formula);
}

fn formatting_weight(weight: &mut Weight, formatting: &Formatting) {
    weight.vector(&formatting.ranges);
    for range in &formatting.ranges {
        weight.string(range);
    }
    weight.vector(&formatting.rules);
    for rule in &formatting.rules {
        rule_weight(weight, rule);
    }
}

fn rule_weight(weight: &mut Weight, rule: &Rule) {
    for buffers in [&rule.formulas, &rule.formula_extras] {
        weight.vector(buffers);
        for buffer in buffers {
            weight.vector(buffer);
        }
    }
    weight.vector(&rule.formula_texts);
    for text in &rule.formula_texts {
        weight.string(text);
    }
    weight.optional_string(&rule.text);
    if let Some(scale) = &rule.color_scale {
        scale_weight(weight, scale);
    }
    if let Some(bar) = &rule.data_bar {
        value_weight(weight, &bar.min_cfvo);
        value_weight(weight, &bar.max_cfvo);
    }
    if let Some(icon_set) = &rule.icon_set {
        values_weight(weight, &icon_set.cfvos);
    }
    if let Some(scale) = &rule.color_scale14 {
        scale_weight(weight, scale);
    }
    if let Some(bar) = &rule.data_bar14 {
        value_weight(weight, &bar.min_cfvo);
        value_weight(weight, &bar.max_cfvo);
    }
    if let Some(icon_set) = &rule.icon_set14 {
        values_weight(weight, &icon_set.cfvos);
        if let Some(icons) = &icon_set.custom_icons {
            weight.vector(icons);
        }
    }
}

fn scale_weight(weight: &mut Weight, scale: &Scale) {
    value_weight(weight, &scale.min_cfvo);
    value_weight(weight, &scale.max_cfvo);
    if let Some(value) = &scale.mid_cfvo {
        value_weight(weight, value);
    }
}

fn values_weight(weight: &mut Weight, values: &Vec<Value>) {
    weight.vector(values);
    for value in values {
        value_weight(weight, value);
    }
}

fn value_weight(weight: &mut Weight, value: &Value) {
    weight.optional_string(&value.value);
    if let Some(formula) = &value.formula_binary {
        weight.formula(formula);
    }
}

fn scenarios_weight(weight: &mut Weight, manager: &Manager) {
    weight.add(manager.retained_bytes());
}

fn slicers_weight(weight: &mut Weight, views: &slicer::Views) {
    weight.vector(&views.items);
    for view in &views.items {
        weight.string(&view.name);
        weight.string(&view.cache_name);
        if let Some(caption) = &view.caption {
            weight.string(caption);
        }
        if let Some(style) = &view.style {
            weight.string(style);
        }
    }
}

fn timelines_weight(weight: &mut Weight, views: &timeline::Views) {
    weight.vector(&views.items);
    for view in &views.items {
        weight.string(&view.name);
        weight.string(&view.cache);
        for text in [&view.caption, &view.scroll_position, &view.style, &view.uid]
            .into_iter()
            .flatten()
        {
            weight.string(text);
        }
    }
}

fn sparklines_weight(weight: &mut Weight, groups: &crate::sparkline::Groups) {
    weight.add(groups.retained_bytes());
}

#[cfg(test)]
mod tests {
    use super::{BTREE_NODE_METADATA_WORDS, BTREE_NODE_SLOT_CAPACITY};
    use crate::package::cell::Cell;
    use crate::sheet::Worksheet;
    use litchi_core::sheet::CellValue;
    use std::mem::size_of;

    #[test]
    fn retained_bytes_charges_spare_cell_string_capacity() {
        let mut spare = String::from("x");
        spare.reserve_exact(1024 * 1024);
        let spare_capacity = spare.capacity();
        let compact_capacity = String::from("x").capacity();

        let mut compact = Worksheet::new(String::from("sheet"));
        compact.add_cell(Cell::new(0, 0, CellValue::String(String::from("x"))));
        let mut padded = Worksheet::new(String::from("sheet"));
        padded.add_cell(Cell::new(0, 0, CellValue::String(spare)));

        let compact_weight = compact.retained_bytes();
        let padded_weight = padded.retained_bytes();
        assert!(padded_weight > compact_weight);
        assert!(
            padded_weight.saturating_sub(compact_weight)
                >= spare_capacity.saturating_sub(compact_capacity)
        );
        assert!(
            padded
                .retained_bytes_within(padded_weight.saturating_sub(1))
                .is_none()
        );
    }

    #[test]
    fn retained_bytes_charges_sparse_btree_node_slots() {
        let empty = Worksheet::new(String::from("sheet"));
        let mut one_cell = Worksheet::new(String::from("sheet"));
        one_cell.add_cell(Cell::new(0, 0, CellValue::Int(1)));

        let delta = one_cell
            .retained_bytes()
            .saturating_sub(empty.retained_bytes());
        let minimum = size_of::<((u32, u32), Cell)>()
            .saturating_mul(BTREE_NODE_SLOT_CAPACITY)
            .saturating_add(size_of::<usize>().saturating_mul(BTREE_NODE_METADATA_WORDS));
        assert!(delta >= minimum);
    }
}
