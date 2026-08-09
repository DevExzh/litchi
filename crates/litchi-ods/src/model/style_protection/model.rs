//! Typed ODF table-cell protection vocabulary and registry storage.

use crate::model::names::formula;
use litchi_core::{Error, Result, xml::escape_xml};
use std::collections::{BTreeMap, HashMap, HashSet};

pub(super) const MAX_CONDITIONAL_STYLES: usize = 65_536;
pub(super) const MAX_RULES_PER_STYLE: usize = 1_024;
pub(super) const MAX_CONDITIONAL_RULES: usize = 262_144;
pub(super) const MAX_CONDITIONAL_ATTRIBUTE_BYTES: usize = 64 * 1024;
pub(super) const MAX_CONDITIONAL_TEXT_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_STYLE_DEPTH: usize = 64;

/// One standard ODF conditional table-cell style.
///
/// Rules are retained in document order. Litchi does not evaluate their
/// conditions or compute an effective style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ConditionalStyle {
    pub style_name: String,
    pub parent_style_name: Option<String>,
    pub rules: Vec<Rule>,
}

/// One inert `style:map` rule belonging to a table-cell style.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Rule {
    /// The decoded condition text. It is never evaluated by litchi.
    pub condition: String,
    /// Namespace bound to a condition prefix, when the condition is qualified.
    pub formula_namespace: Option<formula::Namespace>,
    /// Name of the common table-cell style referenced by the rule.
    pub apply_style_name: String,
    /// Optional lexical base cell address for relative formula references.
    pub base_cell_address: Option<String>,
}

impl ConditionalStyle {
    /// Create an inert conditional table-cell style.
    pub fn new(style_name: impl Into<String>, rules: Vec<Rule>) -> Self {
        Self {
            style_name: style_name.into(),
            parent_style_name: None,
            rules,
        }
    }

    /// Set the optional parent table-cell style name.
    pub fn with_parent_style_name(mut self, parent_style_name: impl Into<String>) -> Self {
        self.parent_style_name = Some(parent_style_name.into());
        self
    }
}

impl Rule {
    /// Create an inert conditional rule without a formula namespace or base address.
    pub fn new(condition: impl Into<String>, apply_style_name: impl Into<String>) -> Self {
        Self {
            condition: condition.into(),
            formula_namespace: None,
            apply_style_name: apply_style_name.into(),
            base_cell_address: None,
        }
    }

    /// Bind the lexical condition prefix to a namespace URI.
    pub fn with_formula_namespace(mut self, namespace: formula::Namespace) -> Self {
        self.formula_namespace = Some(namespace);
        self
    }

    /// Set the optional lexical base cell address.
    pub fn with_base_cell_address(mut self, address: impl Into<String>) -> Self {
        self.base_cell_address = Some(address.into());
        self
    }
}

/// The effective value of the ODF `style:cell-protect` property.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Protection {
    None,
    Protected,
    FormulaHidden,
    ProtectedFormulaHidden,
    HiddenAndProtected,
}

/// One automatic table-cell style with an explicit `style:cell-protect` value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableStyle {
    pub style_name: String,
    pub parent_style_name: Option<String>,
    pub protection: Protection,
}

impl TableStyle {
    pub fn new(style_name: impl Into<String>, protection: Protection) -> Self {
        Self {
            style_name: style_name.into(),
            parent_style_name: None,
            protection,
        }
    }

    pub fn with_parent_style_name(mut self, parent_style_name: impl Into<String>) -> Self {
        self.parent_style_name = Some(parent_style_name.into());
        self
    }
}

impl Protection {
    pub(super) fn parse(value: &str) -> Result<Self> {
        match value {
            "none" => Ok(Self::None),
            "protected" => Ok(Self::Protected),
            "formula-hidden" => Ok(Self::FormulaHidden),
            "protected formula-hidden" | "formula-hidden protected" => {
                Ok(Self::ProtectedFormulaHidden)
            },
            "hidden-and-protected" => Ok(Self::HiddenAndProtected),
            _ => Err(Error::InvalidFormat(format!(
                "invalid style:cell-protect value '{value}'"
            ))),
        }
    }

    /// Whether editing is prohibited while the containing sheet is protected.
    pub fn is_protected(self) -> bool {
        matches!(
            self,
            Self::Protected | Self::ProtectedFormulaHidden | Self::HiddenAndProtected
        )
    }

    /// Whether the cell formula is hidden while protection is active.
    pub fn hides_formula(self) -> bool {
        matches!(self, Self::FormulaHidden | Self::ProtectedFormulaHidden)
    }

    /// Whether both cell content and editing are hidden/protected.
    pub fn hides_content(self) -> bool {
        self == Self::HiddenAndProtected
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Protected => "protected",
            Self::FormulaHidden => "formula-hidden",
            Self::ProtectedFormulaHidden => "protected formula-hidden",
            Self::HiddenAndProtected => "hidden-and-protected",
        }
    }
}

#[derive(Clone, Debug, Default)]
pub struct CellStyleRegistry {
    pub(super) styles: HashMap<String, CellStyleDefinition>,
    pub(super) default: Option<Protection>,
    pub(super) common_table_cell_styles: HashSet<String>,
    pub(super) style_order: Vec<String>,
    pub(super) conditional_styles: Vec<ConditionalStyle>,
    pub(super) conditional_style_index: HashMap<String, usize>,
    pub(super) automatic_protection_styles: Vec<TableStyle>,
    pub(super) parsed_conditional_styles: usize,
    pub(super) parsed_conditional_rules: usize,
    pub(super) conditional_text_bytes: usize,
}

#[derive(Clone, Debug)]
pub struct PreservedXmlFragment {
    pub xml: String,
    pub(super) namespaces: BTreeMap<String, String>,
}

impl PreservedXmlFragment {
    pub(crate) fn namespace_prefixes(&self) -> impl Iterator<Item = &str> {
        self.namespaces.keys().map(String::as_str)
    }

    pub(crate) fn write_missing_namespaces<'a>(
        &self,
        out: &mut String,
        already_declared: impl IntoIterator<Item = &'a str>,
    ) {
        let declared = already_declared.into_iter().collect::<HashSet<_>>();
        for (prefix, uri) in &self.namespaces {
            if declared.contains(prefix.as_str()) {
                continue;
            }
            out.push_str(" xmlns:");
            out.push_str(prefix);
            out.push_str("=\"");
            out.push_str(&escape_xml(uri));
            out.push('"');
        }
    }
}

#[derive(Clone, Debug)]
pub(super) struct CellStyleDefinition {
    pub(super) parent: Option<String>,
    pub(super) protection: Option<Protection>,
    pub(super) conditional_rules: Vec<Rule>,
    pub(super) is_common: bool,
}

pub(super) struct StyleBuilder {
    pub(super) name: Option<String>,
    pub(super) parent: Option<String>,
    pub(super) protection: Option<Protection>,
    pub(super) conditional_rules: Vec<Rule>,
    pub(super) is_default: bool,
    pub(super) is_common: bool,
}
