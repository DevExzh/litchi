//! PivotTable formula metadata.
use super::*;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct View {
    pub(super) cache_id: u32,
    pub(super) sheet_index: usize,
    pub(super) name: String,
}

impl View {
    pub fn try_new(cache_id: u32, sheet_index: usize, name: String) -> Result<Self> {
        validate_pivot_identifier(&name, "PivotTable view name", 255)?;
        Ok(Self {
            cache_id,
            sheet_index,
            name,
        })
    }

    pub fn cache_id(&self) -> u32 {
        self.cache_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn name(&self) -> &str {
        &self.name
    }
}

/// Aggregation encoded by `BrtBeginPName.ifn` for a calculated-field reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Aggregation {
    Sum,
    CountA,
    Average,
    Max,
    Min,
    Product,
    Count,
    StandardDeviation,
    PopulationStandardDeviation,
    Variance,
    PopulationVariance,
}

impl Aggregation {
    fn formula_name(self) -> &'static str {
        match self {
            Self::Sum => "SUM",
            Self::CountA => "COUNTA",
            Self::Average => "AVERAGE",
            Self::Max => "MAX",
            Self::Min => "MIN",
            Self::Product => "PRODUCT",
            Self::Count => "COUNT",
            Self::StandardDeviation => "STDEV",
            Self::PopulationStandardDeviation => "STDEVP",
            Self::Variance => "VAR",
            Self::PopulationVariance => "VARP",
        }
    }
}

/// A calculated-item position encoded by a `BrtBeginPNPair` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Item {
    Name(String),
    AbsolutePosition(u32),
    RelativePosition(i32),
}

/// The formula text represented by one `BrtBeginPName` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Name {
    Field {
        name: String,
        aggregation: Option<Aggregation>,
    },
    Item {
        field_name: String,
        item: Item,
    },
}

impl Name {
    fn validate(&self) -> Result<()> {
        match self {
            Self::Field { name, .. } => {
                validate_pivot_identifier(name, "pivot cache field name", 32_767)
            },
            Self::Item { field_name, item } => {
                validate_pivot_identifier(field_name, "pivot item field name", 32_767)?;
                match item {
                    Item::Name(name) => validate_pivot_identifier(name, "pivot item name", 32_767),
                    Item::AbsolutePosition(position) => {
                        if *position == 0 || *position > i32::MAX as u32 {
                            return Err(invalid(
                                "PtgSxName",
                                format!(
                                    "absolute pivot item position {position} is outside 1..={}",
                                    i32::MAX
                                ),
                            ));
                        }
                        Ok(())
                    },
                    Item::RelativePosition(position) => {
                        if *position == 0 {
                            return Err(invalid(
                                "PtgSxName",
                                "relative pivot item position must not be zero",
                            ));
                        }
                        Ok(())
                    },
                }
            },
        }
    }

    pub(super) fn to_formula_text(&self) -> String {
        match self {
            Self::Field { name, aggregation } => {
                let name = format_pivot_identifier(name);
                match aggregation {
                    Some(aggregation) => format!("{}({name})", aggregation.formula_name()),
                    None => name,
                }
            },
            Self::Item { field_name, item } => {
                let field_name = format_pivot_identifier(field_name);
                let item = match item {
                    Item::Name(name) => format_pivot_identifier(name),
                    Item::AbsolutePosition(position) => position.to_string(),
                    Item::RelativePosition(position) if *position > 0 => {
                        format!("+{position}")
                    },
                    Item::RelativePosition(position) => position.to_string(),
                };
                format!("{field_name}[{item}]")
            },
        }
    }
}

/// Formula-local `BrtBeginPName` collection for one PivotTable view.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Scope {
    pub(super) cache_id: u32,
    pub(super) sheet_index: usize,
    pub(super) view_name: String,
    pub(super) references: std::sync::Arc<[Name]>,
}

impl Scope {
    pub fn try_new(
        cache_id: u32,
        sheet_index: usize,
        view_name: String,
        references: Vec<Name>,
    ) -> Result<Self> {
        validate_pivot_identifier(&view_name, "PivotTable view name", 255)?;
        if references.len() > 16_384 {
            return Err(invalid(
                "BrtBeginPNames",
                format!(
                    "pivot calculated-name count {} exceeds 16384",
                    references.len()
                ),
            ));
        }
        for reference in &references {
            reference.validate()?;
        }
        Ok(Self {
            cache_id,
            sheet_index,
            view_name,
            references: references.into(),
        })
    }

    pub fn cache_id(&self) -> u32 {
        self.cache_id
    }

    pub fn sheet_index(&self) -> usize {
        self.sheet_index
    }

    pub fn view_name(&self) -> &str {
        &self.view_name
    }

    pub fn references(&self) -> &[Name] {
        &self.references
    }
}
