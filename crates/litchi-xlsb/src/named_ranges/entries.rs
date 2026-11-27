//! Defined names as a workbook stores them, with rendered formulas.

use crate::formula::{Compiler, Parser};
use crate::package::error::{Error, Result};
use crate::package::formula::Context;
use crate::raw::{Limits, Records, kind};

/// One defined name of a workbook: its scope and its formula as text.
///
/// ```
/// use litchi_xlsb::builder::{self, Builder};
/// use litchi_xlsb::named_ranges::{Definition, area3d_formula};
///
/// let mut builder = Builder::new();
/// builder.add_worksheet(builder::Worksheet::new("Data"));
/// builder.add_named_range(
///     Definition::new("Totals".to_owned(), None)
///         .with_formula(area3d_formula(0, 0, 4, 1, 1)?),
/// );
/// let package = builder.build()?;
/// let names = package.workbook()?.defined_name_entries()?;
/// assert_eq!(names[0].name, "Totals");
/// assert_eq!(names[0].formula.as_deref(), Some("Data!$B$1:$B$5"));
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct Entry {
    /// The defined name.
    pub name: String,
    /// The complete-catalog position of the scope sheet; `None` for
    /// workbook scope.
    pub scope: Option<usize>,
    /// Whether the name is hidden.
    pub hidden: bool,
    /// The formula as text without a leading `=`, when the name has one that
    /// renders against the workbook's sheet and name tables.
    pub formula: Option<String>,
}

/// Read every `BrtName` record of a workbook stream in record order.
pub(crate) fn read(stream: &[u8], context: &Context, limits: Limits) -> Result<Vec<Entry>> {
    let mut entries = Vec::new();
    for record in Records::new(stream, limits) {
        let record = record?;
        if record.kind() != kind::NAME {
            continue;
        }
        let definition = super::Definition::parse(record.payload(), limits)?;
        let scope = definition
            .sheet_id
            .map(|id| {
                usize::try_from(id)
                    .map_err(|_error| Error::InvalidFormat("name scope exceeds this platform"))
            })
            .transpose()?;
        let formula = definition.formula.as_deref().and_then(|rgce| {
            let tokens = Parser::new(rgce, limits).parse().ok()?;
            let context = match scope {
                Some(sheet) => context.for_sheet(sheet),
                None => context.clone(),
            };
            Compiler::try_tokens_to_string_with_resolution(&tokens, &context).ok()
        });
        entries.try_reserve(1).map_err(|source| Error::Allocation {
            resource: "XLSB defined-name entries",
            source,
        })?;
        entries.push(Entry {
            name: definition.name,
            scope,
            hidden: definition.hidden,
            formula,
        });
    }
    Ok(entries)
}
