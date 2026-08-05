//! Workbook relationship and name resolution for XLSB formulas.
use super::*;

#[derive(Debug, Clone, PartialEq)]
pub struct ExternalBook {
    pub(crate) metadata: Link,
}

impl ExternalBook {
    pub(crate) fn metadata(&self) -> Link {
        self.metadata.clone()
    }

    pub(crate) fn metadata_ref(&self) -> &Link {
        &self.metadata
    }
}

/// Kind of supporting link referenced by `BrtExternSheet` entries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SupportingLink {
    SelfWorkbook,
    SameSheet,
    ExternalWorkbook(u32),
    AddIn,
}

/// Workbook data required to render context-dependent XLSB formula tokens.
///
/// This context is owned once by the workbook and borrowed while worksheets
/// are decoded; it is never cloned per cell.
#[derive(Debug, Clone, Default)]
pub struct Context {
    pub(crate) worksheet_names: std::sync::Arc<[String]>,
    pub(crate) supporting_links: std::sync::Arc<[SupportingLink]>,
    pub(crate) external_sheets: std::sync::Arc<[ExternalSheet]>,
    pub(crate) external_books: std::sync::Arc<[ExternalBook]>,
    pub(crate) defined_names: std::sync::Arc<[String]>,
    pub(crate) tables: std::sync::Arc<[Definition]>,
    pub(crate) pivot_views: std::sync::Arc<[View]>,
    pub(crate) pivot_name_scopes: std::sync::Arc<[Scope]>,
    pub(crate) active_pivot_scope: Option<(u32, usize, String)>,
    pub(crate) current_sheet: Option<usize>,
}

impl Context {
    pub(crate) fn for_sheet(&self, sheet_index: usize) -> Self {
        let mut context = self.clone();
        context.current_sheet = Some(sheet_index);
        context
    }

    /// Whether an XTI resolves to exactly one worksheet in this workbook.
    pub(crate) fn is_internal_single_sheet_xti(&self, index: u16) -> bool {
        let Some(xti) = self.external_sheets.get(usize::from(index)) else {
            return false;
        };
        let Some(SupportingLink::SelfWorkbook) =
            self.supporting_links.get(xti.external_link as usize)
        else {
            return false;
        };
        xti.first_sheet >= 0
            && xti.first_sheet == xti.last_sheet
            && usize::try_from(xti.first_sheet)
                .is_ok_and(|sheet| sheet < self.worksheet_names.len())
    }

    /// Bind formula-local `BrtBeginPName` metadata to an exact PivotTable view.
    pub fn for_pivot_formula(&self, scope: Scope) -> Result<Self> {
        let mut context = self.clone();
        context.current_sheet = Some(scope.sheet_index);
        context.active_pivot_scope =
            Some((scope.cache_id, scope.sheet_index, scope.view_name.clone()));
        context.pivot_name_scopes = vec![scope].into();
        context.validate_active_pivot_scope()?;
        Ok(context)
    }

    fn resolve_table_sheet(&self, index: u16) -> Result<usize> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "structured reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti {index} refers to missing supporting link {}",
                    xti.external_link
                ))
            })?;
        match link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference Xti {index} must select exactly one worksheet"
                    )));
                }
                let sheet = usize::try_from(xti.first_sheet).map_err(|_| {
                    Error::InvalidFormula("table worksheet index overflow".to_string())
                })?;
                if sheet >= self.worksheet_names.len() {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference worksheet {} exceeds {} worksheets",
                        xti.first_sheet,
                        self.worksheet_names.len()
                    )));
                }
                Ok(sheet)
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet structured-reference Xti {index} must use -2/-2"
                    )));
                }
                self.current_sheet.ok_or_else(|| {
                    Error::InvalidFormula(
                        "same-sheet structured reference has no consuming worksheet".to_string(),
                    )
                })
            },
            SupportingLink::ExternalWorkbook(_) => Err(Error::InvalidFormula(
                "resident structured reference points to an external workbook".to_string(),
            )),
            SupportingLink::AddIn => Err(Error::InvalidFormula(
                "structured reference points to an add-in".to_string(),
            )),
        }
    }

    fn resolve_external_table_prefix(&self, index: u16) -> Result<String> {
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti {index} has no supporting link"
                ))
            })?;
        if !matches!(link, SupportingLink::ExternalWorkbook(_)) {
            return Err(Error::InvalidFormula(
                "nonresident structured reference does not point to an external workbook"
                    .to_string(),
            ));
        }
        self.resolve_sheet_prefix(index)
    }

    fn resolve_table_reference(&self, reference: &TableReference) -> Result<String> {
        if let Some(external) = &reference.external {
            if reference.row_type.is_some()
                || reference.columns.is_some()
                || reference.list_index.is_some()
            {
                return Err(Error::InvalidFormula(
                    "nonresident structured reference also contains resident metadata".to_string(),
                ));
            }
            validate_table_name(&external.table)?;
            validate_named_table_columns(&external.columns)?;
            let prefix = self.resolve_external_table_prefix(reference.sheet_index)?;
            return Ok(format!(
                "{prefix}!{}",
                format_structured_reference(
                    &external.table,
                    external.row_type,
                    &external.columns,
                    reference.square_bracket_space,
                    reference.comma_space,
                )
            ));
        }

        let table_id = reference.list_index.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits table ID".to_string())
        })?;
        let row_type = reference.row_type.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits row type".to_string())
        })?;
        let columns = reference.columns.ok_or_else(|| {
            Error::InvalidFormula("resident structured reference omits columns".to_string())
        })?;
        let sheet = self.resolve_table_sheet(reference.sheet_index)?;
        let mut matches = self
            .tables
            .iter()
            .filter(|table| table.table_id == table_id);
        let table = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing table ID {table_id}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "structured reference table ID {table_id} is ambiguous"
            )));
        }
        if table.sheet_index != sheet {
            return Err(Error::InvalidFormula(format!(
                "structured reference locates table ID {table_id} on worksheet {sheet}, but metadata places it on {}",
                table.sheet_index
            )));
        }
        let named_columns = match columns {
            TableColumns::All => TableNamedColumns::All,
            TableColumns::One(index) => {
                let name = table.columns.get(usize::from(index)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference column {index} exceeds {} columns in table {:?}",
                        table.columns.len(),
                        table.display_name
                    ))
                })?;
                TableNamedColumns::One(name.clone())
            },
            TableColumns::Range { first, last } => {
                let first_name = table.columns.get(usize::from(first)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference first column {first} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                let last_name = table.columns.get(usize::from(last)).ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "structured-reference last column {last} exceeds {} columns",
                        table.columns.len()
                    ))
                })?;
                TableNamedColumns::Range {
                    first: first_name.clone(),
                    last: last_name.clone(),
                }
            },
        };
        Ok(format_structured_reference(
            &table.display_name,
            row_type,
            &named_columns,
            reference.square_bracket_space,
            reference.comma_space,
        ))
    }

    fn resolve_sheet_prefix(&self, index: u16) -> Result<String> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "3D reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "Xti index {index} exceeds {} extern-sheet entries",
                    self.external_sheets.len()
                ))
            })?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow".to_string()))?;
        let supporting_link = self.supporting_links.get(link_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {index} refers to missing supporting link {}",
                xti.external_link
            ))
        })?;
        let (first_index, last_index) = match supporting_link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "Xti index {index} has invalid self-reference sheet range {}..={}",
                        xti.first_sheet, xti.last_sheet
                    )));
                }
                (
                    usize::try_from(xti.first_sheet).map_err(|_| {
                        Error::InvalidFormula("first sheet index overflow".to_string())
                    })?,
                    usize::try_from(xti.last_sheet).map_err(|_| {
                        Error::InvalidFormula("last sheet index overflow".to_string())
                    })?,
                )
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet Xti index {index} must use workbook scope -2/-2"
                    )));
                }
                let sheet = self.current_sheet.ok_or_else(|| {
                    Error::UnsupportedFeature(
                        "same-sheet reference requires a consuming worksheet".to_string(),
                    )
                })?;
                (sheet, sheet)
            },
            SupportingLink::ExternalWorkbook(book_index) => {
                return self.resolve_external_sheet_prefix(index, xti, *book_index);
            },
            SupportingLink::AddIn => {
                return Err(Error::UnsupportedFeature(format!(
                    "Xti index {index} refers to an add-in"
                )));
            },
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula(format!(
                "Xti index {index} has invalid sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first = self.worksheet_names.get(first_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti first sheet {} exceeds {} worksheets",
                xti.first_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let last = self.worksheet_names.get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti last sheet {} exceeds {} worksheets",
                xti.last_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let unquoted = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&unquoted))
    }

    fn resolve_external_sheet_prefix(
        &self,
        xti_index: u16,
        xti: &ExternalSheet,
        book_index: u32,
    ) -> Result<String> {
        let book_index = usize::try_from(book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow".to_string()))?;
        let book = self.external_books.get(book_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {xti_index} refers to missing external book {book_index}"
            ))
        })?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(format!(
                "Xti index {xti_index} refers to a DDE or OLE data source"
            )));
        }
        if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
            return Err(Error::InvalidFormula(format!(
                "Xti index {xti_index} has invalid external sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first_index = usize::try_from(xti.first_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let last_index = usize::try_from(xti.last_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let first = book
            .metadata
            .sheet_names()
            .get(first_index)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external sheet {} exceeds {} cached names",
                    xti.first_sheet,
                    book.metadata.sheet_names().len()
                ))
            })?;
        let last = book.metadata.sheet_names().get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external sheet {} exceeds {} cached names",
                xti.last_sheet,
                book.metadata.sheet_names().len()
            ))
        })?;
        let sheets = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&format!(
            "[{}]{sheets}",
            book.metadata.source()
        )))
    }

    fn resolve_external_name(&self, xti_index: u16, name_index: u32) -> Result<String> {
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgNameX name index is one-based and cannot be zero".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(xti_index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "PtgNameX Xti index {xti_index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow".to_string()))?;
        let SupportingLink::ExternalWorkbook(book_index) =
            self.supporting_links.get(link_index).ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "PtgNameX refers to missing supporting link {}",
                    xti.external_link
                ))
            })?
        else {
            return Err(Error::InvalidFormula(
                "PtgNameX does not refer to an external workbook".to_string(),
            ));
        };
        let external_book_index = usize::try_from(*book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow".to_string()))?;
        let book = self
            .external_books
            .get(external_book_index)
            .ok_or_else(|| Error::InvalidFormula(format!("missing external book {book_index}")))?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(
                "PtgNameX refers to a DDE or OLE data source".to_string(),
            ));
        }
        let index = usize::try_from(name_index - 1)
            .map_err(|_| Error::InvalidFormula("external name index overflow".to_string()))?;
        let names = book.metadata.defined_names();
        let name = names.get(index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external name index {name_index} exceeds {} names",
                names.len()
            ))
        })?;
        Ok(format!(
            "{}!{}",
            format_formula_prefix(&format!("[{}]", book.metadata.source())),
            name.name()
        ))
    }
}

fn owner_formula_resolution<T>(result: Result<T>) -> crate::formula::Result<T> {
    result.map_err(|error| match error {
        Error::InvalidFormula(message) => crate::formula::Error::InvalidFormula(message),
        Error::InvalidCellReference(reference) => {
            crate::formula::Error::InvalidCellReference(reference)
        },
        Error::InvalidLength { expected, found } => {
            crate::formula::Error::InvalidLength { expected, found }
        },
        Error::UnsupportedFeature(feature) => crate::formula::Error::UnsupportedFeature(feature),
        Error::Encoding(message) => crate::formula::Error::Encoding(message),
        error => crate::formula::Error::InvalidFormula(error.to_string()),
    })
}

impl Resolution for Context {
    fn sheet_prefix(&self, index: u16) -> crate::formula::Result<String> {
        owner_formula_resolution(self.resolve_sheet_prefix(index))
    }

    fn defined_name(&self, index: u32) -> crate::formula::Result<String> {
        let index = usize::try_from(index)
            .ok()
            .and_then(|index| index.checked_sub(1))
            .ok_or_else(|| {
                crate::formula::Error::InvalidFormula(
                    "PtgName index is one-based and cannot be zero".to_string(),
                )
            })?;
        self.defined_names.get(index).cloned().ok_or_else(|| {
            crate::formula::Error::InvalidFormula(format!(
                "PtgName index {} exceeds {} workbook names",
                index + 1,
                self.defined_names.len()
            ))
        })
    }

    fn external_name(&self, sheet_index: u16, name_index: u32) -> crate::formula::Result<String> {
        owner_formula_resolution(self.resolve_external_name(sheet_index, name_index))
    }

    fn table_reference(&self, reference: &TableReference) -> crate::formula::Result<String> {
        owner_formula_resolution(self.resolve_table_reference(reference))
    }

    fn pivot_name(&self, index: u32) -> crate::formula::Result<String> {
        owner_formula_resolution(self.resolve_pivot_name(index))
    }
}

fn format_formula_prefix(value: &str) -> String {
    if value
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.'))
        && !value.chars().next().is_some_and(|ch| ch.is_ascii_digit())
    {
        value.to_string()
    } else {
        format!("'{}'", value.replace('\'', "''"))
    }
}

impl Context {
    fn validate_active_pivot_scope(&self) -> Result<&Scope> {
        let (cache_id, sheet_index, view_name) =
            self.active_pivot_scope.as_ref().ok_or_else(|| {
                Error::InvalidFormula(
                    "PtgSxName requires an explicit pivot cache, sheet, and view scope".to_string(),
                )
            })?;
        if *sheet_index >= self.worksheet_names.len() {
            return Err(Error::InvalidFormula(format!(
                "pivot sheet index {sheet_index} is outside the workbook sheet range"
            )));
        }
        if self.current_sheet != Some(*sheet_index) {
            return Err(Error::InvalidFormula(format!(
                "pivot scope sheet {sheet_index} does not match the formula sheet {:?}",
                self.current_sheet
            )));
        }

        let mut views = self.pivot_views.iter().filter(|view| {
            view.cache_id == *cache_id
                && view.sheet_index == *sheet_index
                && view.name.eq_ignore_ascii_case(view_name)
        });
        let _view = views.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} does not use cache {cache_id}"
            ))
        })?;
        if views.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "PivotTable view {view_name:?} on sheet {sheet_index} and cache {cache_id} is ambiguous"
            )));
        }

        let mut scopes = self.pivot_name_scopes.iter().filter(|scope| {
            scope.cache_id == *cache_id
                && scope.sheet_index == *sheet_index
                && scope.view_name.eq_ignore_ascii_case(view_name)
        });
        let scope = scopes.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "calculated-name metadata is missing for PivotTable view {view_name:?}"
            ))
        })?;
        if scopes.next().is_some() {
            return Err(Error::InvalidFormula(format!(
                "calculated-name metadata for PivotTable view {view_name:?} is ambiguous"
            )));
        }
        Ok(scope)
    }

    fn resolve_pivot_name(&self, index: u32) -> Result<String> {
        let scope = self.validate_active_pivot_scope()?;
        let index = usize::try_from(index).map_err(|_| {
            Error::InvalidFormula("pivot calculated-name index overflow".to_string())
        })?;
        let reference = scope.references.get(index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "pivot calculated-name index {index} is outside 0..{}",
                scope.references.len()
            ))
        })?;
        Ok(reference.to_formula_text())
    }
}
