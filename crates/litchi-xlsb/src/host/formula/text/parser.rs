#![allow(
    clippy::expect_used,
    clippy::map_err_ignore,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error to this codec boundary"
)]

//! Formula-text parser and context resolver.

use super::super::{
    ArrayValue, BinaryOperator, Definition, Error, ExternalTableReference, Result, SupportingLink,
    TableColumns, TableDataType, TableNamedColumns, TableReference, TableRowType, UnaryOperator,
    validate_named_table_columns, validate_table_name,
};
use super::FORMULA_ERRORS;
use super::ast::{CompileExpr, ParsedStructuredReference, StructuredReferenceItem};
use super::builtin::builtin_function_by_name;
use super::compiler::Compiler;
use super::references::parse_a1_reference;
use super::validation::{excel_name_eq, formula_error_code};

impl<'a> Compiler<'a> {
    pub(super) fn parse_comparison(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_concat()?;
        loop {
            let operator = if self.consume("<>") {
                Some(BinaryOperator::NotEqual)
            } else if self.consume("<=") {
                Some(BinaryOperator::LessEqual)
            } else if self.consume(">=") {
                Some(BinaryOperator::GreaterEqual)
            } else if self.consume("=") {
                Some(BinaryOperator::Equal)
            } else if self.consume("<") {
                Some(BinaryOperator::LessThan)
            } else if self.consume(">") {
                Some(BinaryOperator::GreaterThan)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_concat()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_concat(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_additive()?;
        while self.consume("&") {
            let right = self.parse_additive()?;
            expression = CompileExpr::Binary(
                BinaryOperator::Concat,
                Box::new(expression),
                Box::new(right),
            );
        }
        Ok(expression)
    }

    fn parse_additive(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_multiplicative()?;
        loop {
            let operator = if self.consume("+") {
                Some(BinaryOperator::Add)
            } else if self.consume("-") {
                Some(BinaryOperator::Subtract)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_multiplicative()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_multiplicative(&mut self) -> Result<CompileExpr> {
        let mut expression = self.parse_power()?;
        loop {
            let operator = if self.consume("*") {
                Some(BinaryOperator::Multiply)
            } else if self.consume("/") {
                Some(BinaryOperator::Divide)
            } else {
                None
            };
            let Some(operator) = operator else { break };
            let right = self.parse_power()?;
            expression = CompileExpr::Binary(operator, Box::new(expression), Box::new(right));
        }
        Ok(expression)
    }

    fn parse_power(&mut self) -> Result<CompileExpr> {
        let left = self.parse_unary()?;
        if self.consume("^") {
            let right = self.parse_power()?;
            Ok(CompileExpr::Binary(
                BinaryOperator::Power,
                Box::new(left),
                Box::new(right),
            ))
        } else {
            Ok(left)
        }
    }

    fn parse_unary(&mut self) -> Result<CompileExpr> {
        if self.consume("+") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Plus,
                Box::new(self.parse_unary()?),
            ));
        }
        if self.consume("-") {
            return Ok(CompileExpr::Unary(
                UnaryOperator::Minus,
                Box::new(self.parse_unary()?),
            ));
        }
        let mut expression = self.parse_primary()?;
        while self.consume("%") {
            expression = CompileExpr::Unary(UnaryOperator::Percent, Box::new(expression));
        }
        Ok(expression)
    }

    fn parse_primary(&mut self) -> Result<CompileExpr> {
        self.skip_spaces();
        if self.consume("(") {
            let expression = self.parse_comparison()?;
            if !self.consume(")") {
                return Err(self.error("expected ')'"));
            }
            return Ok(CompileExpr::Parenthesized(Box::new(expression)));
        }
        if self.consume("{") {
            return self.parse_array_constant();
        }
        if self.peek_char() == Some('"') {
            return self.parse_string().map(CompileExpr::String);
        }
        if self.peek_char() == Some('#') {
            return self.parse_error_literal().map(CompileExpr::Error);
        }
        if self
            .peek_char()
            .is_some_and(|ch| ch.is_ascii_digit() || ch == '.')
        {
            return self.parse_number().map(CompileExpr::Number);
        }

        if self.peek_char() == Some('\'') {
            let sheet_qualifier = self.parse_quoted_sheet_name()?;
            if !self.consume("!") {
                return Err(self.error("expected '!' after quoted worksheet name"));
            }
            if sheet_qualifier.starts_with('[') {
                let table = self.parse_identifier()?;
                if self.peek_char() == Some('[') || parse_a1_reference(&table).is_none() {
                    let selection = if self.peek_char() == Some('[') {
                        self.parse_structured_reference()?
                    } else {
                        ParsedStructuredReference {
                            row_type: TableRowType::Data,
                            columns: TableNamedColumns::All,
                            square_bracket_space: false,
                            comma_space: false,
                        }
                    };
                    return self.compile_external_table_reference(
                        &sheet_qualifier,
                        table,
                        selection,
                    );
                }
                return Err(Error::UnsupportedFeature(
                    "external cell references are not supported by this compilation context"
                        .to_string(),
                ));
            }
            let (first_sheet, last_sheet) = Self::split_sheet_qualifier(&sheet_qualifier)?;
            return self.parse_qualified_reference(first_sheet, last_sheet);
        }

        let identifier = self.parse_identifier()?;
        let sheet_range_checkpoint = self.offset;
        if self.consume(":") {
            let last_sheet = self.parse_identifier()?;
            if self.consume("!") {
                return self.parse_qualified_reference(&identifier, Some(&last_sheet));
            }
            self.offset = sheet_range_checkpoint;
        }
        if self.consume("!") {
            return self.parse_qualified_reference(&identifier, None);
        }
        if self.peek_char() == Some('[') {
            let selection = self.parse_structured_reference()?;
            return self.compile_resident_table_reference(&identifier, selection);
        }
        if self.consume("(") {
            let function = builtin_function_by_name(&identifier).ok_or_else(|| {
                Error::UnsupportedFeature(format!(
                    "XLSB formula function {identifier} is not in the supported Ftab set"
                ))
            })?;
            let mut arguments = Vec::new();
            if !self.consume(")") {
                loop {
                    if self.consume(")") {
                        arguments.push(CompileExpr::MissingArg);
                        break;
                    }
                    if self.consume(",") {
                        arguments.push(CompileExpr::MissingArg);
                        continue;
                    }
                    arguments.push(self.parse_comparison()?);
                    if self.consume(")") {
                        break;
                    }
                    if !self.consume(",") {
                        return Err(self.error("expected ',' or ')' in function call"));
                    }
                }
            }
            let argument_count = u8::try_from(arguments.len()).map_err(|_| {
                Error::InvalidFormula(format!("{} has more than 255 arguments", function.name))
            })?;
            if !function.accepts_arg_count(argument_count) {
                return Err(Error::InvalidFormula(format!(
                    "{} does not accept {} arguments (range {}..={})",
                    function.name,
                    arguments.len(),
                    function.min_args,
                    function.max_args,
                )));
            }
            return Ok(CompileExpr::Function(function, arguments));
        }
        if identifier.eq_ignore_ascii_case("TRUE") {
            return Ok(CompileExpr::Bool(true));
        }
        if identifier.eq_ignore_ascii_case("FALSE") {
            return Ok(CompileExpr::Bool(false));
        }

        if let Some(reference) = self.compile_bare_resident_table_reference(&identifier)? {
            return Ok(reference);
        }

        let Some(first) = parse_a1_reference(&identifier) else {
            return self
                .resolve_defined_name(&identifier)
                .map(CompileExpr::Name);
        };
        if self.consume(":") {
            let second_text = self.parse_identifier()?;
            let second = parse_a1_reference(&second_text)
                .ok_or_else(|| self.error("invalid range end reference"))?;
            Ok(CompileExpr::Area(first, second))
        } else {
            Ok(CompileExpr::Ref(first))
        }
    }

    fn parse_structured_reference(&mut self) -> Result<ParsedStructuredReference> {
        debug_assert_eq!(self.peek_char(), Some('['));
        self.offset += 1;
        let leading_space = self.consume_structured_space()?;
        let nested = self.peek_char() == Some('[');
        let mut items = Vec::new();
        let mut separators = Vec::new();
        let mut comma_space = None;
        let mut unwrapped_trailing_space = false;

        if nested {
            loop {
                items.push(self.parse_structured_reference_item(true)?);
                match self.peek_char() {
                    Some(',') => {
                        self.offset += 1;
                        separators.push(',');
                        let spaced = self.consume_structured_space()?;
                        if comma_space
                            .replace(spaced)
                            .is_some_and(|previous| previous != spaced)
                        {
                            return Err(self
                                .error("structured-reference commas use inconsistent whitespace"));
                        }
                    },
                    Some(':') => {
                        self.offset += 1;
                        separators.push(':');
                    },
                    _ => break,
                }
            }
        } else {
            let mut item = self.parse_structured_reference_item(false)?;
            if item.text.ends_with(char::is_whitespace) {
                if !item.text.ends_with(' ')
                    || item
                        .text
                        .strip_suffix(' ')
                        .is_some_and(|text| text.ends_with(char::is_whitespace))
                {
                    return Err(self.error(
                        "structured-reference whitespace cannot be represented by XLSB flags",
                    ));
                }
                item.text.pop();
                if item.text.is_empty() {
                    return Err(self.error("structured-reference item is empty"));
                }
                unwrapped_trailing_space = true;
            }
            items.push(item);
        }

        let trailing_space = if nested {
            self.consume_structured_space()?
        } else {
            unwrapped_trailing_space
        };
        if leading_space != trailing_space {
            return Err(self.error("structured-reference square-bracket whitespace is asymmetric"));
        }
        if self.peek_char() != Some(']') {
            return Err(self.error("expected closing structured-reference bracket"));
        }
        self.offset += 1;
        if nested && items.len() == 1 {
            return Err(self
                .error("redundant nested structured reference cannot be represented faithfully"));
        }

        let (row_type, columns) = Self::classify_structured_reference(items, &separators)?;
        Ok(ParsedStructuredReference {
            row_type,
            columns,
            square_bracket_space: leading_space,
            comma_space: comma_space.unwrap_or(false),
        })
    }

    fn parse_structured_reference_item(
        &mut self,
        bracketed: bool,
    ) -> Result<StructuredReferenceItem> {
        if bracketed {
            if self.peek_char() != Some('[') {
                return Err(self.error("expected nested structured-reference item"));
            }
            self.offset += 1;
        }
        let mut text = String::new();
        let mut first_character_escaped = false;
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated structured reference"));
            };
            if ch == ']' {
                if bracketed {
                    self.offset += 1;
                }
                break;
            }
            self.offset += ch.len_utf8();
            if ch == '\'' {
                let Some(escaped) = self.peek_char() else {
                    return Err(self.error("unterminated structured-reference escape"));
                };
                if !matches!(escaped, '#' | '[' | ']' | '\'' | '@') {
                    return Err(self.error("invalid structured-reference escape"));
                }
                if text.is_empty() {
                    first_character_escaped = true;
                }
                self.offset += escaped.len_utf8();
                text.push(escaped);
            } else {
                text.push(ch);
            }
        }
        if text.is_empty() {
            return Err(self.error("structured-reference item is empty"));
        }
        Ok(StructuredReferenceItem {
            text,
            first_character_escaped,
        })
    }

    fn consume_structured_space(&mut self) -> Result<bool> {
        let start = self.offset;
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
        if self.offset == start {
            return Ok(false);
        }
        if &self.input[start..self.offset] != " " {
            return Err(
                self.error("structured-reference whitespace cannot be represented by XLSB flags")
            );
        }
        Ok(true)
    }

    fn classify_structured_reference(
        items: Vec<StructuredReferenceItem>,
        separators: &[char],
    ) -> Result<(TableRowType, TableNamedColumns)> {
        if separators.len() + 1 != items.len() {
            return Err(Error::InvalidFormula(
                "structured-reference separator count is invalid".to_string(),
            ));
        }

        let mut rows = Vec::new();
        let mut columns = Vec::new();
        let mut item_is_column = Vec::with_capacity(items.len());
        for item in items {
            let row = if item.first_character_escaped {
                None
            } else if item.text.eq_ignore_ascii_case("#All") {
                Some(TableRowType::All)
            } else if item.text.eq_ignore_ascii_case("#Data") {
                Some(TableRowType::DataAlternate)
            } else if item.text.eq_ignore_ascii_case("#Headers") {
                Some(TableRowType::Headers)
            } else if item.text.eq_ignore_ascii_case("#Totals") {
                Some(TableRowType::Totals)
            } else if item.text.eq_ignore_ascii_case("#This Row") {
                Some(TableRowType::Current)
            } else {
                None
            };
            if let Some(row) = row {
                rows.push(row);
                item_is_column.push(false);
            } else if !item.first_character_escaped && item.text.starts_with('#') {
                return Err(Error::InvalidFormula(format!(
                    "unknown structured-reference row selector {:?}",
                    item.text
                )));
            } else if !item.first_character_escaped && item.text.starts_with('@') {
                let column = item.text[1..].to_string();
                if column.is_empty() || !rows.is_empty() {
                    return Err(Error::InvalidFormula(
                        "invalid or duplicate current-row structured reference".to_string(),
                    ));
                }
                rows.push(TableRowType::Current);
                columns.push(column);
                item_is_column.push(true);
            } else {
                columns.push(item.text);
                item_is_column.push(true);
            }
        }

        let mut colon = None;
        for (index, separator) in separators.iter().copied().enumerate() {
            match separator {
                ':' if item_is_column[index] && item_is_column[index + 1] => {
                    if colon.replace(index).is_some() {
                        return Err(Error::InvalidFormula(
                            "structured reference has more than one column range".to_string(),
                        ));
                    }
                },
                ',' if !item_is_column[index] || !item_is_column[index + 1] => {},
                ',' => {
                    return Err(Error::InvalidFormula(
                        "disjoint structured-reference columns cannot fit one PtgList".to_string(),
                    ));
                },
                _ => {
                    return Err(Error::InvalidFormula(
                        "structured-reference separator has invalid operands".to_string(),
                    ));
                },
            }
        }

        let row_type = match rows.as_slice() {
            [] => TableRowType::Data,
            [row] => *row,
            [TableRowType::Headers, TableRowType::DataAlternate] => TableRowType::DataAndHeaders,
            [TableRowType::DataAlternate, TableRowType::Totals] => TableRowType::DataAndTotals,
            _ => {
                return Err(Error::InvalidFormula(
                    "structured-reference row union cannot fit one PtgList".to_string(),
                ));
            },
        };
        let columns = match columns.as_slice() {
            [] => TableNamedColumns::All,
            [column] if colon.is_none() => TableNamedColumns::One(column.clone()),
            [first, last] if colon.is_some() => TableNamedColumns::Range {
                first: first.clone(),
                last: last.clone(),
            },
            _ => {
                return Err(Error::InvalidFormula(
                    "structured-reference columns cannot fit one PtgList".to_string(),
                ));
            },
        };
        validate_named_table_columns(&columns)?;
        Ok((row_type, columns))
    }

    fn compile_bare_resident_table_reference(
        &self,
        table_name: &str,
    ) -> Result<Option<CompileExpr>> {
        let Some(context) = self.context else {
            return Ok(None);
        };
        if !context
            .tables
            .iter()
            .any(|table| excel_name_eq(table.display_name(), table_name))
        {
            return Ok(None);
        }
        self.compile_resident_table_reference(
            table_name,
            ParsedStructuredReference {
                row_type: TableRowType::Data,
                columns: TableNamedColumns::All,
                square_bracket_space: false,
                comma_space: false,
            },
        )
        .map(Some)
    }

    fn compile_resident_table_reference(
        &self,
        table_name: &str,
        selection: ParsedStructuredReference,
    ) -> Result<CompileExpr> {
        let context = self.context.ok_or_else(|| {
            Error::UnresolvedDependency(format!(
                "structured table reference {table_name:?} requires workbook compilation context"
            ))
        })?;
        let mut matches = context
            .tables
            .iter()
            .filter(|table| excel_name_eq(table.display_name(), table_name));
        let table = matches.next().ok_or_else(|| {
            Error::UnresolvedDependency(format!(
                "structured reference names missing table {table_name:?}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::UnresolvedDependency(format!(
                "structured reference table name {table_name:?} is ambiguous"
            )));
        }
        let current_sheet = usize::try_from(context.current_sheet)
            .map_err(|_| Error::InvalidFormula("current worksheet index overflow".to_string()))?;
        if table.sheet_index() != current_sheet {
            return Err(Error::InvalidFormula(format!(
                "table {table_name:?} is on worksheet {}, not the formula worksheet {current_sheet}",
                table.sheet_index()
            )));
        }
        let columns = match selection.columns {
            TableNamedColumns::All => TableColumns::All,
            TableNamedColumns::One(name) => {
                TableColumns::One(Self::resolve_table_column(table, &name)?)
            },
            TableNamedColumns::Range { first, last } => {
                let first = Self::resolve_table_column(table, &first)?;
                let last = Self::resolve_table_column(table, &last)?;
                if first > last {
                    return Err(Error::InvalidFormula(
                        "structured-reference column range is reversed".to_string(),
                    ));
                }
                TableColumns::Range { first, last }
            },
        };
        let sheet_index = u16::try_from(current_sheet)
            .ok()
            .and_then(|index| index.checked_add(2))
            .ok_or_else(|| {
                Error::InvalidFormula(
                    "table worksheet cannot be represented in the extern-sheet table".to_string(),
                )
            })?;
        Ok(CompileExpr::TableReference(TableReference {
            sheet_index,
            row_type: Some(selection.row_type),
            columns: Some(columns),
            square_bracket_space: selection.square_bracket_space,
            comma_space: selection.comma_space,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: Some(table.table_id()),
            external: None,
        }))
    }

    fn resolve_table_column(table: &Definition, name: &str) -> Result<u16> {
        let mut matches = table
            .columns()
            .iter()
            .enumerate()
            .filter(|(_, column)| excel_name_eq(column, name));
        let (index, _) = matches.next().ok_or_else(|| {
            Error::UnresolvedDependency(format!(
                "structured reference names missing column {name:?} in table {:?}",
                table.display_name()
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::UnresolvedDependency(format!(
                "structured-reference column {name:?} is ambiguous"
            )));
        }
        u16::try_from(index).map_err(|_| {
            Error::InvalidFormula("structured-reference column index overflow".to_string())
        })
    }

    fn compile_external_table_reference(
        &self,
        qualifier: &str,
        table: String,
        selection: ParsedStructuredReference,
    ) -> Result<CompileExpr> {
        validate_table_name(&table)?;
        let sheet_index = self.resolve_external_table_xti(qualifier)?;
        Ok(CompileExpr::TableReference(TableReference {
            sheet_index,
            row_type: None,
            columns: None,
            square_bracket_space: selection.square_bracket_space,
            comma_space: selection.comma_space,
            data_type: TableDataType::Reference,
            invalid: false,
            list_index: None,
            external: Some(ExternalTableReference {
                table,
                row_type: selection.row_type,
                columns: selection.columns,
            }),
        }))
    }

    fn resolve_external_table_xti(&self, qualifier: &str) -> Result<u16> {
        let close = qualifier.find(']').ok_or_else(|| {
            Error::InvalidFormula("external structured reference omits ']'".to_string())
        })?;
        if !qualifier.starts_with('[')
            || close == 1
            || close + 1 == qualifier.len()
            || qualifier[1..close].contains('[')
            || qualifier[close + 1..].contains('[')
            || qualifier[close + 1..].contains(']')
        {
            return Err(Error::InvalidFormula(format!(
                "invalid external structured-reference qualifier {qualifier:?}"
            )));
        }
        let target = &qualifier[1..close];
        let sheet = &qualifier[close + 1..];
        if sheet.contains(':') {
            return Err(Error::InvalidFormula(
                "external structured reference must select exactly one worksheet".to_string(),
            ));
        }

        let context = self.context.ok_or_else(|| {
            Error::UnresolvedDependency(
                "external structured reference requires workbook compilation context".to_string(),
            )
        })?;

        let mut found = None;
        let mut unsupported = None;
        for (xti_index, xti) in context.external_sheets.iter().enumerate() {
            let link_index = usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula(format!(
                    "external structured-reference external-link index {} overflows",
                    xti.external_link
                ))
            })?;
            let Some(link) = context.supporting_links.get(link_index) else {
                continue;
            };
            let book_index = match link {
                SupportingLink::ExternalWorkbook(book_index) => usize::try_from(*book_index)
                    .map_err(|_| {
                        Error::InvalidFormula(format!(
                            "external structured-reference book index {book_index} overflows"
                        ))
                    })?,
                SupportingLink::AddIn => continue,
                SupportingLink::SelfWorkbook | SupportingLink::SameSheet => continue,
            };
            let Some(book) = context.external_books.get(book_index) else {
                continue;
            };
            if xti.first_sheet < 0 || xti.last_sheet < 0 {
                return Err(Error::InvalidFormula(format!(
                    "external structured-reference Xti {xti_index} has negative worksheet bounds {}..={}",
                    xti.first_sheet, xti.last_sheet
                )));
            }
            if xti.last_sheet < xti.first_sheet {
                return Err(Error::InvalidFormula(format!(
                    "external structured-reference Xti {xti_index} has reversed worksheet bounds {}..={}",
                    xti.first_sheet, xti.last_sheet
                )));
            }
            if xti.first_sheet != xti.last_sheet {
                return Err(Error::InvalidFormula(format!(
                    "external structured-reference Xti {xti_index} must select exactly one worksheet"
                )));
            }
            if !book.metadata.is_workbook() {
                if excel_name_eq(book.metadata.source(), target) {
                    unsupported.get_or_insert_with(|| {
                        format!(
                            "external structured reference points to an unsupported data source {target:?}"
                        )
                    });
                }
                continue;
            }
            if !excel_name_eq(book.metadata.source(), target) {
                continue;
            }
            let sheet_index = usize::try_from(xti.first_sheet).map_err(|_| {
                Error::InvalidFormula(
                    "external structured-reference worksheet index overflows".to_string(),
                )
            })?;
            let Some(candidate) = book.metadata.sheet_names().get(sheet_index) else {
                continue;
            };
            if !excel_name_eq(candidate, sheet) {
                continue;
            }
            let xti_index = u16::try_from(xti_index).map_err(|_| {
                Error::InvalidFormula("external structured-reference Xti overflow".to_string())
            })?;
            if xti_index == u16::MAX {
                return Err(Error::InvalidFormula(
                    "external structured-reference Xti overflow".to_string(),
                ));
            }
            if found.replace(xti_index).is_some() {
                return Err(Error::UnresolvedDependency(format!(
                    "external structured-reference qualifier {qualifier:?} is ambiguous"
                )));
            }
        }
        if let Some(xti_index) = found {
            return Ok(xti_index);
        }
        if let Some(message) = unsupported {
            return Err(Error::UnsupportedFeature(message));
        }
        Err(Error::UnresolvedDependency(format!(
            "external structured-reference qualifier {qualifier:?} is missing"
        )))
    }

    fn parse_string(&mut self) -> Result<String> {
        debug_assert_eq!(self.peek_char(), Some('"'));
        self.offset += 1;
        let mut value = String::new();
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated string literal"));
            };
            self.offset += ch.len_utf8();
            if ch == '"' {
                if self.peek_char() == Some('"') {
                    self.offset += 1;
                    value.push('"');
                } else {
                    break;
                }
            } else {
                value.push(ch);
            }
        }
        if value.encode_utf16().count() > 255 {
            return Err(Error::InvalidFormula(
                "formula string literal exceeds 255 UTF-16 code units".to_string(),
            ));
        }
        Ok(value)
    }

    fn parse_array_constant(&mut self) -> Result<CompileExpr> {
        let mut values = Vec::new();
        let mut rows = 1_u32;
        let mut cols = 0_u32;
        let mut current_cols = 0_u32;
        loop {
            self.skip_spaces();
            if self.peek_char() == Some('}') {
                return Err(self.error("array rows cannot be empty"));
            }
            let value = if self.peek_char() == Some('"') {
                ArrayValue::String(self.parse_string()?)
            } else if self.peek_char() == Some('#') {
                let start = self.offset;
                while self
                    .peek_char()
                    .is_some_and(|ch| !matches!(ch, ',' | ';' | '}') && !ch.is_whitespace())
                {
                    self.offset += self.peek_char().expect("checked").len_utf8();
                }
                let error = formula_error_code(&self.input[start..self.offset])
                    .ok_or_else(|| self.error("unknown array error literal"))?;
                ArrayValue::Error(error)
            } else if self.input[self.offset..]
                .get(..4)
                .is_some_and(|value| value.eq_ignore_ascii_case("TRUE"))
            {
                self.offset += 4;
                ArrayValue::Bool(true)
            } else if self.input[self.offset..]
                .get(..5)
                .is_some_and(|value| value.eq_ignore_ascii_case("FALSE"))
            {
                self.offset += 5;
                ArrayValue::Bool(false)
            } else {
                let negative = self.consume("-");
                if !negative {
                    self.consume("+");
                }
                let mut number = self.parse_number()?;
                if negative {
                    number = -number;
                }
                ArrayValue::Number(number)
            };
            values.push(value);
            current_cols = current_cols
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormula("array column count overflow".to_string()))?;

            if self.consume(",") {
                continue;
            }
            if self.consume(";") {
                if cols == 0 {
                    cols = current_cols;
                } else if cols != current_cols {
                    return Err(self.error("array rows have different column counts"));
                }
                rows = rows
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormula("array row count overflow".to_string()))?;
                current_cols = 0;
                continue;
            }
            if self.consume("}") {
                if cols == 0 {
                    cols = current_cols;
                } else if cols != current_cols {
                    return Err(self.error("array rows have different column counts"));
                }
                break;
            }
            return Err(self.error("expected ',', ';', or '}' in array constant"));
        }
        if rows > 1_048_576 || cols == 0 || cols > 16_384 {
            return Err(self.error("array dimensions exceed worksheet limits"));
        }
        Ok(CompileExpr::Array { rows, cols, values })
    }

    fn parse_number(&mut self) -> Result<f64> {
        self.skip_spaces();
        let start = self.offset;
        let mut seen_exponent = false;
        while let Some(ch) = self.peek_char() {
            if ch.is_ascii_digit() || ch == '.' {
                self.offset += 1;
            } else if matches!(ch, 'e' | 'E') && !seen_exponent {
                seen_exponent = true;
                self.offset += 1;
                if matches!(self.peek_char(), Some('+' | '-')) {
                    self.offset += 1;
                }
            } else {
                break;
            }
        }
        self.input[start..self.offset]
            .parse::<f64>()
            .map_err(|_| self.error("invalid numeric literal"))
    }

    fn parse_error_literal(&mut self) -> Result<u8> {
        self.skip_spaces();
        let rest = &self.input[self.offset..];
        let Some((literal, code)) = FORMULA_ERRORS.iter().find(|(literal, _)| {
            rest.get(..literal.len())
                .is_some_and(|value| value.eq_ignore_ascii_case(literal))
        }) else {
            return Err(self.error("unknown formula error literal"));
        };
        self.offset += literal.len();
        Ok(*code)
    }

    fn parse_identifier(&mut self) -> Result<String> {
        self.skip_spaces();
        let start = self.offset;
        while let Some(ch) = self.peek_char() {
            if ch.is_alphanumeric() || matches!(ch, '_' | '.' | '$' | '?' | '\\' | '\u{061F}') {
                self.offset += ch.len_utf8();
            } else {
                break;
            }
        }
        if self.offset == start {
            Err(self.error("expected literal, reference, or function"))
        } else {
            Ok(self.input[start..self.offset].to_string())
        }
    }

    fn parse_quoted_sheet_name(&mut self) -> Result<String> {
        self.skip_spaces();
        debug_assert_eq!(self.peek_char(), Some('\''));
        self.offset += 1;
        let mut name = String::new();
        loop {
            let Some(ch) = self.peek_char() else {
                return Err(self.error("unterminated quoted worksheet name"));
            };
            self.offset += ch.len_utf8();
            if ch == '\'' {
                if self.peek_char() == Some('\'') {
                    self.offset += 1;
                    name.push('\'');
                } else {
                    break;
                }
            } else {
                name.push(ch);
            }
        }
        if name.is_empty() {
            return Err(self.error("worksheet name is empty"));
        }
        Ok(name)
    }

    fn split_sheet_qualifier(value: &str) -> Result<(&str, Option<&str>)> {
        let Some((first, last)) = value.split_once(':') else {
            return Ok((value, None));
        };
        if first.is_empty() || last.is_empty() || last.contains(':') {
            return Err(Error::InvalidFormula(format!(
                "invalid worksheet range {value:?}"
            )));
        }
        Ok((first, Some(last)))
    }

    fn parse_qualified_reference(
        &mut self,
        first_sheet: &str,
        last_sheet: Option<&str>,
    ) -> Result<CompileExpr> {
        let first_text = self.parse_identifier()?;
        let first = parse_a1_reference(&first_text)
            .ok_or_else(|| self.error("invalid sheet-qualified cell reference"))?;
        let second = if self.consume(":") {
            let second_text = self.parse_identifier()?;
            Some(
                parse_a1_reference(&second_text)
                    .ok_or_else(|| self.error("invalid sheet-qualified range end"))?,
            )
        } else {
            None
        };
        let sheet_index = self.resolve_sheet_range(first_sheet, last_sheet)?;
        Ok(match second {
            Some(second) => CompileExpr::Area3d(sheet_index, first, second),
            None => CompileExpr::Ref3d(sheet_index, first),
        })
    }

    fn resolve_sheet_range(&self, first_sheet: &str, last_sheet: Option<&str>) -> Result<u16> {
        let context = self.context.ok_or_else(|| {
            Error::UnresolvedDependency(
                "sheet-qualified reference requires workbook compilation context".to_string(),
            )
        })?;

        let resolve_sheet = |sheet: &str| -> Result<usize> {
            let mut matches = context
                .worksheet_names
                .iter()
                .enumerate()
                .filter(|(_, candidate)| excel_name_eq(candidate, sheet));
            let Some((index, _)) = matches.next() else {
                return Err(Error::UnresolvedDependency(format!(
                    "worksheet {sheet:?} is missing from workbook metadata"
                )));
            };
            if matches.next().is_some() {
                return Err(Error::UnresolvedDependency(format!(
                    "worksheet {sheet:?} is ambiguous in workbook metadata"
                )));
            }
            Ok(index)
        };

        let first_index = resolve_sheet(first_sheet)?;
        let last_index = if let Some(last_sheet) = last_sheet {
            resolve_sheet(last_sheet)?
        } else {
            first_index
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula(format!(
                "worksheet range {first_sheet:?}:{last_sheet:?} is in reverse workbook order"
            )));
        }
        if first_index == last_index {
            return u16::try_from(first_index)
                .ok()
                .and_then(|index| index.checked_add(2))
                .ok_or_else(|| {
                    Error::InvalidFormula(format!(
                        "worksheet {first_sheet:?} cannot be represented in the extern-sheet table"
                    ))
                });
        }

        let first = u32::try_from(first_index)
            .map_err(|_| Error::InvalidFormula("first sheet index overflow".to_string()))?;
        let last = u32::try_from(last_index)
            .map_err(|_| Error::InvalidFormula("last sheet index overflow".to_string()))?;
        let mut ranges = context.sheet_ranges.borrow_mut();
        let range_index = if let Some(index) = ranges
            .iter()
            .position(|candidate| *candidate == (first, last))
        {
            index
        } else {
            let base_count = context
                .worksheet_names
                .len()
                .checked_add(2)
                .ok_or_else(|| Error::InvalidFormula("Xti count overflow".to_string()))?;
            if base_count
                .checked_add(ranges.len())
                .is_none_or(|count| count >= usize::from(u16::MAX))
            {
                return Err(Error::InvalidFormula(
                    "formula sheet ranges exceed the XLSB extern-sheet limit".to_string(),
                ));
            }
            ranges.push((first, last));
            ranges.len() - 1
        };
        let xti_index = context
            .worksheet_names
            .len()
            .checked_add(2)
            .and_then(|base| base.checked_add(range_index))
            .ok_or_else(|| Error::InvalidFormula("Xti index overflow".to_string()))?;
        u16::try_from(xti_index)
            .map_err(|_| Error::InvalidFormula("Xti index overflow".to_string()))
    }

    fn resolve_defined_name(&self, name: &str) -> Result<u32> {
        let context = self.context.ok_or_else(|| {
            Error::UnresolvedDependency(format!(
                "defined name {name:?} requires workbook compilation context"
            ))
        })?;

        let mut local = context
            .defined_names
            .iter()
            .enumerate()
            .filter(|(_, candidate)| {
                candidate.sheet_id == Some(context.current_sheet)
                    && excel_name_eq(&candidate.name, name)
            });
        let index = if let Some((index, _)) = local.next() {
            if local.next().is_some() {
                return Err(Error::UnresolvedDependency(format!(
                    "defined name {name:?} is ambiguous for worksheet {}",
                    context.current_sheet
                )));
            }
            index
        } else {
            let mut global = context
                .defined_names
                .iter()
                .enumerate()
                .filter(|(_, candidate)| {
                    candidate.sheet_id.is_none() && excel_name_eq(&candidate.name, name)
                });
            let Some((index, _)) = global.next() else {
                return Err(Error::UnresolvedDependency(format!(
                    "defined name {name:?} is not visible from worksheet {}",
                    context.current_sheet
                )));
            };
            if global.next().is_some() {
                return Err(Error::UnresolvedDependency(format!(
                    "defined name {name:?} is ambiguous in workbook metadata"
                )));
            }
            index
        };
        u32::try_from(index)
            .ok()
            .and_then(|index| index.checked_add(1))
            .ok_or_else(|| Error::InvalidFormula("defined-name index overflow".to_string()))
    }

    fn consume(&mut self, text: &str) -> bool {
        self.skip_spaces();
        if self.input[self.offset..].starts_with(text) {
            self.offset += text.len();
            true
        } else {
            false
        }
    }

    pub(super) fn skip_spaces(&mut self) {
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.offset..].chars().next()
    }

    pub(super) fn error(&self, message: &str) -> Error {
        Error::InvalidFormula(format!("{message} at byte {}", self.offset))
    }
}
