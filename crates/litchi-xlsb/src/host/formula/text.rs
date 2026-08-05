//! XLSB formula-text compiler.
use super::function_table::BUILTIN_FUNCTIONS;
use super::*;
#[derive(Debug, Clone, Copy)]
struct BuiltinFunction {
    index: u16,
    name: &'static str,
    min_args: u8,
    max_args: u8,
}

impl BuiltinFunction {
    fn accepts_arg_count(self, count: u8) -> bool {
        if count < self.min_args || count > self.max_args {
            return false;
        }
        match self.index {
            // GETPIVOTDATA permits the two mandatory arguments, one optional
            // field, or complete field/item pairs thereafter.
            358 => count <= 3 || count.is_multiple_of(2),
            // COUNTIFS is made solely of range/criteria pairs.
            481 => count.is_multiple_of(2),
            // SUMIFS and AVERAGEIFS have one leading aggregate range followed
            // by range/criteria pairs.
            482 | 484 => !count.is_multiple_of(2),
            _ => true,
        }
    }
}

fn builtin_function_by_name(name: &str) -> Option<BuiltinFunction> {
    BUILTIN_FUNCTIONS
        .iter()
        .find_map(|&(index, function_name, min_args, max_args)| {
            function_name
                .eq_ignore_ascii_case(name)
                .then_some(BuiltinFunction {
                    index,
                    name: function_name,
                    min_args,
                    max_args,
                })
        })
}

#[cfg(test)]
pub(super) fn has_builtin_function(index: u16) -> bool {
    let position = BUILTIN_FUNCTIONS
        .binary_search_by_key(&index, |entry| entry.0)
        .ok();
    position.is_some()
}

/// Compiles standards-defined Excel formula text to XLSB RPN tokens.
///
/// The compiler supports literals, A1 references and ranges, parentheses,
/// arithmetic/comparison/concatenation operators, percent, and the built-in
/// non-macro built-in functions from [MS-XLSB]'s `Ftab` table, and typed array
/// constants. Unsupported constructs return an error; they are never replaced
/// by a cached value.
pub struct Compiler<'a> {
    input: &'a str,
    offset: usize,
    context: Option<&'a CompilationContext<'a>>,
}

/// A defined name visible to the XLSB formula text compiler.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct DefinedName {
    pub(crate) name: String,
    pub(crate) sheet_id: Option<u32>,
}

/// Workbook metadata used to compile context-dependent formula operands.
#[derive(Debug, Clone, Copy)]
pub(crate) struct CompilationContext<'a> {
    pub(crate) worksheet_names: &'a [String],
    pub(crate) defined_names: &'a [DefinedName],
    pub(crate) tables: &'a [Definition],
    pub(crate) supporting_links: &'a [SupportingLink],
    pub(crate) external_sheets: &'a [ExternalSheet],
    pub(crate) external_books: &'a [ExternalBook],
    pub(crate) sheet_ranges: &'a std::cell::RefCell<Vec<(u32, u32)>>,
    pub(crate) current_sheet: u32,
}

#[derive(Debug, Clone, Copy)]
enum FormulaEncoding {
    Cell,
    Shared { base_row: u32, base_col: u32 },
}

#[derive(Debug)]
enum CompileExpr {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    MissingArg,
    Parenthesized(Box<CompileExpr>),
    Array {
        rows: u32,
        cols: u32,
        values: Vec<ArrayValue>,
    },
    Ref(A1Reference),
    Area(A1Reference, A1Reference),
    Ref3d(u16, A1Reference),
    Area3d(u16, A1Reference, A1Reference),
    Name(u32),
    TableReference(TableReference),
    Unary(UnaryOperator, Box<CompileExpr>),
    Binary(BinaryOperator, Box<CompileExpr>, Box<CompileExpr>),
    Function(BuiltinFunction, Vec<CompileExpr>),
}

#[derive(Debug)]
struct ParsedStructuredReference {
    row_type: TableRowType,
    columns: TableNamedColumns,
    square_bracket_space: bool,
    comma_space: bool,
}

#[derive(Debug)]
struct StructuredReferenceItem {
    text: String,
    first_character_escaped: bool,
}

#[derive(Debug, Clone, Copy)]
struct A1Reference {
    row: u32,
    col: u32,
    row_relative: bool,
    col_relative: bool,
}

impl<'a> Compiler<'a> {
    pub fn compile(formula: &'a str) -> Result<ParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, None)
    }

    pub(crate) fn compile_with_context(
        formula: &'a str,
        context: &'a CompilationContext<'a>,
    ) -> Result<ParsedFormula> {
        Self::compile_with_encoding(formula, FormulaEncoding::Cell, Some(context))
    }

    /// Compile a shared formula, encoding relative A1 references as
    /// `PtgRefN`/`PtgAreaN` offsets from the first cell in the shared range.
    pub fn compile_shared(formula: &'a str, base_row: u32, base_col: u32) -> Result<ParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, None)
    }

    pub(crate) fn compile_shared_with_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: &'a CompilationContext<'a>,
    ) -> Result<ParsedFormula> {
        Self::compile_shared_with_optional_context(formula, base_row, base_col, Some(context))
    }

    fn compile_shared_with_optional_context(
        formula: &'a str,
        base_row: u32,
        base_col: u32,
        context: Option<&'a CompilationContext<'a>>,
    ) -> Result<ParsedFormula> {
        if base_row >= 1_048_576 || base_col >= 16_384 {
            return Err(Error::InvalidCellReference(format!(
                "shared formula base ({base_row}, {base_col})"
            )));
        }
        Self::compile_with_encoding(
            formula,
            FormulaEncoding::Shared { base_row, base_col },
            context,
        )
    }

    fn compile_with_encoding(
        formula: &'a str,
        encoding: FormulaEncoding,
        context: Option<&'a CompilationContext<'a>>,
    ) -> Result<ParsedFormula> {
        let input = formula.strip_prefix('=').unwrap_or(formula).trim();
        if input.is_empty() {
            return Err(Error::InvalidFormula(
                "formula expression is empty".to_string(),
            ));
        }
        let mut compiler = Self {
            input,
            offset: 0,
            context,
        };
        let expression = compiler.parse_comparison()?;
        compiler.skip_spaces();
        if compiler.offset != compiler.input.len() {
            return Err(compiler.error("unexpected trailing input"));
        }

        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        Self::emit(&expression, &mut rgce, &mut rgcb, encoding)?;
        if rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "compiled formula is {} bytes; maximum is {MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(ParsedFormula { rgce, rgcb })
    }

    fn parse_comparison(&mut self) -> Result<CompileExpr> {
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
                return Err(self.error(
                    "external cell references are not supported by this compilation context",
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
            Error::UnsupportedFeature(format!(
                "structured table reference {table_name:?} requires workbook compilation context"
            ))
        })?;
        let mut matches = context
            .tables
            .iter()
            .filter(|table| excel_name_eq(table.display_name(), table_name));
        let table = matches.next().ok_or_else(|| {
            Error::InvalidFormula(format!(
                "structured reference names missing table {table_name:?}"
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
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
            Error::InvalidFormula(format!(
                "structured reference names missing column {name:?} in table {:?}",
                table.display_name()
            ))
        })?;
        if matches.next().is_some() {
            return Err(Error::InvalidFormula(format!(
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
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(
                "external structured reference requires workbook compilation context".to_string(),
            )
        })?;
        let close = qualifier.find(']').ok_or_else(|| {
            Error::InvalidFormula("external structured reference omits ']'".to_string())
        })?;
        if !qualifier.starts_with('[') || close == 1 || close + 1 == qualifier.len() {
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

        let mut found = None;
        for (xti_index, xti) in context.external_sheets.iter().enumerate() {
            if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                continue;
            }
            let Ok(link_index) = usize::try_from(xti.external_link) else {
                continue;
            };
            let Some(SupportingLink::ExternalWorkbook(book_index)) =
                context.supporting_links.get(link_index)
            else {
                continue;
            };
            let Ok(book_index) = usize::try_from(*book_index) else {
                continue;
            };
            let Some(book) = context.external_books.get(book_index) else {
                continue;
            };
            let Ok(sheet_index) = usize::try_from(xti.first_sheet) else {
                continue;
            };
            if !book.metadata.is_workbook()
                || !excel_name_eq(book.metadata.source(), target)
                || !book
                    .metadata
                    .sheet_names()
                    .get(sheet_index)
                    .is_some_and(|candidate| excel_name_eq(candidate, sheet))
            {
                continue;
            }
            let xti_index = u16::try_from(xti_index).map_err(|_| {
                Error::InvalidFormula("external structured-reference Xti overflow".to_string())
            })?;
            if xti_index == u16::MAX || found.replace(xti_index).is_some() {
                return Err(Error::InvalidFormula(format!(
                    "external structured-reference qualifier {qualifier:?} is ambiguous"
                )));
            }
        }
        found.ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external structured-reference qualifier {qualifier:?} is missing"
            ))
        })
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
        let sheet_index = self.resolve_sheet_range(first_sheet, last_sheet)?;
        let first_text = self.parse_identifier()?;
        let first = parse_a1_reference(&first_text)
            .ok_or_else(|| self.error("invalid sheet-qualified cell reference"))?;
        if self.consume(":") {
            let second_text = self.parse_identifier()?;
            let second = parse_a1_reference(&second_text)
                .ok_or_else(|| self.error("invalid sheet-qualified range end"))?;
            Ok(CompileExpr::Area3d(sheet_index, first, second))
        } else {
            Ok(CompileExpr::Ref3d(sheet_index, first))
        }
    }

    fn resolve_sheet_range(&self, first_sheet: &str, last_sheet: Option<&str>) -> Result<u16> {
        let context = self.context.ok_or_else(|| {
            Error::UnsupportedFeature(
                "sheet-qualified reference requires workbook compilation context".to_string(),
            )
        })?;
        let first_index = context
            .worksheet_names
            .iter()
            .position(|candidate| excel_name_eq(candidate, first_sheet))
            .ok_or_else(|| Error::WorksheetNotFound(first_sheet.to_string()))?;
        let last_index = if let Some(last_sheet) = last_sheet {
            context
                .worksheet_names
                .iter()
                .position(|candidate| excel_name_eq(candidate, last_sheet))
                .ok_or_else(|| Error::WorksheetNotFound(last_sheet.to_string()))?
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
            Error::UnsupportedFeature(format!(
                "defined name {name:?} requires workbook compilation context"
            ))
        })?;
        let local = context.defined_names.iter().position(|candidate| {
            candidate.sheet_id == Some(context.current_sheet)
                && excel_name_eq(&candidate.name, name)
        });
        let index = local.or_else(|| {
            context.defined_names.iter().position(|candidate| {
                candidate.sheet_id.is_none() && excel_name_eq(&candidate.name, name)
            })
        });
        let index = index.ok_or_else(|| {
            Error::InvalidFormula(format!(
                "defined name {name:?} is not visible from worksheet {}",
                context.current_sheet
            ))
        })?;
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

    fn skip_spaces(&mut self) {
        while self.peek_char().is_some_and(char::is_whitespace) {
            self.offset += self.peek_char().expect("checked").len_utf8();
        }
    }

    fn peek_char(&self) -> Option<char> {
        self.input[self.offset..].chars().next()
    }

    fn error(&self, message: &str) -> Error {
        Error::InvalidFormula(format!("{message} at byte {}", self.offset))
    }

    fn emit(
        expression: &CompileExpr,
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        match expression {
            CompileExpr::Number(value) => {
                validate_xnum(*value, "compiled number")?;
                if value.fract() == 0.0 && *value >= 0.0 && *value <= f64::from(u16::MAX) {
                    output.push(ptg_types::PTG_INT);
                    output.extend_from_slice(&(*value as u16).to_le_bytes());
                } else {
                    output.push(ptg_types::PTG_NUM);
                    output.extend_from_slice(&value.to_le_bytes());
                }
            },
            CompileExpr::String(value) => {
                let utf16: Vec<u16> = value.encode_utf16().collect();
                output.push(ptg_types::PTG_STR);
                output.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                for unit in utf16 {
                    output.extend_from_slice(&unit.to_le_bytes());
                }
            },
            CompileExpr::Bool(value) => {
                output.push(ptg_types::PTG_BOOL);
                output.push(u8::from(*value));
            },
            CompileExpr::Error(error) => {
                output.push(ptg_types::PTG_ERR);
                output.push(*error);
            },
            CompileExpr::MissingArg => output.push(ptg_types::PTG_MISSING_ARG),
            CompileExpr::Parenthesized(expression) => {
                Self::emit(expression, output, extra, encoding)?;
                output.push(ptg_types::PTG_PAREN);
            },
            CompileExpr::Array { rows, cols, values } => {
                if matches!(encoding, FormulaEncoding::Shared { .. }) {
                    return Err(Error::InvalidFormula(
                        "shared formulas cannot contain PtgArray".to_string(),
                    ));
                }
                output.push(0x40); // PtgArray, VALUE class
                output.extend_from_slice(&[0; 14]);
                extra.extend_from_slice(&rows.to_le_bytes());
                extra.extend_from_slice(&cols.to_le_bytes());
                for value in values {
                    match value {
                        ArrayValue::Number(value) => {
                            extra.push(0x00);
                            extra.extend_from_slice(&value.to_le_bytes());
                        },
                        ArrayValue::String(value) => {
                            let utf16: Vec<u16> = value.encode_utf16().collect();
                            extra.push(0x01);
                            extra.extend_from_slice(&(utf16.len() as u16).to_le_bytes());
                            for unit in utf16 {
                                extra.extend_from_slice(&unit.to_le_bytes());
                            }
                        },
                        ArrayValue::Bool(value) => {
                            extra.extend_from_slice(&[0x02, u8::from(*value)]);
                        },
                        ArrayValue::Error(error) => {
                            extra.extend_from_slice(&[0x04, *error, 0, 0, 0]);
                        },
                    }
                }
            },
            CompileExpr::Ref(reference) => match encoding {
                FormulaEncoding::Cell => emit_reference(output, 0x44, *reference),
                FormulaEncoding::Shared { base_row, base_col } => {
                    emit_shared_reference(output, 0x4C, *reference, base_row, base_col)?
                },
            },
            CompileExpr::Area(first, last) => {
                match encoding {
                    FormulaEncoding::Cell => {
                        output.push(0x25); // PtgArea, REFERENCE class
                        output.extend_from_slice(&first.row.to_le_bytes());
                        output.extend_from_slice(&last.row.to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                        output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
                    },
                    FormulaEncoding::Shared { base_row, base_col } => {
                        output.push(0x2D); // PtgAreaN, REFERENCE class
                        let (first_row, first_col) =
                            encode_shared_reference(*first, base_row, base_col)?;
                        let (last_row, last_col) =
                            encode_shared_reference(*last, base_row, base_col)?;
                        output.extend_from_slice(&first_row.to_le_bytes());
                        output.extend_from_slice(&last_row.to_le_bytes());
                        output.extend_from_slice(&first_col.to_le_bytes());
                        output.extend_from_slice(&last_col.to_le_bytes());
                    },
                }
            },
            CompileExpr::Ref3d(sheet_index, reference) => {
                output.push(0x5A); // PtgRef3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&reference.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*reference).to_le_bytes());
            },
            CompileExpr::Area3d(sheet_index, first, last) => {
                output.push(0x5B); // PtgArea3d, VALUE class
                output.extend_from_slice(&sheet_index.to_le_bytes());
                output.extend_from_slice(&first.row.to_le_bytes());
                output.extend_from_slice(&last.row.to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*first).to_le_bytes());
                output.extend_from_slice(&reference_column_bits(*last).to_le_bytes());
            },
            CompileExpr::Name(index) => {
                output.push(0x43); // PtgName, VALUE class
                output.extend_from_slice(&index.to_le_bytes());
            },
            CompileExpr::TableReference(reference) => {
                let (token, payload) = reference.to_extended_binary()?;
                output.extend_from_slice(&token);
                extra.extend_from_slice(&payload);
            },
            CompileExpr::Unary(operator, operand) => {
                Self::emit(operand, output, extra, encoding)?;
                output.push(match operator {
                    UnaryOperator::Plus => ptg_types::PTG_UPLUS,
                    UnaryOperator::Minus => ptg_types::PTG_UMINUS,
                    UnaryOperator::Percent => ptg_types::PTG_PERCENT,
                });
            },
            CompileExpr::Binary(operator, left, right) => {
                Self::emit(left, output, extra, encoding)?;
                Self::emit(right, output, extra, encoding)?;
                output.push(match operator {
                    BinaryOperator::Add => ptg_types::PTG_ADD,
                    BinaryOperator::Subtract => ptg_types::PTG_SUB,
                    BinaryOperator::Multiply => ptg_types::PTG_MUL,
                    BinaryOperator::Divide => ptg_types::PTG_DIV,
                    BinaryOperator::Power => ptg_types::PTG_POWER,
                    BinaryOperator::Concat => ptg_types::PTG_CONCAT,
                    BinaryOperator::LessThan => ptg_types::PTG_LT,
                    BinaryOperator::LessEqual => ptg_types::PTG_LE,
                    BinaryOperator::Equal => ptg_types::PTG_EQ,
                    BinaryOperator::GreaterEqual => ptg_types::PTG_GE,
                    BinaryOperator::GreaterThan => ptg_types::PTG_GT,
                    BinaryOperator::NotEqual => ptg_types::PTG_NE,
                    BinaryOperator::Intersection => ptg_types::PTG_ISECT,
                    BinaryOperator::Union => ptg_types::PTG_UNION,
                    BinaryOperator::Range => ptg_types::PTG_RANGE,
                });
            },
            CompileExpr::Function(function, arguments) => {
                if function.index == 1 {
                    return Self::emit_if(arguments, output, extra, encoding);
                }
                if function.index == 100 {
                    return Self::emit_choose(arguments, output, extra, encoding);
                }
                if function.index == 480 {
                    return Self::emit_iferror(arguments, output, extra, encoding);
                }
                for argument in arguments {
                    Self::emit(argument, output, extra, encoding)?;
                }
                if function.min_args == function.max_args {
                    output.push(0x41); // PtgFunc, VALUE class
                    output.extend_from_slice(&function.index.to_le_bytes());
                } else {
                    output.push(0x42); // PtgFuncVar, VALUE class
                    output.push(arguments.len() as u8);
                    output.extend_from_slice(&function.index.to_le_bytes());
                }
            },
        }
        Ok(())
    }

    fn emit_if(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!(matches!(arguments.len(), 2 | 3));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if = append_attribute(output, 0x02, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto_true = append_attribute(output, 0x08, 0);
        let goto_false = if arguments.len() == 3 {
            Self::emit(&arguments[2], output, extra, encoding)?;
            Some(append_attribute(output, 0x08, 0))
        } else {
            None
        };
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x01, 0x00]);

        patch_attribute_offset(output, attr_if, goto_true + 4 - (attr_if + 4))?;
        patch_skip_to_end(output, goto_true)?;
        if let Some(position) = goto_false {
            patch_skip_to_end(output, position)?;
        }
        Ok(())
    }

    fn emit_iferror(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert_eq!(arguments.len(), 2);
        Self::emit(&arguments[0], output, extra, encoding)?;
        let attr_if_error = append_attribute(output, 0x80, 0);
        Self::emit(&arguments[1], output, extra, encoding)?;
        let goto = append_attribute(output, 0x08, 0);
        output.extend_from_slice(&[0x41, 0xE0, 0x01]);

        patch_attribute_offset(output, attr_if_error, goto + 4 - (attr_if_error + 4))?;
        patch_skip_to_end(output, goto)?;
        Ok(())
    }

    fn emit_choose(
        arguments: &[CompileExpr],
        output: &mut Vec<u8>,
        extra: &mut Vec<u8>,
        encoding: FormulaEncoding,
    ) -> Result<()> {
        debug_assert!((2..=255).contains(&arguments.len()));
        Self::emit(&arguments[0], output, extra, encoding)?;
        let choice_count = arguments.len() - 1;
        let attr_choose = output.len();
        output.extend_from_slice(&[ptg_types::PTG_ATTR, 0x04]);
        output.extend_from_slice(&(choice_count as u16).to_le_bytes());
        output.resize(output.len() + (choice_count + 1) * 2, 0);
        let attr_size = output.len() - attr_choose;
        patch_u16(
            output,
            attr_choose + 4,
            attr_size - 4,
            "PtgAttrChoose first offset",
        )?;

        let mut gotos = Vec::with_capacity(choice_count);
        for (index, argument) in arguments[1..].iter().enumerate() {
            Self::emit(argument, output, extra, encoding)?;
            gotos.push(append_attribute(output, 0x08, 0));
            let cumulative = output.len() - (attr_choose + attr_size);
            patch_u16(
                output,
                attr_choose + 6 + index * 2,
                cumulative,
                "PtgAttrChoose branch offset",
            )?;
        }
        output.extend_from_slice(&[0x42, arguments.len() as u8, 0x64, 0x00]);
        for goto in gotos {
            patch_skip_to_end(output, goto)?;
        }
        Ok(())
    }
}

fn append_attribute(output: &mut Vec<u8>, selector: u8, offset: u16) -> usize {
    let position = output.len();
    output.extend_from_slice(&[ptg_types::PTG_ATTR, selector]);
    output.extend_from_slice(&offset.to_le_bytes());
    position
}

fn patch_attribute_offset(output: &mut [u8], position: usize, offset: usize) -> Result<()> {
    patch_u16(output, position + 2, offset, "PtgAttr offset")
}

fn patch_skip_to_end(output: &mut [u8], position: usize) -> Result<()> {
    let remaining = output
        .len()
        .checked_sub(position + 4)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo position exceeds formula".to_string()))?;
    let offset = remaining
        .checked_sub(1)
        .ok_or_else(|| Error::InvalidFormula("PtgAttrGoTo has no following token".to_string()))?;
    patch_attribute_offset(output, position, offset)
}

fn patch_u16(output: &mut [u8], position: usize, value: usize, context: &str) -> Result<()> {
    let value = u16::try_from(value)
        .map_err(|_| Error::InvalidFormula(format!("{context} exceeds 65,535 bytes")))?;
    let target = output
        .get_mut(position..position + 2)
        .ok_or_else(|| Error::InvalidFormula(format!("{context} position is outside formula")))?;
    target.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn validate_xnum(value: f64, context: &str) -> Result<()> {
    if !value.is_finite()
        || (value == 0.0 && value.is_sign_negative())
        || (value != 0.0 && !value.is_normal())
    {
        return Err(Error::InvalidFormula(format!(
            "{context} contains a non-finite, denormalized, or negative-zero Xnum"
        )));
    }
    Ok(())
}

pub(crate) const FORMULA_ERRORS: &[(&str, u8)] = &[
    ("#GETTING_DATA", 0x2B),
    ("#DIV/0!", 0x07),
    ("#VALUE!", 0x0F),
    ("#NULL!", 0x00),
    ("#NAME?", 0x1D),
    ("#REF!", 0x17),
    ("#NUM!", 0x24),
    ("#N/A", 0x2A),
];

fn formula_error_code(value: &str) -> Option<u8> {
    FORMULA_ERRORS
        .iter()
        .find_map(|(literal, code)| literal.eq_ignore_ascii_case(value).then_some(*code))
}

pub(crate) fn excel_name_eq(left: &str, right: &str) -> bool {
    left.chars()
        .flat_map(char::to_lowercase)
        .eq(right.chars().flat_map(char::to_lowercase))
}

fn parse_a1_reference(value: &str) -> Option<A1Reference> {
    let bytes = value.as_bytes();
    let mut offset = 0;
    let col_relative = bytes.get(offset) != Some(&b'$');
    if !col_relative {
        offset += 1;
    }
    let col_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_alphabetic) {
        offset += 1;
    }
    if offset == col_start {
        return None;
    }
    let mut col = 0u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
    }
    if col == 0 || col > 16_384 {
        return None;
    }

    let row_relative = bytes.get(offset) != Some(&b'$');
    if !row_relative {
        offset += 1;
    }
    let row_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_digit) {
        offset += 1;
    }
    if offset == row_start || offset != bytes.len() {
        return None;
    }
    let row = value[row_start..offset].parse::<u32>().ok()?;
    if row == 0 || row > 1_048_576 {
        return None;
    }
    Some(A1Reference {
        row: row - 1,
        col: col - 1,
        row_relative,
        col_relative,
    })
}

fn reference_column_bits(reference: A1Reference) -> u16 {
    let mut bits = reference.col as u16;
    if reference.col_relative {
        bits |= 0x4000;
    }
    if reference.row_relative {
        bits |= 0x8000;
    }
    bits
}

fn emit_reference(output: &mut Vec<u8>, token: u8, reference: A1Reference) {
    output.push(token);
    output.extend_from_slice(&reference.row.to_le_bytes());
    output.extend_from_slice(&reference_column_bits(reference).to_le_bytes());
}

fn emit_shared_reference(
    output: &mut Vec<u8>,
    token: u8,
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<()> {
    let (row, col) = encode_shared_reference(reference, base_row, base_col)?;
    output.push(token);
    output.extend_from_slice(&row.to_le_bytes());
    output.extend_from_slice(&col.to_le_bytes());
    Ok(())
}

fn encode_shared_reference(
    reference: A1Reference,
    base_row: u32,
    base_col: u32,
) -> Result<(u32, u16)> {
    let row = if reference.row_relative {
        let offset = i64::from(reference.row) - i64::from(base_row);
        i32::try_from(offset)
            .map_err(|_| Error::InvalidFormula("shared row offset overflow".to_string()))?
            as u32
    } else {
        reference.row
    };
    let col_value = if reference.col_relative {
        let offset = i64::from(reference.col) - i64::from(base_col);
        if !(-16_383..=16_383).contains(&offset) {
            return Err(Error::InvalidFormula(format!(
                "shared column offset {offset} is outside the XLSB range"
            )));
        }
        (offset as i32 as u16) & 0x3FFF
    } else {
        reference.col as u16
    };
    let mut col = col_value;
    if reference.col_relative {
        col |= 0x4000;
    }
    if reference.row_relative {
        col |= 0x8000;
    }
    Ok((row, col))
}
