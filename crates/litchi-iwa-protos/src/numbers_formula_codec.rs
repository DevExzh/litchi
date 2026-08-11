//! Strict generated-free streaming reader for the table-local scalar formula subset.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Wire helpers stay beside the generated-free formula model."
)]

use core::{fmt, str};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_formula_generated::LitchiIwaFormulaProjection as projection;

const MAX_DEPTH: u32 = 32;
const MAX_FIELD_NUMBER: u32 = 0x1fff_ffff;

/// Finite aggregate policy for both the preflight and callback passes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_bytes: usize,
    max_fields: usize,
    max_work: usize,
    recursion_limit: u32,
    max_nodes: usize,
    max_text_bytes: usize,
}

impl DecodeOptions {
    #[must_use]
    pub const fn new(
        max_bytes: usize,
        max_fields: usize,
        max_work: usize,
        recursion_limit: u32,
        max_nodes: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_bytes,
            max_fields,
            max_work,
            recursion_limit,
            max_nodes,
            max_text_bytes,
        }
    }
    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }
}

/// Caller-authorized table-local formula context.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct FormulaContext {
    owner: u32,
    host_row: u32,
    host_column: u32,
    rows: u32,
    columns: u32,
}

impl FormulaContext {
    #[must_use]
    pub const fn new(owner: u32, host_row: u32, host_column: u32, rows: u32, columns: u32) -> Self {
        Self {
            owner,
            host_row,
            host_column,
            rows,
            columns,
        }
    }
}

/// Supported postfix binary operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
}

/// One validated table-local cell coordinate.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellCoordinate {
    row: u32,
    column: u32,
}

impl CellCoordinate {
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// One source-ordered evaluator node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FormulaNode {
    Binary(BinaryOperator),
    Negation,
    PlusSign,
    Percent,
    Function {
        identifier: u32,
        argument_count: u32,
    },
    Number {
        bits: u64,
    },
    Boolean(bool),
    Empty,
    Token(bool),
    LocalCell {
        coordinate: CellCoordinate,
        row_is_sticky: u32,
        column_is_sticky: u32,
    },
    CellReference {
        coordinate: CellCoordinate,
    },
    Colon,
    ColonWithUids,
    AppendWhitespace,
    PrependWhitespace,
}

/// Owner-aware local precedent emitted beside a reference node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct LocalPrecedent {
    owner: u32,
    coordinate: CellCoordinate,
}

/// Canonical table-local node that the scalar evaluator cannot execute.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct UnsupportedLocal {
    node_type: u32,
    function_identifier: Option<u32>,
}

impl UnsupportedLocal {
    #[must_use]
    pub const fn node_type(self) -> u32 {
        self.node_type
    }
    #[must_use]
    pub const fn function_identifier(self) -> Option<u32> {
        self.function_identifier
    }
}

impl LocalPrecedent {
    #[must_use]
    pub const fn owner(self) -> u32 {
        self.owner
    }
    #[must_use]
    pub const fn coordinate(self) -> CellCoordinate {
        self.coordinate
    }
}

/// Fallible source-order sink. Observations must be discarded on decode error.
pub trait FormulaVisitor {
    fn visit_node(&mut self, _node: FormulaNode) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_unsupported_local(&mut self, _node: UnsupportedLocal) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl FormulaVisitor for () {}

/// Dependency-only sink. Supported evaluator nodes are intentionally omitted.
pub trait FormulaDependencyVisitor {
    fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
        Ok(())
    }
    fn visit_unsupported_local(&mut self, _node: UnsupportedLocal) -> Result<(), DecodeError> {
        Ok(())
    }
}

impl FormulaDependencyVisitor for () {}

/// Exact successful aggregate report.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeReport {
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    text_bytes: usize,
    node_count: usize,
    precedent_count: usize,
    unsupported_local_count: usize,
    allocations: usize,
}

macro_rules! accessors {
    ($(($name:ident, $ty:ty)),+ $(,)?) => {$(
        #[must_use]
        pub const fn $name(self) -> $ty { self.$name }
    )+};
}

impl DecodeReport {
    accessors!(
        (bytes, usize),
        (fields, usize),
        (work, usize),
        (max_depth, u32),
        (text_bytes, usize),
        (node_count, usize),
        (precedent_count, usize),
        (unsupported_local_count, usize),
        (allocations, usize)
    );
}

/// Typed aggregate refusal.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DecodeLimit {
    Bytes { observed: usize, maximum: usize },
    Fields { observed: usize, maximum: usize },
    Work { observed: usize, maximum: usize },
    Nesting { observed: u32, maximum: u32 },
    Nodes { observed: usize, maximum: usize },
    Text { observed: usize, maximum: usize },
}

/// Content-free malformed or unsupported classification.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum InvalidReason {
    MalformedWire,
    MissingRequired,
    DuplicateField,
    UnexpectedField,
    InvalidCoordinate,
    UnsupportedFormula,
    ExternalOwner,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    limit: Option<DecodeLimit>,
    reason: Option<InvalidReason>,
}

impl DecodeError {
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        self.limit
    }
    #[must_use]
    pub const fn invalid_reason(&self) -> Option<InvalidReason> {
        self.reason
    }
    const fn invalid(reason: InvalidReason) -> Self {
        Self {
            limit: None,
            reason: Some(reason),
        }
    }
    const fn limited(limit: DecodeLimit) -> Self {
        Self {
            limit: Some(limit),
            reason: None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or unsupported FormulaArchive")
    }
}

impl std::error::Error for DecodeError {}

/// Strict allocation-free inspection for sizing an evaluator stack before
/// requesting callbacks. Its report covers this one complete validation pass.
pub fn inspect_formula_archive(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Evaluator,
    )?;
    Ok(budget.report())
}

/// Strict dependency-only inspection for global AST-to-edge proof.
///
/// Canonical evaluator-unsupported local nodes are tagged rather than refused;
/// malformed, noncanonical, and external-owner forms remain hard failures.
pub fn inspect_formula_dependencies_with_visitor<V: FormulaDependencyVisitor>(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Dependencies,
    )?;
    budget.preflight_callback_pass()?;
    let mut adapter = DependencyAdapter(visitor);
    decode_formula(
        source,
        context,
        &mut budget,
        &mut adapter,
        true,
        DecodeMode::Dependencies,
    )?;
    Ok(budget.report())
}

struct DependencyAdapter<'visitor, V>(&'visitor mut V);

impl<V: FormulaDependencyVisitor> FormulaVisitor for DependencyAdapter<'_, V> {
    fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
        self.0.visit_precedent(precedent)
    }
    fn visit_unsupported_local(&mut self, node: UnsupportedLocal) -> Result<(), DecodeError> {
        self.0.visit_unsupported_local(node)
    }
}

/// Strictly preflight then stream one table-local FormulaArchive.
pub fn decode_formula_archive_with_visitor<V: FormulaVisitor>(
    source: &[u8],
    context: FormulaContext,
    options: DecodeOptions,
    visitor: &mut V,
) -> Result<DecodeReport, DecodeError> {
    validate_context(context)?;
    let mut budget = Budget::new(source, options)?;
    decode_formula(
        source,
        context,
        &mut budget,
        &mut (),
        false,
        DecodeMode::Evaluator,
    )?;
    budget.preflight_callback_pass()?;
    decode_formula(
        source,
        context,
        &mut budget,
        visitor,
        true,
        DecodeMode::Evaluator,
    )?;
    Ok(budget.report())
}

fn validate_context(context: FormulaContext) -> Result<(), DecodeError> {
    if context.owner == 0
        || context.rows == 0
        || context.columns == 0
        || context.host_row >= context.rows
        || context.host_column >= context.columns
    {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DecodeMode {
    Evaluator,
    Dependencies,
}

fn decode_formula<V: FormulaVisitor>(
    source: &[u8],
    context: FormulaContext,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
) -> Result<(), DecodeError> {
    budget.message(source, 1)?;
    let mut ast = None;
    let mut host_column = None;
    let mut host_row = None;
    let mut column_negative = None;
    let mut row_negative = None;
    let mut opaque = [None; 4];
    let mut selected = [false; 9];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, 1)? {
        if !(1..=9).contains(&field.number) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        singular(&mut selected, field.number as usize - 1)?;
        match field.number {
            1 => ast = Some(field.bytes()?),
            2 => host_column = Some(field.varint_u32()?),
            3 => host_row = Some(field.varint_u32()?),
            4 => column_negative = Some(field.boolean()?),
            5 => row_negative = Some(field.boolean()?),
            6..=9 => opaque[field.number as usize - 6] = Some(field.bytes()?),
            _ => unreachable!(),
        }
    }
    let ast = ast.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    budget.message(source, 1)?;
    let view: projection::FormulaArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_ast_node_array()
        || view.ast_node_array != ast
        || view.host_column != host_column
        || view.host_row != host_row
        || view.host_column_is_negative != column_negative
        || view.host_row_is_negative != row_negative
        || view.translation_flags != opaque[0]
        || view.host_table_uid != opaque[1]
        || view.host_column_uid != opaque[2]
        || view.host_row_uid != opaque[3]
    {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    decode_node_array(ast, context, budget, visitor, emit, mode, 2)
}

fn decode_node_array<V: FormulaVisitor>(
    source: &[u8],
    context: FormulaContext,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let child = depth + 1;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number != 1 {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        if !emit {
            budget.node()?;
        }
        decode_node(field.bytes()?, context, budget, visitor, emit, mode, child)?;
    }
    Ok(())
}

#[derive(Default)]
struct NodeFields<'source> {
    kind: Option<u32>,
    function: Option<u32>,
    arguments: Option<u32>,
    number: Option<u64>,
    boolean: Option<bool>,
    token: Option<bool>,
    local: Option<&'source [u8]>,
    cross: Option<&'source [u8]>,
    column: Option<&'source [u8]>,
    row: Option<&'source [u8]>,
    cross_extra: Option<&'source [u8]>,
    uid: Option<&'source [u8]>,
    tract: Option<&'source [u8]>,
    whitespace: Option<&'source str>,
    unmodeled: bool,
    present: u64,
}

fn decode_node<V: FormulaVisitor>(
    source: &[u8],
    context: FormulaContext,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
    mode: DecodeMode,
    depth: u32,
) -> Result<(), DecodeError> {
    budget.message(source, depth)?;
    let mut fields = NodeFields::default();
    let mut seen = [false; 48];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if field.number == 0 || field.number > 47 {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        singular(&mut seen, field.number as usize)?;
        fields.present |= field_bit(field.number);
        match field.number {
            1 => fields.kind = Some(field.varint_u32()?),
            2 => fields.function = Some(field.varint_u32()?),
            3 => fields.arguments = Some(field.varint_u32()?),
            4 => fields.number = Some(field.fixed64()?),
            5 => fields.boolean = Some(field.boolean()?),
            10 => fields.token = Some(field.boolean()?),
            14 => {
                let _thunk = field.bytes()?;
                fields.unmodeled = true;
            },
            15 => fields.local = Some(field.bytes()?),
            16 => fields.cross = Some(field.bytes()?),
            25 => {
                let text = strict_utf8(field.bytes()?)?;
                budget.text(text.len())?;
                fields.whitespace = Some(text);
            },
            26 => fields.column = Some(field.bytes()?),
            27 => fields.row = Some(field.bytes()?),
            28 => fields.cross_extra = Some(field.bytes()?),
            30 => fields.uid = Some(field.bytes()?),
            38 => fields.tract = Some(field.bytes()?),
            _ => {
                validate_unmodeled_field(field, budget)?;
                fields.unmodeled = true;
            },
        }
    }
    let kind = fields
        .kind
        .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?;
    budget.message(source, depth)?;
    let view: projection::ASTNodeArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_node_type()
        || u32::try_from(view.node_type).ok() != Some(kind)
        || view.function_index != fields.function
        || view.function_num_args != fields.arguments
        || view.number.map(f64::to_bits) != fields.number
        || view.boolean != fields.boolean
        || view.token_boolean != fields.token
        || view.local_cell_reference != fields.local
        || view.cross_table_cell_reference != fields.cross
        || view.whitespace != fields.whitespace
        || view.column != fields.column
        || view.row != fields.row
        || view.cross_table_extra != fields.cross_extra
        || view.uid_coordinate != fields.uid
        || view.tract_list != fields.tract
    {
        return Err(DecodeError::invalid(InvalidReason::MalformedWire));
    }
    if fields.cross.is_some()
        || fields.cross_extra.is_some()
        || fields.uid.is_some()
        || matches!(kind, 28 | 48 | 63..=68)
    {
        return Err(DecodeError::invalid(InvalidReason::ExternalOwner));
    }
    let unsupported_function = kind == 16
        && fields
            .function
            .is_some_and(|identifier| !matches!(identifier, 15 | 30 | 84 | 88 | 168));
    let present = fields.present;
    if fields.unmodeled || (fields.tract.is_some() && kind != 45) || unsupported_function {
        if mode == DecodeMode::Dependencies
            && dependency_tolerates(kind, present, unsupported_function)
        {
            return unsupported_local(kind, fields.function, mode, budget, visitor, emit);
        }
        return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula));
    }
    let (node, precedent) = match classify_node(kind, fields, context, budget, depth + 1) {
        Ok(value) => value,
        Err(error)
            if mode == DecodeMode::Dependencies
                && error.invalid_reason() == Some(InvalidReason::UnsupportedFormula) =>
        {
            if dependency_tolerates(kind, present, false) {
                return unsupported_local(kind, None, mode, budget, visitor, emit);
            }
            return Err(error);
        },
        Err(error) => return Err(error),
    };
    if emit {
        visitor.visit_node(node)?;
        if let Some(precedent) = precedent {
            visitor.visit_precedent(precedent)?;
        }
    } else if precedent.is_some() {
        budget.precedent()?;
    }
    Ok(())
}

const fn field_bit(number: u32) -> u64 {
    1u64 << number
}

fn dependency_tolerates(kind: u32, present: u64, unsupported_function: bool) -> bool {
    let whitespace = field_bit(1) | field_bit(25);
    let allowed = match kind {
        16 if unsupported_function => whitespace | field_bit(2) | field_bit(3),
        6 | 30 | 34 | 35 | 54 | 56 | 57 | 69 | 70 => whitespace,
        19 => whitespace | field_bit(6),
        17 => whitespace | field_bit(4) | field_bit(42) | field_bit(43),
        20 => whitespace | field_bit(7) | field_bit(19) | field_bit(20) | field_bit(21),
        21 => {
            whitespace
                | field_bit(8)
                | field_bit(9)
                | field_bit(22)
                | field_bit(23)
                | field_bit(24)
                | field_bit(29)
        },
        24 => whitespace | field_bit(11) | field_bit(12),
        25 => whitespace | field_bit(13),
        31 => whitespace | field_bit(17) | field_bit(18),
        52 => whitespace | field_bit(34) | field_bit(35) | field_bit(36),
        53 => whitespace | field_bit(37),
        _ => return false,
    };
    present & !allowed == 0
}

fn unsupported_local<V: FormulaVisitor>(
    node_type: u32,
    function_identifier: Option<u32>,
    mode: DecodeMode,
    budget: &mut Budget,
    visitor: &mut V,
    emit: bool,
) -> Result<(), DecodeError> {
    if mode == DecodeMode::Evaluator {
        return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula));
    }
    let unsupported = UnsupportedLocal {
        node_type,
        function_identifier,
    };
    if emit {
        visitor.visit_unsupported_local(unsupported)
    } else {
        budget.unsupported_local()
    }
}

fn classify_node(
    kind: u32,
    fields: NodeFields<'_>,
    context: FormulaContext,
    budget: &mut Budget,
    depth: u32,
) -> Result<(FormulaNode, Option<LocalPrecedent>), DecodeError> {
    const FUNCTION: u16 = 1 << 0;
    const ARGUMENTS: u16 = 1 << 1;
    const NUMBER: u16 = 1 << 2;
    const BOOLEAN: u16 = 1 << 3;
    const TOKEN: u16 = 1 << 4;
    const LOCAL: u16 = 1 << 5;
    const COLUMN: u16 = 1 << 6;
    const ROW: u16 = 1 << 7;
    const TRACT: u16 = 1 << 8;
    let payload = present(fields.function.is_some(), FUNCTION)
        | present(fields.arguments.is_some(), ARGUMENTS)
        | present(fields.number.is_some(), NUMBER)
        | present(fields.boolean.is_some(), BOOLEAN)
        | present(fields.token.is_some(), TOKEN)
        | present(fields.local.is_some(), LOCAL)
        | present(fields.column.is_some(), COLUMN)
        | present(fields.row.is_some(), ROW)
        | present(fields.tract.is_some(), TRACT);
    let no_payload = || payload == 0;
    let binary = match kind {
        1 => Some(BinaryOperator::Add),
        2 => Some(BinaryOperator::Subtract),
        3 => Some(BinaryOperator::Multiply),
        4 => Some(BinaryOperator::Divide),
        5 => Some(BinaryOperator::Power),
        7 => Some(BinaryOperator::GreaterThan),
        8 => Some(BinaryOperator::GreaterThanOrEqual),
        9 => Some(BinaryOperator::LessThan),
        10 => Some(BinaryOperator::LessThanOrEqual),
        11 => Some(BinaryOperator::Equal),
        12 => Some(BinaryOperator::NotEqual),
        _ => None,
    };
    if let Some(operator) = binary {
        require(no_payload())?;
        return Ok((FormulaNode::Binary(operator), None));
    }
    let (node, coordinate) = match kind {
        13 if no_payload() => (FormulaNode::Negation, None),
        14 if no_payload() => (FormulaNode::PlusSign, None),
        15 if no_payload() => (FormulaNode::Percent, None),
        16 => (
            FormulaNode::Function {
                identifier: fields
                    .function
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                argument_count: fields.arguments.unwrap_or(0),
            },
            None,
        ),
        17 => (
            FormulaNode::Number {
                bits: fields
                    .number
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            },
            None,
        ),
        18 => (
            FormulaNode::Boolean(
                fields
                    .boolean
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            ),
            None,
        ),
        22 if no_payload() => (FormulaNode::Empty, None),
        23 => (
            FormulaNode::Token(
                fields
                    .token
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
            ),
            None,
        ),
        27 => {
            let local = decode_local(
                fields
                    .local
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                context,
                budget,
                depth,
            )?;
            (
                FormulaNode::LocalCell {
                    coordinate: local.0,
                    row_is_sticky: local.1,
                    column_is_sticky: local.2,
                },
                Some(local.0),
            )
        },
        29 if no_payload() => (FormulaNode::Colon, None),
        32 => (FormulaNode::AppendWhitespace, None),
        33 => (FormulaNode::PrependWhitespace, None),
        36 => {
            let row = decode_axis(
                fields
                    .row
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                false,
                budget,
                depth,
            )?;
            let column = decode_axis(
                fields
                    .column
                    .ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))?,
                true,
                budget,
                depth,
            )?;
            let coordinate = CellCoordinate {
                row: resolve(context.host_row, row.0, row.1)?,
                column: resolve(context.host_column, column.0, column.1)?,
            };
            validate_coordinate(coordinate, context)?;
            (FormulaNode::CellReference { coordinate }, Some(coordinate))
        },
        45 if fields.tract.is_none() => (FormulaNode::ColonWithUids, None),
        28 | 48 | 63..=68 => return Err(DecodeError::invalid(InvalidReason::ExternalOwner)),
        _ => return Err(DecodeError::invalid(InvalidReason::UnsupportedFormula)),
    };
    let allowed = match kind {
        16 => FUNCTION | ARGUMENTS,
        17 => NUMBER,
        18 => BOOLEAN,
        23 => TOKEN,
        27 => LOCAL,
        36 => COLUMN | ROW,
        45 => TRACT,
        _ => 0,
    };
    require(payload & !allowed == 0)?;
    let precedent = coordinate.map(|coordinate| LocalPrecedent {
        owner: context.owner,
        coordinate,
    });
    Ok((node, precedent))
}

fn decode_local(
    source: &[u8],
    context: FormulaContext,
    budget: &mut Budget,
    depth: u32,
) -> Result<(CellCoordinate, u32, u32), DecodeError> {
    budget.message(source, depth)?;
    let mut values = [None; 4];
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        if !(1..=4).contains(&field.number) {
            return Err(DecodeError::invalid(InvalidReason::UnexpectedField));
        }
        let slot = &mut values[field.number as usize - 1];
        set_once(slot, field.varint_u32()?)?;
    }
    let coordinate = CellCoordinate {
        row: required(values[0])?,
        column: required(values[1])?,
    };
    let row_is_sticky = required(values[2])?;
    let column_is_sticky = required(values[3])?;
    validate_coordinate(coordinate, context)?;
    budget.message(source, depth)?;
    let view: projection::LocalCellReferenceArchiveLazyView<'_> = budget
        .options
        .buffa()
        .decode_lazy_view(source)
        .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
    if !view.has_row_handle()
        || !view.has_column_handle()
        || !view.has_row_is_sticky()
        || !view.has_column_is_sticky()
        || view.row_handle != coordinate.row
        || view.column_handle != coordinate.column
        || view.row_is_sticky != row_is_sticky
        || view.column_is_sticky != column_is_sticky
    {
        return Err(DecodeError::invalid(InvalidReason::MissingRequired));
    }
    Ok((coordinate, row_is_sticky, column_is_sticky))
}

fn decode_axis(
    source: &[u8],
    column: bool,
    budget: &mut Budget,
    depth: u32,
) -> Result<(i32, bool), DecodeError> {
    budget.message(source, depth)?;
    let mut coordinate = None;
    let mut absolute = None;
    let mut remaining = source;
    while let Some(field) = next_field(&mut remaining, budget, depth)? {
        match field.number {
            1 => set_once(&mut coordinate, zigzag32(field.varint_u32()?))?,
            2 => set_once(&mut absolute, field.boolean()?)?,
            _ => return Err(DecodeError::invalid(InvalidReason::UnexpectedField)),
        }
    }
    let coordinate = required(coordinate)?;
    budget.message(source, depth)?;
    if column {
        let view: projection::ColumnCoordinateArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if !view.has_column() || view.column != coordinate || view.absolute != absolute {
            return Err(DecodeError::invalid(InvalidReason::MalformedWire));
        }
    } else {
        let view: projection::RowCoordinateArchiveLazyView<'_> = budget
            .options
            .buffa()
            .decode_lazy_view(source)
            .map_err(|_error| DecodeError::invalid(InvalidReason::MalformedWire))?;
        if !view.has_row() || view.row != coordinate || view.absolute != absolute {
            return Err(DecodeError::invalid(InvalidReason::MalformedWire));
        }
    }
    Ok((coordinate, absolute.unwrap_or(false)))
}

fn resolve(host: u32, coordinate: i32, absolute: bool) -> Result<u32, DecodeError> {
    let value = if absolute {
        i64::from(coordinate)
    } else {
        i64::from(host) + i64::from(coordinate)
    };
    u32::try_from(value).map_err(|_error| DecodeError::invalid(InvalidReason::InvalidCoordinate))
}

fn validate_coordinate(value: CellCoordinate, context: FormulaContext) -> Result<(), DecodeError> {
    if value.row >= context.rows || value.column >= context.columns {
        return Err(DecodeError::invalid(InvalidReason::InvalidCoordinate));
    }
    Ok(())
}

fn require(condition: bool) -> Result<(), DecodeError> {
    if condition {
        Ok(())
    } else {
        Err(DecodeError::invalid(InvalidReason::UnexpectedField))
    }
}

const fn present(value: bool, mask: u16) -> u16 {
    if value { mask } else { 0 }
}

fn required<T: Copy>(value: Option<T>) -> Result<T, DecodeError> {
    value.ok_or_else(|| DecodeError::invalid(InvalidReason::MissingRequired))
}

fn singular(seen: &mut [bool], index: usize) -> Result<(), DecodeError> {
    let value = seen
        .get_mut(index)
        .ok_or_else(|| DecodeError::invalid(InvalidReason::UnexpectedField))?;
    if *value {
        return Err(DecodeError::invalid(InvalidReason::DuplicateField));
    }
    *value = true;
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T) -> Result<(), DecodeError> {
    if slot.is_some() {
        return Err(DecodeError::invalid(InvalidReason::DuplicateField));
    }
    *slot = Some(value);
    Ok(())
}

fn validate_unmodeled_field(field: Field<'_>, budget: &mut Budget) -> Result<(), DecodeError> {
    match field.number {
        9 | 11..=13 | 18 | 22..=24 | 37 | 42..=43 | 46..=47 => {
            let _value = field.varint()?;
        },
        19..=20 | 29 | 36 => {
            let _value = field.boolean()?;
        },
        7..=8 => {
            let _value = field.fixed64()?;
        },
        6 | 17 | 21 | 34..=35 => {
            let value = strict_utf8(field.bytes()?)?;
            budget.text(value.len())?;
        },
        33 | 39..=41 | 44..=45 => {
            let _value = field.bytes()?;
        },
        _ => return Err(DecodeError::invalid(InvalidReason::UnexpectedField)),
    }
    Ok(())
}

#[derive(Clone)]
struct Budget {
    options: DecodeOptions,
    bytes: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    text: usize,
    nodes: usize,
    precedents: usize,
    unsupported_local: usize,
}

impl Budget {
    fn new(source: &[u8], options: DecodeOptions) -> Result<Self, DecodeError> {
        if source.len() > options.max_bytes {
            return Err(DecodeError::limited(DecodeLimit::Bytes {
                observed: source.len(),
                maximum: options.max_bytes,
            }));
        }
        if options.recursion_limit == 0 || options.recursion_limit > MAX_DEPTH {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_DEPTH,
            }));
        }
        Ok(Self {
            options,
            bytes: source.len(),
            fields: 0,
            work: 0,
            max_depth: 0,
            text: 0,
            nodes: 0,
            precedents: 0,
            unsupported_local: 0,
        })
    }
    fn message(&mut self, source: &[u8], depth: u32) -> Result<(), DecodeError> {
        self.depth(depth)?;
        self.work(source.len())
    }
    fn field(&mut self) -> Result<(), DecodeError> {
        self.fields = checked(self.fields, 1)?;
        if self.fields > self.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: self.fields,
                maximum: self.options.max_fields,
            }));
        }
        Ok(())
    }
    fn work(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.work = checked(self.work, amount)?;
        if self.work > self.options.max_work {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: self.work,
                maximum: self.options.max_work,
            }));
        }
        Ok(())
    }
    fn depth(&mut self, depth: u32) -> Result<(), DecodeError> {
        if depth > self.options.recursion_limit {
            return Err(DecodeError::limited(DecodeLimit::Nesting {
                observed: depth,
                maximum: self.options.recursion_limit,
            }));
        }
        self.max_depth = self.max_depth.max(depth);
        Ok(())
    }
    fn node(&mut self) -> Result<(), DecodeError> {
        self.nodes = checked(self.nodes, 1)?;
        if self.nodes > self.options.max_nodes {
            return Err(DecodeError::limited(DecodeLimit::Nodes {
                observed: self.nodes,
                maximum: self.options.max_nodes,
            }));
        }
        Ok(())
    }
    fn precedent(&mut self) -> Result<(), DecodeError> {
        self.precedents = checked(self.precedents, 1)?;
        Ok(())
    }
    fn unsupported_local(&mut self) -> Result<(), DecodeError> {
        self.unsupported_local = checked(self.unsupported_local, 1)?;
        Ok(())
    }
    fn text(&mut self, amount: usize) -> Result<(), DecodeError> {
        self.text = checked(self.text, amount)?;
        if self.text > self.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: self.text,
                maximum: self.options.max_text_bytes,
            }));
        }
        Ok(())
    }
    fn preflight_callback_pass(&self) -> Result<(), DecodeError> {
        let fields = self.fields.checked_mul(2);
        let work = self.work.checked_mul(2);
        let text = self.text.checked_mul(2);
        let probe = Self {
            options: self.options,
            bytes: self.bytes,
            fields: fields.ok_or_else(malformed)?,
            work: work.ok_or_else(malformed)?,
            max_depth: self.max_depth,
            text: text.ok_or_else(malformed)?,
            nodes: self.nodes,
            precedents: self.precedents,
            unsupported_local: self.unsupported_local,
        };
        if probe.fields > probe.options.max_fields {
            return Err(DecodeError::limited(DecodeLimit::Fields {
                observed: probe.fields,
                maximum: probe.options.max_fields,
            }));
        }
        if probe.work > probe.options.max_work {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: probe.work,
                maximum: probe.options.max_work,
            }));
        }
        if probe.text > probe.options.max_text_bytes {
            return Err(DecodeError::limited(DecodeLimit::Text {
                observed: probe.text,
                maximum: probe.options.max_text_bytes,
            }));
        }
        Ok(())
    }
    const fn report(&self) -> DecodeReport {
        DecodeReport {
            bytes: self.bytes,
            fields: self.fields,
            work: self.work,
            max_depth: self.max_depth,
            text_bytes: self.text,
            node_count: self.nodes,
            precedent_count: self.precedents,
            unsupported_local_count: self.unsupported_local,
            allocations: 0,
        }
    }
}

fn checked(left: usize, right: usize) -> Result<usize, DecodeError> {
    left.checked_add(right).ok_or_else(malformed)
}

const fn malformed() -> DecodeError {
    DecodeError::invalid(InvalidReason::MalformedWire)
}

#[derive(Clone, Copy)]
struct Field<'source> {
    number: u32,
    wire: u8,
    value: Value<'source>,
}

#[derive(Clone, Copy)]
enum Value<'source> {
    Varint(u64),
    Fixed64(u64),
    Bytes(&'source [u8]),
    Fixed32,
}

impl Field<'_> {
    fn varint(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Varint(value) if self.wire == 0 => Ok(value),
            _ => Err(malformed()),
        }
    }
    fn varint_u32(self) -> Result<u32, DecodeError> {
        u32::try_from(self.varint()?).map_err(|_error| malformed())
    }
    fn boolean(self) -> Result<bool, DecodeError> {
        match self.varint()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => Err(malformed()),
        }
    }
    fn fixed64(self) -> Result<u64, DecodeError> {
        match self.value {
            Value::Fixed64(value) if self.wire == 1 => Ok(value),
            _ => Err(malformed()),
        }
    }
}

impl<'source> Field<'source> {
    fn bytes(self) -> Result<&'source [u8], DecodeError> {
        match self.value {
            Value::Bytes(value) if self.wire == 2 => Ok(value),
            _ => Err(malformed()),
        }
    }
}

fn next_field<'source>(
    source: &mut &'source [u8],
    budget: &mut Budget,
    depth: u32,
) -> Result<Option<Field<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    budget.depth(depth)?;
    budget.field()?;
    let tag = take_varint(source)?;
    let number = u32::try_from(tag >> 3).map_err(|_error| malformed())?;
    let wire = u8::try_from(tag & 7).map_err(|_error| malformed())?;
    if number == 0 || number > MAX_FIELD_NUMBER {
        return Err(malformed());
    }
    let value = match wire {
        0 => Value::Varint(take_varint(source)?),
        1 => {
            let raw = take(source, 8)?;
            Value::Fixed64(u64::from_le_bytes(
                raw.try_into().map_err(|_error| malformed())?,
            ))
        },
        2 => {
            let length = usize::try_from(take_varint(source)?).map_err(|_error| malformed())?;
            Value::Bytes(take(source, length)?)
        },
        5 => {
            let _raw = take(source, 4)?;
            Value::Fixed32
        },
        _ => return Err(malformed()),
    };
    Ok(Some(Field {
        number,
        wire,
        value,
    }))
}

fn take<'source>(source: &mut &'source [u8], amount: usize) -> Result<&'source [u8], DecodeError> {
    if source.len() < amount {
        return Err(malformed());
    }
    let (value, rest) = source.split_at(amount);
    *source = rest;
    Ok(value)
}

fn take_varint(source: &mut &[u8]) -> Result<u64, DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original.get(index).ok_or_else(malformed)?;
        if index == 9 && byte > 1 {
            return Err(malformed());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            if varint_len(value) != consumed {
                return Err(malformed());
            }
            *source = &original[consumed..];
            return Ok(value);
        }
    }
    Err(malformed())
}

const fn varint_len(value: u64) -> usize {
    if value == 0 {
        1
    } else {
        (64usize - value.leading_zeros() as usize).div_ceil(7)
    }
}

fn strict_utf8(source: &[u8]) -> Result<&str, DecodeError> {
    str::from_utf8(source).map_err(|_error| malformed())
}

const fn zigzag32(value: u32) -> i32 {
    ((value >> 1) as i32) ^ (-((value & 1) as i32))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn put_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = value as u8 & 0x7f;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    fn key(output: &mut Vec<u8>, field: u32, wire: u8) {
        put_varint(output, (u64::from(field) << 3) | u64::from(wire));
    }

    fn varint(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 0);
        put_varint(output, value);
    }

    fn bytes(output: &mut Vec<u8>, field: u32, value: &[u8]) {
        key(output, field, 2);
        put_varint(output, value.len() as u64);
        output.extend_from_slice(value);
    }

    fn fixed64(output: &mut Vec<u8>, field: u32, value: u64) {
        key(output, field, 1);
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn zigzag(value: i32) -> u32 {
        ((value << 1) ^ (value >> 31)) as u32
    }

    fn node(kind: u32, fields: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
        let mut output = Vec::new();
        varint(&mut output, 1, u64::from(kind));
        fields(&mut output);
        output
    }

    fn formula(nodes: &[Vec<u8>]) -> Vec<u8> {
        let mut ast = Vec::new();
        for node in nodes {
            bytes(&mut ast, 1, node);
        }
        let mut formula = Vec::new();
        bytes(&mut formula, 1, &ast);
        formula
    }

    fn context() -> FormulaContext {
        FormulaContext::new(7, 4, 5, 20, 20)
    }

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(source.len(), 2_000_000, 100_000_000, 16, 100_000, 1_000_000)
    }

    #[derive(Default)]
    struct Facts {
        nodes: Vec<FormulaNode>,
        precedents: Vec<LocalPrecedent>,
    }

    impl FormulaVisitor for Facts {
        fn visit_node(&mut self, node: FormulaNode) -> Result<(), DecodeError> {
            self.nodes.push(node);
            Ok(())
        }
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.precedents.push(precedent);
            Ok(())
        }
    }

    #[test]
    fn supported_postfix_nodes_and_owner_aware_precedents_stream_in_order() {
        let number = node(17, |node| fixed64(node, 4, 3.5f64.to_bits()));
        let local = node(27, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 1);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        });
        let relative = node(36, |node| {
            let mut row = Vec::new();
            varint(&mut row, 1, u64::from(zigzag(-1)));
            let mut column = Vec::new();
            varint(&mut column, 1, u64::from(zigzag(2)));
            bytes(node, 26, &column);
            bytes(node, 27, &row);
        });
        let add = node(1, |_| {});
        let function = node(16, |node| {
            varint(node, 2, 168);
            varint(node, 3, 2);
        });
        let source = formula(&[number, local, relative, add, function]);
        let mut facts = Facts::default();
        let report =
            decode_formula_archive_with_visitor(&source, context(), options(&source), &mut facts)
                .unwrap();
        assert_eq!(report.node_count(), 5);
        assert_eq!(report.precedent_count(), 2);
        assert_eq!(report.allocations(), 0);
        assert_eq!(
            facts.nodes[0],
            FormulaNode::Number {
                bits: 3.5f64.to_bits()
            }
        );
        assert_eq!(
            facts.precedents,
            vec![
                LocalPrecedent {
                    owner: 7,
                    coordinate: CellCoordinate { row: 2, column: 3 }
                },
                LocalPrecedent {
                    owner: 7,
                    coordinate: CellCoordinate { row: 3, column: 7 }
                }
            ]
        );
        assert_eq!(
            facts.nodes.last(),
            Some(&FormulaNode::Function {
                identifier: 168,
                argument_count: 2
            })
        );
    }

    fn reason(error: DecodeError) -> InvalidReason {
        error.invalid_reason().unwrap()
    }

    #[test]
    fn malformed_duplicate_unknown_and_external_forms_fail_closed() {
        let mut duplicate = node(17, |node| fixed64(node, 4, 1.0f64.to_bits()));
        varint(&mut duplicate, 1, 17);
        let duplicate = formula(&[duplicate]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &duplicate,
                    context(),
                    options(&duplicate),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::DuplicateField
        );

        let mut unknown = node(17, |node| fixed64(node, 4, 1.0f64.to_bits()));
        varint(&mut unknown, 48, 1);
        let unknown = formula(&[unknown]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &unknown,
                    context(),
                    options(&unknown),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::UnexpectedField
        );

        let wrong_wire = formula(&[node(19, |node| varint(node, 6, 1))]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &wrong_wire,
                    context(),
                    options(&wrong_wire),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let noncanonical = [0x0a, 0x02, 0x0a, 0x80];
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &noncanonical,
                    context(),
                    options(&noncanonical),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::MalformedWire
        );

        let cross = formula(&[node(28, |_| {})]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(&cross, context(), options(&cross), &mut ())
                    .unwrap_err()
            ),
            InvalidReason::ExternalOwner
        );
        let unsupported = formula(&[node(19, |_| {})]);
        assert_eq!(
            reason(
                decode_formula_archive_with_visitor(
                    &unsupported,
                    context(),
                    options(&unsupported),
                    &mut ()
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
    }

    #[derive(Default)]
    struct DependencyFacts {
        precedents: Vec<LocalPrecedent>,
        unsupported: Vec<UnsupportedLocal>,
        calls: usize,
    }

    impl FormulaDependencyVisitor for DependencyFacts {
        fn visit_precedent(&mut self, precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.calls += 1;
            self.precedents.push(precedent);
            Ok(())
        }
        fn visit_unsupported_local(&mut self, node: UnsupportedLocal) -> Result<(), DecodeError> {
            self.calls += 1;
            self.unsupported.push(node);
            Ok(())
        }
    }

    #[test]
    fn unsupported_local_unreachable_is_dependency_visible_but_impacted_evaluation_refuses() {
        let string = node(19, |node| bytes(node, 6, b"preserve"));
        let local = node(27, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        });
        let unknown_function = node(16, |node| {
            varint(node, 2, 999);
            varint(node, 3, 1);
        });
        let source = formula(&[string, local, unknown_function]);
        let mut facts = DependencyFacts::default();
        let report = inspect_formula_dependencies_with_visitor(
            &source,
            context(),
            options(&source),
            &mut facts,
        )
        .unwrap();
        assert_eq!(report.node_count(), 3);
        assert_eq!(report.precedent_count(), 1);
        assert_eq!(report.unsupported_local_count(), 2);
        assert_eq!(facts.precedents[0].owner(), 7);
        assert_eq!(
            facts.unsupported,
            vec![
                UnsupportedLocal {
                    node_type: 19,
                    function_identifier: None
                },
                UnsupportedLocal {
                    node_type: 16,
                    function_identifier: Some(999)
                }
            ]
        );
        assert_eq!(
            reason(inspect_formula_archive(&source, context(), options(&source)).unwrap_err()),
            InvalidReason::UnsupportedFormula
        );

        let external = formula(&[node(28, |_| {})]);
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &external,
                    context(),
                    options(&external),
                    &mut DependencyFacts::default()
                )
                .unwrap_err()
            ),
            InvalidReason::ExternalOwner
        );

        let hidden_local = formula(&[node(20, |node| {
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        })]);
        let mut hidden_facts = DependencyFacts::default();
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &hidden_local,
                    context(),
                    options(&hidden_local),
                    &mut hidden_facts
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
        assert_eq!(hidden_facts.calls, 0);
    }

    #[test]
    fn dependency_max_minus_one_preempts_unsupported_and_precedent_callbacks() {
        let source = formula(&[
            node(19, |node| bytes(node, 6, b"x")),
            node(17, |node| fixed64(node, 4, 1.0f64.to_bits())),
        ]);
        let report = inspect_formula_dependencies_with_visitor(
            &source,
            context(),
            options(&source),
            &mut DependencyFacts::default(),
        )
        .unwrap();
        for limited in [
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work() - 1,
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count() - 1,
                report.text_bytes(),
            ),
        ] {
            let mut facts = DependencyFacts::default();
            let error =
                inspect_formula_dependencies_with_visitor(&source, context(), limited, &mut facts)
                    .unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(facts.calls, 0);
        }
    }

    #[test]
    fn native_decimal_number_shape_is_local_but_reference_bearing_near_shape_refuses() {
        let decimal = formula(&[node(17, |node| {
            fixed64(node, 4, 1.25f64.to_bits());
            varint(node, 42, 1);
            varint(node, 43, 2);
        })]);
        let mut facts = DependencyFacts::default();
        let report = inspect_formula_dependencies_with_visitor(
            &decimal,
            context(),
            options(&decimal),
            &mut facts,
        )
        .unwrap();
        assert_eq!(report.unsupported_local_count(), 1);
        assert_eq!(facts.unsupported[0].node_type(), 17);
        assert_eq!(
            reason(inspect_formula_archive(&decimal, context(), options(&decimal)).unwrap_err()),
            InvalidReason::UnsupportedFormula
        );

        let hidden_reference = formula(&[node(17, |node| {
            fixed64(node, 4, 1.25f64.to_bits());
            varint(node, 42, 1);
            varint(node, 43, 2);
            let mut reference = Vec::new();
            varint(&mut reference, 1, 2);
            varint(&mut reference, 2, 3);
            varint(&mut reference, 3, 0);
            varint(&mut reference, 4, 0);
            bytes(node, 15, &reference);
        })]);
        let mut hidden_facts = DependencyFacts::default();
        assert_eq!(
            reason(
                inspect_formula_dependencies_with_visitor(
                    &hidden_reference,
                    context(),
                    options(&hidden_reference),
                    &mut hidden_facts
                )
                .unwrap_err()
            ),
            InvalidReason::UnsupportedFormula
        );
        assert_eq!(hidden_facts.calls, 0);
    }

    #[derive(Default)]
    struct Calls(usize);

    impl FormulaVisitor for Calls {
        fn visit_node(&mut self, _node: FormulaNode) -> Result<(), DecodeError> {
            self.0 += 1;
            Ok(())
        }
        fn visit_precedent(&mut self, _precedent: LocalPrecedent) -> Result<(), DecodeError> {
            self.0 += 1;
            Ok(())
        }
    }

    #[test]
    fn exact_limits_are_inclusive_and_max_minus_one_preempts_callbacks() {
        let source = formula(&[
            node(17, |node| fixed64(node, 4, 1.0f64.to_bits())),
            node(32, |node| bytes(node, 25, b" ")),
        ]);
        let report = decode_formula_archive_with_visitor(
            &source,
            context(),
            options(&source),
            &mut Calls::default(),
        )
        .unwrap();
        let inspection = inspect_formula_archive(&source, context(), options(&source)).unwrap();
        assert_eq!(inspection.node_count(), report.node_count());
        assert_eq!(inspection.fields() * 2, report.fields());
        assert_eq!(inspection.work() * 2, report.work());
        let exact = DecodeOptions::new(
            source.len(),
            report.fields(),
            report.work(),
            report.max_depth(),
            report.node_count(),
            report.text_bytes(),
        );
        assert!(
            decode_formula_archive_with_visitor(&source, context(), exact, &mut Calls::default())
                .is_ok()
        );
        let limited = [
            DecodeOptions::new(
                source.len() - 1,
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields() - 1,
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work() - 1,
                report.max_depth(),
                report.node_count(),
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count() - 1,
                report.text_bytes(),
            ),
            DecodeOptions::new(
                source.len(),
                report.fields(),
                report.work(),
                report.max_depth(),
                report.node_count(),
                report.text_bytes() - 1,
            ),
        ];
        for options in limited {
            let mut calls = Calls::default();
            let error =
                decode_formula_archive_with_visitor(&source, context(), options, &mut calls)
                    .unwrap_err();
            assert!(error.resource_limit().is_some());
            assert_eq!(calls.0, 0);
        }
    }

    #[test]
    fn four_thousand_to_eight_thousand_nodes_scale_below_two_point_two() {
        fn run(count: usize) -> DecodeReport {
            let nodes = (0..count)
                .map(|index| node(17, |node| fixed64(node, 4, (index as f64).to_bits())))
                .collect::<Vec<_>>();
            let source = formula(&nodes);
            decode_formula_archive_with_visitor(
                &source,
                context(),
                options(&source),
                &mut Calls::default(),
            )
            .unwrap()
        }
        let four = run(4_096);
        let eight = run(8_192);
        assert_eq!(eight.node_count(), four.node_count() * 2);
        assert!(eight.fields() * 100 <= four.fields() * 220);
        assert!(eight.work() * 100 <= four.work() * 220);
        assert!(eight.text_bytes() * 100 <= four.text_bytes().max(1) * 220);
        assert_eq!(eight.allocations(), 0);
    }
}
