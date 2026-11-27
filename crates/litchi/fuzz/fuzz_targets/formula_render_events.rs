#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_protos::numbers_formula_codec::{
    BinaryOperator, FormulaRenderAxis, FormulaRenderCategoryReference, FormulaRenderCellReference,
    FormulaRenderCfuuid, FormulaRenderColonTract, FormulaRenderCoordinatePair,
    FormulaRenderCrossTableCellReference, FormulaRenderEvent, FormulaRenderLocalCellReference,
    FormulaRenderRangeSummary, FormulaRenderStickyBits, FormulaRenderUuid,
};
use litchi_numbers_wire::formula_render::{
    CompatibilityFormulaVisitor, FormulaCategoryId, FormulaEventRenderBudget, FormulaTablePrefix,
    ReferenceResolver,
};

const MAX_EVENTS: usize = 96;
const MAX_OUTPUT_BYTES: usize = 1024;
const MAX_NODES: usize = 96;
const MAX_PARTS: usize = 384;
const MAX_RENDER_DEPTH: usize = 8;

#[derive(Debug, Clone, PartialEq, Eq)]
enum EventError {
    OutputLimit {
        observed: usize,
        maximum: usize,
    },
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    Invalid(&'static str),
    Structure {
        nodes: usize,
        parts: usize,
    },
    RenderDepth {
        observed: usize,
        maximum: usize,
    },
    Parse,
    InvalidFormat,
}

#[derive(Debug, Clone, Copy)]
struct EventBudget {
    maximum_output: usize,
    maximum_nodes: usize,
    maximum_parts: usize,
    maximum_render_depth: usize,
    charged_output: usize,
}

impl EventBudget {
    const fn new(
        maximum_output: usize,
        maximum_nodes: usize,
        maximum_parts: usize,
        maximum_render_depth: usize,
    ) -> Self {
        Self {
            maximum_output,
            maximum_nodes,
            maximum_parts,
            maximum_render_depth,
            charged_output: 0,
        }
    }
}

impl FormulaRenderBudget for EventBudget {
    type Error = EventError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        EventError::OutputLimit {
            observed,
            maximum: self.maximum_output,
        }
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        EventError::Allocation { resource, amount }
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        EventError::Invalid(message)
    }

    fn check(&self, amount: usize) -> Result<(), Self::Error> {
        if amount > self.maximum_output {
            Err(self.output_limit(amount))
        } else {
            Ok(())
        }
    }

    fn check_structure(&self, nodes: usize, parts: usize) -> Result<(), Self::Error> {
        if nodes > self.maximum_nodes || parts > self.maximum_parts {
            Err(EventError::Structure { nodes, parts })
        } else {
            Ok(())
        }
    }

    fn charge(&mut self, amount: usize) -> Result<(), Self::Error> {
        let observed = self
            .charged_output
            .checked_add(amount)
            .unwrap_or(usize::MAX);
        if observed > self.maximum_output {
            Err(self.output_limit(observed))
        } else {
            self.charged_output = observed;
            Ok(())
        }
    }
}

impl FormulaEventRenderBudget for EventBudget {
    fn check_render_depth(&self, depth: usize) -> Result<(), Self::Error> {
        if depth > self.maximum_render_depth {
            Err(EventError::RenderDepth {
                observed: depth,
                maximum: self.maximum_render_depth,
            })
        } else {
            Ok(())
        }
    }

    fn parse_error(&self, _message: String) -> Self::Error {
        EventError::Parse
    }

    fn invalid_format(&self, _message: String) -> Self::Error {
        EventError::InvalidFormat
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct FuzzResolver;

impl ReferenceResolver for FuzzResolver {
    fn table_prefix(&self, id: &FormulaRenderCfuuid) -> Option<FormulaTablePrefix<'_>> {
        id.is_complete().then_some(FormulaTablePrefix {
            sheet: "Fuzz Sheet",
            table: "Fuzz Table",
        })
    }

    fn category_name(&self, id: FormulaCategoryId) -> Option<&str> {
        ((id.lower ^ id.upper) % 2 == 0).then_some("Category")
    }

    fn function_name(&self, index: u32) -> Option<&str> {
        match index {
            0 => Some("SUM"),
            1 => Some("MAX"),
            _ => None,
        }
    }
}

struct Cursor<'source> {
    source: &'source [u8],
    offset: usize,
}

impl<'source> Cursor<'source> {
    const fn new(source: &'source [u8]) -> Self {
        Self { source, offset: 0 }
    }

    fn byte(&mut self) -> u8 {
        let value = self.source.get(self.offset).copied().unwrap_or(0);
        self.offset = self.offset.saturating_add(1);
        value
    }

    fn bounded(&mut self, maximum: usize) -> usize {
        if maximum == 0 {
            0
        } else {
            usize::from(self.byte()) % maximum
        }
    }

    fn u32(&mut self) -> u32 {
        u32::from_le_bytes([self.byte(), self.byte(), self.byte(), self.byte()])
    }

    fn f64(&mut self) -> f64 {
        let bits = u64::from_le_bytes([
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
            self.byte(),
        ]);
        let value = f64::from_bits(bits);
        if value.is_finite() {
            value
        } else {
            f64::from(self.byte()) / 4.0
        }
    }
}

fn complete_uuid() -> FormulaRenderCfuuid {
    FormulaRenderCfuuid {
        has_uuid_bytes: false,
        word0: Some(1),
        word1: Some(2),
        word2: Some(3),
        word3: Some(4),
    }
}

fn maybe_uuid(cursor: &mut Cursor<'_>) -> FormulaRenderCfuuid {
    let mut id = complete_uuid();
    if cursor.byte() & 1 == 0 {
        id.word3 = None;
    }
    id
}

fn local_reference(cursor: &mut Cursor<'_>) -> FormulaRenderLocalCellReference {
    FormulaRenderLocalCellReference {
        row_handle: cursor.u32() % 128,
        column_handle: cursor.u32() % 64,
        row_is_sticky: u32::from(cursor.byte() & 1 != 0),
        column_is_sticky: u32::from(cursor.byte() & 1 != 0),
    }
}

fn cell_reference(cursor: &mut Cursor<'_>) -> FormulaRenderCellReference {
    match cursor.byte() % 4 {
        0 => FormulaRenderCellReference {
            coordinates: Some(FormulaRenderCoordinatePair {
                column: FormulaRenderAxis {
                    coordinate: i32::from(cursor.byte()) - 96,
                    absolute: cursor.byte() & 1 != 0,
                },
                row: FormulaRenderAxis {
                    coordinate: i32::from(cursor.byte()) - 96,
                    absolute: cursor.byte() & 1 != 0,
                },
            }),
            local: None,
            cross_table: None,
            cross_table_extra: None,
        },
        1 => FormulaRenderCellReference {
            coordinates: None,
            local: Some(local_reference(cursor)),
            cross_table: None,
            cross_table_extra: None,
        },
        2 => FormulaRenderCellReference {
            coordinates: None,
            local: None,
            cross_table: Some(FormulaRenderCrossTableCellReference {
                table_id: maybe_uuid(cursor),
                row_handle: cursor.u32() % 128,
                column_handle: cursor.u32() % 64,
                row_is_sticky: u32::from(cursor.byte() & 1 != 0),
                column_is_sticky: u32::from(cursor.byte() & 1 != 0),
            }),
            cross_table_extra: None,
        },
        _ => FormulaRenderCellReference {
            coordinates: None,
            local: None,
            cross_table: None,
            cross_table_extra: None,
        },
    }
}

fn category_reference(cursor: &mut Cursor<'_>) -> FormulaRenderCategoryReference {
    let uuid = FormulaRenderUuid {
        lower: u64::from(cursor.u32()),
        upper: u64::from(cursor.u32()),
    };
    FormulaRenderCategoryReference {
        group_by_uid: Some(uuid),
        column_uid: None,
        absolute_group_uid: (cursor.byte() & 1 != 0).then_some(uuid),
        relative_group_uid: None,
        last_group_uid: None,
        group_uid_count: 1,
    }
}

fn colon_tract(cursor: &mut Cursor<'_>) -> FormulaRenderColonTract {
    let begin_column = i64::from(cursor.byte() % 8);
    let begin_row = i64::from(cursor.byte() % 8);
    FormulaRenderColonTract {
        relative_column: FormulaRenderRangeSummary {
            count: 1,
            first_begin: Some(begin_column),
            first_end: Some(begin_column + 1),
        },
        relative_row: FormulaRenderRangeSummary {
            count: 1,
            first_begin: Some(begin_row),
            first_end: Some(begin_row + 1),
        },
        absolute_column: FormulaRenderRangeSummary {
            count: 0,
            first_begin: None,
            first_end: None,
        },
        absolute_row: FormulaRenderRangeSummary {
            count: 0,
            first_begin: None,
            first_end: None,
        },
        preserve_rectangular: cursor.byte() & 1 != 0,
        sticky: FormulaRenderStickyBits {
            begin_row_is_absolute: cursor.byte() & 1 != 0,
            begin_column_is_absolute: cursor.byte() & 1 != 0,
            end_row_is_absolute: cursor.byte() & 1 != 0,
            end_column_is_absolute: cursor.byte() & 1 != 0,
        },
        cross_table_extra: None,
    }
}

fn borrowed_text<'source>(cursor: &mut Cursor<'source>, maximum: usize) -> &'source str {
    let length = cursor.bounded(maximum.saturating_add(1));
    let start = cursor.offset;
    let end = start.saturating_add(length).min(cursor.source.len());
    cursor.offset = end;
    std::str::from_utf8(&cursor.source[start..end]).unwrap_or("")
}

fn send_event<R, B>(
    visitor: &mut CompatibilityFormulaVisitor<'_, '_, R, B>,
    event: FormulaRenderEvent<'_>,
) -> bool
where
    R: ReferenceResolver,
    B: FormulaEventRenderBudget,
{
    match visitor.visit_event(event) {
        Ok(()) => true,
        Err(error) => {
            let _ = black_box(error);
            false
        },
    }
}

fn run_scalar_sequence(resolver: &FuzzResolver, budget: EventBudget) {
    let mut budget = budget;
    let mut visitor = CompatibilityFormulaVisitor::new(3, 4, resolver, &mut budget);
    let events = [
        FormulaRenderEvent::BeginArray { depth: 1 },
        FormulaRenderEvent::Number { value: 1.0 },
        FormulaRenderEvent::String("value"),
        FormulaRenderEvent::Function {
            identifier: 0,
            argument_count: 2,
        },
        FormulaRenderEvent::EndArray,
    ];
    for event in events {
        if !send_event(&mut visitor, event) {
            break;
        }
    }
    let _ = black_box(visitor.finish());
    black_box(budget.charged_output);
}

fn run_array_and_reference_sequences(resolver: &FuzzResolver, budget: EventBudget) {
    let mut budget = budget;
    let mut visitor = CompatibilityFormulaVisitor::new(2, 5, resolver, &mut budget);
    let events = [
        FormulaRenderEvent::BeginArray { depth: 1 },
        FormulaRenderEvent::Number { value: 1.0 },
        FormulaRenderEvent::Number { value: 2.0 },
        FormulaRenderEvent::Array {
            columns: 2,
            rows: 1,
        },
        FormulaRenderEvent::EndArray,
    ];
    for event in events {
        if !send_event(&mut visitor, event) {
            break;
        }
    }
    let _ = black_box(visitor.finish());
    black_box(budget.charged_output);

    let mut budget = budget;
    let mut visitor = CompatibilityFormulaVisitor::new(2, 5, resolver, &mut budget);
    let events = [
        FormulaRenderEvent::BeginArray { depth: 1 },
        FormulaRenderEvent::CellReference(FormulaRenderCellReference {
            coordinates: None,
            local: Some(FormulaRenderLocalCellReference {
                row_handle: 0,
                column_handle: 0,
                row_is_sticky: 1,
                column_is_sticky: 1,
            }),
            cross_table: None,
            cross_table_extra: None,
        }),
        FormulaRenderEvent::LocalCellReference(None),
        FormulaRenderEvent::Binary(BinaryOperator::Add),
        FormulaRenderEvent::EndArray,
    ];
    for event in events {
        if !send_event(&mut visitor, event) {
            break;
        }
    }
    let _ = black_box(visitor.finish());
    black_box(budget.charged_output);
}

fn arbitrary_events(source: &[u8], resolver: &FuzzResolver, budget: EventBudget) {
    let mut cursor = Cursor::new(source);
    let mut budget = budget;
    let mut visitor = CompatibilityFormulaVisitor::new(
        cursor.u32() % 32,
        cursor.u32() % 32,
        resolver,
        &mut budget,
    );
    for _ in 0..MAX_EVENTS {
        let event = match cursor.byte() % 24 {
            0 => FormulaRenderEvent::BeginArray {
                depth: u32::from(cursor.byte() % 12),
            },
            1 => FormulaRenderEvent::EndArray,
            2 => FormulaRenderEvent::ThunkBegin,
            3 => FormulaRenderEvent::ThunkEnd,
            4 => FormulaRenderEvent::Binary(match cursor.byte() % 12 {
                0 => BinaryOperator::Add,
                1 => BinaryOperator::Subtract,
                2 => BinaryOperator::Multiply,
                3 => BinaryOperator::Divide,
                4 => BinaryOperator::Power,
                5 => BinaryOperator::Concatenate,
                6 => BinaryOperator::GreaterThan,
                7 => BinaryOperator::GreaterThanOrEqual,
                8 => BinaryOperator::LessThan,
                9 => BinaryOperator::LessThanOrEqual,
                10 => BinaryOperator::Equal,
                _ => BinaryOperator::NotEqual,
            }),
            5 => FormulaRenderEvent::Negation,
            6 => FormulaRenderEvent::Percent,
            7 => FormulaRenderEvent::Number {
                value: cursor.f64(),
            },
            8 => FormulaRenderEvent::String(borrowed_text(&mut cursor, 16)),
            9 => FormulaRenderEvent::Boolean(cursor.byte() & 1 != 0),
            10 => FormulaRenderEvent::Token(cursor.byte() & 1 != 0),
            11 => FormulaRenderEvent::Date {
                value: cursor.f64(),
            },
            12 => FormulaRenderEvent::Duration {
                value: cursor.f64(),
            },
            13 => FormulaRenderEvent::EmptyArgument,
            14 => FormulaRenderEvent::Function {
                identifier: cursor.u32() % 4,
                argument_count: u32::from(cursor.byte() % 8),
            },
            15 => FormulaRenderEvent::List {
                argument_count: u32::from(cursor.byte() % 8),
            },
            16 => FormulaRenderEvent::Array {
                columns: u32::from(cursor.byte() % 5),
                rows: u32::from(cursor.byte() % 5),
            },
            17 => FormulaRenderEvent::UnknownFunction {
                name: (cursor.byte() & 1 != 0).then(|| borrowed_text(&mut cursor, 12)),
                argument_count: u32::from(cursor.byte() % 8),
            },
            18 => FormulaRenderEvent::CellReference(cell_reference(&mut cursor)),
            19 => FormulaRenderEvent::LocalCellReference(
                (cursor.byte() & 1 != 0).then(|| local_reference(&mut cursor)),
            ),
            20 => {
                FormulaRenderEvent::CrossTableCellReference((cursor.byte() & 1 != 0).then(|| {
                    FormulaRenderCrossTableCellReference {
                        table_id: maybe_uuid(&mut cursor),
                        row_handle: cursor.u32() % 64,
                        column_handle: cursor.u32() % 32,
                        row_is_sticky: u32::from(cursor.byte() & 1 != 0),
                        column_is_sticky: u32::from(cursor.byte() & 1 != 0),
                    }
                }))
            },
            21 => match cursor.byte() % 4 {
                0 => FormulaRenderEvent::Colon,
                1 => FormulaRenderEvent::ColonWithUids,
                2 => FormulaRenderEvent::ColonTract(colon_tract(&mut cursor)),
                _ => FormulaRenderEvent::CategoryReference(Some(category_reference(&mut cursor))),
            },
            22 => FormulaRenderEvent::ReferenceError,
            _ => FormulaRenderEvent::Ignored {
                raw_node_type: u32::from(cursor.byte()),
            },
        };
        if !send_event(&mut visitor, event) {
            break;
        }
    }
    let _ = black_box(visitor.finish());
    black_box((budget.charged_output, cursor.offset));
}

fuzz_target!(|data: &[u8]| {
    let resolver = FuzzResolver;
    run_scalar_sequence(
        &resolver,
        EventBudget::new(MAX_OUTPUT_BYTES, MAX_NODES, MAX_PARTS, MAX_RENDER_DEPTH),
    );
    run_array_and_reference_sequences(&resolver, EventBudget::new(256, 32, 128, 4));
    arbitrary_events(data, &resolver, EventBudget::new(64, 16, 64, 3));
});
