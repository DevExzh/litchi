#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi_iwa_common::formula::render::{FormulaExpr, FormulaRenderBudget, FormulaRenderer};

const MAX_OPERATIONS: usize = 128;
const MAX_HANDLES: usize = 96;
const MAX_OUTPUT_BYTES: usize = 512;
const MAX_NODES: usize = 64;
const MAX_PARTS: usize = 256;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ArenaError {
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
}

#[derive(Debug, Clone, Copy)]
struct ArenaBudget {
    maximum_output: usize,
    maximum_nodes: usize,
    maximum_parts: usize,
    charged_output: usize,
}

impl ArenaBudget {
    const fn new(maximum_output: usize, maximum_nodes: usize, maximum_parts: usize) -> Self {
        Self {
            maximum_output,
            maximum_nodes,
            maximum_parts,
            charged_output: 0,
        }
    }
}

impl FormulaRenderBudget for ArenaBudget {
    type Error = ArenaError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        ArenaError::OutputLimit {
            observed,
            maximum: self.maximum_output,
        }
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        ArenaError::Allocation { resource, amount }
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        ArenaError::Invalid(message)
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
            Err(ArenaError::Structure { nodes, parts })
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
}

fn retain(handles: &mut Vec<FormulaExpr>, expression: FormulaExpr) {
    if handles.len() == MAX_HANDLES {
        handles.remove(0);
    }
    handles.push(expression);
}

fn take_arguments(handles: &mut Vec<FormulaExpr>, count: usize) -> Vec<FormulaExpr> {
    let count = count.min(handles.len()).min(8);
    let start = handles.len().saturating_sub(count);
    handles.drain(start..).collect()
}

fn owned_text(cursor: &mut Cursor<'_>, maximum: usize) -> String {
    let length = cursor.bounded(maximum.saturating_add(1));
    let mut text = String::new();
    text.reserve(length);
    for _ in 0..length {
        let character = b'a'.saturating_add(cursor.byte() % 26) as char;
        text.push(character);
    }
    text
}

fn record(handles: &mut Vec<FormulaExpr>, result: Result<FormulaExpr, ArenaError>) {
    if let Ok(expression) = result {
        retain(handles, expression);
    }
}

fn seeded_operations(
    renderer: &mut FormulaRenderer,
    budget: &ArenaBudget,
    handles: &mut Vec<FormulaExpr>,
) {
    record(handles, renderer.static_expr("1", budget));
    record(handles, renderer.owned_expr("2".to_owned(), budget));
    if handles.len() >= 2 {
        let right = handles[handles.len() - 1];
        let left = handles[handles.len() - 2];
        record(handles, renderer.binary(left, "+", right, true, budget));
    }
    if let Some(expression) = handles.last().copied() {
        record(handles, renderer.unary("-", expression, "", budget));
        record(
            handles,
            renderer.comma_joined(Some("SUM".to_owned()), vec![expression], "(", ")", budget),
        );
        record(handles, renderer.array(vec![expression], 1, budget));
    }
    record(handles, renderer.array(Vec::new(), 0, budget));
}

fn arbitrary_operations(
    source: &[u8],
    renderer: &mut FormulaRenderer,
    budget: &ArenaBudget,
    handles: &mut Vec<FormulaExpr>,
) {
    let mut cursor = Cursor::new(source);
    for _ in 0..MAX_OPERATIONS {
        let operation = cursor.byte() % 8;
        match operation {
            0 => record(
                handles,
                renderer.static_expr(["", "x", "1", "TRUE", "A1"][cursor.bounded(5)], budget),
            ),
            1 => record(
                handles,
                renderer.owned_expr(owned_text(&mut cursor, 16), budget),
            ),
            2 => {
                if let Some(expression) = handles.last().copied() {
                    record(
                        handles,
                        renderer.unary(
                            if cursor.byte() & 1 == 0 { "(" } else { "-" },
                            expression,
                            if cursor.byte() & 1 == 0 { ")" } else { "" },
                            budget,
                        ),
                    );
                }
            },
            3 => {
                if handles.len() >= 2 {
                    if let (Some(right), Some(left)) = (handles.pop(), handles.pop()) {
                        let operator = ["+", "-", "*", "/", "&", "="][cursor.bounded(6)];
                        record(
                            handles,
                            renderer.binary(left, operator, right, cursor.byte() & 1 == 0, budget),
                        );
                    }
                }
            },
            4 => {
                let arguments = take_arguments(handles, cursor.bounded(6));
                let name = owned_text(&mut cursor, 8);
                record(
                    handles,
                    renderer.comma_joined(Some(name), arguments, "(", ")", budget),
                );
            },
            5 => {
                let values = take_arguments(handles, cursor.bounded(6));
                record(handles, renderer.array(values, cursor.bounded(4), budget));
            },
            6 => {
                let index = cursor.bounded(handles.len());
                if let Some(expression) = handles.get(index).copied() {
                    retain(handles, expression);
                }
            },
            _ => {
                let index = cursor.bounded(handles.len());
                if let Some(expression) = handles.get(index).copied() {
                    let mut render_budget = *budget;
                    let _ = black_box(renderer.render(expression, &mut render_budget));
                }
            },
        }
    }
}

fn exercise_foreign_and_empty_handles() {
    let composition_budget = ArenaBudget::new(MAX_OUTPUT_BYTES, MAX_NODES, MAX_PARTS);
    let mut source = FormulaRenderer::default();
    let foreign = source.static_expr("foreign", &composition_budget);
    let mut target = FormulaRenderer::default();
    let mut render_budget = composition_budget;

    // Rendering against a fresh, empty arena is the empty-DAG boundary. The
    // only available handle is foreign, so it must be rejected before charge.
    if let Ok(foreign) = foreign {
        let _ = black_box(target.render(foreign, &mut render_budget));
        let local = target.static_expr("local", &composition_budget);
        if let Ok(local) = local {
            let _ = black_box(target.binary(foreign, "+", local, false, &composition_budget));
        }
    }
    black_box((target.node_count(), target.part_count()));
}

fn exercise_arena(source: &[u8], budget: ArenaBudget) {
    let mut renderer = FormulaRenderer::default();
    let mut handles = Vec::new();
    handles.reserve(16);
    seeded_operations(&mut renderer, &budget, &mut handles);
    arbitrary_operations(source, &mut renderer, &budget, &mut handles);

    if let Some(expression) = handles.last().copied() {
        let mut render_budget = budget;
        let _ = black_box(renderer.render(expression, &mut render_budget));
        black_box(render_budget.charged_output);
    }
    black_box((renderer.node_count(), renderer.part_count(), handles.len()));
}

fuzz_target!(|data: &[u8]| {
    exercise_arena(
        data,
        ArenaBudget::new(MAX_OUTPUT_BYTES, MAX_NODES, MAX_PARTS),
    );
    exercise_arena(data, ArenaBudget::new(16, 4, 12));
    exercise_foreign_and_empty_handles();
});
