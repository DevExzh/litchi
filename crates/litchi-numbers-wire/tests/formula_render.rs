//! Independent contract tests for the generated-free Numbers formula event renderer.
//!
//! These vectors exercise the wire-facing event stream directly.  The resolver and
//! budget below intentionally do not use a Numbers package implementation, so the
//! tests cover the shared renderer's output and failure boundaries independently of
//! any archive census or semantic adapter.

use litchi_iwa_common::formula::render::FormulaRenderBudget;
use litchi_iwa_protos::numbers_formula_codec::{
    BinaryOperator, FormulaNode, FormulaRenderAxis, FormulaRenderCategoryReference,
    FormulaRenderCellReference, FormulaRenderCfuuid, FormulaRenderColonTract,
    FormulaRenderCoordinatePair, FormulaRenderCrossTableCellReference,
    FormulaRenderCrossTableExtra, FormulaRenderEvent, FormulaRenderRangeSummary,
    FormulaRenderStickyBits, FormulaRenderUuid, FormulaRenderVisitor,
};
use litchi_numbers_wire::formula_render::{
    CompatibilityFormulaVisitor, FormulaCategoryId, FormulaEventRenderBudget,
    FormulaRenderCodecVisitor, FormulaTablePrefix, ReferenceResolver, render_scalar_formula_nodes,
};

#[derive(Debug, PartialEq, Eq)]
enum RenderError {
    OutputLimit {
        observed: usize,
        maximum: usize,
    },
    StructureLimit {
        resource: &'static str,
        observed: usize,
        maximum: usize,
    },
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    Invalid(&'static str),
    Depth {
        observed: usize,
        maximum: usize,
    },
    Parse(String),
    InvalidFormat(String),
}

#[derive(Debug, Clone, Copy)]
struct RenderBudget {
    maximum_output: usize,
    maximum_nodes: usize,
    maximum_parts: usize,
    maximum_depth: usize,
    charged_output: usize,
}

impl RenderBudget {
    const fn unlimited() -> Self {
        Self {
            maximum_output: usize::MAX,
            maximum_nodes: 100_000,
            maximum_parts: 800_000,
            maximum_depth: 64,
            charged_output: 0,
        }
    }

    const fn with_output_limit(maximum_output: usize) -> Self {
        Self {
            maximum_output,
            ..Self::unlimited()
        }
    }

    const fn with_depth_limit(mut self, maximum_depth: usize) -> Self {
        self.maximum_depth = maximum_depth;
        self
    }
}

impl FormulaRenderBudget for RenderBudget {
    type Error = RenderError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        RenderError::OutputLimit {
            observed,
            maximum: self.maximum_output,
        }
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        RenderError::Allocation { resource, amount }
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        RenderError::Invalid(message)
    }

    fn check(&self, amount: usize) -> Result<(), Self::Error> {
        if amount > self.maximum_output {
            Err(self.output_limit(amount))
        } else {
            Ok(())
        }
    }

    fn check_structure(&self, nodes: usize, parts: usize) -> Result<(), Self::Error> {
        if nodes > self.maximum_nodes {
            return Err(RenderError::StructureLimit {
                resource: "nodes",
                observed: nodes,
                maximum: self.maximum_nodes,
            });
        }
        if parts > self.maximum_parts {
            return Err(RenderError::StructureLimit {
                resource: "parts",
                observed: parts,
                maximum: self.maximum_parts,
            });
        }
        Ok(())
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

impl FormulaEventRenderBudget for RenderBudget {
    fn check_render_depth(&self, depth: usize) -> Result<(), Self::Error> {
        if depth > self.maximum_depth {
            Err(RenderError::Depth {
                observed: depth,
                maximum: self.maximum_depth,
            })
        } else {
            Ok(())
        }
    }

    fn parse_error(&self, message: String) -> Self::Error {
        RenderError::Parse(message)
    }

    fn invalid_format(&self, message: String) -> Self::Error {
        RenderError::InvalidFormat(message)
    }
}

#[derive(Debug, Default, Clone, Copy)]
struct Resolver;

const fn known_table_id() -> FormulaRenderCfuuid {
    FormulaRenderCfuuid {
        has_uuid_bytes: false,
        word0: Some(7),
        word1: Some(8),
        word2: Some(9),
        word3: Some(10),
    }
}

impl ReferenceResolver for Resolver {
    fn table_prefix(&self, id: &FormulaRenderCfuuid) -> Option<FormulaTablePrefix<'_>> {
        (*id == known_table_id()).then_some(FormulaTablePrefix {
            sheet: "Sheet 2",
            table: "Input",
        })
    }

    fn category_name(&self, id: FormulaCategoryId) -> Option<&str> {
        (id == FormulaCategoryId {
            lower: 11,
            upper: 12,
        })
        .then_some("North]\\Region")
    }

    fn function_name(&self, index: u32) -> Option<&str> {
        (index == 42).then_some("SUM")
    }
}

#[derive(Debug, Clone, Copy)]
struct TableOnlyResolver;

impl ReferenceResolver for TableOnlyResolver {
    fn table_prefix(&self, _id: &FormulaRenderCfuuid) -> Option<FormulaTablePrefix<'_>> {
        None
    }

    fn table_only_name(&self, id: &FormulaRenderCfuuid) -> Option<&str> {
        (*id == known_table_id()).then_some("Body")
    }

    fn category_name(&self, _id: FormulaCategoryId) -> Option<&str> {
        None
    }

    fn function_name(&self, _index: u32) -> Option<&str> {
        None
    }
}

#[derive(Debug, Clone, Copy)]
struct ConflictingTableResolver;

impl ReferenceResolver for ConflictingTableResolver {
    fn table_prefix(&self, id: &FormulaRenderCfuuid) -> Option<FormulaTablePrefix<'_>> {
        (*id == known_table_id()).then_some(FormulaTablePrefix {
            sheet: "Invented sheet",
            table: "Body",
        })
    }

    fn table_only_name(&self, id: &FormulaRenderCfuuid) -> Option<&str> {
        (*id == known_table_id()).then_some("Body")
    }

    fn category_name(&self, _id: FormulaCategoryId) -> Option<&str> {
        None
    }

    fn function_name(&self, _index: u32) -> Option<&str> {
        None
    }
}

fn render_events_with_resolver<R>(
    events: impl IntoIterator<Item = FormulaRenderEvent<'static>>,
    resolver: &R,
    budget: &mut RenderBudget,
) -> Result<String, RenderError>
where
    R: ReferenceResolver,
{
    let mut visitor = CompatibilityFormulaVisitor::new(0, 0, resolver, budget);
    for event in events {
        visitor.visit_event(event)?;
    }
    visitor.finish()
}

fn render_events(
    events: impl IntoIterator<Item = FormulaRenderEvent<'static>>,
    budget: &mut RenderBudget,
) -> Result<String, RenderError> {
    let resolver = Resolver;
    render_events_with_resolver(events, &resolver, budget)
}

fn category_event() -> FormulaRenderEvent<'static> {
    FormulaRenderEvent::CategoryReference(Some(FormulaRenderCategoryReference {
        group_by_uid: Some(FormulaRenderUuid { lower: 1, upper: 2 }),
        column_uid: Some(FormulaRenderUuid { lower: 3, upper: 4 }),
        absolute_group_uid: Some(FormulaRenderUuid {
            lower: 11,
            upper: 12,
        }),
        relative_group_uid: None,
        last_group_uid: None,
        group_uid_count: 1,
    }))
}

#[test]
fn events_render_values_strings_operators_and_references() -> Result<(), RenderError> {
    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::Number { value: 1.5 },
                FormulaRenderEvent::Number { value: 2.0 },
                FormulaRenderEvent::Binary(BinaryOperator::Add),
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=(1.5+2)"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::String("Café北京\"x"),
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=\"Café北京\"\"x\""
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::CrossTableCellReference(Some(
                    FormulaRenderCrossTableCellReference {
                        row_handle: 2,
                        column_handle: 2,
                        row_is_sticky: 0,
                        column_is_sticky: 0,
                        table_id: known_table_id(),
                    },
                )),
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=Sheet 2::Input::C3"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                category_event(),
                FormulaRenderEvent::EndArray
            ],
            &mut budget
        )?,
        "=#CATEGORY![North\\]\\\\Region]"
    );
    Ok(())
}

#[test]
fn table_only_names_render_cross_cell_coordinate_and_range_references() -> Result<(), RenderError> {
    let resolver = TableOnlyResolver;
    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events_with_resolver(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::CrossTableCellReference(Some(
                    FormulaRenderCrossTableCellReference {
                        row_handle: 2,
                        column_handle: 2,
                        row_is_sticky: 0,
                        column_is_sticky: 0,
                        table_id: known_table_id(),
                    },
                )),
                FormulaRenderEvent::EndArray,
            ],
            &resolver,
            &mut budget,
        )?,
        "=Body::C3"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events_with_resolver(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::CellReference(FormulaRenderCellReference {
                    coordinates: Some(FormulaRenderCoordinatePair {
                        column: FormulaRenderAxis {
                            coordinate: 0,
                            absolute: false,
                        },
                        row: FormulaRenderAxis {
                            coordinate: 0,
                            absolute: false,
                        },
                    }),
                    local: None,
                    cross_table: None,
                    cross_table_extra: Some(FormulaRenderCrossTableExtra {
                        table_id: known_table_id(),
                    }),
                }),
                FormulaRenderEvent::EndArray,
            ],
            &resolver,
            &mut budget,
        )?,
        "=Body::A1"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events_with_resolver(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::ColonTract(FormulaRenderColonTract {
                    relative_column: FormulaRenderRangeSummary {
                        count: 1,
                        first_begin: Some(0),
                        first_end: Some(2),
                    },
                    relative_row: FormulaRenderRangeSummary {
                        count: 1,
                        first_begin: Some(0),
                        first_end: Some(2),
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
                    preserve_rectangular: true,
                    sticky: FormulaRenderStickyBits {
                        begin_row_is_absolute: false,
                        begin_column_is_absolute: false,
                        end_row_is_absolute: false,
                        end_column_is_absolute: false,
                    },
                    cross_table_extra: Some(FormulaRenderCrossTableExtra {
                        table_id: known_table_id(),
                    }),
                }),
                FormulaRenderEvent::EndArray,
            ],
            &resolver,
            &mut budget,
        )?,
        "=Body::A1:C3"
    );
    Ok(())
}

#[test]
fn table_only_name_takes_precedence_over_sheet_table_prefix() -> Result<(), RenderError> {
    let resolver = ConflictingTableResolver;
    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events_with_resolver(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::CrossTableCellReference(Some(
                    FormulaRenderCrossTableCellReference {
                        row_handle: 0,
                        column_handle: 0,
                        row_is_sticky: 0,
                        column_is_sticky: 0,
                        table_id: known_table_id(),
                    },
                )),
                FormulaRenderEvent::EndArray,
            ],
            &resolver,
            &mut budget,
        )?,
        "=Body::A1"
    );
    Ok(())
}

#[test]
fn events_render_functions_arrays_thunks_and_empty_root() -> Result<(), RenderError> {
    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::Number { value: 1.0 },
                FormulaRenderEvent::Number { value: 2.0 },
                FormulaRenderEvent::Function {
                    identifier: 42,
                    argument_count: 2,
                },
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=SUM(1,2)"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::Number { value: 1.0 },
                FormulaRenderEvent::Number { value: 2.0 },
                FormulaRenderEvent::Number { value: 3.0 },
                FormulaRenderEvent::Number { value: 4.0 },
                FormulaRenderEvent::Array {
                    columns: 2,
                    rows: 2,
                },
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "={1,2;3,4}"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::ThunkBegin,
                FormulaRenderEvent::BeginArray { depth: 2 },
                FormulaRenderEvent::Number { value: 7.0 },
                FormulaRenderEvent::EndArray,
                FormulaRenderEvent::ThunkEnd,
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=7"
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "="
    );

    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_events(
            [
                FormulaRenderEvent::BeginArray { depth: 1 },
                FormulaRenderEvent::UnknownFunction {
                    name: Some("自定义"),
                    argument_count: 0,
                },
                FormulaRenderEvent::EndArray,
            ],
            &mut budget,
        )?,
        "=自定义()"
    );
    Ok(())
}

#[test]
fn scalar_renderer_keeps_empty_and_postfix_compatibility() -> Result<(), RenderError> {
    let resolver = Resolver;
    let mut budget = RenderBudget::unlimited();
    assert_eq!(
        render_scalar_formula_nodes(&[], &resolver, &mut budget)?,
        "="
    );

    let mut budget = RenderBudget::unlimited();
    let nodes = [
        FormulaNode::Number {
            bits: 1.0_f64.to_bits(),
        },
        FormulaNode::Number {
            bits: 2.0_f64.to_bits(),
        },
        FormulaNode::Binary(BinaryOperator::Add),
    ];
    assert_eq!(
        render_scalar_formula_nodes(&nodes, &resolver, &mut budget)?,
        "=(1+2)"
    );

    let mut budget = RenderBudget::unlimited();
    let ignored = [FormulaNode::PlusSign, FormulaNode::AppendWhitespace];
    assert_eq!(
        render_scalar_formula_nodes(&ignored, &resolver, &mut budget)?,
        "=FORMULA()"
    );
    Ok(())
}

#[test]
fn malformed_events_and_scalar_nodes_report_typed_parse_errors() {
    let resolver = Resolver;
    let mut budget = RenderBudget::unlimited();
    let mut visitor = CompatibilityFormulaVisitor::new(0, 0, &resolver, &mut budget);
    let error = visitor
        .visit_event(FormulaRenderEvent::Binary(BinaryOperator::Add))
        .expect_err("the malformed event must produce a typed error");
    assert!(matches!(error, RenderError::Parse(_)));

    let mut budget = RenderBudget::unlimited();
    let nodes = [FormulaNode::Binary(BinaryOperator::Add)];
    let error = render_scalar_formula_nodes(&nodes, &resolver, &mut budget)
        .expect_err("a binary operator without operands is malformed");
    assert!(matches!(error, RenderError::Parse(_)));
}

#[test]
fn event_depth_must_match_the_active_array_stack() {
    let resolver = Resolver;
    let mut budget = RenderBudget::unlimited();
    let mut visitor = CompatibilityFormulaVisitor::new(0, 0, &resolver, &mut budget);
    let error = visitor
        .visit_event(FormulaRenderEvent::BeginArray { depth: 2 })
        .expect_err("the mismatched depth must produce a typed error");
    assert!(matches!(error, RenderError::Parse(message) if message.contains("reported depth")));
}

#[test]
fn depth_and_output_limits_are_checked_before_publishing() {
    let resolver = Resolver;
    let mut budget = RenderBudget::unlimited().with_depth_limit(0);
    let mut visitor = CompatibilityFormulaVisitor::new(0, 0, &resolver, &mut budget);
    let error = visitor
        .visit_event(FormulaRenderEvent::BeginArray { depth: 1 })
        .expect_err("the depth refusal must produce a typed error");
    assert_eq!(
        error,
        RenderError::Depth {
            observed: 1,
            maximum: 0,
        }
    );

    let mut budget = RenderBudget::with_output_limit(1);
    let nodes = [FormulaNode::Number {
        bits: 1.0_f64.to_bits(),
    }];
    let error = render_scalar_formula_nodes(&nodes, &resolver, &mut budget)
        .expect_err("the formula prefix and scalar must exceed one byte");
    assert!(matches!(error, RenderError::OutputLimit { maximum: 1, .. }));

    let mut budget = RenderBudget::with_output_limit(1);
    let error = render_events(
        [
            FormulaRenderEvent::BeginArray { depth: 1 },
            FormulaRenderEvent::String("x"),
            FormulaRenderEvent::EndArray,
        ],
        &mut budget,
    )
    .expect_err("a quoted string plus formula prefix must exceed one byte");
    assert!(matches!(error, RenderError::OutputLimit { maximum: 1, .. }));
}

#[test]
fn codec_visitor_bridge_keeps_typed_errors_until_finish() {
    let resolver = Resolver;
    let mut budget = RenderBudget::unlimited();
    let mut visitor = FormulaRenderCodecVisitor::new(0, 0, &resolver, &mut budget);
    assert!(
        FormulaRenderVisitor::visit(
            &mut visitor,
            FormulaRenderEvent::Binary(BinaryOperator::Add)
        )
        .is_err()
    );
    let error = visitor
        .finish()
        .expect_err("the codec bridge must publish its captured semantic error");
    assert!(matches!(error, RenderError::Parse(_)));
}
