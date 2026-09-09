use litchi_iwa_common::formula::render::{FormulaRenderBudget, FormulaRenderer};

#[derive(Debug, PartialEq, Eq)]
enum BudgetError {
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
}

#[derive(Debug, Clone, Copy)]
struct TestBudget {
    maximum_output: usize,
    maximum_nodes: usize,
    maximum_parts: usize,
    charged_output: usize,
}

impl TestBudget {
    const fn unlimited() -> Self {
        Self {
            maximum_output: usize::MAX,
            maximum_nodes: 100_000,
            maximum_parts: 800_000,
            charged_output: 0,
        }
    }

    const fn with_output_limit(maximum_output: usize) -> Self {
        Self {
            maximum_output,
            maximum_nodes: 100_000,
            maximum_parts: 800_000,
            charged_output: 0,
        }
    }

    const fn with_structure_limits(mut self, maximum_nodes: usize, maximum_parts: usize) -> Self {
        self.maximum_nodes = maximum_nodes;
        self.maximum_parts = maximum_parts;
        self
    }
}

impl FormulaRenderBudget for TestBudget {
    type Error = BudgetError;

    fn output_limit(&self, observed: usize) -> Self::Error {
        BudgetError::OutputLimit {
            observed,
            maximum: self.maximum_output,
        }
    }

    fn allocation(&self, resource: &'static str, amount: usize) -> Self::Error {
        BudgetError::Allocation { resource, amount }
    }

    fn invalid(&self, message: &'static str) -> Self::Error {
        BudgetError::Invalid(message)
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
            return Err(BudgetError::StructureLimit {
                resource: "nodes",
                observed: nodes,
                maximum: self.maximum_nodes,
            });
        }
        if parts > self.maximum_parts {
            return Err(BudgetError::StructureLimit {
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

#[test]
fn expression_composition_preserves_postfix_parts_and_exact_output() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::unlimited();
    let left = renderer.static_expr("1", &budget)?;
    let right = renderer.static_expr("2", &budget)?;
    let sum = renderer.binary(left, "+", right, true, &budget)?;
    let expression = renderer.unary("-", sum, "", &budget)?;

    let mut render_budget = TestBudget::unlimited();
    assert_eq!(renderer.render(expression, &mut render_budget)?, "=-(1+2)");
    assert_eq!(render_budget.charged_output, "=-(1+2)".len());
    assert_eq!(renderer.node_count(), 4);
    assert_eq!(renderer.part_count(), 10);
    Ok(())
}

#[test]
fn shared_nonempty_subexpression_is_rendered_for_each_reference() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::unlimited();
    let leaf = renderer.static_expr("x", &budget)?;
    let repeated = renderer.binary(leaf, "+", leaf, false, &budget)?;

    let mut render_budget = TestBudget::unlimited();
    assert_eq!(renderer.render(repeated, &mut render_budget)?, "=x+x");
    assert_eq!(render_budget.charged_output, "=x+x".len());
    Ok(())
}

#[test]
fn function_and_array_composition_uses_expected_separators() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::unlimited();
    let a = renderer.static_expr("A", &budget)?;
    let b = renderer.static_expr("B", &budget)?;
    let c = renderer.static_expr("C", &budget)?;
    let d = renderer.static_expr("D", &budget)?;

    let function =
        renderer.comma_joined(Some("SUM".to_owned()), vec![a, b, c], "(", ")", &budget)?;
    let array = renderer.array(vec![a, b, c, d], 2, &budget)?;
    let flat_array = renderer.array(vec![a, b, c], 0, &budget)?;

    let mut render_budget = TestBudget::unlimited();
    assert_eq!(
        renderer.render(function, &mut render_budget)?,
        "=SUM(A,B,C)"
    );
    assert_eq!(renderer.render(array, &mut render_budget)?, "={A,B;C,D}");
    assert_eq!(renderer.render(flat_array, &mut render_budget)?, "={A,B,C}");
    Ok(())
}

#[test]
fn owned_unicode_text_is_measured_in_bytes_and_boundary_is_inclusive() -> Result<(), BudgetError> {
    let text = "利润🧪é";
    let output_len = 1 + text.len();
    let mut renderer = FormulaRenderer::default();
    let exact_budget = TestBudget::with_output_limit(output_len);
    let expression = renderer.owned_expr(text.to_owned(), &exact_budget)?;

    let before_failed_growth = (renderer.node_count(), renderer.part_count());
    let growth_error = renderer
        .owned_expr("x".to_owned(), &exact_budget)
        .expect_err("retained owned text must count toward the output budget");
    assert_eq!(
        growth_error,
        BudgetError::OutputLimit {
            observed: output_len + 1,
            maximum: output_len,
        }
    );
    assert_eq!(
        (renderer.node_count(), renderer.part_count()),
        before_failed_growth
    );

    let mut render_budget = exact_budget;
    let rendered = renderer.render(expression, &mut render_budget)?;
    assert_eq!(rendered, format!("={text}"));
    assert_eq!(rendered.len(), output_len);
    assert_eq!(render_budget.charged_output, output_len);

    let mut one_short = TestBudget::with_output_limit(output_len - 1);
    let render_error = renderer
        .render(expression, &mut one_short)
        .expect_err("one byte below the rendered formula must be rejected");
    assert_eq!(
        render_error,
        BudgetError::OutputLimit {
            observed: output_len,
            maximum: output_len - 1,
        }
    );
    assert_eq!(one_short.charged_output, 0);
    Ok(())
}

#[test]
fn arena_growth_renders_deep_composition_iteratively() -> Result<(), BudgetError> {
    let depth = 10_000;
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::unlimited();
    let mut expression = renderer.static_expr("x", &budget)?;
    for _ in 0..depth {
        expression = renderer.unary("(", expression, ")", &budget)?;
    }

    let mut render_budget = TestBudget::unlimited();
    let rendered = renderer.render(expression, &mut render_budget)?;
    assert_eq!(rendered.len(), 1 + 1 + 2 * depth);
    assert_eq!(renderer.node_count(), depth + 1);
    assert_eq!(renderer.part_count(), 1 + 3 * depth);
    Ok(())
}

#[test]
fn shared_zero_width_dag_renders_once_under_one_byte_budget() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let composition_budget = TestBudget::with_output_limit(1);
    let empty = renderer.static_expr("", &composition_budget)?;
    let mut expression = empty;
    for _ in 0..64 {
        expression = renderer.binary(expression, "", expression, false, &composition_budget)?;
    }

    let mut render_budget = TestBudget::with_output_limit(1);
    assert_eq!(renderer.render(expression, &mut render_budget)?, "=");
    assert_eq!(render_budget.charged_output, 1);
    assert_eq!(renderer.node_count(), 65);
    assert_eq!(renderer.part_count(), 193);
    Ok(())
}

#[test]
fn zero_width_node_growth_stops_at_the_node_ceiling_before_mutation() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::with_output_limit(1).with_structure_limits(1, 800_000);
    let empty = renderer.static_expr("", &budget)?;
    let before = (renderer.node_count(), renderer.part_count());

    let error = renderer
        .binary(empty, "", empty, false, &budget)
        .expect_err("the second zero-width node must exceed the node ceiling");
    assert_eq!(
        error,
        BudgetError::StructureLimit {
            resource: "nodes",
            observed: 2,
            maximum: 1,
        }
    );
    assert_eq!((renderer.node_count(), renderer.part_count()), before);
    Ok(())
}

#[test]
fn dynamic_part_limits_refuse_function_and_array_before_arena_mutation() -> Result<(), BudgetError>
{
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::with_output_limit(1).with_structure_limits(100_000, 1);
    let _empty = renderer.static_expr("", &budget)?;
    let before = (renderer.node_count(), renderer.part_count());

    let function_error = renderer
        .comma_joined(Some("F".to_owned()), Vec::new(), "(", ")", &budget)
        .expect_err("function parts must be checked before the temporary Vec allocation");
    assert_eq!(
        function_error,
        BudgetError::StructureLimit {
            resource: "parts",
            observed: 4,
            maximum: 1,
        }
    );
    assert_eq!((renderer.node_count(), renderer.part_count()), before);

    let array_error = renderer
        .array(Vec::new(), 0, &budget)
        .expect_err("array parts must be checked before the temporary Vec allocation");
    assert_eq!(
        array_error,
        BudgetError::StructureLimit {
            resource: "parts",
            observed: 3,
            maximum: 1,
        }
    );
    assert_eq!((renderer.node_count(), renderer.part_count()), before);
    Ok(())
}

#[test]
fn shared_dag_traversal_is_bounded_even_when_output_fits() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let composition_budget = TestBudget::unlimited();
    let leaf = renderer.static_expr("x", &composition_budget)?;
    let repeated = renderer.binary(leaf, "+", leaf, false, &composition_budget)?;

    let mut traversal_budget = TestBudget::with_output_limit(4).with_structure_limits(100_000, 8);
    let error = renderer
        .render(repeated, &mut traversal_budget)
        .expect_err("the output fits, but the traversal structure exceeds its ceiling");
    assert_eq!(
        error,
        BudgetError::StructureLimit {
            resource: "parts",
            observed: 9,
            maximum: 8,
        }
    );
    assert_eq!(traversal_budget.charged_output, 4);
    Ok(())
}

#[test]
fn output_limit_refusal_is_atomic_for_composition_and_render() -> Result<(), BudgetError> {
    let mut renderer = FormulaRenderer::default();
    let budget = TestBudget::with_output_limit(4);
    let expression = renderer.owned_expr("abc".to_owned(), &budget)?;
    let before = (renderer.node_count(), renderer.part_count());

    let error = renderer
        .owned_expr("d".to_owned(), &budget)
        .expect_err("the second owned value exceeds retained text capacity");
    assert!(matches!(
        error,
        BudgetError::OutputLimit {
            observed: 5,
            maximum: 4,
        }
    ));
    assert_eq!((renderer.node_count(), renderer.part_count()), before);

    let mut render_budget = TestBudget::with_output_limit(4);
    assert_eq!(renderer.render(expression, &mut render_budget)?, "=abc");
    assert_eq!(render_budget.charged_output, 4);

    let mut one_short = TestBudget::with_output_limit(3);
    let error = renderer
        .render(expression, &mut one_short)
        .expect_err("render charging must reject before publishing output");
    assert!(matches!(
        error,
        BudgetError::OutputLimit {
            observed: 4,
            maximum: 3,
        }
    ));
    assert_eq!(one_short.charged_output, 0);
    Ok(())
}

#[test]
fn foreign_handles_are_rejected_without_mutating_the_target_arena() -> Result<(), BudgetError> {
    let mut source = FormulaRenderer::default();
    let source_budget = TestBudget::unlimited();
    let foreign = source.static_expr("foreign", &source_budget)?;

    let mut target = FormulaRenderer::default();
    let target_budget = TestBudget::unlimited();
    let local = target.static_expr("local", &target_budget)?;
    let before = (target.node_count(), target.part_count());

    let error = target
        .binary(foreign, "+", local, false, &target_budget)
        .expect_err("an expression from another arena must not be accepted");
    assert!(matches!(error, BudgetError::Invalid(_)));
    assert_eq!((target.node_count(), target.part_count()), before);

    let mut render_budget = TestBudget::unlimited();
    let error = target
        .render(foreign, &mut render_budget)
        .expect_err("a foreign handle must not be renderable");
    assert!(matches!(error, BudgetError::Invalid(_)));
    assert_eq!(render_budget.charged_output, 0);
    Ok(())
}
