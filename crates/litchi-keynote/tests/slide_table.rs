use litchi_keynote::slide::table::{formula, sort};
use litchi_numbers::formula as numbers_formula;
use litchi_numbers::table::sort as numbers_sort;

fn into_keynote_expression(
    value: numbers_formula::FormulaExpression,
) -> formula::FormulaExpression {
    value
}

fn into_keynote_order(value: numbers_sort::Order) -> sort::Order {
    value
}

#[test]
fn formula_leaf_reuses_numbers_types_and_constructors() {
    let reference = formula::FormulaCellReference::mixed(2, 3, true, false);
    let expression = formula::FormulaExpression::binary(
        formula::FormulaBinaryOperator::Add,
        formula::FormulaExpression::cell(reference),
        formula::FormulaExpression::relative_cell(0, 1),
    );

    let from_keynote = into_keynote_expression(expression.clone());
    assert_eq!(from_keynote, expression);
    let canonical: numbers_formula::FormulaExpression =
        numbers_formula::FormulaExpression::table_cell(
            7,
            numbers_formula::FormulaCellReference::relative(1, 1),
        );
    let from_numbers = into_keynote_expression(canonical);
    assert_eq!(
        from_numbers,
        formula::FormulaExpression::table_cell(7, formula::FormulaCellReference::relative(1, 1),)
    );

    assert_eq!(
        expression,
        formula::FormulaExpression::Binary {
            operator: formula::FormulaBinaryOperator::Add,
            left: Box::new(formula::FormulaExpression::Cell(reference)),
            right: Box::new(formula::FormulaExpression::Cell(
                formula::FormulaCellReference::relative(0, 1),
            )),
        }
    );
    assert_eq!(
        formula::FormulaCachedValue::Boolean(true).into_value(),
        litchi_numbers::cell::Value::Boolean(true)
    );
}

#[test]
fn formula_leaf_exposes_axis_and_pivot_constructors() {
    let pivot = formula::FormulaPivotCategoryReference::new(
        formula::FormulaUuid::new(1, 2),
        formula::FormulaUuid::new(3, 4),
        formula::FormulaUuid::new(5, 6),
        2,
        1,
    );

    assert_eq!(
        formula::FormulaExpression::table_rows(
            7,
            formula::FormulaAxisReference::relative(1),
            formula::FormulaAxisReference::absolute(3),
        ),
        formula::FormulaExpression::TableRows {
            table_id: 7,
            start: formula::FormulaAxisReference::relative(1),
            end: formula::FormulaAxisReference::absolute(3),
        }
    );
    assert_eq!(
        formula::FormulaExpression::pivot_category(pivot),
        formula::FormulaExpression::PivotCategory(pivot)
    );
}

#[test]
fn sort_leaf_reuses_numbers_types_and_validates_orders() -> sort::Result<()> {
    let column = sort::ColumnIndex::new(3)?;
    let rule = sort::Rule::new(column, sort::Direction::Descending);
    let order = sort::Order::selected_rows([rule])?;

    let from_keynote = into_keynote_order(order.clone());
    assert_eq!(from_keynote, order);
    let canonical = numbers_sort::Order::new([numbers_sort::Rule::new(
        numbers_sort::ColumnIndex::new(1)?,
        numbers_sort::Direction::Ascending,
    )])?;
    let from_numbers = into_keynote_order(canonical);
    assert_eq!(from_numbers.scope(), sort::Scope::EntireTable);

    assert_eq!(order.scope(), sort::Scope::SelectedRows);
    assert_eq!(order.rules(), &[rule]);
    assert_eq!(sort::RowRange::new(2, 5)?.len(), 3);
    assert!(matches!(sort::Order::new([]), Err(sort::Error::EmptyOrder)));
    assert!(matches!(
        sort::Order::new([rule, rule]),
        Err(sort::Error::DuplicateColumn { column: 3 })
    ));
    assert!(matches!(
        sort::RowRange::new(5, 5),
        Err(sort::Error::InvalidRowRange { start: 5, end: 5 })
    ));
    Ok(())
}
