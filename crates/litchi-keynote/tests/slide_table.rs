use litchi_iwa_common::{formula as common_formula, table::sort as common_sort};
use litchi_keynote::slide::table::{formula, sort};

fn into_keynote_expression(value: common_formula::FormulaExpression) -> formula::FormulaExpression {
    value
}

fn into_keynote_order(value: common_sort::Order) -> sort::Order {
    value
}

#[test]
fn formula_leaf_uses_neutral_common_types_and_constructors() {
    let reference = formula::FormulaCellReference::mixed(2, 3, true, false);
    let expression = formula::FormulaExpression::binary(
        formula::FormulaBinaryOperator::Add,
        formula::FormulaExpression::cell(reference),
        formula::FormulaExpression::relative_cell(0, 1),
    );

    let from_keynote = into_keynote_expression(expression.clone());
    assert_eq!(from_keynote, expression);
    let canonical: common_formula::FormulaExpression =
        common_formula::FormulaExpression::table_cell(
            7,
            common_formula::FormulaCellReference::relative(1, 1),
        );
    let from_common = into_keynote_expression(canonical);
    assert_eq!(
        from_common,
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
fn sort_leaf_uses_neutral_common_types_and_validates_orders() -> sort::Result<()> {
    let column = sort::ColumnIndex::new(3)?;
    let rule = sort::Rule::new(column, sort::Direction::Descending);
    let order = sort::Order::selected_rows([rule])?;

    let from_keynote = into_keynote_order(order.clone());
    assert_eq!(from_keynote, order);
    let canonical = common_sort::Order::new([common_sort::Rule::new(
        common_sort::ColumnIndex::new(1)?,
        common_sort::Direction::Ascending,
    )])?;
    let from_common = into_keynote_order(canonical);
    assert_eq!(from_common.scope(), sort::Scope::EntireTable);

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
