use super::{
    CellRead, Context, EvaluationContext, Evaluator, Limits, Position, Resolver, Shape,
    SheetExtent, ValueEvaluator,
};
use crate::codec::formula::expression::Expression;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits as CoreExecutionLimits,
    Limits as CoreLimits, Profile, Resource,
};
use std::{
    mem::size_of,
    num::{NonZeroU64, NonZeroUsize},
};

#[derive(Debug)]
struct NoReadResolver;

impl Resolver for NoReadResolver {
    fn sheet_extent(
        &self,
        _sheet: &str,
        _execution: &ExecutionContext,
    ) -> super::EvaluationResult<Option<SheetExtent>> {
        Ok(None)
    }

    fn read_cell<'a>(
        &'a self,
        _sheet: &str,
        _row: usize,
        _column: usize,
        _execution: &ExecutionContext,
    ) -> super::EvaluationResult<CellRead<'a>> {
        Ok(CellRead::Empty)
    }

    fn sheet_index(
        &self,
        _sheet: &str,
        _execution: &ExecutionContext,
    ) -> super::EvaluationResult<Option<usize>> {
        Ok(None)
    }

    fn sheet_name_at(
        &self,
        _index: usize,
        _execution: &ExecutionContext,
    ) -> super::EvaluationResult<Option<&str>> {
        Ok(None)
    }

    fn sheet_count(&self, _execution: &ExecutionContext) -> super::EvaluationResult<usize> {
        Ok(0)
    }
}

fn execution() -> ExecutionContext {
    let (_cancellation, token) = CancellationSource::pair();
    ExecutionContext::new(
        Budget::root(
            "value-shape-mask-vp15-2",
            CoreLimits::for_profile(Profile::Desktop),
        ),
        token,
        CoreExecutionLimits::new(
            NonZeroUsize::MIN,
            NonZeroUsize::MIN,
            NonZeroU64::new(1 << 30).expect("nonzero in-flight byte limit"),
            1 << 20,
        )
        .expect("valid execution limits"),
    )
}

#[test]
fn shape_planner_retains_mask_capacity_reservation_across_reuse() {
    let expression =
        Expression::parse("=IF({TRUE();FALSE()};IF({TRUE();FALSE()};{1;2};{3;4});{5;6})")
            .expect("nested lazy inline-array fixture should parse");
    let resolver = NoReadResolver;
    let execution = execution();
    let value_context = Context::new(&execution, Position::new("Sheet", 0, 0));
    let scalar_context = EvaluationContext::new(&execution);
    let limits = Limits::default();
    let scalar = Evaluator::new(
        &expression,
        &scalar_context,
        limits.scalar_limits(),
        execution.budget().clone(),
    );
    let mut evaluator = ValueEvaluator::new(
        &expression,
        &resolver,
        &value_context,
        limits,
        scalar,
        execution.budget().clone(),
    );
    let demand = Shape::new(1, 1).expect("nonempty demand shape");
    let expected_shape = Shape::new(1, 2).expect("nonempty result shape");

    let first = evaluator
        .shape_hint_demand(expression.root(), demand, None)
        .expect("first shape-planning pass should succeed");
    assert_eq!(first, Some(expected_shape));
    let first_capacity = evaluator.shape_masks.capacity();
    assert!(
        first_capacity > 0,
        "nested lazy branches should install masks"
    );
    let first_mask_bytes = u64::try_from(
        first_capacity
            .checked_mul(size_of::<super::ShapeMask>())
            .expect("shape-mask capacity byte count should fit usize"),
    )
    .expect("shape-mask capacity byte count should fit u64");
    let first_reservation = evaluator
        .shape_mask_reservation
        .as_ref()
        .expect("retained shape-mask capacity must retain its reservation");
    assert_eq!(first_reservation.resource(), Resource::Memory);
    assert_eq!(first_reservation.amount(), first_mask_bytes);
    let memory_between_passes = execution.budget().used(Resource::Memory);
    assert!(
        memory_between_passes >= first_mask_bytes,
        "the aggregate storage budget must remain charged while mask capacity is retained"
    );

    let second = evaluator
        .shape_hint_demand(expression.root(), demand, None)
        .expect("reused shape-planning pass should succeed");
    assert_eq!(second, Some(expected_shape));
    let second_capacity = evaluator.shape_masks.capacity();
    assert!(
        second_capacity > 0,
        "mask capacity should remain retained after reuse"
    );
    let second_mask_bytes = u64::try_from(
        second_capacity
            .checked_mul(size_of::<super::ShapeMask>())
            .expect("shape-mask capacity byte count should fit usize"),
    )
    .expect("shape-mask capacity byte count should fit u64");
    let second_reservation = evaluator
        .shape_mask_reservation
        .as_ref()
        .expect("reused shape-mask capacity must retain its reservation");
    assert_eq!(second_reservation.resource(), Resource::Memory);
    assert_eq!(second_reservation.amount(), second_mask_bytes);
    assert!(
        execution.budget().used(Resource::Memory) >= second_mask_bytes,
        "the aggregate storage budget must remain charged after reuse"
    );

    drop(evaluator);
    assert_eq!(
        execution.budget().used(Resource::Memory),
        0,
        "dropping the VM must release retained planner reservations"
    );
}

#[test]
fn demand_cache_entry_has_a_bounded_inline_footprint() {
    assert!(size_of::<super::DemandCacheEntry>() <= 32);
}
