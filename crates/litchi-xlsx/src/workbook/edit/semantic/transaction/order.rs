//! Typed pending tab-order operations for one transaction.

use crate::error::{Result, allocation, invalid};

use super::super::super::validation::{FinalOrder, MoveIntent, OrderPlan, Target};
use super::Edit;

pub(super) fn move_to(edit: &mut Edit, identity: usize, to: usize) -> Result<()> {
    let order = plan(edit)?;
    if to >= order.positions.len() {
        return Err(invalid("tab move destination exceeds the workbook order"));
    }
    let from = order
        .positions
        .iter()
        .position(|candidate| *candidate == identity)
        .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
    if from == to {
        return Ok(());
    }
    order
        .moves
        .try_reserve(1)
        .map_err(|source| allocation("tab move", source))?;
    let identity = order.positions.remove(from);
    order.positions.insert(to, identity);
    order.moves.push(MoveIntent {
        sheet: identity,
        from,
        to,
    });
    Ok(())
}

pub(super) fn move_relative(
    edit: &mut Edit,
    identity: usize,
    anchor: usize,
    after: bool,
) -> Result<()> {
    if identity == anchor {
        return Ok(());
    }
    let order = plan(edit)?;
    let from = order
        .positions
        .iter()
        .position(|candidate| *candidate == identity)
        .ok_or_else(|| invalid("selected tab disappeared from the pending order"))?;
    if !order.positions.contains(&anchor) {
        return Err(invalid("anchor tab disappeared from the pending order"));
    }
    order
        .moves
        .try_reserve(1)
        .map_err(|source| allocation("tab move", source))?;
    let identity = order.positions.remove(from);
    let anchor = order
        .positions
        .iter()
        .position(|candidate| *candidate == anchor)
        .ok_or_else(|| invalid("anchor tab disappeared during reorder"))?;
    let to = if after {
        anchor
            .checked_add(1)
            .ok_or_else(|| invalid("tab move position overflow"))?
    } else {
        anchor
    };
    order.positions.insert(to, identity);
    if from != to {
        order.moves.push(MoveIntent {
            sheet: identity,
            from,
            to,
        });
    }
    Ok(())
}

pub(super) fn projected_position(edit: &Edit, target: Target) -> Option<usize> {
    FinalOrder::plan(
        edit.base.len(),
        edit.order.as_ref().filter(|order| order.is_effective()),
        &edit.added,
    )
    .ok()?
    .position(target)
}

fn plan(edit: &mut Edit) -> Result<&mut OrderPlan> {
    if edit.order.is_none() {
        let len = edit.base.inner.sheets.len();
        let mut positions = Vec::new();
        positions
            .try_reserve_exact(len)
            .map_err(|source| allocation("tab-order plan", source))?;
        positions.extend(0..len);
        edit.order = Some(OrderPlan {
            positions,
            moves: Vec::new(),
        });
    }
    edit.order
        .as_mut()
        .ok_or_else(|| invalid("tab-order plan initialization failed"))
}
