//! A checked allocation budget for operations that build derived font data.
//!
//! The ordinary Allsorts APIs remain unchanged. A caller that needs an
//! allocation bound enters a scope before invoking a budget-aware operation;
//! budget-aware buffers and containers charge capacity before reserving it.
//! The scope is thread-local so concurrent font operations do not share a
//! mutable counter. A scope and its [`Budget`] retain the first refusal as an
//! immutable [`BudgetFailure`] snapshot for typed error mapping.

use std::cell::RefCell;
use std::mem::size_of;
use std::rc::Rc;

/// The independent counter that rejected a budget request.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BudgetDomain {
    /// Cumulative temporary allocation bytes.
    WorkingBytes,
    /// Cumulative retained output bytes.
    OutputBytes,
    /// Bounded parser or traversal work units.
    WorkUnits,
}

/// An immutable snapshot of the first request refused by a budget.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct BudgetFailure {
    /// The counter that rejected the request.
    pub domain: BudgetDomain,
    /// The configured maximum for that counter.
    pub limit: usize,
    /// The cumulative amount that the rejected request would have reached.
    pub requested: usize,
}

/// A failed allocation-budget charge.
///
/// When this came from an active budget, the authoritative immutable record is
/// [`Budget::failure`] or [`Scope::failure`]. A reservation can also surface
/// this shape for a structural arithmetic or allocator error; in that case no
/// budget failure snapshot is present and callers must not classify it as a
/// policy refusal.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct BudgetExceeded {
    /// The counter that rejected the request.
    pub domain: BudgetDomain,
    /// The configured maximum number of budget units.
    pub limit: usize,
    /// The cumulative units requested by the operation after this charge.
    pub requested: usize,
}
impl std::fmt::Display for BudgetExceeded {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            formatter,
            "allocation budget exceeded: requested {} budget units, limit {} budget units",
            self.requested, self.limit
        )
    }
}

impl std::error::Error for BudgetExceeded {}

#[derive(Debug)]
struct State {
    working_limit: usize,
    working_requested: usize,
    output_limit: usize,
    output_requested: usize,
    work: usize,
    failed: bool,
    failure: Option<BudgetFailure>,
}

/// A cloneable handle to the current operation's allocation budget.
#[derive(Clone, Debug)]
pub struct Budget(Rc<RefCell<State>>);

/// The guard returned by [`enter`].
pub struct Scope {
    previous: Option<Budget>,
    budget: Budget,
}

thread_local! {
    static ACTIVE: RefCell<Option<Budget>> = const { RefCell::new(None) };
}

/// Enter cumulative working and retained-output budgets for the current thread.
///
/// The working counter measures bytes requested by budget-aware temporary
/// buffers and container reservations. It is deliberately cumulative: a
/// temporary buffer that is dropped and replaced still consumed an allocation
/// request, so a caller cannot bypass the bound by forcing repeated growth.
/// Retained output has its own cumulative counter because it remains live after
/// the operation returns. Traversal work uses the working limit as a separate
/// unit counter. The byte counters describe these requested payload bytes; they
/// do not claim to model allocator metadata, rounding, or the exact physical
/// process peak.
pub fn enter(working_limit: usize, output_limit: usize) -> Scope {
    ACTIVE.with(|active| {
        let budget = Budget(Rc::new(RefCell::new(State {
            working_limit,
            working_requested: 0,
            output_limit,
            output_requested: 0,
            work: 0,
            failed: false,
            failure: None,
        })));
        let previous = active.replace(Some(budget.clone()));
        Scope { previous, budget }
    })
}

/// Return the active budget, if the caller entered one.
pub fn current() -> Option<Budget> {
    ACTIVE.with(|active| active.borrow().clone())
}

/// Return the first request refused by the active budget, if any.
pub fn failure() -> Option<BudgetFailure> {
    current().and_then(|budget| budget.failure())
}

/// Charge an allocation whose owner is not a `Vec` or `String`.
///
/// This is used for exact-sized boxes and other owned containers that cannot
/// be passed through [`reserve_vec`] directly.
pub fn charge_allocation(bytes: usize) -> Result<(), BudgetExceeded> {
    current().map_or(Ok(()), |budget| budget.charge(bytes))
}

/// Charge the exact inflate workspace that the selected flate2 backend asks
/// zlib-rs to allocate when a decoder is constructed.
///
/// zlib-rs keeps the layout calculation beside its allocator and exposes the
/// compiled size through this small pinned-dependency bridge. Other flate2
/// backends retain their ordinary behaviour; they do not use this hook.
#[cfg(feature = "flate2_zlib-rs")]
pub fn charge_zlib_workspace() -> Result<(), BudgetExceeded> {
    charge_allocation(zlib_rs::inflate_workspace_size())
}

/// No zlib-rs workspace is constructed when another flate2 backend is active.
#[cfg(not(feature = "flate2_zlib-rs"))]
pub fn charge_zlib_workspace() -> Result<(), BudgetExceeded> {
    Ok(())
}

/// Return whether a budgeted allocator has already refused a request.
pub fn failed() -> bool {
    current().is_some_and(|budget| budget.0.borrow().failed)
}

/// Mark the active budget as failed when a checked arithmetic operation cannot
/// represent the allocation request. This is used by foreign allocators whose
/// allocation trait has no `Result` channel. This is structural failure
/// bookkeeping, not a configured-domain refusal, so it does not create a
/// [`BudgetFailure`] snapshot.
pub(crate) fn fail() {
    if let Some(budget) = current() {
        budget.0.borrow_mut().failed = true;
    }
}

impl Budget {
    fn reject_state(state: &mut State, candidate: BudgetFailure) -> BudgetExceeded {
        state.failed = true;
        let failure = match state.failure {
            Some(failure) => failure,
            None => {
                state.failure = Some(candidate);
                candidate
            },
        };
        BudgetExceeded {
            domain: failure.domain,
            limit: failure.limit,
            requested: failure.requested,
        }
    }

    fn existing_failure(state: &State) -> Option<BudgetExceeded> {
        state.failure.map(|failure| BudgetExceeded {
            domain: failure.domain,
            limit: failure.limit,
            requested: failure.requested,
        })
    }

    fn structural_state(state: &mut State, candidate: BudgetExceeded) -> BudgetExceeded {
        state.failed = true;
        Self::existing_failure(state).unwrap_or(candidate)
    }

    /// Charge a working-byte request before performing the corresponding
    /// temporary allocation.
    pub fn charge(&self, bytes: usize) -> Result<(), BudgetExceeded> {
        let mut state = self.0.borrow_mut();
        if let Some(failure) = Self::existing_failure(&state) {
            return Err(failure);
        }
        let requested = match state.working_requested.checked_add(bytes) {
            Some(requested) => requested,
            None => {
                let limit = state.working_limit;
                return Err(Self::structural_state(
                    &mut state,
                    BudgetExceeded {
                        domain: BudgetDomain::WorkingBytes,
                        limit,
                        requested: usize::MAX,
                    },
                ));
            },
        };
        if requested > state.working_limit {
            let limit = state.working_limit;
            return Err(Self::reject_state(
                &mut state,
                BudgetFailure {
                    domain: BudgetDomain::WorkingBytes,
                    limit,
                    requested,
                },
            ));
        }
        state.working_requested = requested;
        Ok(())
    }

    /// Charge a byte request before performing the retained output allocation.
    pub fn charge_output(&self, bytes: usize) -> Result<(), BudgetExceeded> {
        let mut state = self.0.borrow_mut();
        if let Some(failure) = Self::existing_failure(&state) {
            return Err(failure);
        }
        let requested = match state.output_requested.checked_add(bytes) {
            Some(requested) => requested,
            None => {
                let limit = state.output_limit;
                return Err(Self::structural_state(
                    &mut state,
                    BudgetExceeded {
                        domain: BudgetDomain::OutputBytes,
                        limit,
                        requested: usize::MAX,
                    },
                ));
            },
        };
        if requested > state.output_limit {
            let limit = state.output_limit;
            return Err(Self::reject_state(
                &mut state,
                BudgetFailure {
                    domain: BudgetDomain::OutputBytes,
                    limit,
                    requested,
                },
            ));
        }
        state.output_requested = requested;
        Ok(())
    }

    /// Return the first request refused by this budget, if any.
    pub fn failure(&self) -> Option<BudgetFailure> {
        self.0.borrow().failure
    }

    /// Return the cumulative working bytes charged so far.
    pub fn requested(&self) -> usize {
        self.0.borrow().working_requested
    }

    /// Return the configured working-byte maximum.
    pub fn limit(&self) -> usize {
        self.0.borrow().working_limit
    }

    /// Return the uncharged portion of the working-byte maximum.
    pub fn remaining_working(&self) -> usize {
        let state = self.0.borrow();
        state.working_limit.saturating_sub(state.working_requested)
    }

    /// Return the cumulative retained output bytes charged so far.
    pub fn output_requested(&self) -> usize {
        self.0.borrow().output_requested
    }

    /// Return the configured retained-output maximum.
    pub fn output_limit(&self) -> usize {
        self.0.borrow().output_limit
    }

    /// Return the uncharged portion of the retained-output maximum.
    pub fn remaining_output(&self) -> usize {
        let state = self.0.borrow();
        state.output_limit.saturating_sub(state.output_requested)
    }

    /// Charge one unit of bounded parser/traversal work.
    pub fn charge_work(&self, units: usize) -> Result<(), BudgetExceeded> {
        let mut state = self.0.borrow_mut();
        if let Some(failure) = Self::existing_failure(&state) {
            return Err(failure);
        }
        let work = match state.work.checked_add(units) {
            Some(work) => work,
            None => {
                let limit = state.working_limit;
                return Err(Self::structural_state(
                    &mut state,
                    BudgetExceeded {
                        domain: BudgetDomain::WorkUnits,
                        limit,
                        requested: usize::MAX,
                    },
                ));
            },
        };
        if work > state.working_limit {
            let limit = state.working_limit;
            return Err(Self::reject_state(
                &mut state,
                BudgetFailure {
                    domain: BudgetDomain::WorkUnits,
                    limit,
                    requested: work,
                },
            ));
        }
        state.work = work;
        Ok(())
    }
}

fn arithmetic_failure(domain: BudgetDomain, limit: usize) -> BudgetExceeded {
    let candidate = BudgetExceeded {
        domain,
        limit,
        requested: usize::MAX,
    };
    current().map_or(candidate, |budget| {
        let mut state = budget.0.borrow_mut();
        Budget::structural_state(&mut state, candidate)
    })
}

/// Reserve additional capacity in a vector.
///
/// `Vec::try_reserve_exact` accepts an amount relative to the current length,
/// not the currently available spare capacity. When growth is needed, the
/// complete requested new length is charged. That charge is cumulative, so it
/// records both the old allocation and the replacement request while the new
/// allocation is being installed; charging only the capacity delta would
/// understate the requested bytes.
pub fn reserve_vec<T>(vec: &mut Vec<T>, additional: usize) -> Result<(), BudgetExceeded> {
    let spare = vec.capacity().checked_sub(vec.len()).unwrap_or(0);
    if additional <= spare {
        return Ok(());
    }

    let budget = current();
    let limit = budget.as_ref().map_or(usize::MAX, Budget::limit);
    let new_len = vec
        .len()
        .checked_add(additional)
        .ok_or_else(|| arithmetic_failure(BudgetDomain::WorkingBytes, limit))?;
    let bytes = new_len
        .checked_mul(size_of::<T>())
        .ok_or_else(|| arithmetic_failure(BudgetDomain::WorkingBytes, limit))?;
    if let Some(budget) = &budget {
        budget.charge(bytes)?;
    }
    vec.try_reserve_exact(additional)
        .map_err(|_| BudgetExceeded {
            domain: BudgetDomain::WorkingBytes,
            limit,
            requested: budget.as_ref().map_or(bytes, |budget| budget.requested()),
        })
}

/// Reserve the next amortized growth step for a vector before pushing one item.
///
/// Callers that append one item at a time should use this helper instead of
/// reserving one element on every push. The requested capacity doubles (with a
/// checked fallback for an empty vector), preserving the allocation behaviour
/// of the ordinary `Vec` growth path while charging each reservation first.
pub fn reserve_for_push<T>(vec: &mut Vec<T>) -> Result<(), BudgetExceeded> {
    if vec.len() == vec.capacity() {
        let additional = vec.capacity().max(1);
        reserve_vec(vec, additional)?;
    }
    Ok(())
}

/// Reserve UTF-8 string storage before appending to a `String`.
pub fn reserve_string(string: &mut String, additional: usize) -> Result<(), BudgetExceeded> {
    let spare = string.capacity().checked_sub(string.len()).unwrap_or(0);
    if additional <= spare {
        return Ok(());
    }

    let budget = current();
    let limit = budget.as_ref().map_or(usize::MAX, Budget::limit);
    let new_len = string
        .len()
        .checked_add(additional)
        .ok_or_else(|| arithmetic_failure(BudgetDomain::WorkingBytes, limit))?;
    if let Some(budget) = &budget {
        budget.charge(new_len)?;
    }
    string
        .try_reserve_exact(additional)
        .map_err(|_| BudgetExceeded {
            domain: BudgetDomain::WorkingBytes,
            limit,
            requested: budget.as_ref().map_or(new_len, |budget| budget.requested()),
        })
}

impl Scope {
    /// Return the first request refused while this scope was active.
    ///
    /// The snapshot remains available after the scope is dropped, so callers
    /// can classify a flattened parser or writer error without inspecting the
    /// error text or guessing which counter rejected it.
    pub fn failure(&self) -> Option<BudgetFailure> {
        self.budget.failure()
    }
}

impl Drop for Scope {
    fn drop(&mut self) {
        ACTIVE.with(|active| {
            active.replace(self.previous.take());
        });
    }
}
// End of the budget scope implementation.

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn partial_spare_charges_requested_vec_capacity() {
        let _scope = enter(23, usize::MAX);
        let budget = current().unwrap();
        let mut values = Vec::new();
        reserve_vec(&mut values, 8).unwrap();
        values.extend_from_slice(&[0_u8; 5]);
        reserve_vec(&mut values, 10).unwrap();

        assert!(values.capacity() >= 15);
        // The initial eight-byte allocation remains live while the
        // fifteen-byte replacement is requested, so both requests count.
        assert_eq!(budget.requested(), 8 + 15);
    }

    #[test]
    fn partial_spare_charges_requested_string_capacity() {
        let _scope = enter(23, usize::MAX);
        let budget = current().unwrap();
        let mut value = String::new();
        reserve_string(&mut value, 8).unwrap();
        value.push_str("12345");
        reserve_string(&mut value, 10).unwrap();

        assert!(value.capacity() >= 15);
        assert_eq!(budget.requested(), 8 + 15);
    }

    #[test]
    fn failure_snapshot_preserves_first_working_refusal() {
        let scope = enter(4, 8);
        let budget = current().unwrap();
        let expected = BudgetFailure {
            domain: BudgetDomain::WorkingBytes,
            limit: 4,
            requested: 5,
        };

        assert!(budget.charge(5).is_err());
        assert_eq!(budget.failure(), Some(expected));
        assert_eq!(scope.failure(), Some(expected));
        assert_eq!(failure(), Some(expected));

        // A later refusal cannot overwrite the first domain or request.
        assert!(budget.charge_output(9).is_err());
        assert_eq!(budget.failure(), Some(expected));
    }

    #[test]
    fn failure_snapshot_identifies_output_bytes() {
        let _scope = enter(usize::MAX, 3);
        let budget = current().unwrap();
        assert!(budget.charge_output(4).is_err());
        assert_eq!(
            budget.failure(),
            Some(BudgetFailure {
                domain: BudgetDomain::OutputBytes,
                limit: 3,
                requested: 4,
            })
        );
    }

    #[test]
    fn failure_snapshot_identifies_work_units() {
        let _scope = enter(3, usize::MAX);
        let budget = current().unwrap();
        assert!(budget.charge_work(4).is_err());
        assert_eq!(
            budget.failure(),
            Some(BudgetFailure {
                domain: BudgetDomain::WorkUnits,
                limit: 3,
                requested: 4,
            })
        );
    }

    #[test]
    fn structural_counter_overflow_has_no_policy_snapshot() {
        let _scope = enter(usize::MAX, usize::MAX);
        let budget = current().unwrap();
        budget.charge(usize::MAX).unwrap();

        let error = budget.charge(1).unwrap_err();
        assert_eq!(error.domain, BudgetDomain::WorkingBytes);
        assert_eq!(error.requested, usize::MAX);
        assert!(budget.failure().is_none());
        assert!(failed());
    }

    #[test]
    fn structural_foreign_allocator_failure_has_no_policy_snapshot() {
        let _scope = enter(8, usize::MAX);
        fail();
        assert!(failed());
        assert!(failure().is_none());
    }

    #[test]
    fn structural_overflow_preserves_an_earlier_policy_snapshot() {
        let _scope = enter(4, 8);
        let budget = current().unwrap();
        let expected = BudgetFailure {
            domain: BudgetDomain::WorkingBytes,
            limit: 4,
            requested: 5,
        };
        assert!(budget.charge(5).is_err());

        let error = arithmetic_failure(BudgetDomain::OutputBytes, 8);
        assert_eq!(error.domain, expected.domain);
        assert_eq!(error.limit, expected.limit);
        assert_eq!(error.requested, expected.requested);
        assert_eq!(budget.failure(), Some(expected));
    }
}
