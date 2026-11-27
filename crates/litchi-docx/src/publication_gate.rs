//! Apply-once publication gate shared by every clone of a reversible patch.

use std::{
    fmt,
    sync::{
        Arc,
        atomic::{AtomicU8, Ordering},
    },
};

/// Lifecycle of one patch publication, stored in an [`AtomicU8`].
///
/// | From       | Event                                  | To         |
/// |------------|----------------------------------------|------------|
/// | `Ready`    | [`PublicationGate::claim`] succeeds    | `InFlight` |
/// | `InFlight` | [`PublicationClaim::finalize`]         | `Applied`  |
/// | `InFlight` | the claim is dropped without finalize  | `Ready`    |
///
/// `Applied` is terminal. Only the holder of the single live
/// [`PublicationClaim`] leaves `InFlight`, so no other transition exists and
/// no value other than these three is ever stored.
#[derive(Clone, Copy)]
#[repr(u8)]
enum GateState {
    Ready = 0,
    InFlight = 1,
    Applied = 2,
}

/// Why [`PublicationGate::claim`] refused a publication attempt.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ClaimRefusal {
    /// Another attempt holds the claim; retry once it is released.
    InFlight,
    /// The patch, or one of its clones, was already applied.
    Applied,
}

/// The gate shared by a patch and all of its clones.
#[derive(Clone)]
pub(crate) struct PublicationGate(Arc<AtomicU8>);

impl PublicationGate {
    pub(crate) fn new() -> Self {
        Self(Arc::new(AtomicU8::new(GateState::Ready as u8)))
    }

    pub(crate) fn is_applied(&self) -> bool {
        self.0.load(Ordering::Acquire) == GateState::Applied as u8
    }

    /// `Ready` to `InFlight`. A refusal leaves the gate unchanged.
    pub(crate) fn claim(&self) -> Result<PublicationClaim, ClaimRefusal> {
        match self.0.compare_exchange(
            GateState::Ready as u8,
            GateState::InFlight as u8,
            Ordering::AcqRel,
            Ordering::Acquire,
        ) {
            Ok(_) => Ok(PublicationClaim {
                gate: self.clone(),
                finalized: false,
            }),
            Err(observed) if observed == GateState::InFlight as u8 => Err(ClaimRefusal::InFlight),
            // The failed exchange observed a value other than `Ready` and
            // `InFlight`; `Applied` is the only other value ever stored.
            Err(_) => Err(ClaimRefusal::Applied),
        }
    }
}

impl fmt::Debug for PublicationGate {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Debug::fmt(&self.0, formatter)
    }
}

/// Retry-safe ownership of one publication attempt.
pub(crate) struct PublicationClaim {
    gate: PublicationGate,
    finalized: bool,
}

impl PublicationClaim {
    /// `InFlight` to `Applied`: permanently consume the patch and its clones.
    pub(crate) fn finalize(mut self) {
        self.gate
            .0
            .store(GateState::Applied as u8, Ordering::Release);
        self.finalized = true;
    }
}

impl Drop for PublicationClaim {
    /// `InFlight` to `Ready` unless the claim was finalized.
    #[expect(
        clippy::let_underscore_must_use,
        reason = "only this claim leaves `InFlight`, so the exchange cannot fail"
    )]
    fn drop(&mut self) {
        if !self.finalized {
            let _ = self.gate.0.compare_exchange(
                GateState::InFlight as u8,
                GateState::Ready as u8,
                Ordering::AcqRel,
                Ordering::Acquire,
            );
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{ClaimRefusal, PublicationGate};

    #[test]
    fn claim_is_exclusive_retryable_and_final_once_applied() {
        let gate = PublicationGate::new();
        let clone = gate.clone();

        let claim = gate.claim().expect("ready gate");
        assert_eq!(clone.claim().err(), Some(ClaimRefusal::InFlight));
        assert!(!clone.is_applied());
        drop(claim);

        let claim = clone.claim().expect("released gate");
        claim.finalize();
        assert!(gate.is_applied());
        assert_eq!(gate.claim().err(), Some(ClaimRefusal::Applied));
        assert_eq!(format!("{gate:?}"), "2");
    }
}
