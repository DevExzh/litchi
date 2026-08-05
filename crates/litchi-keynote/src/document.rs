//! Cheaply shareable semantic Keynote document snapshots.

use std::sync::Arc;

use crate::Show;

#[derive(Debug)]
struct State {
    show: Show,
}

/// An immutable, cheaply clonable semantic Keynote document snapshot.
#[derive(Debug, Clone)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    /// Create a snapshot from an already decoded semantic show.
    #[must_use]
    pub fn from_show(show: Show) -> Self {
        Self {
            state: Arc::new(State { show }),
        }
    }

    /// Capture another cheap handle to the same snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the immutable semantic show.
    #[must_use]
    pub fn show(&self) -> &Show {
        &self.state.show
    }

    /// Borrow the slides without copying the snapshot.
    #[must_use]
    pub fn slides(&self) -> &[crate::Slide] {
        self.state.show.slides()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn snapshots_are_send_sync_and_shareable() {
        assert_send_sync::<Document>();
        let document = Document::from_show(Show::builder().build());
        let snapshot = document.snapshot();
        assert_eq!(document.show(), snapshot.show());
    }
}
