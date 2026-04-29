//! Per-slide timing support for PPT files.
//!
//! Controls automatic slide advance timing independently of transitions.
//! Timing is encoded in the `SSSlideInfoAtom` (type=1017) record within
//! each slide container.
//!
//! # Binary Structure
//!
//! The `SSSlideInfoAtom` is a 16-byte record:
//!
//! ```text
//! SSSlideInfoAtom (type=1017)
//! ├── slideTime (u32): auto-advance time in milliseconds
//! ├── soundIdRef (u32): sound reference (0 = none)
//! ├── effectDirection (u8): transition direction
//! ├── effectType (u8): transition type
//! ├── effectTransitionFlags (u16):
//! │   ├── bit 0: manual advance (on click)
//! │   ├── bit 2: hidden slide
//! │   ├── bit 4: sound
//! │   ├── bit 6: loop sound
//! │   ├── bit 8: stop sound
//! │   ├── bit 10: auto advance
//! │   └── bit 12: cursor visible
//! ├── speed (u8): transition speed (0=slow, 1=medium, 2=fast)
//! └── unused (3 bytes)
//! ```

use super::records::{PptError, RecordBuilder, record_type};

/// Per-slide timing configuration.
///
/// Controls how and when a slide advances during a presentation.
#[derive(Debug, Clone)]
pub struct SlideTiming {
    /// Auto-advance time in milliseconds (0 = no auto-advance).
    pub advance_time_ms: u32,
    /// Whether the slide can advance on mouse click.
    pub advance_on_click: bool,
    /// Whether the slide is hidden during presentation.
    pub hidden: bool,
}

impl Default for SlideTiming {
    fn default() -> Self {
        Self {
            advance_time_ms: 0,
            advance_on_click: true,
            hidden: false,
        }
    }
}

impl SlideTiming {
    /// Create a timing that advances automatically after the given milliseconds.
    ///
    /// # Arguments
    ///
    /// * `ms` - Auto-advance time in milliseconds
    ///
    /// # Example
    ///
    /// ```
    /// use litchi::ole::ppt::writer::slide_timing::SlideTiming;
    /// // Auto-advance after 5 seconds, also allow click
    /// let timing = SlideTiming::auto_advance(5000);
    /// ```
    pub fn auto_advance(ms: u32) -> Self {
        Self {
            advance_time_ms: ms,
            advance_on_click: true,
            hidden: false,
        }
    }

    /// Create a timing that only advances on click (no auto-advance).
    pub fn on_click_only() -> Self {
        Self::default()
    }

    /// Create a timing for a hidden slide.
    pub fn hidden() -> Self {
        Self {
            advance_time_ms: 0,
            advance_on_click: true,
            hidden: true,
        }
    }

    /// Set whether clicking advances the slide.
    pub fn with_click_advance(mut self, enabled: bool) -> Self {
        self.advance_on_click = enabled;
        self
    }
}

/// Build an SSSlideInfoAtom record for per-slide timing.
///
/// This is used when the slide has timing but no transition effect.
/// If the slide also has a transition, the transition writer handles
/// the SSSlideInfoAtom instead (since it shares the same record).
pub fn build_slide_timing(timing: &SlideTiming) -> Result<Vec<u8>, PptError> {
    let mut data = Vec::with_capacity(16);

    // slideTime (u32): auto-advance time in ms
    data.extend_from_slice(&timing.advance_time_ms.to_le_bytes());

    // soundIdRef (u32): 0 = no sound
    data.extend_from_slice(&0u32.to_le_bytes());

    // effectDirection (u8): 0 = no direction
    data.push(0);

    // effectType (u8): 0 = no transition
    data.push(0);

    // effectTransitionFlags (u16)
    let mut flags: u16 = 0;
    if timing.advance_on_click {
        flags |= 1 << 0; // MANUAL_ADVANCE_BIT
    }
    if timing.hidden {
        flags |= 1 << 2; // HIDDEN_BIT
    }
    if timing.advance_time_ms > 0 {
        flags |= 1 << 10; // AUTO_ADVANCE_BIT
    }
    data.extend_from_slice(&flags.to_le_bytes());

    // speed (u8): 1 = medium (default)
    data.push(1);

    // unused (3 bytes)
    data.extend_from_slice(&[0u8; 3]);

    let mut builder = RecordBuilder::new(0x00, 0, record_type::SSSLIDEINFO_ATOM);
    builder.write_data(&data);
    builder.build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_build_slide_timing_default() {
        let timing = SlideTiming::default();
        let data = build_slide_timing(&timing).unwrap();
        // 8 bytes header + 16 bytes data
        assert_eq!(data.len(), 24);

        // Verify record type = 1017
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1017);

        // Verify slideTime = 0
        let slide_time = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        assert_eq!(slide_time, 0);
    }

    #[test]
    fn test_build_slide_timing_auto_advance() {
        let timing = SlideTiming::auto_advance(5000);
        let data = build_slide_timing(&timing).unwrap();
        assert_eq!(data.len(), 24);

        // Verify slideTime = 5000
        let slide_time = u32::from_le_bytes([data[8], data[9], data[10], data[11]]);
        assert_eq!(slide_time, 5000);

        // Verify AUTO_ADVANCE_BIT and MANUAL_ADVANCE_BIT set
        let flags = u16::from_le_bytes([data[18], data[19]]);
        assert_ne!(flags & (1 << 0), 0, "MANUAL_ADVANCE_BIT should be set");
        assert_ne!(flags & (1 << 10), 0, "AUTO_ADVANCE_BIT should be set");
    }

    #[test]
    fn test_build_slide_timing_hidden() {
        let timing = SlideTiming::hidden();
        let data = build_slide_timing(&timing).unwrap();

        // Verify HIDDEN_BIT set
        let flags = u16::from_le_bytes([data[18], data[19]]);
        assert_ne!(flags & (1 << 2), 0, "HIDDEN_BIT should be set");
    }
}
