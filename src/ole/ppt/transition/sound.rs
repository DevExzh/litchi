//! Sound support for slide transitions.
//!
//! Provides structures for transition sounds (similar to animation sounds but specific to transitions).

use crate::ole::ppt::animation::sound::{BuiltinSound, SoundType};

/// Sound information for slide transition.
#[derive(Debug, Clone, PartialEq)]
pub struct TransitionSound {
    /// Sound type and data
    pub sound_type: SoundType,
    /// Sound reference ID
    pub sound_ref: u32,
    /// Loop sound until next sound
    pub loop_sound: bool,
    /// Stop previous sound
    pub stop_previous: bool,
}

impl Default for TransitionSound {
    fn default() -> Self {
        Self::new()
    }
}

impl TransitionSound {
    /// Create a new empty transition sound with default Click sound.
    pub fn new() -> Self {
        Self {
            sound_type: SoundType::Builtin(BuiltinSound::Click),
            sound_ref: BuiltinSound::Click.id(),
            loop_sound: false,
            stop_previous: false,
        }
    }

    /// Create a built-in sound.
    pub fn builtin(sound: BuiltinSound) -> Self {
        Self {
            sound_type: SoundType::Builtin(sound),
            sound_ref: sound.id(),
            loop_sound: false,
            stop_previous: false,
        }
    }

    /// Create an embedded sound from data.
    pub fn embedded(name: impl Into<String>, data: Vec<u8>, sound_ref: u32) -> Self {
        Self {
            sound_type: SoundType::Embedded {
                name: name.into(),
                data,
            },
            sound_ref,
            loop_sound: false,
            stop_previous: false,
        }
    }

    /// Create a linked sound from file path.
    pub fn linked(name: impl Into<String>, file_path: impl Into<String>, sound_ref: u32) -> Self {
        Self {
            sound_type: SoundType::Linked {
                name: name.into(),
                file_path: file_path.into(),
            },
            sound_ref,
            loop_sound: false,
            stop_previous: false,
        }
    }

    /// Check if this is a built-in sound.
    pub fn is_builtin(&self) -> bool {
        matches!(self.sound_type, SoundType::Builtin(_))
    }

    /// Check if this is an embedded sound.
    pub fn is_embedded(&self) -> bool {
        matches!(self.sound_type, SoundType::Embedded { .. })
    }

    /// Check if this is a linked sound.
    pub fn is_linked(&self) -> bool {
        matches!(self.sound_type, SoundType::Linked { .. })
    }

    /// Get the sound name.
    pub fn name(&self) -> &str {
        match &self.sound_type {
            SoundType::Builtin(sound) => sound.name(),
            SoundType::Embedded { name, .. } => name,
            SoundType::Linked { name, .. } => name,
        }
    }

    /// Set loop flag.
    pub fn with_loop(mut self, loop_sound: bool) -> Self {
        self.loop_sound = loop_sound;
        self
    }

    /// Set stop previous sound flag.
    pub fn with_stop_previous(mut self, stop: bool) -> Self {
        self.stop_previous = stop;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_builtin_sound() {
        let sound = TransitionSound::builtin(BuiltinSound::Chime);
        assert_eq!(sound.sound_ref, 7);
        assert!(sound.is_builtin());
        assert!(!sound.is_embedded());
        assert!(!sound.is_linked());
        assert_eq!(sound.name(), "Chime");
    }

    #[test]
    fn test_embedded_sound() {
        let data = vec![0u8; 512];
        let sound = TransitionSound::embedded("CustomTransition", data.clone(), 50);
        assert!(sound.is_embedded());
        assert!(!sound.is_linked());
        assert!(!sound.is_builtin());
        assert_eq!(sound.name(), "CustomTransition");
        match &sound.sound_type {
            SoundType::Embedded { data: d, .. } => assert_eq!(d, &data),
            _ => panic!("Expected embedded sound"),
        }
    }

    #[test]
    fn test_linked_sound() {
        let sound = TransitionSound::linked("External", "/sounds/whoosh.wav", 100);
        assert!(sound.is_linked());
        assert!(!sound.is_embedded());
        assert!(!sound.is_builtin());
        assert_eq!(sound.name(), "External");
        match &sound.sound_type {
            SoundType::Linked { file_path, .. } => {
                assert_eq!(file_path, "/sounds/whoosh.wav");
            },
            _ => panic!("Expected linked sound"),
        }
    }

    #[test]
    fn test_sound_flags() {
        let sound = TransitionSound::builtin(BuiltinSound::Click)
            .with_loop(true)
            .with_stop_previous(true);
        assert!(sound.loop_sound);
        assert!(sound.stop_previous);
    }
}
