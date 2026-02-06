//! Sound support for animations.
//!
//! Provides structures for embedded and external sounds in animations.

/// Built-in PowerPoint sound types.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BuiltinSound {
    Applause,
    Arrow,
    Bomb,
    Breeze,
    Camera,
    CashRegister,
    Chime,
    Click,
    Coin,
    DrumRoll,
    Explosion,
    Hammer,
    Laser,
    Push,
    Suction,
    Swoosh,
    Typewriter,
    Voltage,
    Whoosh,
    Wind,
}

impl BuiltinSound {
    /// Get the sound ID for this built-in sound.
    pub fn id(self) -> u32 {
        match self {
            Self::Applause => 1,
            Self::Arrow => 2,
            Self::Bomb => 3,
            Self::Breeze => 4,
            Self::Camera => 5,
            Self::CashRegister => 6,
            Self::Chime => 7,
            Self::Click => 8,
            Self::Coin => 9,
            Self::DrumRoll => 10,
            Self::Explosion => 11,
            Self::Hammer => 12,
            Self::Laser => 13,
            Self::Push => 14,
            Self::Suction => 15,
            Self::Swoosh => 16,
            Self::Typewriter => 17,
            Self::Voltage => 18,
            Self::Whoosh => 19,
            Self::Wind => 20,
        }
    }

    /// Get the display name for this built-in sound.
    pub fn name(self) -> &'static str {
        match self {
            Self::Applause => "Applause",
            Self::Arrow => "Arrow",
            Self::Bomb => "Bomb",
            Self::Breeze => "Breeze",
            Self::Camera => "Camera",
            Self::CashRegister => "Cash Register",
            Self::Chime => "Chime",
            Self::Click => "Click",
            Self::Coin => "Coin",
            Self::DrumRoll => "Drum Roll",
            Self::Explosion => "Explosion",
            Self::Hammer => "Hammer",
            Self::Laser => "Laser",
            Self::Push => "Push",
            Self::Suction => "Suction",
            Self::Swoosh => "Swoosh",
            Self::Typewriter => "Typewriter",
            Self::Voltage => "Voltage",
            Self::Whoosh => "Whoosh",
            Self::Wind => "Wind",
        }
    }

    /// Parse a built-in sound from name (case-insensitive).
    pub fn from_name(name: &str) -> Option<Self> {
        match name.to_lowercase().replace(' ', "").as_str() {
            "applause" => Some(Self::Applause),
            "arrow" => Some(Self::Arrow),
            "bomb" => Some(Self::Bomb),
            "breeze" => Some(Self::Breeze),
            "camera" => Some(Self::Camera),
            "cashregister" => Some(Self::CashRegister),
            "chime" => Some(Self::Chime),
            "click" => Some(Self::Click),
            "coin" => Some(Self::Coin),
            "drumroll" => Some(Self::DrumRoll),
            "explosion" => Some(Self::Explosion),
            "hammer" => Some(Self::Hammer),
            "laser" => Some(Self::Laser),
            "push" => Some(Self::Push),
            "suction" => Some(Self::Suction),
            "swoosh" => Some(Self::Swoosh),
            "typewriter" => Some(Self::Typewriter),
            "voltage" => Some(Self::Voltage),
            "whoosh" => Some(Self::Whoosh),
            "wind" => Some(Self::Wind),
            _ => None,
        }
    }

    /// Parse a built-in sound from ID.
    pub fn from_id(id: u32) -> Option<Self> {
        match id {
            1 => Some(Self::Applause),
            2 => Some(Self::Arrow),
            3 => Some(Self::Bomb),
            4 => Some(Self::Breeze),
            5 => Some(Self::Camera),
            6 => Some(Self::CashRegister),
            7 => Some(Self::Chime),
            8 => Some(Self::Click),
            9 => Some(Self::Coin),
            10 => Some(Self::DrumRoll),
            11 => Some(Self::Explosion),
            12 => Some(Self::Hammer),
            13 => Some(Self::Laser),
            14 => Some(Self::Push),
            15 => Some(Self::Suction),
            16 => Some(Self::Swoosh),
            17 => Some(Self::Typewriter),
            18 => Some(Self::Voltage),
            19 => Some(Self::Whoosh),
            20 => Some(Self::Wind),
            _ => None,
        }
    }

    /// List all available built-in sounds.
    pub fn all() -> &'static [Self] {
        &[
            Self::Applause,
            Self::Arrow,
            Self::Bomb,
            Self::Breeze,
            Self::Camera,
            Self::CashRegister,
            Self::Chime,
            Self::Click,
            Self::Coin,
            Self::DrumRoll,
            Self::Explosion,
            Self::Hammer,
            Self::Laser,
            Self::Push,
            Self::Suction,
            Self::Swoosh,
            Self::Typewriter,
            Self::Voltage,
            Self::Whoosh,
            Self::Wind,
        ]
    }
}

/// Sound type for animations.
#[derive(Debug, Clone, PartialEq)]
pub enum SoundType {
    /// Built-in PowerPoint sound
    Builtin(BuiltinSound),
    /// Embedded sound data
    Embedded { name: String, data: Vec<u8> },
    /// Linked external sound file
    Linked { name: String, file_path: String },
}

/// Sound information for animation.
#[derive(Debug, Clone, PartialEq)]
pub struct AnimationSound {
    /// Sound type and data
    pub sound_type: SoundType,
    /// Sound reference ID
    pub sound_ref: u32,
    /// Loop until next sound
    pub loop_sound: bool,
    /// Stop previous sound when this plays
    pub stop_previous: bool,
}

impl Default for AnimationSound {
    fn default() -> Self {
        Self::new()
    }
}

impl AnimationSound {
    /// Create a new empty sound with default Click sound.
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

    /// Set loop sound flag.
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
        let sound = AnimationSound::builtin(BuiltinSound::Click);
        assert_eq!(sound.sound_ref, 8);
        assert_eq!(sound.name(), "Click");
        assert!(sound.is_builtin());
        assert!(!sound.is_embedded());
        assert!(!sound.is_linked());
    }

    #[test]
    fn test_embedded_sound() {
        let data = vec![0u8; 1024];
        let sound = AnimationSound::embedded("CustomSound", data.clone(), 100);
        assert_eq!(sound.sound_ref, 100);
        assert!(!sound.is_builtin());
        assert!(sound.is_embedded());
        assert!(!sound.is_linked());
        assert_eq!(sound.name(), "CustomSound");
        match &sound.sound_type {
            SoundType::Embedded { data: d, .. } => assert_eq!(d, &data),
            _ => panic!("Expected embedded sound"),
        }
    }

    #[test]
    fn test_linked_sound() {
        let sound = AnimationSound::linked("External", "/path/to/sound.wav", 200);
        assert_eq!(sound.sound_ref, 200);
        assert!(!sound.is_builtin());
        assert!(!sound.is_embedded());
        assert!(sound.is_linked());
        assert_eq!(sound.name(), "External");
        match &sound.sound_type {
            SoundType::Linked { file_path, .. } => {
                assert_eq!(file_path, "/path/to/sound.wav");
            },
            _ => panic!("Expected linked sound"),
        }
    }

    #[test]
    fn test_sound_flags() {
        let sound = AnimationSound::builtin(BuiltinSound::Click)
            .with_loop(true)
            .with_stop_previous(true);
        assert!(sound.loop_sound);
        assert!(sound.stop_previous);
    }

    #[test]
    fn test_builtin_sound_enum_lookup() {
        assert_eq!(BuiltinSound::from_name("click"), Some(BuiltinSound::Click));
        assert_eq!(BuiltinSound::from_name("Click"), Some(BuiltinSound::Click));
        assert_eq!(BuiltinSound::from_name("CLICK"), Some(BuiltinSound::Click));
        assert_eq!(BuiltinSound::from_name("unknown"), None);
    }

    #[test]
    fn test_builtin_sound_enum_id() {
        assert_eq!(BuiltinSound::Click.id(), 8);
        assert_eq!(BuiltinSound::Applause.id(), 1);
        assert_eq!(BuiltinSound::from_id(8), Some(BuiltinSound::Click));
        assert_eq!(BuiltinSound::from_id(1), Some(BuiltinSound::Applause));
        assert_eq!(BuiltinSound::from_id(999), None);
    }

    #[test]
    fn test_builtin_sound_enum_name() {
        assert_eq!(BuiltinSound::Click.name(), "Click");
        assert_eq!(BuiltinSound::Applause.name(), "Applause");
        assert_eq!(BuiltinSound::CashRegister.name(), "Cash Register");
    }

    #[test]
    fn test_builtin_sounds_list() {
        let sounds = BuiltinSound::all();
        assert_eq!(sounds.len(), 20);
        assert!(sounds.contains(&BuiltinSound::Click));
        assert!(sounds.contains(&BuiltinSound::Applause));
    }
}
