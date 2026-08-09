//! Contextual semantic models for one `PowerPoint` sound collection.

/// The MS-PPT `SoundBuiltinIdAtom` description domain.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u16)]
pub enum BuiltinId {
    CashRegister = 100,
    Typewriter = 101,
    ScreechingBrakes = 102,
    Whoosh = 103,
    Laser = 104,
    Camera = 105,
    Chime = 106,
    Clapping = 107,
    Applause = 108,
    DriveBy = 109,
    DrumRoll = 110,
    Explosion = 111,
    BreakingGlass = 112,
    Gunshot = 113,
    SlideProjector = 114,
    Ricochet = 115,
    Arrow = 116,
    Bomb = 117,
    Breeze = 118,
    Click = 119,
    Coin = 120,
    Hammer = 121,
    Push = 122,
    Suction = 123,
    Voltage = 124,
    Wind = 125,
}

impl BuiltinId {
    /// Return the native `SoundBuiltinIdAtom` value.
    #[must_use]
    pub const fn value(self) -> u16 {
        self as u16
    }
}

/// One inert embedded sound whose media payload borrows the presentation
/// stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Sound<'a> {
    /// Positive identifier used by `SoundIdRef` fields.
    pub id: u32,
    /// Producer-visible sound name from `SoundNameAtom`.
    pub name: String,
    /// Optional source-file extension from `SoundExtensionAtom`.
    pub extension: Option<String>,
    /// Optional built-in description from `SoundBuiltinIdAtom`.
    pub builtin_id: Option<BuiltinId>,
    /// Borrowed WAV or AIFF bytes from `SoundDataBlob`.
    pub data: &'a [u8],
}

/// A validated MS-PPT `SoundCollectionContainer` in source order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Collection<'a> {
    /// Positive seed used by `PowerPoint` for new sound identifiers.
    pub sound_id_seed: u32,
    /// Embedded sounds in their native record order.
    pub sounds: Vec<Sound<'a>>,
}

impl<'a> Collection<'a> {
    /// Find a sound by its checked native identifier.
    #[must_use]
    pub fn get(&self, id: u32) -> Option<&Sound<'a>> {
        self.sounds.iter().find(|sound| sound.id == id)
    }
}
