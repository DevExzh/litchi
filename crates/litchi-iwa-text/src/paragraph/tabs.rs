//! Archive-free paragraph ruler tab-stop values.

/// Validation failures produced while constructing paragraph ruler values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The implicit tab interval is not finite.
    DefaultIntervalNonFinite,
    /// The implicit tab interval is zero or negative.
    DefaultIntervalNonPositive,
    /// A decimal-tab character is a Unicode control character.
    DecimalCharacterControl,
    /// A tab-stop position is not finite.
    PositionNonFinite,
    /// A tab-stop position is negative.
    PositionNegative,
    /// A tab leader is empty.
    LeaderEmpty,
    /// A tab leader contains a Unicode control character.
    LeaderControl,
    /// Explicit tab-stop positions are not strictly increasing.
    StopsNotIncreasing,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::DefaultIntervalNonFinite | Self::DefaultIntervalNonPositive => {
                "paragraph default tab interval must be finite and positive"
            },
            Self::DecimalCharacterControl => {
                "paragraph decimal-tab character must not be a control character"
            },
            Self::PositionNonFinite | Self::PositionNegative => {
                "paragraph tab-stop position must be finite and nonnegative"
            },
            Self::LeaderEmpty | Self::LeaderControl => {
                "paragraph tab leader must be nonempty and contain no control characters"
            },
            Self::StopsNotIncreasing => "paragraph tab-stop positions must be strictly increasing",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for paragraph ruler value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Default distance between implicit paragraph tab stops.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct DefaultInterval(f32);

impl DefaultInterval {
    /// Native iWork default: half an inch, or 36 typographic points.
    pub const DEFAULT: Self = Self(36.0);

    /// Construct a finite, positive interval in typographic points.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DefaultIntervalNonFinite`] for NaN or infinity and
    /// [`Error::DefaultIntervalNonPositive`] for zero or negative input.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::DefaultIntervalNonFinite);
        }
        if points <= 0.0 {
            return Err(Error::DefaultIntervalNonPositive);
        }
        Ok(Self(points))
    }

    /// Return the interval in typographic points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl Default for DefaultInterval {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Character used to align decimal tab stops.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct DecimalCharacter(char);

impl DecimalCharacter {
    /// The conventional period decimal separator.
    pub const PERIOD: Self = Self('.');
    /// The conventional comma decimal separator.
    pub const COMMA: Self = Self(',');

    /// Construct a decimal-tab character from one non-control Unicode scalar.
    ///
    /// # Errors
    ///
    /// Returns [`Error::DecimalCharacterControl`] for a Unicode control
    /// character.
    pub fn new(character: char) -> Result<Self> {
        if character.is_control() {
            return Err(Error::DecimalCharacterControl);
        }
        Ok(Self(character))
    }

    /// Return the Unicode scalar used for decimal alignment.
    #[must_use]
    pub const fn character(self) -> char {
        self.0
    }
}

impl Default for DecimalCharacter {
    fn default() -> Self {
        Self::PERIOD
    }
}

/// Nonnegative tab-stop position measured from the left text boundary.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct Position(f32);

impl Position {
    /// The left text boundary.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative tab-stop position in typographic points.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PositionNonFinite`] for NaN or infinity and
    /// [`Error::PositionNegative`] for a finite value below zero.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::PositionNonFinite);
        }
        if points < 0.0 {
            return Err(Error::PositionNegative);
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the position in typographic points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native iWork tab-stop alignment.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Alignment {
    #[default]
    Left,
    Center,
    Right,
    Decimal,
}

/// Nonempty leader text repeated between content and a tab stop.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Leader(Box<str>);

impl Leader {
    /// Construct a leader from nonempty, non-control text.
    ///
    /// # Errors
    ///
    /// Returns [`Error::LeaderEmpty`] for empty text and
    /// [`Error::LeaderControl`] when the text contains a Unicode control
    /// character.
    pub fn new(text: &str) -> Result<Self> {
        validate_leader(text)?;
        Ok(Self(text.into()))
    }

    /// Borrow the leader text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl TryFrom<String> for Leader {
    type Error = Error;

    fn try_from(text: String) -> Result<Self> {
        validate_leader(&text)?;
        Ok(Self(text.into_boxed_str()))
    }
}

impl TryFrom<&str> for Leader {
    type Error = Error;

    fn try_from(text: &str) -> Result<Self> {
        Self::new(text)
    }
}

/// One explicit paragraph ruler tab stop.
#[derive(Debug, Clone, PartialEq)]
pub struct Stop {
    /// Position measured from the left text boundary.
    pub position: Position,
    /// Alignment used at this position.
    pub alignment: Alignment,
    /// Optional leader text.
    pub leader: Option<Leader>,
}

impl Stop {
    /// Construct a tab stop without a leader.
    #[must_use]
    pub const fn new(position: Position, alignment: Alignment) -> Self {
        Self {
            position,
            alignment,
            leader: None,
        }
    }

    /// Add a leader to this tab stop.
    #[must_use]
    pub fn with_leader(mut self, leader: Leader) -> Self {
        self.leader = Some(leader);
        self
    }
}

/// Ordered explicit paragraph tab stops.
///
/// Native iWork ruler state requires positions to be strictly increasing.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Stops(Box<[Stop]>);

impl Stops {
    /// Validate and store ordered explicit tab stops.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StopsNotIncreasing`] when adjacent positions are not
    /// strictly increasing.
    pub fn new(stops: Vec<Stop>) -> Result<Self> {
        if stops
            .windows(2)
            .any(|pair| pair[0].position >= pair[1].position)
        {
            return Err(Error::StopsNotIncreasing);
        }
        Ok(Self(stops.into_boxed_slice()))
    }

    /// Borrow the ordered tab stops.
    #[must_use]
    pub fn as_slice(&self) -> &[Stop] {
        &self.0
    }

    /// Return whether no explicit tab stops are present.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

impl AsRef<[Stop]> for Stops {
    fn as_ref(&self) -> &[Stop] {
        self.as_slice()
    }
}

fn validate_leader(text: &str) -> Result<()> {
    if text.is_empty() {
        return Err(Error::LeaderEmpty);
    }
    if text.chars().any(char::is_control) {
        return Err(Error::LeaderControl);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn scalar_values_are_strict_and_normalized() {
        assert_eq!(DefaultInterval::default().points(), 36.0);
        assert_eq!(DefaultInterval::from_points(54.0).unwrap().points(), 54.0);
        assert_eq!(
            DefaultInterval::from_points(0.0),
            Err(Error::DefaultIntervalNonPositive)
        );
        assert_eq!(
            DefaultInterval::from_points(f32::INFINITY),
            Err(Error::DefaultIntervalNonFinite)
        );
        assert_eq!(DecimalCharacter::COMMA.character(), ',');
        assert_eq!(DecimalCharacter::new('٫').unwrap().character(), '٫');
        assert_eq!(
            DecimalCharacter::new('\n'),
            Err(Error::DecimalCharacterControl)
        );
        assert_eq!(
            Position::from_points(-0.0).unwrap().points().to_bits(),
            0.0_f32.to_bits()
        );
        assert_eq!(Position::from_points(-0.1), Err(Error::PositionNegative));
        assert_eq!(
            Position::from_points(f32::NAN),
            Err(Error::PositionNonFinite)
        );
        assert_eq!(Leader::new(""), Err(Error::LeaderEmpty));
        assert_eq!(Leader::new("\t"), Err(Error::LeaderControl));
    }

    #[test]
    fn tab_stop_values_round_trip_without_native_state() {
        let stops = Stops::new(vec![
            Stop::new(Position::from_points(48.5).unwrap(), Alignment::Left),
            Stop::new(Position::from_points(72.0).unwrap(), Alignment::Decimal)
                .with_leader(Leader::new(".").unwrap()),
        ])
        .unwrap();
        assert_eq!(stops.as_slice().len(), 2);
        assert_eq!(stops.as_slice()[1].leader.as_ref().unwrap().as_str(), ".");
        assert!(matches!(
            Stops::new(vec![
                Stop::new(Position::from_points(12.0).unwrap(), Alignment::Center),
                Stop::new(Position::from_points(12.0).unwrap(), Alignment::Right),
            ]),
            Err(Error::StopsNotIncreasing)
        ));
    }
}
