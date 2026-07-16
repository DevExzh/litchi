//! Strict public paragraph ruler tab-stop types.

/// Nonnegative tab-stop position measured from the left text boundary.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd, Default)]
pub struct ParagraphTabPosition(f32);

impl ParagraphTabPosition {
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite, nonnegative tab-stop position in typographic points.
    pub fn from_points(points: f32) -> crate::Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(crate::Error::InvalidFormat(
                "paragraph tab-stop position must be finite and nonnegative".to_owned(),
            ));
        }
        Ok(if points == 0.0 {
            Self::ZERO
        } else {
            Self(points)
        })
    }

    /// Return the position in typographic points.
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native iWork tab-stop alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ParagraphTabAlignment {
    #[default]
    Left,
    Center,
    Right,
    Decimal,
}

impl ParagraphTabAlignment {
    pub(crate) const fn native_value(self) -> i32 {
        match self {
            Self::Left => 0,
            Self::Center => 1,
            Self::Right => 2,
            Self::Decimal => 3,
        }
    }

    pub(crate) fn from_native_value(value: i32) -> crate::Result<Self> {
        match value {
            0 => Ok(Self::Left),
            1 => Ok(Self::Center),
            2 => Ok(Self::Right),
            3 => Ok(Self::Decimal),
            _ => Err(crate::Error::InvalidFormat(format!(
                "unsupported native iWork tab-stop alignment {value}"
            ))),
        }
    }
}

/// Nonempty leader text repeated between content and a tab stop.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ParagraphTabLeader(Box<str>);

impl ParagraphTabLeader {
    pub fn new(text: impl Into<String>) -> crate::Result<Self> {
        let text = text.into();
        if text.is_empty() || text.chars().any(char::is_control) {
            return Err(crate::Error::InvalidFormat(
                "paragraph tab leader must be nonempty and contain no control characters"
                    .to_owned(),
            ));
        }
        Ok(Self(text.into_boxed_str()))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// One explicit paragraph ruler tab stop.
#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphTabStop {
    pub position: ParagraphTabPosition,
    pub alignment: ParagraphTabAlignment,
    pub leader: Option<ParagraphTabLeader>,
}

impl ParagraphTabStop {
    pub const fn new(position: ParagraphTabPosition, alignment: ParagraphTabAlignment) -> Self {
        Self {
            position,
            alignment,
            leader: None,
        }
    }

    pub fn with_leader(mut self, leader: ParagraphTabLeader) -> Self {
        self.leader = Some(leader);
        self
    }
}

/// Ordered explicit paragraph tab stops.
///
/// Native iWork ruler state requires positions to be strictly increasing.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParagraphTabStops(Box<[ParagraphTabStop]>);

impl ParagraphTabStops {
    pub fn new(stops: Vec<ParagraphTabStop>) -> crate::Result<Self> {
        if stops
            .windows(2)
            .any(|pair| pair[0].position >= pair[1].position)
        {
            return Err(crate::Error::InvalidFormat(
                "paragraph tab-stop positions must be strictly increasing".to_owned(),
            ));
        }
        Ok(Self(stops.into_boxed_slice()))
    }

    pub fn as_slice(&self) -> &[ParagraphTabStop] {
        &self.0
    }

    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }
}

impl AsRef<[ParagraphTabStop]> for ParagraphTabStops {
    fn as_ref(&self) -> &[ParagraphTabStop] {
        self.as_slice()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn tab_stops_are_strict_typed_and_ordered() {
        let stops = ParagraphTabStops::new(vec![
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(48.5).unwrap(),
                ParagraphTabAlignment::Left,
            ),
            ParagraphTabStop::new(
                ParagraphTabPosition::from_points(72.0).unwrap(),
                ParagraphTabAlignment::Decimal,
            )
            .with_leader(ParagraphTabLeader::new(".").unwrap()),
        ])
        .unwrap();
        assert_eq!(stops.as_slice().len(), 2);
        assert_eq!(stops.as_slice()[1].leader.as_ref().unwrap().as_str(), ".");
        assert_eq!(
            ParagraphTabPosition::from_points(-0.0)
                .unwrap()
                .points()
                .to_bits(),
            0.0_f32.to_bits()
        );
        assert!(ParagraphTabPosition::from_points(-0.1).is_err());
        assert!(ParagraphTabPosition::from_points(f32::NAN).is_err());
        assert!(ParagraphTabLeader::new("").is_err());
        assert!(ParagraphTabLeader::new("\t").is_err());
        assert!(
            ParagraphTabStops::new(vec![
                ParagraphTabStop::new(
                    ParagraphTabPosition::from_points(12.0).unwrap(),
                    ParagraphTabAlignment::Center,
                ),
                ParagraphTabStop::new(
                    ParagraphTabPosition::from_points(12.0).unwrap(),
                    ParagraphTabAlignment::Right,
                ),
            ])
            .is_err()
        );
        assert!(ParagraphTabAlignment::from_native_value(4).is_err());
    }
}
