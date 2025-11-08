use crate::Result;
use std::fmt;
use std::str::FromStr;

/// RGB color representation.
///
/// Represents a color using red, green, and blue components, each in the range 0-255.
/// Supports parsing from hex strings (#RRGGBB, #RGB) and CSS3 named colors.
///
/// # Examples
///
/// ```rust
/// use litchi::common::RGBColor;
///
/// // Create a red color
/// let red = RGBColor::new(255, 0, 0);
///
/// // Create from hex string
/// let blue = RGBColor::from_hex("0000FF").unwrap();
/// let green: RGBColor = "#0f0".parse().unwrap();
///
/// // Create from CSS3 named color
/// let yellow = RGBColor::from_name("yellow").unwrap();
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RGBColor {
    /// Red component (0-255)
    pub r: u8,
    /// Green component (0-255)
    pub g: u8,
    /// Blue component (0-255)
    pub b: u8,
}

impl RGBColor {
    /// Create a new RGB color.
    ///
    /// # Arguments
    ///
    /// * `r` - Red component (0-255)
    /// * `g` - Green component (0-255)
    /// * `b` - Blue component (0-255)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color = RGBColor::new(255, 128, 0); // Orange
    /// ```
    #[inline]
    pub const fn new(r: u8, g: u8, b: u8) -> Self {
        Self { r, g, b }
    }

    /// Create an RGB color from a hex string.
    ///
    /// Supports both #RRGGBB and #RGB formats.
    ///
    /// # Arguments
    ///
    /// * `hex` - Hex color string (e.g., "FF0000", "#FF0000", or "#F00")
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let red = RGBColor::from_hex("FF0000").unwrap();
    /// let blue = RGBColor::from_hex("#0000FF").unwrap();
    /// let green = RGBColor::from_hex("#0f0").unwrap(); // Short form
    /// ```
    pub fn from_hex(hex: &str) -> Result<Self> {
        let hex = hex.trim().trim_start_matches('#');

        match hex.len() {
            3 => {
                // #RGB format - expand to #RRGGBB
                let r = u8::from_str_radix(&hex[0..1], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex digit for red".to_string()))?;
                let g = u8::from_str_radix(&hex[1..2], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex digit for green".to_string()))?;
                let b = u8::from_str_radix(&hex[2..3], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex digit for blue".to_string()))?;

                // Expand single digit to double (e.g., F -> FF)
                Ok(Self::new(r * 17, g * 17, b * 17))
            },
            6 => {
                // #RRGGBB format
                let r = u8::from_str_radix(&hex[0..2], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex value for red".to_string()))?;
                let g = u8::from_str_radix(&hex[2..4], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex value for green".to_string()))?;
                let b = u8::from_str_radix(&hex[4..6], 16)
                    .map_err(|_| crate::Error::Other("Invalid hex value for blue".to_string()))?;

                Ok(Self::new(r, g, b))
            },
            _ => Err(crate::Error::Other(format!(
                "Invalid hex color format '{}', expected #RGB or #RRGGBB",
                hex
            ))),
        }
    }

    /// Create from CSS3 named color.
    ///
    /// # Arguments
    ///
    /// * `name` - CSS3 color name (case-insensitive)
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let red = RGBColor::from_name("red").unwrap();
    /// let blue = RGBColor::from_name("blue").unwrap();
    /// assert_eq!(red, RGBColor::new(255, 0, 0));
    /// ```
    pub fn from_name(name: &str) -> Result<Self> {
        // Basic CSS3 named colors - commonly used across all Office formats
        let color_lower = name.to_lowercase();
        match color_lower.as_str() {
            // Primary colors
            "black" => Ok(Self::new(0, 0, 0)),
            "white" => Ok(Self::new(255, 255, 255)),
            "red" => Ok(Self::new(255, 0, 0)),
            "green" => Ok(Self::new(0, 128, 0)),
            "blue" => Ok(Self::new(0, 0, 255)),
            "lime" => Ok(Self::new(0, 255, 0)),
            "yellow" => Ok(Self::new(255, 255, 0)),
            "cyan" | "aqua" => Ok(Self::new(0, 255, 255)),
            "magenta" | "fuchsia" => Ok(Self::new(255, 0, 255)),
            // Grays
            "silver" => Ok(Self::new(192, 192, 192)),
            "gray" | "grey" => Ok(Self::new(128, 128, 128)),
            // Extended colors
            "maroon" => Ok(Self::new(128, 0, 0)),
            "olive" => Ok(Self::new(128, 128, 0)),
            "navy" => Ok(Self::new(0, 0, 128)),
            "purple" => Ok(Self::new(128, 0, 128)),
            "teal" => Ok(Self::new(0, 128, 128)),
            "orange" => Ok(Self::new(255, 165, 0)),
            // More common colors
            "pink" => Ok(Self::new(255, 192, 203)),
            "brown" => Ok(Self::new(165, 42, 42)),
            "gold" => Ok(Self::new(255, 215, 0)),
            "transparent" => Ok(Self::new(0, 0, 0)), // Treat as black
            _ => Err(crate::Error::Other(format!(
                "Unknown color name '{}'",
                name
            ))),
        }
    }

    /// Convert to hex string (without # prefix, uppercase).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color = RGBColor::new(255, 0, 0);
    /// assert_eq!(color.to_hex(), "FF0000");
    /// ```
    pub fn to_hex(&self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// Convert to hex string with # prefix (lowercase).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color = RGBColor::new(255, 0, 0);
    /// assert_eq!(color.to_hex_string(), "#ff0000");
    /// ```
    #[inline]
    pub fn to_hex_string(&self) -> String {
        format!("#{:02x}{:02x}{:02x}", self.r, self.g, self.b)
    }

    /// Get red component
    #[inline]
    pub const fn red(&self) -> u8 {
        self.r
    }

    /// Get green component
    #[inline]
    pub const fn green(&self) -> u8 {
        self.g
    }

    /// Get blue component
    #[inline]
    pub const fn blue(&self) -> u8 {
        self.b
    }
}

// Predefined common colors as constants
impl RGBColor {
    /// Black color (#000000)
    pub const BLACK: Self = Self::new(0, 0, 0);
    /// White color (#ffffff)
    pub const WHITE: Self = Self::new(255, 255, 255);
    /// Red color (#ff0000)
    pub const RED: Self = Self::new(255, 0, 0);
    /// Green color (#008000)
    pub const GREEN: Self = Self::new(0, 128, 0);
    /// Blue color (#0000ff)
    pub const BLUE: Self = Self::new(0, 0, 255);
    /// Lime color (#00ff00)
    pub const LIME: Self = Self::new(0, 255, 0);
    /// Yellow color (#ffff00)
    pub const YELLOW: Self = Self::new(255, 255, 0);
    /// Cyan/Aqua color (#00ffff)
    pub const CYAN: Self = Self::new(0, 255, 255);
    /// Magenta/Fuchsia color (#ff00ff)
    pub const MAGENTA: Self = Self::new(255, 0, 255);
    /// Silver color (#c0c0c0)
    pub const SILVER: Self = Self::new(192, 192, 192);
    /// Gray color (#808080)
    pub const GRAY: Self = Self::new(128, 128, 128);
}

impl FromStr for RGBColor {
    type Err = crate::Error;

    /// Parse color from string (hex or named).
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi::common::RGBColor;
    ///
    /// let color: RGBColor = "#ff0000".parse().unwrap();
    /// let color2: RGBColor = "red".parse().unwrap();
    /// let color3: RGBColor = "#f00".parse().unwrap();
    /// ```
    fn from_str(s: &str) -> Result<Self> {
        let trimmed = s.trim();

        // Try hex first
        if trimmed.starts_with('#') || trimmed.chars().all(|c| c.is_ascii_hexdigit()) {
            Self::from_hex(trimmed)
        } else {
            // Try named color
            Self::from_name(trimmed)
        }
    }
}

impl fmt::Display for RGBColor {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "#{}", self.to_hex())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_new_color() {
        let color = RGBColor::new(255, 128, 64);
        assert_eq!(color.red(), 255);
        assert_eq!(color.green(), 128);
        assert_eq!(color.blue(), 64);
    }

    #[test]
    fn test_to_hex() {
        let color = RGBColor::new(255, 128, 64);
        assert_eq!(color.to_hex(), "FF8040");

        let color = RGBColor::new(0, 0, 0);
        assert_eq!(color.to_hex(), "000000");

        let color = RGBColor::new(255, 255, 255);
        assert_eq!(color.to_hex(), "FFFFFF");
    }

    #[test]
    fn test_to_hex_string() {
        let color = RGBColor::new(255, 0, 0);
        assert_eq!(color.to_hex_string(), "#ff0000");
    }

    #[test]
    fn test_from_hex_rrggbb() {
        let color = RGBColor::from_hex("#ff0000").unwrap();
        assert_eq!(color, RGBColor::new(255, 0, 0));

        let color = RGBColor::from_hex("ff0000").unwrap(); // Without #
        assert_eq!(color, RGBColor::new(255, 0, 0));

        let color = RGBColor::from_hex("#00ff00").unwrap();
        assert_eq!(color, RGBColor::new(0, 255, 0));
    }

    #[test]
    fn test_from_hex_rgb() {
        let color = RGBColor::from_hex("#f00").unwrap();
        assert_eq!(color, RGBColor::new(255, 0, 0));

        let color = RGBColor::from_hex("#0f0").unwrap();
        assert_eq!(color, RGBColor::new(0, 255, 0));

        let color = RGBColor::from_hex("#abc").unwrap();
        assert_eq!(color, RGBColor::new(170, 187, 204));
    }

    #[test]
    fn test_from_name() {
        assert_eq!(
            RGBColor::from_name("red").unwrap(),
            RGBColor::new(255, 0, 0)
        );
        assert_eq!(
            RGBColor::from_name("green").unwrap(),
            RGBColor::new(0, 128, 0)
        );
        assert_eq!(
            RGBColor::from_name("blue").unwrap(),
            RGBColor::new(0, 0, 255)
        );
        assert_eq!(
            RGBColor::from_name("white").unwrap(),
            RGBColor::new(255, 255, 255)
        );
        assert_eq!(
            RGBColor::from_name("black").unwrap(),
            RGBColor::new(0, 0, 0)
        );

        // Case insensitive
        assert_eq!(
            RGBColor::from_name("RED").unwrap(),
            RGBColor::new(255, 0, 0)
        );
        assert_eq!(
            RGBColor::from_name("Red").unwrap(),
            RGBColor::new(255, 0, 0)
        );
    }

    #[test]
    fn test_from_str() {
        let color: RGBColor = "#ff0000".parse().unwrap();
        assert_eq!(color, RGBColor::new(255, 0, 0));

        let color: RGBColor = "red".parse().unwrap();
        assert_eq!(color, RGBColor::new(255, 0, 0));

        let color: RGBColor = "#f00".parse().unwrap();
        assert_eq!(color, RGBColor::new(255, 0, 0));
    }

    #[test]
    fn test_display() {
        let color = RGBColor::new(255, 128, 64);
        assert_eq!(color.to_string(), "#FF8040");
    }

    #[test]
    fn test_constants() {
        assert_eq!(RGBColor::BLACK, RGBColor::new(0, 0, 0));
        assert_eq!(RGBColor::WHITE, RGBColor::new(255, 255, 255));
        assert_eq!(RGBColor::RED, RGBColor::new(255, 0, 0));
        assert_eq!(RGBColor::GREEN, RGBColor::new(0, 128, 0));
        assert_eq!(RGBColor::BLUE, RGBColor::new(0, 0, 255));
    }

    #[test]
    fn test_invalid_hex() {
        assert!(RGBColor::from_hex("#gg0000").is_err());
        assert!(RGBColor::from_hex("#ff00").is_err()); // Wrong length
        assert!(RGBColor::from_hex("#ff00000").is_err()); // Wrong length
    }

    #[test]
    fn test_invalid_name() {
        assert!(RGBColor::from_name("notacolor").is_err());
    }
}
