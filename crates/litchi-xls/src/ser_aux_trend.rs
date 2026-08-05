//! BIFF8 `SerAuxTrend` record (0x104B, MS-XLS 2.4.250) of the Chart Sheet
//! substream (MS-XLS 2.1): a trendline.
//!
//! Everything in this module is INERT: values are stored verbatim and no
//! trendline is computed or rendered. The `ordUser` bound for moving-average
//! trendlines depends on the preceding `SerParent`/`Series` records, and the
//! ignore rules of `fEquation`/`fRSquared` depend on preceding `ObjectLink`
//! records; those cross-record constraints are documented here, not enforced
//! by the record reader.
//!
//! # References
//!
//! - MS-XLS 2.4.250 (SerAuxTrend), 2.5.14 (Boolean), 2.5.40
//!   (ChartNumNillable), 2.5.184 (NilChartNum), 2.5.342 (Xnum)

use super::{Error, Result};

/// Record type of the `SerAuxTrend` record (MS-XLS 2.4.250).
pub(crate) const SER_AUX_TREND_RECORD_TYPE: u16 = 0x104B;

/// Byte length of a `SerAuxTrend` record payload.
const PAYLOAD_LEN: usize = 28;
/// `ChartNumNillable` most-significant bytes marking a `NilChartNum`
/// non-numeric value (MS-XLS 2.5.40).
const NIL_CHART_NUM_MARKER: [u8; 2] = [0xFF, 0xFF];
/// Minimum `ordUser` value for polynomial trendlines (MS-XLS 2.4.250).
const MIN_ORDER: u8 = 0x02;
/// Maximum `ordUser` value for polynomial trendlines (MS-XLS 2.4.250).
const MAX_ORDER: u8 = 0x06;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: SER_AUX_TREND_RECORD_TYPE,
        message: message.into(),
    }
}

/// Parse a `Boolean` byte (MS-XLS 2.5.14).
fn parse_bool1(value: u8, field: &str) -> Result<bool> {
    match value {
        0x00 => Ok(false),
        0x01 => Ok(true),
        other => Err(invalid(format!(
            "SerAuxTrend {field} {other:#04X} is not a Boolean"
        ))),
    }
}

/// The `regt` trendline type (MS-XLS 2.4.250).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum TrendlineKind {
    /// 0x00: polynomial.
    Polynomial = 0x00,
    /// 0x01: exponential.
    Exponential = 0x01,
    /// 0x02: logarithmic.
    Logarithmic = 0x02,
    /// 0x03: power.
    Power = 0x03,
    /// 0x04: moving average.
    MovingAverage = 0x04,
}

impl TrendlineKind {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0x00 => Ok(Self::Polynomial),
            0x01 => Ok(Self::Exponential),
            0x02 => Ok(Self::Logarithmic),
            0x03 => Ok(Self::Power),
            0x04 => Ok(Self::MovingAverage),
            other => Err(invalid(format!(
                "SerAuxTrend regt {other:#04X} is not a defined trendline type"
            ))),
        }
    }
}

/// Typed `SerAuxTrend` record content (MS-XLS 2.4.250): a trendline.
///
/// The `numIntercept` bytes are preserved verbatim as a `ChartNumNillable`
/// union (MS-XLS 2.5.40); [`Self::intercept`] decodes it.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SerAuxTrend {
    /// The trendline type (`regt`).
    kind: TrendlineKind,
    /// Polynomial order or moving average period (`ordUser`); for polynomial
    /// trendlines in 0x02..=0x06.
    order: u8,
    /// Raw `numIntercept` bytes: a `ChartNumNillable` union.
    intercept: [u8; 8],
    /// Whether the trendline equation is displayed (`fEquation`).
    show_equation: bool,
    /// Whether the R-squared value is displayed (`fRSquared`).
    show_r_squared: bool,
    /// Number of periods to forecast forward (`numForecast`).
    forecast: f64,
    /// Number of periods to forecast backward (`numBackcast`).
    backcast: f64,
}

impl SerAuxTrend {
    /// Parse a `SerAuxTrend` record payload.
    pub fn parse(data: &[u8]) -> Result<Self> {
        if data.len() != PAYLOAD_LEN {
            return Err(Error::InvalidLength {
                expected: PAYLOAD_LEN,
                found: data.len(),
            });
        }
        let kind = TrendlineKind::parse(data[0])?;
        let order = data[1];
        // MS-XLS 2.4.250: ordUser MUST be in 0x02..=0x06 for polynomial
        // trendlines; it is ignored (and preserved) for moving average, and
        // ignored for all other types.
        if kind == TrendlineKind::Polynomial && !(MIN_ORDER..=MAX_ORDER).contains(&order) {
            return Err(invalid(format!(
                "SerAuxTrend ordUser {order:#04X} is outside {MIN_ORDER:#04X}..={MAX_ORDER:#04X}"
            )));
        }
        Ok(Self {
            kind,
            order,
            intercept: data[2..10].try_into().expect("length checked"),
            show_equation: parse_bool1(data[10], "fEquation")?,
            show_r_squared: parse_bool1(data[11], "fRSquared")?,
            forecast: f64::from_le_bytes(data[12..20].try_into().expect("length checked")),
            backcast: f64::from_le_bytes(data[20..28].try_into().expect("length checked")),
        })
    }

    /// Serialize back to a complete `SerAuxTrend` record payload.
    pub fn to_payload(&self) -> Vec<u8> {
        let mut payload = Vec::with_capacity(PAYLOAD_LEN);
        payload.push(self.kind as u8);
        payload.push(self.order);
        payload.extend_from_slice(&self.intercept);
        payload.push(u8::from(self.show_equation));
        payload.push(u8::from(self.show_r_squared));
        payload.extend_from_slice(&self.forecast.to_le_bytes());
        payload.extend_from_slice(&self.backcast.to_le_bytes());
        payload
    }

    /// The trendline type (`regt`).
    pub fn kind(&self) -> TrendlineKind {
        self.kind
    }

    /// Polynomial order or moving average period (`ordUser`); preserved
    /// verbatim when ignored for the trendline type.
    pub fn order(&self) -> u8 {
        self.order
    }

    /// Where the trendline intersects the value axis (`numIntercept`), or
    /// `None` when the `ChartNumNillable` union holds a `NilChartNum`
    /// non-numeric value (MS-XLS 2.5.40).
    pub fn intercept(&self) -> Option<f64> {
        if self.intercept[6..8] == NIL_CHART_NUM_MARKER {
            None
        } else {
            Some(f64::from_le_bytes(self.intercept))
        }
    }

    /// The raw `numIntercept` bytes (a `ChartNumNillable` union), preserved
    /// verbatim.
    pub fn intercept_bytes(&self) -> [u8; 8] {
        self.intercept
    }

    /// Whether the trendline equation is displayed (`fEquation`).
    pub fn show_equation(&self) -> bool {
        self.show_equation
    }

    /// Whether the R-squared value is displayed (`fRSquared`).
    pub fn show_r_squared(&self) -> bool {
        self.show_r_squared
    }

    /// Number of periods to forecast forward (`numForecast`).
    pub fn forecast(&self) -> f64 {
        self.forecast
    }

    /// Number of periods to forecast backward (`numBackcast`).
    pub fn backcast(&self) -> f64 {
        self.backcast
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(
        kind: u8,
        order: u8,
        intercept: [u8; 8],
        equation: u8,
        r_squared: u8,
        forecast: f64,
        backcast: f64,
    ) -> Vec<u8> {
        let mut data = Vec::new();
        data.push(kind);
        data.push(order);
        data.extend_from_slice(&intercept);
        data.push(equation);
        data.push(r_squared);
        data.extend_from_slice(&forecast.to_le_bytes());
        data.extend_from_slice(&backcast.to_le_bytes());
        data
    }

    fn nil_intercept() -> [u8; 8] {
        // NilChartNum with type 0x0100 (MS-XLS 2.4.250) and the 0xFFFF marker.
        [0, 0, 0, 0, 0x00, 0x01, 0xFF, 0xFF]
    }

    #[test]
    fn round_trip_all_trendline_kinds() {
        for (kind, expected) in [
            (0x00, TrendlineKind::Polynomial),
            (0x01, TrendlineKind::Exponential),
            (0x02, TrendlineKind::Logarithmic),
            (0x03, TrendlineKind::Power),
            (0x04, TrendlineKind::MovingAverage),
        ] {
            let bytes = record(kind, 0x03, nil_intercept(), 0x01, 0x00, 1.5, -0.5);
            let parsed = SerAuxTrend::parse(&bytes).unwrap();
            assert_eq!(parsed.kind(), expected);
            assert_eq!(parsed.order(), 0x03);
            assert_eq!(parsed.intercept(), None);
            assert_eq!(parsed.intercept_bytes(), nil_intercept());
            assert!(parsed.show_equation());
            assert!(!parsed.show_r_squared());
            assert_eq!(parsed.forecast(), 1.5);
            assert_eq!(parsed.backcast(), -0.5);
            assert_eq!(parsed.to_payload(), bytes);
        }
    }

    #[test]
    fn numeric_intercept_decodes() {
        let bytes = record(0x01, 0x00, 2.75f64.to_le_bytes(), 0x00, 0x01, 0.0, 0.0);
        let parsed = SerAuxTrend::parse(&bytes).unwrap();
        assert_eq!(parsed.intercept(), Some(2.75));
        assert_eq!(parsed.to_payload(), bytes);
    }

    #[test]
    fn rejects_malformed_records() {
        let bytes = record(0x00, 0x02, nil_intercept(), 0x01, 0x01, 0.0, 0.0);
        // Truncated and overlong payloads.
        assert!(SerAuxTrend::parse(&bytes[..26]).is_err());
        assert!(SerAuxTrend::parse(&[bytes.as_slice(), &[0]].concat()).is_err());
        // Undefined regt.
        assert!(SerAuxTrend::parse(&record(0x05, 0x02, nil_intercept(), 0, 0, 0.0, 0.0)).is_err());
        // Polynomial order outside 0x02..=0x06.
        assert!(SerAuxTrend::parse(&record(0x00, 0x01, nil_intercept(), 0, 0, 0.0, 0.0)).is_err());
        assert!(SerAuxTrend::parse(&record(0x00, 0x07, nil_intercept(), 0, 0, 0.0, 0.0)).is_err());
        // Non-Boolean fEquation / fRSquared.
        assert!(
            SerAuxTrend::parse(&record(0x01, 0x00, nil_intercept(), 0x02, 0, 0.0, 0.0)).is_err()
        );
        assert!(
            SerAuxTrend::parse(&record(0x01, 0x00, nil_intercept(), 0, 0xFF, 0.0, 0.0)).is_err()
        );
    }
}
