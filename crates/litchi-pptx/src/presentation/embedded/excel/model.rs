use crate::{Error, Result};

const MAX_SERIES: usize = 4_096;
const MAX_POINTS: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 64 * 1024;
const MAX_WORKBOOK_BYTES: usize = 32 * 1024 * 1024;

/// Chart family understood by the embedded-workbook generator.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
    Bubble,
    Doughnut,
    Radar,
    Surface,
    Stock,
    Unknown,
}

/// One immutable chart series.
#[derive(Debug, Clone, PartialEq)]
pub struct Series {
    pub name: String,
    pub values: Vec<f64>,
    pub categories: Vec<String>,
    pub x_values: Vec<f64>,
    pub bubble_sizes: Vec<f64>,
}

impl Series {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            values: Vec::new(),
            categories: Vec::new(),
            x_values: Vec::new(),
            bubble_sizes: Vec::new(),
        }
    }

    #[must_use]
    pub fn with_values(mut self, values: Vec<f64>) -> Self {
        self.values = values;
        self
    }

    #[must_use]
    pub fn with_categories(mut self, categories: Vec<String>) -> Self {
        self.categories = categories;
        self
    }

    #[must_use]
    pub fn with_x_values(mut self, values: Vec<f64>) -> Self {
        self.x_values = values;
        self
    }

    #[must_use]
    pub fn with_bubble_sizes(mut self, values: Vec<f64>) -> Self {
        self.bubble_sizes = values;
        self
    }
}

/// Bounded chart data independent of any old host chart-part type.
#[derive(Debug, Clone, PartialEq)]
pub struct Chart {
    pub kind: Kind,
    pub title: Option<String>,
    pub series: Vec<Series>,
    pub show_legend: bool,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

impl Chart {
    pub fn new(kind: Kind, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            kind,
            title: None,
            series: Vec::new(),
            show_legend: true,
            x,
            y,
            width,
            height,
        }
    }

    #[must_use]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    #[must_use]
    pub fn add_series(mut self, series: Series) -> Self {
        self.series.push(series);
        self
    }

    #[must_use]
    pub fn with_legend(mut self, show: bool) -> Self {
        self.show_legend = show;
        self
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.kind == Kind::Unknown {
            return Err(Error::Invalid(
                "embedded workbook chart kind is unknown".to_string(),
            ));
        }
        if self.width <= 0 || self.height <= 0 {
            return Err(Error::Invalid(
                "embedded workbook chart extents must be positive".to_string(),
            ));
        }
        if self.series.len() > MAX_SERIES {
            return Err(Error::Limit {
                resource: "embedded workbook series",
                limit: MAX_SERIES,
            });
        }
        if self
            .title
            .as_ref()
            .is_some_and(|value| value.len() > MAX_TEXT_BYTES)
        {
            return Err(Error::Limit {
                resource: "embedded workbook title bytes",
                limit: MAX_TEXT_BYTES,
            });
        }
        let mut points = 0usize;
        for series in &self.series {
            if series.name.len() > MAX_TEXT_BYTES
                || series
                    .categories
                    .iter()
                    .any(|value| value.len() > MAX_TEXT_BYTES)
            {
                return Err(Error::Limit {
                    resource: "embedded workbook text bytes",
                    limit: MAX_TEXT_BYTES,
                });
            }
            if series
                .values
                .iter()
                .chain(&series.x_values)
                .chain(&series.bubble_sizes)
                .any(|value| !value.is_finite())
            {
                return Err(Error::Invalid(format!(
                    "series '{}' contains a non-finite number",
                    series.name
                )));
            }
            points = points
                .checked_add(series.values.len())
                .and_then(|value| value.checked_add(series.x_values.len()))
                .and_then(|value| value.checked_add(series.bubble_sizes.len()))
                .ok_or_else(|| Error::Limit {
                    resource: "embedded workbook points",
                    limit: MAX_POINTS,
                })?;
            if points > MAX_POINTS {
                return Err(Error::Limit {
                    resource: "embedded workbook points",
                    limit: MAX_POINTS,
                });
            }
        }
        if self.kind == Kind::Bubble {
            for series in &self.series {
                if series.x_values.len() != series.values.len()
                    || series.bubble_sizes.len() != series.values.len()
                    || series.bubble_sizes.iter().any(|value| *value < 0.0)
                {
                    return Err(Error::Invalid(format!(
                        "bubble series '{}' has invalid lengths",
                        series.name
                    )));
                }
            }
        }
        if self.kind == Kind::Scatter {
            for series in &self.series {
                if !series.x_values.is_empty() && series.x_values.len() != series.values.len() {
                    return Err(Error::Invalid(format!(
                        "scatter series '{}' has invalid lengths",
                        series.name
                    )));
                }
            }
        }
        if self.kind == Kind::Stock && !matches!(self.series.len(), 3 | 4) {
            return Err(Error::Invalid(
                "stock charts require three or four series".to_string(),
            ));
        }
        Ok(())
    }
}

/// An opaque generated workbook ready to be embedded as an OPC part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Workbook {
    bytes: Vec<u8>,
}

impl Workbook {
    /// Generate a workbook from independent chart data.
    pub fn from_chart(chart: &Chart) -> Result<Self> {
        let bytes = super::package::generate(chart)?;
        if bytes.len() > MAX_WORKBOOK_BYTES {
            return Err(Error::Limit {
                resource: "embedded workbook bytes",
                limit: MAX_WORKBOOK_BYTES,
            });
        }
        Ok(Self { bytes })
    }

    /// Borrow the generated XLSX bytes without copying.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Move the generated XLSX bytes out of the wrapper.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}
