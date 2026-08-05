//! Contextual chart values used by the PresentationML reader and writer.

/// A supported DrawingML chart family.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Type {
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

/// Basic information discovered in a chart part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    pub chart_type: Type,
    pub title: Option<String>,
    pub has_legend: bool,
}

/// One chart data series.
#[derive(Debug, Clone)]
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

    pub fn with_values(mut self, values: Vec<f64>) -> Self {
        self.values = values;
        self
    }

    pub fn with_categories(mut self, categories: Vec<String>) -> Self {
        self.categories = categories;
        self
    }

    pub fn with_x_values(mut self, x_values: Vec<f64>) -> Self {
        self.x_values = x_values;
        self
    }

    pub fn with_bubble_sizes(mut self, bubble_sizes: Vec<f64>) -> Self {
        self.bubble_sizes = bubble_sizes;
        self
    }
}

/// Chart data for PresentationML authoring.
#[derive(Debug, Clone)]
pub struct Chart {
    pub chart_type: Type,
    pub title: Option<String>,
    pub series: Vec<Series>,
    pub show_legend: bool,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

impl Chart {
    pub fn new(chart_type: Type, x: i64, y: i64, width: i64, height: i64) -> Self {
        Self {
            chart_type,
            title: None,
            series: Vec::new(),
            show_legend: true,
            x,
            y,
            width,
            height,
        }
    }

    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    pub fn add_series(mut self, series: Series) -> Self {
        self.series.push(series);
        self
    }

    pub fn with_legend(mut self, show: bool) -> Self {
        self.show_legend = show;
        self
    }
}
