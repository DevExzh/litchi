/// Optimized SVG Path Builder
///
/// Creates minimal SVG path strings similar to SVGO optimization
use std::fmt::Write;

/// Path command type
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum PathCommand {
    MoveTo {
        x: f64,
        y: f64,
    },
    LineTo {
        x: f64,
        y: f64,
    },
    CubicTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    QuadTo {
        x1: f64,
        y1: f64,
        x: f64,
        y: f64,
    },
    Arc {
        rx: f64,
        ry: f64,
        rotation: f64,
        large_arc: bool,
        sweep: bool,
        x: f64,
        y: f64,
    },
    Close,
}

/// SVG Path Builder with optimization
pub struct PathBuilder {
    commands: Vec<PathCommand>,
    current_pos: Option<(f64, f64)>,
    precision: u32,
}

impl PathBuilder {
    /// Create new path builder
    pub fn new() -> Self {
        Self {
            commands: Vec::new(),
            current_pos: None,
            precision: 2, // Default to 2 decimal places like SVGO
        }
    }

    /// Set decimal precision
    pub fn with_precision(mut self, precision: u32) -> Self {
        self.precision = precision;
        self
    }

    /// Add MoveTo command
    pub fn move_to(&mut self, x: f64, y: f64) {
        self.commands.push(PathCommand::MoveTo { x, y });
        self.current_pos = Some((x, y));
    }

    /// Add LineTo command
    pub fn line_to(&mut self, x: f64, y: f64) {
        self.commands.push(PathCommand::LineTo { x, y });
        self.current_pos = Some((x, y));
    }

    /// Add CubicTo (Bezier) command
    pub fn cubic_to(&mut self, x1: f64, y1: f64, x2: f64, y2: f64, x: f64, y: f64) {
        self.commands.push(PathCommand::CubicTo {
            x1,
            y1,
            x2,
            y2,
            x,
            y,
        });
        self.current_pos = Some((x, y));
    }

    /// Add QuadTo (Quadratic Bezier) command
    pub fn quad_to(&mut self, x1: f64, y1: f64, x: f64, y: f64) {
        self.commands.push(PathCommand::QuadTo { x1, y1, x, y });
        self.current_pos = Some((x, y));
    }

    /// Add Arc command
    pub fn arc(
        &mut self,
        rx: f64,
        ry: f64,
        rotation: f64,
        large_arc: bool,
        sweep: bool,
        x: f64,
        y: f64,
    ) {
        self.commands.push(PathCommand::Arc {
            rx,
            ry,
            rotation,
            large_arc,
            sweep,
            x,
            y,
        });
        self.current_pos = Some((x, y));
    }

    /// Add Close command
    pub fn close(&mut self) {
        self.commands.push(PathCommand::Close);
    }

    /// Build optimized SVG path string
    pub fn build(&self) -> String {
        if self.commands.is_empty() {
            return String::new();
        }

        let mut result = String::with_capacity(self.commands.len() * 10);
        let mut prev_cmd = None;

        for cmd in &self.commands {
            match cmd {
                PathCommand::MoveTo { x, y } => {
                    self.write_command(&mut result, 'M', prev_cmd);
                    self.write_coords(&mut result, &[*x, *y]);
                    prev_cmd = Some('M');
                },
                PathCommand::LineTo { x, y } => {
                    self.write_command(&mut result, 'L', prev_cmd);
                    self.write_coords(&mut result, &[*x, *y]);
                    prev_cmd = Some('L');
                },
                PathCommand::CubicTo {
                    x1,
                    y1,
                    x2,
                    y2,
                    x,
                    y,
                } => {
                    self.write_command(&mut result, 'C', prev_cmd);
                    self.write_coords(&mut result, &[*x1, *y1, *x2, *y2, *x, *y]);
                    prev_cmd = Some('C');
                },
                PathCommand::QuadTo { x1, y1, x, y } => {
                    self.write_command(&mut result, 'Q', prev_cmd);
                    self.write_coords(&mut result, &[*x1, *y1, *x, *y]);
                    prev_cmd = Some('Q');
                },
                PathCommand::Arc {
                    rx,
                    ry,
                    rotation,
                    large_arc,
                    sweep,
                    x,
                    y,
                } => {
                    self.write_command(&mut result, 'A', prev_cmd);
                    self.write_coords(&mut result, &[*rx, *ry, *rotation]);
                    write!(
                        &mut result,
                        " {}{}",
                        if *large_arc { '1' } else { '0' },
                        if *sweep { '1' } else { '0' }
                    )
                    .ok();
                    self.write_coords(&mut result, &[*x, *y]);
                    prev_cmd = Some('A');
                },
                PathCommand::Close => {
                    result.push('z');
                    prev_cmd = Some('z');
                },
            }
        }

        result
    }

    /// Write command letter (omit if repeating same command)
    fn write_command(&self, out: &mut String, cmd: char, prev: Option<char>) {
        if prev.is_none() || prev != Some(cmd) {
            out.push(cmd);
        }
    }

    /// Write coordinates with optimal precision
    fn write_coords(&self, out: &mut String, coords: &[f64]) {
        for (i, &val) in coords.iter().enumerate() {
            if i > 0 {
                out.push(' ');
            }
            self.write_number(out, val);
        }
    }

    /// Write number with minimal digits (SVGO-style)
    fn write_number(&self, out: &mut String, val: f64) {
        // Round to precision
        let factor = 10f64.powi(self.precision as i32);
        let rounded = (val * factor).round() / factor;

        // Check if it's an integer
        if rounded.fract().abs() < 1e-10 {
            write!(out, "{}", rounded as i64).ok();
        } else {
            // Format with precision, then trim trailing zeros
            let formatted = format!("{:.prec$}", rounded, prec = self.precision as usize);
            let trimmed = formatted.trim_end_matches('0').trim_end_matches('.');
            out.push_str(trimmed);
        }
    }

    /// Optimize path (remove redundant commands, simplify)
    pub fn optimize(&mut self) {
        if self.commands.len() < 2 {
            return;
        }

        let mut optimized = Vec::with_capacity(self.commands.len());
        let mut last_pos = (0.0, 0.0);

        for cmd in self.commands.drain(..) {
            match cmd {
                PathCommand::MoveTo { x, y } => {
                    // Skip redundant MoveTo to same position
                    if optimized.is_empty()
                        || (x - last_pos.0).abs() > 1e-6
                        || (y - last_pos.1).abs() > 1e-6
                    {
                        optimized.push(cmd);
                        last_pos = (x, y);
                    }
                },
                PathCommand::LineTo { x, y } => {
                    // Skip zero-length lines
                    if (x - last_pos.0).abs() > 1e-6 || (y - last_pos.1).abs() > 1e-6 {
                        optimized.push(cmd);
                        last_pos = (x, y);
                    }
                },
                _ => {
                    optimized.push(cmd);
                    if let Some(pos) = cmd.end_point() {
                        last_pos = pos;
                    }
                },
            }
        }

        self.commands = optimized;
    }
}

impl Default for PathBuilder {
    fn default() -> Self {
        Self::new()
    }
}

impl PathCommand {
    /// Get the end point of this command
    fn end_point(&self) -> Option<(f64, f64)> {
        match self {
            Self::MoveTo { x, y } | Self::LineTo { x, y } => Some((*x, *y)),
            Self::CubicTo { x, y, .. } => Some((*x, *y)),
            Self::QuadTo { x, y, .. } => Some((*x, *y)),
            Self::Arc { x, y, .. } => Some((*x, *y)),
            Self::Close => None,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_simple_path() {
        let mut builder = PathBuilder::new();
        builder.move_to(10.0, 20.0);
        builder.line_to(30.0, 40.0);
        builder.close();

        let path = builder.build();
        assert_eq!(path, "M10 20L30 40z");
    }

    #[test]
    fn test_precision() {
        let mut builder = PathBuilder::new().with_precision(1);
        builder.move_to(10.123, 20.567);
        builder.line_to(30.89, 40.12);

        let path = builder.build();
        assert_eq!(path, "M10.1 20.6L30.9 40.1");
    }

    #[test]
    fn test_optimization() {
        let mut builder = PathBuilder::new();
        builder.move_to(10.0, 20.0);
        builder.move_to(10.0, 20.0); // Redundant
        builder.line_to(10.0, 20.0); // Zero length
        builder.line_to(30.0, 40.0);

        builder.optimize();
        let path = builder.build();
        assert_eq!(path, "M10 20L30 40");
    }
}
