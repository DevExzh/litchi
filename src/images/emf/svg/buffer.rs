/// SVG element buffer for in-place optimization
///
/// This module provides buffering and optimization of SVG elements during conversion
/// to minimize output size without post-processing.
use super::state::DeviceContext;

/// Buffer for grouping and optimizing SVG elements
pub struct ElementBuffer {
    /// Finalized elements ready for output
    pub elements: Vec<String>,
    /// Pending lines with same stroke for merging into polyline
    pending_lines: Vec<(f64, f64, f64, f64)>, // x1, y1, x2, y2
    /// Current stroke style for pending lines
    current_stroke: Option<String>,
}

impl ElementBuffer {
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            pending_lines: Vec::new(),
            current_stroke: None,
        }
    }

    /// Add an element, possibly merging with buffered elements
    pub fn add_element(&mut self, element: String, dc: &DeviceContext) {
        // Check if this is a line element that can be merged
        if element.starts_with("<line ") {
            self.try_buffer_line(element, dc);
        } else {
            // Flush any pending lines before adding different element
            self.flush();
            self.elements.push(element);
        }
    }

    /// Try to buffer a line for potential merging into polyline
    fn try_buffer_line(&mut self, line: String, dc: &DeviceContext) {
        // Parse line coordinates (simple extraction)
        if let Some((x1, y1, x2, y2, _stroke)) = Self::parse_line(&line) {
            let current_stroke_attrs = dc.get_stroke_attrs();

            // Check if stroke matches current buffer
            if let Some(ref buffered_stroke) = self.current_stroke
                && buffered_stroke == &current_stroke_attrs
            {
                // Same stroke - buffer this line
                self.pending_lines.push((x1, y1, x2, y2));

                // Flush when buffer gets large enough
                if self.pending_lines.len() >= 10 {
                    self.flush_lines_as_path();
                }
                return;
            }

            // Different stroke - flush old lines and start new buffer
            self.flush();
            self.current_stroke = Some(current_stroke_attrs);
            self.pending_lines.push((x1, y1, x2, y2));
        } else {
            // Can't parse - flush and add as-is
            self.flush();
            self.elements.push(line);
        }
    }

    /// Parse line element to extract coordinates
    fn parse_line(line: &str) -> Option<(f64, f64, f64, f64, String)> {
        // Quick parse for x1, y1, x2, y2 from line element
        let mut x1 = None;
        let mut y1 = None;
        let mut x2 = None;
        let mut y2 = None;

        for part in line.split_whitespace() {
            if let Some(val) = part.strip_prefix("x1=\"") {
                x1 = val.trim_end_matches('"').parse().ok();
            } else if let Some(val) = part.strip_prefix("y1=\"") {
                y1 = val.trim_end_matches('"').parse().ok();
            } else if let Some(val) = part.strip_prefix("x2=\"") {
                x2 = val.trim_end_matches('"').parse().ok();
            } else if let Some(val) = part.strip_prefix("y2=\"") {
                y2 = val.trim_end_matches('"').parse().ok();
            }
        }

        if let (Some(x1), Some(y1), Some(x2), Some(y2)) = (x1, y1, x2, y2) {
            // Extract stroke attributes (everything after coordinates)
            let stroke_start = line.find("stroke")?;
            let stroke = line[stroke_start..line.len() - 2].to_string(); // Remove />
            Some((x1, y1, x2, y2, stroke))
        } else {
            None
        }
    }

    /// Flush pending lines as path or individual line elements
    fn flush_lines_as_path(&mut self) {
        if self.pending_lines.is_empty() {
            return;
        }

        // Group lines into continuous paths
        let mut current_path: Vec<(f64, f64, f64, f64)> = Vec::new();
        let stroke = self.current_stroke.as_ref().unwrap().clone();

        for &line in &self.pending_lines {
            let (x1, y1, _x2, _y2) = line;

            if current_path.is_empty() {
                // Start new path
                current_path.push(line);
            } else {
                // Check if this line connects to the previous one
                let (_, _, prev_x2, prev_y2) = current_path.last().unwrap();
                let connects = (x1 - prev_x2).abs() < 0.1 && (y1 - prev_y2).abs() < 0.1;

                if connects {
                    // Continues path
                    current_path.push(line);
                } else {
                    // Disconnected - flush current path and start new one
                    Self::output_path_or_lines(&mut self.elements, &current_path, &stroke);
                    current_path.clear();
                    current_path.push(line);
                }
            }
        }

        // Flush remaining path
        if !current_path.is_empty() {
            Self::output_path_or_lines(&mut self.elements, &current_path, &stroke);
        }

        self.pending_lines.clear();
    }

    /// Output a path or individual lines depending on what's more efficient
    fn output_path_or_lines(
        elements: &mut Vec<String>,
        lines: &[(f64, f64, f64, f64)],
        stroke: &str,
    ) {
        if lines.is_empty() {
            return;
        }

        if lines.len() == 1 {
            // Single line - output as line element
            let (x1, y1, x2, y2) = lines[0];
            elements.push(format!(
                "<line x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\" {}/>",
                Self::fmt(x1),
                Self::fmt(y1),
                Self::fmt(x2),
                Self::fmt(y2),
                stroke
            ));
        } else if lines.len() >= 3 {
            // Multiple connected lines - use path (more compact)
            let mut path_data = String::with_capacity(lines.len() * 15);
            let (x1, y1, _, _) = lines[0];
            path_data.push_str(&format!("M{} {}", Self::fmt(x1), Self::fmt(y1)));

            for &(_, _, x2, y2) in lines.iter() {
                path_data.push_str(&format!("L{} {}", Self::fmt(x2), Self::fmt(y2)));
            }

            elements.push(format!(
                "<path d=\"{}\" fill=\"none\" {}/>",
                path_data, stroke
            ));
        } else {
            // 2 lines - output as individual lines (similar size to path)
            for &(x1, y1, x2, y2) in lines {
                elements.push(format!(
                    "<line x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\" {}/>",
                    Self::fmt(x1),
                    Self::fmt(y1),
                    Self::fmt(x2),
                    Self::fmt(y2),
                    stroke
                ));
            }
        }
    }

    /// Flush all pending buffered elements
    pub fn flush(&mut self) {
        if !self.pending_lines.is_empty() {
            self.flush_lines_as_path();
        }
        self.current_stroke = None;
    }

    /// Format number minimally (remove trailing zeros and unnecessary decimals)
    #[inline]
    fn fmt(n: f64) -> String {
        if n.fract().abs() < 0.01 {
            format!("{:.0}", n)
        } else {
            let s = format!("{:.2}", n);
            s.trim_end_matches('0').trim_end_matches('.').to_string()
        }
    }
}

impl Default for ElementBuffer {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_number_formatting() {
        assert_eq!(ElementBuffer::fmt(10.0), "10");
        assert_eq!(ElementBuffer::fmt(10.5), "10.5");
        assert_eq!(ElementBuffer::fmt(10.50), "10.5");
        assert_eq!(ElementBuffer::fmt(10.123), "10.12");
    }
}
