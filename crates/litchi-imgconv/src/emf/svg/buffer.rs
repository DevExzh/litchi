/// SVG element buffer for in-place optimization
///
/// This module provides buffering and optimization of SVG elements during conversion
/// to minimize output size without post-processing.
use super::state::DeviceContext;
use crate::svg_utils::write_num;

/// One buffered line segment: `(x1, y1, x2, y2)`.
type Segment = (f64, f64, f64, f64);

/// Lines are merged into one path once this many are pending.
const RUN_FLUSH_THRESHOLD: usize = 10;

/// The pending run of `<line>` elements that may still merge into one path.
///
/// A run holds only lines drawn with the same stroke under the same clip
/// region, so the attributes of the run are the attributes of every line in
/// it.
///
/// ```text
/// Empty --line--------------------------> Run
/// Run   --line, same stroke and clip----> Run   (emits the run at the threshold)
/// Run   --line, other stroke or clip----> Run   (emits the old run first)
/// Run   --any other element, or end-----> Empty (emits the run)
/// ```
enum LineRunState {
    /// No line is pending.
    Empty,
    /// Lines that share `stroke` and `clip`. `lines` is empty right after a
    /// threshold flush, which keeps the run open for the next line.
    Run {
        /// Serialized stroke attributes of the run.
        stroke: String,
        /// Serialized `clip-path` attribute of the run; empty when unclipped.
        clip: String,
        /// Pending segments in drawing order.
        lines: Vec<Segment>,
    },
}

/// Buffer for grouping and optimizing SVG elements
pub(super) struct ElementBuffer {
    /// Finalized elements ready for output
    pub elements: Vec<String>,
    /// Lines waiting to be merged into a path
    run: LineRunState,
}

impl ElementBuffer {
    pub(super) fn new() -> Self {
        Self {
            elements: Vec::new(),
            run: LineRunState::Empty,
        }
    }

    /// Add an element, possibly merging with buffered elements
    pub(super) fn add_element(&mut self, element: String, dc: &DeviceContext) {
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
        let Some((x1, y1, x2, y2, _stroke)) = Self::parse_line(&line) else {
            // Can't parse - flush and add as-is
            self.flush();
            self.elements.push(line);
            return;
        };
        let line_stroke = dc.get_stroke_attrs();
        let line_clip = dc.clip_attr();

        match &mut self.run {
            LineRunState::Run {
                stroke,
                clip,
                lines,
            } if *stroke == line_stroke && *clip == line_clip => {
                // Same stroke and clip - buffer this line
                lines.push((x1, y1, x2, y2));

                // Flush when buffer gets large enough
                if lines.len() >= RUN_FLUSH_THRESHOLD {
                    Self::output_run(&mut self.elements, lines, stroke, clip);
                    lines.clear();
                }
            },
            LineRunState::Empty | LineRunState::Run { .. } => {
                // Different stroke or clip - flush old lines and start a new run
                self.flush();
                self.run = LineRunState::Run {
                    stroke: line_stroke,
                    clip: line_clip,
                    lines: vec![(x1, y1, x2, y2)],
                };
            },
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

    /// Emit the lines of one run as paths or individual line elements
    fn output_run(elements: &mut Vec<String>, lines: &[Segment], stroke: &str, clip: &str) {
        // Group lines into continuous paths
        let mut start = 0;
        for index in 1..lines.len() {
            let (_, _, prev_x2, prev_y2) = lines[index - 1];
            let (x1, y1, _, _) = lines[index];
            // Check if this line connects to the previous one
            let connects = (x1 - prev_x2).abs() < 0.1 && (y1 - prev_y2).abs() < 0.1;
            if !connects {
                // Disconnected - flush current path and start new one
                Self::output_path_or_lines(elements, &lines[start..index], stroke, clip);
                start = index;
            }
        }
        // Flush remaining path
        Self::output_path_or_lines(elements, &lines[start..], stroke, clip);
    }

    /// Output a path or individual lines depending on what's more efficient (optimized)
    fn output_path_or_lines(
        elements: &mut Vec<String>,
        lines: &[Segment],
        stroke: &str,
        clip: &str,
    ) {
        if lines.len() >= 3 {
            // Multiple connected lines - use path (more compact)
            let mut path_data = String::with_capacity(lines.len() * 15);
            let (x1, y1, _, _) = lines[0];
            path_data.push('M');
            write_num(&mut path_data, x1);
            path_data.push(' ');
            write_num(&mut path_data, y1);

            for &(_, _, x2, y2) in lines {
                path_data.push('L');
                write_num(&mut path_data, x2);
                path_data.push(' ');
                write_num(&mut path_data, y2);
            }

            let mut s = String::with_capacity(path_data.len() + stroke.len() + clip.len() + 32);
            s.push_str("<path d=\"");
            s.push_str(&path_data);
            s.push_str("\" fill=\"none\" ");
            Self::close_element(&mut s, stroke, clip);
            elements.push(s);
        } else {
            // 1 or 2 lines - output as individual lines (similar size to path)
            for &(x1, y1, x2, y2) in lines {
                let mut s = String::with_capacity(96);
                s.push_str("<line x1=\"");
                write_num(&mut s, x1);
                s.push_str("\" y1=\"");
                write_num(&mut s, y1);
                s.push_str("\" x2=\"");
                write_num(&mut s, x2);
                s.push_str("\" y2=\"");
                write_num(&mut s, y2);
                s.push_str("\" ");
                Self::close_element(&mut s, stroke, clip);
                elements.push(s);
            }
        }
    }

    /// Append the run attributes and close the element.
    fn close_element(element: &mut String, stroke: &str, clip: &str) {
        element.push_str(stroke);
        if !clip.is_empty() {
            element.push(' ');
            element.push_str(clip);
        }
        element.push_str("/>");
    }

    /// Flush all pending buffered elements
    pub(super) fn flush(&mut self) {
        match std::mem::replace(&mut self.run, LineRunState::Empty) {
            LineRunState::Empty => {},
            LineRunState::Run {
                stroke,
                clip,
                lines,
            } => Self::output_run(&mut self.elements, &lines, &stroke, &clip),
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
    use super::ElementBuffer;
    use crate::svg_utils::fmt_num;

    #[test]
    fn test_number_formatting() {
        // Test using the shared fmt_num function
        assert_eq!(fmt_num(10.0), "10");
        assert_eq!(fmt_num(10.5), "10.5");
        assert_eq!(fmt_num(10.50), "10.5");
        assert_eq!(fmt_num(10.123), "10.12");
    }

    #[test]
    fn malformed_line_is_not_buffered() {
        assert!(
            ElementBuffer::parse_line(
                r#"<line x1="1" y1="2" x2="not-a-number" y2="4" stroke="black"/>"#
            )
            .is_none()
        );
    }

    #[test]
    fn flushing_without_pending_lines_emits_nothing() {
        let mut buffer = ElementBuffer::new();

        buffer.flush();

        assert!(buffer.elements.is_empty());
    }
}
