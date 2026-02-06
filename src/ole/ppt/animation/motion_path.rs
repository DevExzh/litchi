//! Motion path support for animations.
//!
//! Provides structures for custom and predefined motion paths.

/// Motion path definition for animation.
#[derive(Debug, Clone, PartialEq)]
pub struct MotionPath {
    /// Path type
    pub path_type: MotionPathType,
    /// Custom path commands (for custom paths)
    pub commands: Vec<PathCommand>,
    /// Path editing mode
    pub edit_mode: PathEditMode,
    /// Origin point (relative to shape)
    pub origin_x: f64,
    pub origin_y: f64,
}

impl Default for MotionPath {
    fn default() -> Self {
        Self::new()
    }
}

impl MotionPath {
    /// Create a new empty motion path.
    pub fn new() -> Self {
        Self {
            path_type: MotionPathType::Custom,
            commands: Vec::new(),
            edit_mode: PathEditMode::Relative,
            origin_x: 0.0,
            origin_y: 0.0,
        }
    }

    /// Create a predefined motion path.
    pub fn predefined(path_type: MotionPathType) -> Self {
        Self {
            path_type,
            commands: Vec::new(),
            edit_mode: PathEditMode::Relative,
            origin_x: 0.0,
            origin_y: 0.0,
        }
    }

    /// Create a custom motion path from commands.
    pub fn custom(commands: Vec<PathCommand>) -> Self {
        Self {
            path_type: MotionPathType::Custom,
            commands,
            edit_mode: PathEditMode::Relative,
            origin_x: 0.0,
            origin_y: 0.0,
        }
    }

    /// Add a path command.
    pub fn add_command(&mut self, command: PathCommand) {
        self.commands.push(command);
    }

    /// Set the origin point.
    pub fn set_origin(&mut self, x: f64, y: f64) {
        self.origin_x = x;
        self.origin_y = y;
    }

    /// Check if this is a custom path.
    pub fn is_custom(&self) -> bool {
        self.path_type == MotionPathType::Custom
    }
}

/// Motion path type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MotionPathType {
    /// Custom path
    Custom,
    /// Line paths
    Line,
    /// Arc paths
    Arc,
    /// Turn paths
    Turn,
    /// Shape paths
    Circle,
    /// Diamond shape
    Diamond,
    /// Heart shape
    Heart,
    /// Hexagon shape
    Hexagon,
    /// Octagon shape
    Octagon,
    /// Pentagon shape
    Pentagon,
    /// Square shape
    Square,
    /// Star 4 points
    Star4,
    /// Star 5 points
    Star5,
    /// Star 6 points
    Star6,
    /// Star 8 points
    Star8,
    /// Triangle shape
    Triangle,
    /// Loop de loop
    LoopDeLoop,
    /// Curved X
    CurvedX,
    /// S curve 1
    SCurve1,
    /// S curve 2
    SCurve2,
    /// Sine wave
    SineWave,
    /// Spiral left
    SpiralLeft,
    /// Spiral right
    SpiralRight,
    /// Spring
    Spring,
    /// Zigzag
    Zigzag,
}

/// Path command for custom motion paths.
#[derive(Debug, Clone, PartialEq)]
pub enum PathCommand {
    /// Move to point
    MoveTo { x: f64, y: f64 },
    /// Line to point
    LineTo { x: f64, y: f64 },
    /// Cubic Bezier curve
    CurveTo {
        x1: f64,
        y1: f64,
        x2: f64,
        y2: f64,
        x: f64,
        y: f64,
    },
    /// Quadratic Bezier curve
    QuadTo { x1: f64, y1: f64, x: f64, y: f64 },
    /// Arc
    Arc {
        rx: f64,
        ry: f64,
        rotation: f64,
        large_arc: bool,
        sweep: bool,
        x: f64,
        y: f64,
    },
    /// Close path
    Close,
}

/// Path editing mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum PathEditMode {
    /// Relative coordinates (to shape position)
    #[default]
    Relative,
    /// Absolute coordinates (on slide)
    Absolute,
    /// Fixed (cannot be edited)
    Fixed,
}

/// Helper functions to create common motion paths.
pub struct MotionPathBuilder;

impl MotionPathBuilder {
    /// Create a straight line path.
    pub fn line(dx: f64, dy: f64) -> MotionPath {
        MotionPath::custom(vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo { x: dx, y: dy },
        ])
    }

    /// Create a circular path.
    pub fn circle(radius: f64) -> MotionPath {
        let mut path = MotionPath::predefined(MotionPathType::Circle);
        path.commands = vec![
            PathCommand::MoveTo { x: radius, y: 0.0 },
            PathCommand::Arc {
                rx: radius,
                ry: radius,
                rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: -radius,
                y: 0.0,
            },
            PathCommand::Arc {
                rx: radius,
                ry: radius,
                rotation: 0.0,
                large_arc: false,
                sweep: true,
                x: radius,
                y: 0.0,
            },
        ];
        path
    }

    /// Create a zigzag path.
    pub fn zigzag(width: f64, height: f64, segments: usize) -> MotionPath {
        let mut commands = vec![PathCommand::MoveTo { x: 0.0, y: 0.0 }];
        let segment_width = width / (segments as f64);

        for i in 1..=segments {
            let x = segment_width * (i as f64);
            let y = if i % 2 == 0 { 0.0 } else { height };
            commands.push(PathCommand::LineTo { x, y });
        }

        MotionPath::custom(commands)
    }

    /// Create an S-curve path.
    pub fn s_curve(width: f64, height: f64) -> MotionPath {
        MotionPath::custom(vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::CurveTo {
                x1: width * 0.25,
                y1: 0.0,
                x2: width * 0.25,
                y2: height * 0.5,
                x: width * 0.5,
                y: height * 0.5,
            },
            PathCommand::CurveTo {
                x1: width * 0.75,
                y1: height * 0.5,
                x2: width * 0.75,
                y2: height,
                x: width,
                y: height,
            },
        ])
    }

    /// Create a spiral path.
    pub fn spiral(radius: f64, turns: f64, clockwise: bool) -> MotionPath {
        let mut commands = vec![PathCommand::MoveTo { x: 0.0, y: 0.0 }];
        let steps = (turns * 36.0) as usize; // 36 steps per turn
        let angle_step = std::f64::consts::PI * 2.0 / 36.0;
        let radius_step = radius / (steps as f64);

        for i in 1..=steps {
            let angle = if clockwise {
                angle_step * (i as f64)
            } else {
                -angle_step * (i as f64)
            };
            let r = radius_step * (i as f64);
            let x = r * angle.cos();
            let y = r * angle.sin();
            commands.push(PathCommand::LineTo { x, y });
        }

        MotionPath::custom(commands)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_motion_path_default() {
        let path = MotionPath::default();
        assert_eq!(path.path_type, MotionPathType::Custom);
        assert!(path.commands.is_empty());
        assert!(path.is_custom());
    }

    #[test]
    fn test_motion_path_predefined() {
        let path = MotionPath::predefined(MotionPathType::Circle);
        assert_eq!(path.path_type, MotionPathType::Circle);
        assert!(!path.is_custom());
    }

    #[test]
    fn test_motion_path_custom() {
        let commands = vec![
            PathCommand::MoveTo { x: 0.0, y: 0.0 },
            PathCommand::LineTo { x: 100.0, y: 100.0 },
        ];
        let path = MotionPath::custom(commands.clone());
        assert_eq!(path.commands.len(), 2);
        assert!(path.is_custom());
    }

    #[test]
    fn test_motion_path_add_command() {
        let mut path = MotionPath::new();
        path.add_command(PathCommand::MoveTo { x: 0.0, y: 0.0 });
        path.add_command(PathCommand::LineTo { x: 50.0, y: 50.0 });
        assert_eq!(path.commands.len(), 2);
    }

    #[test]
    fn test_motion_path_set_origin() {
        let mut path = MotionPath::new();
        path.set_origin(10.0, 20.0);
        assert_eq!(path.origin_x, 10.0);
        assert_eq!(path.origin_y, 20.0);
    }

    #[test]
    fn test_motion_path_builder_line() {
        let path = MotionPathBuilder::line(100.0, 50.0);
        assert_eq!(path.commands.len(), 2);
        assert!(matches!(path.commands[0], PathCommand::MoveTo { .. }));
        assert!(matches!(path.commands[1], PathCommand::LineTo { .. }));
    }

    #[test]
    fn test_motion_path_builder_circle() {
        let path = MotionPathBuilder::circle(50.0);
        assert!(!path.commands.is_empty());
    }

    #[test]
    fn test_motion_path_builder_zigzag() {
        let path = MotionPathBuilder::zigzag(100.0, 50.0, 5);
        assert_eq!(path.commands.len(), 6); // Move + 5 lines
    }

    #[test]
    fn test_motion_path_builder_s_curve() {
        let path = MotionPathBuilder::s_curve(100.0, 100.0);
        assert_eq!(path.commands.len(), 3); // Move + 2 curves
    }

    #[test]
    fn test_motion_path_builder_spiral() {
        let path = MotionPathBuilder::spiral(100.0, 2.0, true);
        assert!(!path.commands.is_empty());
    }

    #[test]
    fn test_path_edit_mode_default() {
        assert_eq!(PathEditMode::default(), PathEditMode::Relative);
    }
}
