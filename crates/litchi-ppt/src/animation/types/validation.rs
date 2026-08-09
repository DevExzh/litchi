//! Validation rules for typed animation properties and formulas.

use super::time::{TimeAnimateValueType, TimeNodeFill};

pub(crate) fn is_valid_runtime_context(value: &str) -> bool {
    fn valid_version(value: &str) -> bool {
        let mut components = value.split('.');
        let first = components.next().unwrap_or_default();
        !first.is_empty()
            && first.bytes().all(|byte| byte.is_ascii_digit())
            && components.next().is_none_or(|second| {
                !second.is_empty()
                    && second.bytes().all(|byte| byte.is_ascii_digit())
                    && components.next().is_none()
            })
    }

    fn valid_atom(atom: &str) -> bool {
        fn valid_relation(value: &str) -> bool {
            value == "!"
                || ["gte", "gt", "lte", "lt"]
                    .iter()
                    .any(|relation| value.eq_ignore_ascii_case(relation))
        }

        if atom.is_empty()
            || atom.starts_with(' ')
            || atom.ends_with(' ')
            || atom
                .chars()
                .any(|character| character.is_whitespace() && character != ' ')
        {
            return false;
        }
        let fields = atom.split_ascii_whitespace().collect::<Vec<_>>();
        match fields.as_slice() {
            [app] => app.eq_ignore_ascii_case("ppt"),
            [first, second] => {
                (first.eq_ignore_ascii_case("ppt") && valid_version(second))
                    || (valid_relation(first) && second.eq_ignore_ascii_case("ppt"))
            },
            [relation, app, version] => {
                valid_relation(relation)
                    && app.eq_ignore_ascii_case("ppt")
                    && valid_version(version)
            },
            _ => false,
        }
    }

    let sequence = value.strip_suffix(';').unwrap_or(value);
    !sequence.is_empty() && sequence.split(';').all(valid_atom)
}

pub(crate) fn is_valid_time_points_types(value: &str) -> bool {
    value
        .bytes()
        .all(|byte| matches!(byte, b'A' | b'a' | b'F' | b'f' | b'T' | b't' | b'S' | b's'))
}

pub(crate) fn is_valid_time_filter(value: &str) -> bool {
    fn normalized_time(value: &str) -> bool {
        value == "1.0"
            || value.strip_prefix("0.").is_some_and(|digits| {
                !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit())
            })
    }

    !value.is_empty()
        && value.split(';').all(|entry| {
            let mut fields = entry.split(',');
            matches!(
                (fields.next(), fields.next(), fields.next()),
                (Some(time), Some(transformed), None)
                    if normalized_time(time) && normalized_time(transformed)
            )
        })
}

pub(crate) fn is_valid_motion_path(value: &str) -> bool {
    fn spaces(bytes: &[u8], position: &mut usize) -> bool {
        let start = *position;
        while bytes.get(*position) == Some(&b' ') {
            *position += 1;
        }
        *position != start
    }

    fn coordinate(bytes: &[u8], position: &mut usize) -> bool {
        if bytes.get(*position) == Some(&b'(') {
            let start = *position + 1;
            let mut depth = 1usize;
            let mut end = start;
            while end < bytes.len() && depth != 0 {
                match bytes[end] {
                    b'(' => depth += 1,
                    b')' => depth -= 1,
                    _ => {},
                }
                end += 1;
            }
            if depth != 0 || !is_valid_time_formula_bytes(&bytes[start..end - 1]) {
                return false;
            }
            *position = end;
            true
        } else {
            let start = *position;
            while bytes.get(*position).is_some_and(u8::is_ascii_digit) {
                *position += 1;
            }
            if *position == start {
                return false;
            }
            if bytes.get(*position) == Some(&b'.') {
                *position += 1;
                let fraction = *position;
                while bytes.get(*position).is_some_and(u8::is_ascii_digit) {
                    *position += 1;
                }
                if *position == fraction {
                    return false;
                }
            }
            true
        }
    }

    let bytes = value.as_bytes();
    let mut position = 0usize;
    let mut commands = 0usize;
    spaces(bytes, &mut position);
    while let Some(command) = bytes.get(position).copied() {
        position += 1;
        commands += 1;
        let coordinates = match command {
            b'm' | b'M' | b'l' | b'L' => 2,
            b'c' | b'C' => 6,
            b'z' | b'Z' => 0,
            b'e' | b'E' => return true,
            _ => return false,
        };
        for _ in 0..coordinates {
            if !spaces(bytes, &mut position) || !coordinate(bytes, &mut position) {
                return false;
            }
        }
        spaces(bytes, &mut position);
    }
    commands != 0
}

fn is_valid_time_formula_bytes(value: &[u8]) -> bool {
    struct Parser<'a> {
        value: &'a [u8],
        position: usize,
    }

    impl Parser<'_> {
        fn expression(&mut self) -> bool {
            if !self.term() {
                return false;
            }
            while matches!(self.peek(), Some(b'+' | b'-')) {
                self.position += 1;
                if !self.term() {
                    return false;
                }
            }
            true
        }

        fn term(&mut self) -> bool {
            if !self.power() {
                return false;
            }
            while matches!(self.peek(), Some(b'*' | b'/' | b'%')) {
                self.position += 1;
                if !self.power() {
                    return false;
                }
            }
            true
        }

        fn power(&mut self) -> bool {
            if !self.unary() {
                return false;
            }
            while self.peek() == Some(b'^') {
                self.position += 1;
                if !self.unary() {
                    return false;
                }
            }
            true
        }

        fn unary(&mut self) -> bool {
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.position += 1;
            }
            self.factor()
        }

        fn factor(&mut self) -> bool {
            if self.peek() == Some(b'$') {
                self.position += 1;
                return true;
            }
            if self.peek() == Some(b'(') {
                self.position += 1;
                return self.expression() && self.take(b')');
            }
            if self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                return self.number();
            }
            let prefixed = self.peek() == Some(b'#');
            if prefixed {
                self.position += 1;
            }
            let start = self.position;
            while self
                .peek()
                .is_some_and(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'.' | b'_'))
            {
                self.position += 1;
            }
            if self.position == start {
                return false;
            }
            let Ok(name) = std::str::from_utf8(&self.value[start..self.position]) else {
                return false;
            };
            if !prefixed && self.peek() == Some(b'(') && is_time_formula_function(name) {
                self.position += 1;
                if !self.expression() {
                    return false;
                }
                if self.peek() == Some(b',') {
                    self.position += 1;
                    if !self.expression() {
                        return false;
                    }
                }
                self.take(b')')
            } else {
                matches!(name, "pi" | "e") && !prefixed || is_time_formula_attribute(name)
            }
        }

        fn number(&mut self) -> bool {
            let start = self.position;
            while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                self.position += 1;
            }
            if self.position == start {
                return false;
            }
            if self.peek() == Some(b'.') {
                self.position += 1;
                let fraction = self.position;
                while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                    self.position += 1;
                }
                if self.position == fraction {
                    return false;
                }
            }
            if matches!(self.peek(), Some(b'e' | b'E')) {
                self.position += 1;
                if self.peek() == Some(b'-') {
                    self.position += 1;
                }
                let exponent = self.position;
                while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                    self.position += 1;
                }
                if self.position == exponent {
                    return false;
                }
            }
            true
        }

        fn peek(&self) -> Option<u8> {
            self.value.get(self.position).copied()
        }

        fn take(&mut self, expected: u8) -> bool {
            if self.peek() == Some(expected) {
                self.position += 1;
                true
            } else {
                false
            }
        }
    }

    let mut parser = Parser { value, position: 0 };
    !value.is_empty() && parser.expression() && parser.position == value.len()
}

fn is_time_formula_function(value: &str) -> bool {
    matches!(
        value,
        "abs"
            | "acos"
            | "asin"
            | "atan"
            | "ceil"
            | "cos"
            | "cosh"
            | "deg"
            | "exp"
            | "floor"
            | "ln"
            | "max"
            | "min"
            | "rad"
            | "rand"
            | "sin"
            | "sinh"
            | "sqrt"
            | "tan"
            | "tanh"
    )
}

fn is_time_formula_attribute(value: &str) -> bool {
    matches!(
        value,
        "ppt_x"
            | "ppt_y"
            | "ppt_w"
            | "ppt_h"
            | "ScaleX"
            | "ScaleY"
            | "stype.rotation"
            | "style.opacity"
            | "style.visibility"
            | "ppt_r"
            | "r"
            | "style.fontSize"
            | "style.fontWeight"
            | "style.fontStyle"
            | "style.fontFamily"
            | "style.textEffectEmboss"
            | "style.textShadow"
            | "style.textTransform"
            | "style.textDecorationUnderline"
            | "style.textEffectOutline"
            | "style.textDecorationLineThrough"
            | "style.sRotation"
            | "imageData.cropTop"
            | "imageData.cropBottom"
            | "imageData.cropLeft"
            | "imageData.cropRight"
            | "imageData.gain"
            | "imageData.blackleve"
            | "imageData.gamma"
            | "imageData.grayscale"
            | "fill.on"
            | "fill.type"
            | "fill.opacity"
            | "fill.method"
            | "fill.opacity2"
            | "fill.angle"
            | "fill.focus"
            | "fill.focusposition.x"
            | "fill.focusposition.y"
            | "fill.focussize.x"
            | "fill.focussize.y"
            | "stroke.on"
            | "stroke.weight"
            | "stroke.opacity"
            | "stroke.linestyle"
            | "stroke.dashstyle"
            | "stroke.filltype"
            | "stroke.imagesize.x"
            | "stroke.imagesize.y"
            | "stroke.startArrow"
            | "stroke.endArrow"
            | "stroke.startArrowWidth"
            | "stroke.startArrowLength"
            | "stroke.endArrowWidth"
            | "stroke.endArrowLength"
            | "shadow.on"
            | "shadow.type"
            | "shadow.opacity"
            | "shadow.offset.x"
            | "shadow.offset.y"
            | "shadow.offset2.x"
            | "shadow.offset2.y"
            | "shadow.origin.x"
            | "shadow.origin.y"
            | "shadow.matrix.xtox"
            | "shadow.matrix.ytox"
            | "shadow.matrix.ytoy"
            | "shadow.matrix.perspectiveX"
            | "shadow.matrix.perspectiveY"
            | "skew.on"
            | "skew.offset.x"
            | "skew.offset.y"
            | "skew.origin.x"
            | "skew.origin.y"
            | "skew.matrix.xtox"
            | "skew.matrix.ytox"
            | "skew.matrix.ytoy"
            | "skew.matrix.perspectiveX"
            | "skew.matrix.perspectiveY"
            | "extrusion.on"
            | "extrusion.type"
            | "extrusion.render"
            | "extrusion.viewpointorigin.x"
            | "extrusion.viewpointorigin.y"
            | "extrusion.viewpoint.x"
            | "extrusion.viewpoint.y"
            | "extrusion.viewpoint.z"
            | "extrusion.plane"
            | "extrusion.skewangle"
            | "extrusion.skewamt"
            | "extrusion.backdepth"
            | "extrusion.foredepth"
            | "extrusion.orientation.x"
            | "extrusion.orientation.y"
            | "extrusion.orientation.z"
            | "extrusion.orientationangle"
            | "extrusion.rotationangle.x"
            | "extrusion.rotationangle.y"
            | "extrusion.lockrotationcenter"
            | "extrusion.autorotationcenter"
            | "extrusion.rotationcenter.x"
            | "extrusion.rotationcenter.y"
            | "extrusion.rotationcenter.z"
            | "extrusion.colormode"
    )
}

pub(crate) fn is_valid_animation_attribute_name(value: &str) -> bool {
    (value != "stype.rotation"
        && value != "imageData.blackleve"
        && is_time_formula_attribute(value))
        || matches!(
            value,
            "ppt_c"
                | "xshear"
                | "yshear"
                | "image"
                | "fillcolor"
                | "style.rotation"
                | "style.color"
                | "imageData.blacklevel"
                | "imageData.chromakey"
                | "fill.color"
                | "fill.color2"
                | "stroke.color"
                | "stroke.src"
                | "stroke.color2"
                | "shadow.color"
                | "shadow.color2"
                | "extrusion.color"
        )
}

pub(crate) fn time_animation_attribute_value_type(attribute: &str) -> Option<TimeAnimateValueType> {
    time_set_attribute_value_type(attribute).or_else(|| {
        is_valid_animation_attribute_name(attribute).then_some(TimeAnimateValueType::String)
    })
}

pub(crate) fn is_valid_time_animate_value(
    attribute: &str,
    value_type: TimeAnimateValueType,
    value: &str,
) -> bool {
    match value_type {
        TimeAnimateValueType::String => true,
        TimeAnimateValueType::Number | TimeAnimateValueType::Color => {
            is_valid_time_set_value(attribute, value)
        },
    }
}

pub(crate) fn is_valid_time_formula(value: &str) -> bool {
    is_valid_time_formula_bytes(value.as_bytes())
}

pub(crate) fn time_set_attribute_value_type(attribute: &str) -> Option<TimeAnimateValueType> {
    if is_time_set_preset_attribute(attribute) || is_time_set_numeric_attribute(attribute) {
        Some(TimeAnimateValueType::Number)
    } else if matches!(
        attribute,
        "ppt_c"
            | "fillcolor"
            | "style.color"
            | "imageData.chromakey"
            | "fill.color"
            | "fill.color2"
            | "stroke.color"
            | "stroke.color2"
            | "shadow.color"
            | "shadow.color2"
            | "extrusion.color"
    ) {
        Some(TimeAnimateValueType::Color)
    } else {
        None
    }
}

pub(crate) fn is_valid_time_set_value(attribute: &str, value: &str) -> bool {
    if is_time_set_numeric_attribute(attribute) {
        return is_valid_formula_or_number(value);
    }
    if time_set_attribute_value_type(attribute) == Some(TimeAnimateValueType::Color) {
        return value.len() == 7
            && value.starts_with('#')
            && value[1..].bytes().all(|byte| byte.is_ascii_hexdigit());
    }
    match attribute {
        "style.visibility" => matches!(value, "hidden" | "visible"),
        "style.fontWeight" => matches!(value, "none" | "normal" | "bold"),
        "style.fontStyle" => matches!(value, "none" | "normal" | "italic"),
        "style.textEffectEmboss" => matches!(value, "none" | "normal" | "emboss"),
        "style.textShadow" => matches!(value, "none" | "normal" | "auto"),
        "style.textTransform" => matches!(value, "none" | "normal" | "sub" | "super"),
        "style.textDecorationUnderline"
        | "style.textEffectOutline"
        | "style.textDecorationLineThrough"
        | "imageData.grayscale"
        | "extrusion.lockrotationcenter"
        | "extrusion.autorotationcenter"
        | "extrusion.colormode" => matches!(value, "false" | "true"),
        "fill.on" | "stroke.on" | "shadow.on" | "skew.on" | "extrusion.on" => {
            matches!(value, "false" | "f" | "t" | "true")
        },
        "fill.type" => matches!(
            value,
            "solid"
                | "pattern"
                | "tile"
                | "frame"
                | "gradientUnscaled"
                | "gradient"
                | "gradientCenter"
                | "gradientRadial"
                | "gradientTile"
                | "background"
        ),
        "fill.method" => matches!(value, "none" | "linear" | "sigma" | "any"),
        "stroke.linestyle" => matches!(
            value,
            "single" | "thinThin" | "thinThick" | "thickThin" | "thickBetweenThin"
        ),
        "stroke.dashstyle" => matches!(
            value,
            "solid" | "dot" | "dash" | "dashDot" | "longDash" | "longDashDot" | "longDashDotDot"
        ),
        "stroke.filltype" => matches!(value, "solid" | "tile" | "pattern" | "frame"),
        "stroke.startArrow" | "stroke.endArrow" => matches!(
            value,
            "none"
                | "block"
                | "classic"
                | "diamond"
                | "oval"
                | "open"
                | "chevron"
                | "doublechevron"
        ),
        "stroke.startArrowWidth" | "stroke.endArrowWidth" => {
            matches!(value, "narrow" | "medium" | "wide")
        },
        "stroke.startArrowLength" | "stroke.endArrowLength" => {
            matches!(value, "short" | "medium" | "long")
        },
        "shadow.type" => matches!(value, "single" | "double" | "emboss" | "perspective"),
        "extrusion.type" => matches!(value, "parallel" | "perspective"),
        "extrusion.render" => matches!(value, "solid" | "wireframe" | "boundingcube"),
        "extrusion.plane" => matches!(value, "xy" | "zx" | "yz"),
        _ => false,
    }
}

fn is_time_set_preset_attribute(attribute: &str) -> bool {
    matches!(
        attribute,
        "style.visibility"
            | "style.fontWeight"
            | "style.fontStyle"
            | "style.textEffectEmboss"
            | "style.textShadow"
            | "style.textTransform"
            | "style.textDecorationUnderline"
            | "style.textEffectOutline"
            | "style.textDecorationLineThrough"
            | "imageData.grayscale"
            | "fill.on"
            | "fill.type"
            | "fill.method"
            | "stroke.on"
            | "stroke.linestyle"
            | "stroke.dashstyle"
            | "stroke.filltype"
            | "stroke.startArrow"
            | "stroke.endArrow"
            | "stroke.startArrowWidth"
            | "stroke.startArrowLength"
            | "stroke.endArrowWidth"
            | "stroke.endArrowLength"
            | "shadow.on"
            | "shadow.type"
            | "skew.on"
            | "extrusion.on"
            | "extrusion.type"
            | "extrusion.render"
            | "extrusion.plane"
            | "extrusion.lockrotationcenter"
            | "extrusion.autorotationcenter"
            | "extrusion.colormode"
    )
}

fn is_time_set_numeric_attribute(attribute: &str) -> bool {
    matches!(
        attribute,
        "ppt_x"
            | "ppt_y"
            | "ppt_w"
            | "ppt_h"
            | "ppt_r"
            | "xshear"
            | "yshear"
            | "ScaleX"
            | "ScaleY"
            | "r"
            | "style.opacity"
            | "style.rotation"
            | "style.fontSize"
            | "style.sRotation"
            | "imageData.cropTop"
            | "imageData.cropBottom"
            | "imageData.cropLeft"
            | "imageData.cropRight"
            | "imageData.gain"
            | "imageData.blacklevel"
            | "imageData.gamma"
            | "fill.opacity"
            | "fill.opacity2"
            | "fill.angle"
            | "fill.focus"
            | "fill.focusposition.x"
            | "fill.focusposition.y"
            | "fill.focussize.x"
            | "fill.focussize.y"
            | "stroke.weight"
            | "stroke.opacity"
            | "stroke.imagesize.x"
            | "stroke.imagesize.y"
            | "shadow.opacity"
            | "shadow.offset.x"
            | "shadow.offset.y"
            | "shadow.offset2.x"
            | "shadow.offset2.y"
            | "shadow.origin.x"
            | "shadow.origin.y"
            | "shadow.matrix.xtox"
            | "shadow.matrix.ytox"
            | "shadow.matrix.ytoy"
            | "shadow.matrix.perspectiveX"
            | "shadow.matrix.perspectiveY"
            | "skew.offset.x"
            | "skew.offset.y"
            | "skew.origin.x"
            | "skew.origin.y"
            | "skew.matrix.xtox"
            | "skew.matrix.ytox"
            | "skew.matrix.ytoy"
            | "skew.matrix.perspectiveX"
            | "skew.matrix.perspectiveY"
            | "extrusion.viewpointorigin.x"
            | "extrusion.viewpointorigin.y"
            | "extrusion.viewpoint.x"
            | "extrusion.viewpoint.y"
            | "extrusion.viewpoint.z"
            | "extrusion.skewangle"
            | "extrusion.skewamt"
            | "extrusion.backdepth"
            | "extrusion.foredepth"
            | "extrusion.orientation.x"
            | "extrusion.orientation.y"
            | "extrusion.orientation.z"
            | "extrusion.orientationangle"
            | "extrusion.rotationangle.x"
            | "extrusion.rotationangle.y"
            | "extrusion.rotationcenter.x"
            | "extrusion.rotationcenter.y"
            | "extrusion.rotationcenter.z"
    )
}

impl TimeNodeFill {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::HoldUntilParentEnds),
            1 => Some(Self::ResetWhenInactive),
            2 => Some(Self::HoldUntilNext),
            3 => Some(Self::HoldUntilParentEndsLegacy),
            4 => Some(Self::ResetWhenInactiveLegacy),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::HoldUntilParentEnds => 0,
            Self::ResetWhenInactive => 1,
            Self::HoldUntilNext => 2,
            Self::HoldUntilParentEndsLegacy => 3,
            Self::ResetWhenInactiveLegacy => 4,
        }
    }
}

fn is_valid_formula_or_number(value: &str) -> bool {
    if let Some(formula) = value
        .strip_prefix('(')
        .and_then(|rest| rest.strip_suffix(')'))
    {
        return is_valid_time_formula_bytes(formula.as_bytes());
    }
    let bytes = value.as_bytes();
    let mut position = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(position) == Some(&b'.') {
        position += 1;
        let start = position;
        while bytes.get(position).is_some_and(u8::is_ascii_digit) {
            position += 1;
        }
        if position == start {
            return false;
        }
    } else {
        let start = position;
        while bytes.get(position).is_some_and(u8::is_ascii_digit) {
            position += 1;
        }
        if position == start {
            return false;
        }
        if bytes.get(position) == Some(&b'.') {
            position += 1;
            while bytes.get(position).is_some_and(u8::is_ascii_digit) {
                position += 1;
            }
        }
    }
    if position < bytes.len() {
        if bytes.get(position) == Some(&b'-') {
            position += 1;
        }
        if !matches!(bytes.get(position), Some(b'e' | b'E')) {
            return false;
        }
        position += 1;
        let start = position;
        while bytes.get(position).is_some_and(u8::is_ascii_digit) {
            position += 1;
        }
        if position == start {
            return false;
        }
    }
    position == bytes.len()
}
