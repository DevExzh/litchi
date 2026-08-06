//! PresentationML wire conversions for the semantic animation model.

use super::super::codec::{
    MAX_NORMALIZED_TIME_DECIMALS, MAX_TIME_FILTER_BYTES, MAX_TIME_FILTER_POINTS,
};
use super::super::invalid;
use super::{
    ConditionEvent, DiagramBuildType, Duration, Effect, EventFilter, Fill, GraphicChartBuildType,
    GraphicDiagramBuildType, NormalizedTime, OleChartBuildType, ParagraphBuildType, PresetClass,
    Repeat, Restart, RuntimeTrigger, SyncBehavior, TimeFilter, TimeNodeType, TimePoint,
};
use crate::Result;

impl Effect {
    pub(in crate::animations) fn from_preset_parts(class: &str, id: u32) -> Self {
        match (class, id) {
            ("entr", 1) => Self::Appear,
            ("entr", 2) => Self::FlyIn,
            ("entr", 10) => Self::Fade,
            ("entr", 16) => Self::Split,
            ("entr", 22) => Self::Wipe,
            ("entr", 23) => Self::Zoom,
            ("entr", 24) => Self::Bounce,
            ("entr", 42) => Self::FloatIn,
            ("emph", 6) => Self::GrowShrink,
            ("emph", 8) => Self::Spin,
            _ => Self::Custom(format!("{class}:{id}")),
        }
    }
}

impl ParagraphBuildType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "p" => Ok(Self::Paragraph),
            "cust" => Ok(Self::Custom),
            "whole" => Ok(Self::Whole),
            _ => Err(invalid("invalid paragraph build type")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Paragraph => "p",
            Self::Custom => "cust",
            Self::Whole => "whole",
        }
    }
}

impl DiagramBuildType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "whole" => Ok(Self::Whole),
            "depthByNode" => Ok(Self::DepthByNode),
            "depthByBranch" => Ok(Self::DepthByBranch),
            "breadthByNode" => Ok(Self::BreadthByNode),
            "breadthByLvl" => Ok(Self::BreadthByLevel),
            "cw" => Ok(Self::Clockwise),
            "cwIn" => Ok(Self::ClockwiseIn),
            "cwOut" => Ok(Self::ClockwiseOut),
            "ccw" => Ok(Self::CounterClockwise),
            "ccwIn" => Ok(Self::CounterClockwiseIn),
            "ccwOut" => Ok(Self::CounterClockwiseOut),
            "inByRing" => Ok(Self::InByRing),
            "outByRing" => Ok(Self::OutByRing),
            "up" => Ok(Self::Up),
            "down" => Ok(Self::Down),
            "allAtOnce" => Ok(Self::AllAtOnce),
            "cust" => Ok(Self::Custom),
            _ => Err(invalid("invalid diagram build type")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::Whole => "whole",
            Self::DepthByNode => "depthByNode",
            Self::DepthByBranch => "depthByBranch",
            Self::BreadthByNode => "breadthByNode",
            Self::BreadthByLevel => "breadthByLvl",
            Self::Clockwise => "cw",
            Self::ClockwiseIn => "cwIn",
            Self::ClockwiseOut => "cwOut",
            Self::CounterClockwise => "ccw",
            Self::CounterClockwiseIn => "ccwIn",
            Self::CounterClockwiseOut => "ccwOut",
            Self::InByRing => "inByRing",
            Self::OutByRing => "outByRing",
            Self::Up => "up",
            Self::Down => "down",
            Self::AllAtOnce => "allAtOnce",
            Self::Custom => "cust",
        }
    }
}

impl GraphicDiagramBuildType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "one" => Ok(Self::One),
            "lvlOne" => Ok(Self::LevelOne),
            "lvlAtOnce" => Ok(Self::LevelAtOnce),
            _ => Err(invalid("invalid graphical-object diagram build type")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::One => "one",
            Self::LevelOne => "lvlOne",
            Self::LevelAtOnce => "lvlAtOnce",
        }
    }
}

impl GraphicChartBuildType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid graphical-object chart build type")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Series => "series",
            Self::Category => "category",
            Self::SeriesElement => "seriesEl",
            Self::CategoryElement => "categoryEl",
        }
    }
}

impl OleChartBuildType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "allAtOnce" => Ok(Self::AllAtOnce),
            "series" => Ok(Self::Series),
            "category" => Ok(Self::Category),
            "seriesEl" => Ok(Self::SeriesElement),
            "categoryEl" => Ok(Self::CategoryElement),
            _ => Err(invalid("invalid OLE chart build type")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::AllAtOnce => "allAtOnce",
            Self::Series => "series",
            Self::Category => "category",
            Self::SeriesElement => "seriesEl",
            Self::CategoryElement => "categoryEl",
        }
    }
}

impl EventFilter {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "cancelBubble" => Ok(Self::CancelBubble),
            _ => Err(invalid("invalid PowerPoint animation event filter")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::CancelBubble => "cancelBubble",
        }
    }
}

impl Fill {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "remove" => Ok(Self::Remove),
            "freeze" => Ok(Self::Freeze),
            "hold" => Ok(Self::Hold),
            "transition" => Ok(Self::Transition),
            _ => Err(invalid("invalid animation fill behavior")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::Remove => "remove",
            Self::Freeze => "freeze",
            Self::Hold => "hold",
            Self::Transition => "transition",
        }
    }
}

impl Restart {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "always" => Ok(Self::Always),
            "whenNotActive" => Ok(Self::WhenNotActive),
            "never" => Ok(Self::Never),
            _ => Err(invalid("invalid animation restart behavior")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::Always => "always",
            Self::WhenNotActive => "whenNotActive",
            Self::Never => "never",
        }
    }
}

impl NormalizedTime {
    fn parse(value: &str) -> Result<Self> {
        if value.is_empty() {
            return Err(invalid("normalized time is empty"));
        }
        let (whole, fraction) = match value.split_once('.') {
            Some((whole, fraction)) => {
                if fraction.is_empty() || fraction.contains('.') {
                    return Err(invalid("invalid normalized time decimal"));
                }
                (whole, Some(fraction))
            },
            None => (value, None),
        };
        if !matches!(whole, "0" | "1") {
            return Err(invalid("normalized time must be between 0 and 1"));
        }
        let Some(fraction) = fraction else {
            return Ok(if whole == "1" {
                Self {
                    numerator: 1,
                    scale: 1,
                }
            } else {
                Self {
                    numerator: 0,
                    scale: 1,
                }
            });
        };
        if fraction.len() > MAX_NORMALIZED_TIME_DECIMALS
            || !fraction.bytes().all(|byte| byte.is_ascii_digit())
        {
            return Err(invalid("invalid or over-precise normalized time"));
        }
        if whole == "1" && fraction.bytes().any(|byte| byte != b'0') {
            return Err(invalid("normalized time exceeds 1.0"));
        }
        if whole == "1" {
            return Ok(Self {
                numerator: 1,
                scale: 1,
            });
        }
        let numerator = fraction
            .parse::<u64>()
            .map_err(|_| invalid("normalized time decimal overflows"))?;
        let scale = 10u64
            .checked_pow(
                u32::try_from(fraction.len())
                    .map_err(|_| invalid("normalized time precision overflows"))?,
            )
            .ok_or_else(|| invalid("normalized time scale overflows"))?;
        Ok(Self::normalized(numerator, scale))
    }

    fn write_value(self) -> String {
        if self.numerator == 0 {
            return "0".to_string();
        }
        if self.numerator == self.scale {
            return "1".to_string();
        }
        let decimals = self.scale.ilog10() as usize;
        format!("0.{:0decimals$}", self.numerator)
    }
}

impl TimeFilter {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        if value.len() > MAX_TIME_FILTER_BYTES {
            return Err(invalid("animation time filter exceeds safety limit"));
        }
        let mut points = Vec::new();
        for pair in value.split(';') {
            if points.len() >= MAX_TIME_FILTER_POINTS {
                return Err(invalid(
                    "animation time filter point count exceeds safety limit",
                ));
            }
            let pair = pair.trim();
            let (local, warped) = pair
                .split_once(',')
                .ok_or_else(|| invalid("animation time filter point is missing a comma"))?;
            if warped.contains(',') {
                return Err(invalid("animation time filter point has too many values"));
            }
            points.push(TimePoint::new(
                NormalizedTime::parse(local.trim())?,
                NormalizedTime::parse(warped.trim())?,
            ));
        }
        Self::new(points)
    }

    pub(in crate::animations) fn write_value(&self) -> String {
        let mut output = String::new();
        for (index, point) in self.points.iter().enumerate() {
            if index != 0 {
                output.push(';');
            }
            output.push_str(&point.local_time.write_value());
            output.push(',');
            output.push_str(&point.warped_time.write_value());
        }
        output
    }
}

impl SyncBehavior {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "canSlip" => Ok(Self::CanSlip),
            "locked" => Ok(Self::Locked),
            "none" => Ok(Self::None),
            _ => Err(invalid("invalid animation synchronization behavior")),
        }
    }

    pub(in crate::animations) const fn as_str(self) -> &'static str {
        match self {
            Self::CanSlip => "canSlip",
            Self::Locked => "locked",
            Self::None => "none",
        }
    }
}

impl Repeat {
    pub(in crate::animations) fn write_value(self) -> String {
        match self {
            Self::Finite(value) => value.to_string(),
            Self::Indefinite => "indefinite".to_string(),
        }
    }
}

impl Duration {
    pub(in crate::animations) fn write_value(self) -> String {
        match self {
            Self::Finite(value) => value.to_string(),
            Self::Indefinite => "indefinite".to_string(),
        }
    }
}

impl ConditionEvent {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "onBegin" => Ok(Self::OnBegin),
            "onEnd" => Ok(Self::OnEnd),
            "begin" => Ok(Self::Begin),
            "end" => Ok(Self::End),
            "onClick" => Ok(Self::OnClick),
            "onDblClick" => Ok(Self::OnDoubleClick),
            "onMouseOver" => Ok(Self::OnMouseOver),
            "onMouseOut" => Ok(Self::OnMouseOut),
            "onNext" => Ok(Self::OnNext),
            "onPrev" => Ok(Self::OnPrevious),
            "onStopAudio" => Ok(Self::OnStopAudio),
            _ => Err(invalid("invalid animation condition event")),
        }
    }
    pub(in crate::animations) fn as_str(self) -> &'static str {
        match self {
            Self::OnBegin => "onBegin",
            Self::OnEnd => "onEnd",
            Self::Begin => "begin",
            Self::End => "end",
            Self::OnClick => "onClick",
            Self::OnDoubleClick => "onDblClick",
            Self::OnMouseOver => "onMouseOver",
            Self::OnMouseOut => "onMouseOut",
            Self::OnNext => "onNext",
            Self::OnPrevious => "onPrev",
            Self::OnStopAudio => "onStopAudio",
        }
    }
}

impl RuntimeTrigger {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "first" => Ok(Self::First),
            "last" => Ok(Self::Last),
            "all" => Ok(Self::All),
            _ => Err(invalid("invalid animation runtime trigger")),
        }
    }
    pub(in crate::animations) fn as_str(self) -> &'static str {
        match self {
            Self::First => "first",
            Self::Last => "last",
            Self::All => "all",
        }
    }
}

impl PresetClass {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "entr" => Ok(Self::Entrance),
            "exit" => Ok(Self::Exit),
            "emph" => Ok(Self::Emphasis),
            "path" => Ok(Self::MotionPath),
            "verb" => Ok(Self::Verb),
            "mediacall" => Ok(Self::MediaCall),
            _ => Err(invalid("invalid animation preset class")),
        }
    }
    pub(in crate::animations) fn as_str(self) -> &'static str {
        match self {
            Self::Entrance => "entr",
            Self::Exit => "exit",
            Self::Emphasis => "emph",
            Self::MotionPath => "path",
            Self::Verb => "verb",
            Self::MediaCall => "mediacall",
        }
    }
}

impl TimeNodeType {
    pub(in crate::animations) fn parse(value: &str) -> Result<Self> {
        match value {
            "clickEffect" => Ok(Self::ClickEffect),
            "withEffect" => Ok(Self::WithEffect),
            "afterEffect" => Ok(Self::AfterEffect),
            "mainSeq" => Ok(Self::MainSequence),
            "interactiveSeq" => Ok(Self::InteractiveSequence),
            "clickPar" => Ok(Self::ClickParallel),
            "withGroup" => Ok(Self::WithGroup),
            "afterGroup" => Ok(Self::AfterGroup),
            "tmRoot" => Ok(Self::TimingRoot),
            _ => Err(invalid("invalid animation time-node type")),
        }
    }
    pub(in crate::animations) fn as_str(self) -> &'static str {
        match self {
            Self::ClickEffect => "clickEffect",
            Self::WithEffect => "withEffect",
            Self::AfterEffect => "afterEffect",
            Self::MainSequence => "mainSeq",
            Self::InteractiveSequence => "interactiveSeq",
            Self::ClickParallel => "clickPar",
            Self::WithGroup => "withGroup",
            Self::AfterGroup => "afterGroup",
            Self::TimingRoot => "tmRoot",
        }
    }
}
