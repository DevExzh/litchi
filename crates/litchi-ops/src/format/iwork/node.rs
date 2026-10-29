//! The typed node of an object tree.

use std::borrow::Cow;

use litchi_core::FileFormat;

use super::children::Children;
use super::tree::Item;
use super::{At, Iwork};
use crate::error::Result;
use crate::node::{
    Children as ViewChildren, Format, NodePath, Preview, Project, Source, Value, ViewOptions,
};

fn chart_length(points: f32) -> Result<crate::value::Length> {
    let emu = f64::from(points) * 12_700.0;
    let rounded = emu.round();
    #[allow(
        clippy::cast_precision_loss,
        reason = "the comparison deliberately bounds the f64 EMU value by i64"
    )]
    // `i64::MAX as f64` rounds to 2^63, and Rust's float-to-int cast would
    // saturate that boundary back to i64::MAX.  Reject the rounded 2^63
    // value explicitly instead of admitting a silent saturation.
    #[allow(
        clippy::cast_precision_loss,
        reason = "the comparison deliberately bounds the f64 EMU value below 2^63"
    )]
    let max_emu_exclusive = (1_u64 << 63) as f64;
    if !emu.is_finite() || rounded <= 0.0 || rounded >= max_emu_exclusive {
        return Err(crate::error::Error::new(
            crate::error::Code::LimitExceeded,
            "chart font size is outside the unified length range",
        )
        .with_property("font.size".to_owned()));
    }
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the checked range above is exactly the target i64 range"
    )]
    Ok(crate::value::Length::from_emu(rounded as i64))
}

/// One element of an Apple Pages document, Keynote presentation or Numbers
/// workbook: a borrowed position in the handle's object tree.
#[derive(Clone, Copy)]
pub struct Node<'f> {
    file: &'f Iwork,
    at: At,
}

impl<'f> Node<'f> {
    pub(crate) const fn new(file: &'f Iwork, at: At) -> Self {
        Self { file, at }
    }

    /// Where the node's item lives.
    pub(crate) const fn at(&self) -> At {
        self.at
    }

    pub(crate) fn item(&self) -> &'f Item {
        self.file.item(self.at)
    }

    /// The kind ordinal, without materializing a lazily built item.
    pub(crate) fn kind_ordinal(&self) -> u8 {
        self.file.kind_at(self.at)
    }

    /// The node's children, every kind.
    #[cfg(feature = "keynote")]
    pub(crate) fn child_nodes(&self) -> impl Iterator<Item = Node<'f>> + 'f {
        let file = self.file;
        file.kids(self.at)
            .iter()
            .map(move |child| Node::new(file, child))
    }

    /// The OfficeCLI type name.
    #[must_use]
    pub fn kind(&self) -> &'static str {
        self.item().type_name()
    }

    /// The format of the node's document.
    #[must_use]
    pub fn file_format(&self) -> FileFormat {
        self.file.format()
    }

    /// The canonical positional path.
    #[must_use]
    pub fn path(&self) -> NodePath<'f> {
        NodePath::new(self.item().path())
    }

    /// The element's text, if it has any.
    ///
    /// # Errors
    ///
    /// None: the text was read at open; the `Result` keeps node reads
    /// uniform across formats.
    pub fn text(&self) -> Result<Option<Cow<'f, str>>> {
        Ok(self.item().text().map(Cow::Borrowed))
    }

    /// The value of property `key`, if set.
    pub fn value(&self, key: &str) -> Result<Option<&'f Value<'static>>> {
        let item = self.item();
        if item.type_name() == "chart" {
            let values = self.chart_values(item)?;
            let index = item
                .props()
                .iter()
                .position(|property| property.name == key);
            return Ok(index.and_then(|index| values.get(index)?.as_ref()));
        }
        Ok(item.value(key))
    }

    /// Resolve the chart's deferred point-size property only when a view
    /// requests the chart format.  Tree construction therefore discovers
    /// chart nodes without scanning every chart's style graph, while a
    /// focused leaf refusal still reaches the caller as a typed error.
    fn chart_values(&self, item: &'f Item) -> Result<&'f [Option<Value<'static>>]> {
        if item.type_name() != "chart" {
            return Ok(item.values());
        }
        if let Some(values) = item.lazy_values().get() {
            return Ok(values);
        }
        let leaf = item.leaf();
        let value = match self.file.owner() {
            #[cfg(feature = "pages")]
            super::Owner::Pages(package) => chart_length(
                package
                    .body_chart_font_size(litchi_pages::BodyChartSelector::index(leaf[0] as usize))
                    .map_err(|error| {
                        super::leaf_error(
                            super::LeafError::PagesChartFont(error),
                            FileFormat::Pages,
                        )
                    })?
                    .points(),
            )
            .map(Value::Length),
            #[cfg(feature = "keynote")]
            super::Owner::Keynote(package) => chart_length(
                package
                    .slide_chart_font_size(
                        litchi_keynote::SlideSelector::index(leaf[0] as usize),
                        litchi_keynote::ChartSelector::index(leaf[2] as usize),
                    )
                    .map_err(|error| {
                        super::leaf_error(
                            super::LeafError::KeynoteChartFont(error),
                            FileFormat::Keynote,
                        )
                    })?
                    .points(),
            )
            .map(Value::Length),
            #[cfg(feature = "numbers")]
            super::Owner::Numbers(package) => {
                let property = package
                    .sheet_chart_font(
                        litchi_numbers::SheetSelector::index(leaf[0] as usize),
                        litchi_numbers::ChartSelector::index(leaf[1] as usize),
                        litchi_numbers::sheet::chart::font::PropertyKind::Size,
                    )
                    .map_err(|error| {
                        super::leaf_error(
                            super::LeafError::NumbersChartFont(error),
                            FileFormat::Numbers,
                        )
                    })?;
                match property {
                    litchi_numbers::sheet::chart::font::Property::Size(size) => {
                        chart_length(size.points()).map(Value::Length)
                    },
                    litchi_numbers::sheet::chart::font::Property::Font(_) => {
                        Err(super::leaf_error(
                            super::LeafError::NumbersChartFont(
                                litchi_numbers::sheet::chart::font::Error::InvalidSource,
                            ),
                            FileFormat::Numbers,
                        ))
                    },
                    _ => Err(super::leaf_error(
                        super::LeafError::NumbersChartFont(
                            litchi_numbers::sheet::chart::font::Error::InvalidSource,
                        ),
                        FileFormat::Numbers,
                    )),
                }
            },
        }?;
        let mut values = item.values().to_vec().into_boxed_slice();
        let index = item
            .props()
            .iter()
            .position(|property| property.name == "font.size")
            .ok_or_else(|| {
                crate::error::Error::new(
                    crate::error::Code::InternalError,
                    "chart font-size property is missing",
                )
            })?;
        values[index] = Some(value);
        let _ = item.lazy_values().set(values);
        Ok(item
            .lazy_values()
            .get()
            .expect("deferred chart values were just initialized"))
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("iwork::Node")
            .field("kind", &self.kind())
            .finish_non_exhaustive()
    }
}

impl<'f> Project<'f> for Node<'f> {
    fn kind(&self) -> &'static str {
        Node::kind(self)
    }

    fn file_format(&self) -> FileFormat {
        Node::file_format(self)
    }

    fn path(&self) -> NodePath<'_> {
        Node::path(self)
    }

    fn text(&self) -> Result<Option<Cow<'_, str>>> {
        Node::text(self)
    }

    fn preview(&self) -> Result<Option<Preview<'_>>> {
        Ok(self
            .item()
            .text()
            .map(|text| Preview::Prefix(Cow::Borrowed(text))))
    }

    fn child_count(&self, _: &ViewOptions) -> Result<Option<u32>> {
        Ok(Some(
            u32::try_from(self.file.kids(self.at).len()).unwrap_or(u32::MAX),
        ))
    }

    fn format(&self) -> Result<Format<'_>> {
        let item = self.item();
        Ok(Format::new(
            item.props(),
            Source::Tree(self.chart_values(item)?),
        ))
    }

    fn children(&self, _: &ViewOptions) -> ViewChildren<'_, 'f> {
        ViewChildren::iwork(Children::filtered(self.file, self.at, u128::MAX))
    }
}

const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Node<'static>>();
    assert!(size_of::<Node<'static>>() <= 16);
};

#[cfg(test)]
mod tests {
    use super::chart_length;
    use crate::error::Code;

    #[test]
    fn chart_length_rejects_unrepresentable_positive_sizes() {
        assert_eq!(chart_length(12.0).expect("normal size").emu(), 152_400);
        assert_eq!(
            chart_length(0.000_001).unwrap_err().code(),
            Code::LimitExceeded
        );
        assert_eq!(
            chart_length(f32::MAX).unwrap_err().code(),
            Code::LimitExceeded
        );
        #[allow(
            clippy::cast_possible_truncation,
            clippy::cast_precision_loss,
            reason = "the f32 neighbors are the actual leaf-representable boundary inputs"
        )]
        let threshold = (2_f64.powi(63) / 12_700.0) as f32;
        for points in [
            f32::from_bits(threshold.to_bits().saturating_sub(1)),
            threshold,
            f32::from_bits(threshold.to_bits().saturating_add(1)),
        ] {
            let rounded = (f64::from(points) * 12_700.0).round();
            let expected = rounded.is_finite() && rounded > 0.0 && rounded < 2_f64.powi(63);
            assert_eq!(chart_length(points).is_ok(), expected, "{points:?}");
        }
    }
}
