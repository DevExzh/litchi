//! Compile-time coverage for the selector-first Keynote value-axis surface.
//!
//! This test intentionally uses only the public semantic vocabulary. Native
//! object identifiers, archive values, and generated protobuf types must not
//! be required to construct or inspect a value-axis settings value.

use std::mem::size_of;

use litchi_keynote::{
    ChartValueAxisCommit, ChartValueAxisDiagnostics, ChartValueAxisEdit, ChartValueAxisError,
    ChartValueAxisLimitKind, ChartValueAxisPatch,
    chart::axis::{
        Axis, Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
    },
};

#[test]
fn value_axis_api_is_typed_and_archive_free() {
    let minimum = Bound::new(-10.0).expect("finite minimum");
    let maximum = Bound::new(100.0).expect("finite maximum");
    let bounds = Bounds::fixed(minimum, maximum).expect("ordered bounds");
    let steps = Steps::fixed(
        MajorStepCount::new(5).expect("positive major steps"),
        MinorStepCount::new(2).expect("non-negative minor steps"),
    );
    let settings = ValueAxisSettings::new(bounds, steps, Scale::Logarithmic);

    assert_eq!(Axis::Value.as_str(), "value");
    assert_eq!(settings.bounds().minimum(), Some(minimum));
    assert_eq!(settings.bounds().maximum(), Some(maximum));
    assert_eq!(settings.steps().major().map(MajorStepCount::value), Some(5));
    assert_eq!(settings.steps().minor().map(MinorStepCount::value), Some(2));
    assert_eq!(settings.scale(), Scale::Logarithmic);

    // Keep the transaction family in the public API even though its native
    // state is intentionally opaque and can only be obtained from Package.
    let _ = [
        size_of::<ChartValueAxisCommit>(),
        size_of::<ChartValueAxisDiagnostics>(),
        size_of::<ChartValueAxisEdit<'static>>(),
        size_of::<ChartValueAxisError>(),
        size_of::<ChartValueAxisLimitKind>(),
        size_of::<ChartValueAxisPatch>(),
    ];
}
