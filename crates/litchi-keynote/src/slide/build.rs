//! Slide builds: create, update and remove the build animations of a
//! slide's drawables (record 0372).
//!
//! [`crate::Package::slide_build_list`] reads a slide's builds in build
//! order; [`crate::Package::edit_slide_builds`] stages [`Edit::add`],
//! [`Edit::set`] and [`Edit::remove`] and commits them as one verified
//! package. Effects, actions and emphasis use the [`crate::build`] values.

pub use crate::package::slide_build::{
    Build, Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal, Settings, Stage,
};
