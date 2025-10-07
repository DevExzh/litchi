// OMML element handlers
//
// This module contains handlers for specific OMML elements that require
// complex parsing logic. Each handler is organized into separate modules
// for better maintainability.

mod delim;
mod bar;
mod accent;
mod matrix;
mod fraction;
mod nary;
mod function;
mod radical;
mod script;
mod box_handler;
mod phantom;
mod group_char;
mod border_box;
mod eq_arr;
mod spacing;
mod char_handler;
mod components;
mod matrix_cell;
mod eq_arr_pr;
mod limit;
mod pre_script;
mod post_script;
mod run_props;
mod ctrl_props;

pub use delim::DelimiterHandler;
pub use bar::BarHandler;
pub use accent::AccentHandler;
pub use matrix::{MatrixHandler, MatrixRowHandler};
pub use fraction::FractionHandler;
pub use nary::NaryHandler;
pub use function::{FunctionHandler, FunctionNameHandler};
pub use radical::RadicalHandler;
pub use script::{SuperscriptHandler, SubscriptHandler, SubSupHandler, SuperscriptElementHandler, SubscriptElementHandler};
pub use box_handler::BoxHandler;
pub use phantom::PhantomHandler;
pub use group_char::GroupCharHandler;
pub use border_box::BorderBoxHandler;
pub use eq_arr::EqArrHandler;
pub use spacing::SpacingHandler;
pub use char_handler::CharHandler;
pub use components::{NumeratorHandler, DenominatorHandler, DegreeHandler, BaseHandler, LowerLimitHandler, UpperLimitHandler, IntegrandHandler, LimUppHandler, LimLowHandler};
pub use matrix_cell::MatrixCellHandler;
pub use eq_arr_pr::EqArrPrHandler;
pub use limit::LimitHandler;
pub use pre_script::PreScriptHandler;
pub use post_script::PostScriptHandler;
pub use run_props::RunPropsHandler;
pub use ctrl_props::CtrlPropsHandler;
