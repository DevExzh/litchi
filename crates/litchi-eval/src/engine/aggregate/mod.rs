pub(crate) mod average;
pub(crate) mod count;
pub(crate) mod extrema;
pub(crate) mod sum;

pub(crate) use average::{eval_avedev, eval_average, eval_averagea};
pub(crate) use count::{eval_count, eval_counta, eval_countblank};
pub(crate) use extrema::{eval_max, eval_maxa, eval_min, eval_mina};
pub(crate) use sum::{eval_product, eval_sum, eval_sumproduct};
