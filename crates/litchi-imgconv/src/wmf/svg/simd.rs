//! Bounded batch coordinate transformation.
//!
//! This deliberately keeps the scalar implementation as the portability and
//! correctness baseline. LLVM vectorizes the tight loop on supported targets;
//! unlike the former intrinsic implementations, it cannot read past a short
//! or mismatched input/output slice.

#[derive(Debug, Clone, Copy)]
pub(super) struct SimdTransform {
    scale_x: f64,
    scale_y: f64,
    translate_x: f64,
    translate_y: f64,
}

impl SimdTransform {
    pub(super) fn new(scale_x: f64, scale_y: f64, translate_x: f64, translate_y: f64) -> Self {
        Self {
            scale_x,
            scale_y,
            translate_x,
            translate_y,
        }
    }

    #[inline]
    pub(super) fn transform_batch(
        &self,
        xs: &[i16],
        ys: &[i16],
        out_x: &mut [f64],
        out_y: &mut [f64],
    ) -> usize {
        let len = xs.len().min(ys.len()).min(out_x.len()).min(out_y.len());
        for index in 0..len {
            out_x[index] = f64::from(xs[index]).mul_add(self.scale_x, self.translate_x);
            out_y[index] = f64::from(ys[index]).mul_add(self.scale_y, self.translate_y);
        }
        len
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn batch_transform_is_bounded_by_shortest_slice() {
        let transform = SimdTransform::new(2.0, -3.0, 1.0, 7.0);
        let mut x = [0.0; 2];
        let mut y = [0.0; 3];
        assert_eq!(
            transform.transform_batch(&[1, 2, 3], &[4, 5, 6], &mut x, &mut y),
            2
        );
        assert_eq!(x, [3.0, 5.0]);
        assert_eq!(&y[..2], &[-5.0, -8.0]);
    }
}
