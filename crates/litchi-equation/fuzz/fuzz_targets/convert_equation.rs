#![no_main]

//! Every conversion entry point of `litchi-equation`, selected by the first
//! input byte: binary MTEF, OMML and MathML to LaTeX, and LaTeX to OMML for
//! each host. A conversion may fail; it must not panic or exceed its limits.

use libfuzzer_sys::fuzz_target;
use litchi_equation::{Limits, OmmlHost};

fuzz_target!(|data: &[u8]| {
    let Some((&selector, input)) = data.split_first() else {
        return;
    };
    let limits = Limits::DEFAULT;
    match selector % 4 {
        0 => {
            let _ = litchi_equation::read_mtef(input, &limits);
        },
        selector => {
            let Ok(text) = std::str::from_utf8(input) else {
                return;
            };
            match selector {
                1 => {
                    let _ = litchi_equation::omml_to_latex(text, &limits);
                },
                2 => {
                    let _ = litchi_equation::mathml_to_latex(text, &limits);
                },
                _ => {
                    for host in [
                        OmmlHost::Bare,
                        OmmlHost::WordprocessingMl,
                        OmmlHost::DrawingMl,
                    ] {
                        if let Ok(omml) = litchi_equation::latex_to_omml(text, host, &limits) {
                            // What the writer produced reads back.
                            let _ = litchi_equation::omml_to_latex(omml.xml(), &limits);
                        }
                    }
                },
            }
        },
    }
});
