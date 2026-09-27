# Offline analysis corrections

The raw paired capture completed without run failures. Initial analysis refused the inherited `docx_styles_benign` case because both legs returned `style numPr is missing numId`. The report now explicitly treats that case as a refusal-path control, not admitted benign DOCX evidence. Raw captures are unchanged. Every other timed case must succeed; all input/outcome identities still gate.

The first independent triage stopped because the classifier omitted Error Display's `non-conformant markup compatibility XML` prefix. Its script, partial generated inputs, source/build log and executable hash are retained. The executable was removed before rerunning the corrected classifier. No production or timing data changed.
