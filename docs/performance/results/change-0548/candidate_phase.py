"""Build reviewed candidate and finish normal/guard ABBA without source edits."""
import run as R
import guard_run as G
import inspect_assembly as A
R.configure('candidate');R.freeze()
G.build();G.smoke();R.build('normal');R.build('alloc')
R.capture('native',1);G.capture(1)
R.capture('native',2);G.capture(2)
R.configure('baseline','candidate');R.capture('native',2);G.capture(2)
R.configure('candidate');A.inspect('candidate')
R.capture('profile');R.capture('hardware');R.capture('alloc')
