"""Capture complete baseline before applying isolated production candidate."""
import run as R
import guard_run as G
import inspect_assembly as A
R.TARGET.mkdir(exist_ok=False)
(R.TARGET/'tmp').mkdir()
R.configure('baseline');R.freeze()
G.build();R.build('normal');R.build('alloc')
A.inspect('baseline')
R.capture('profile');R.capture('hardware');R.capture('alloc')
R.capture('native',1);G.capture(1)
