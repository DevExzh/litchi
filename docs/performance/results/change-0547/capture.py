"""One serial baseline build, four scoped profiles, and measured assembly."""
import run as R
import inspect_assembly
R.TARGET.mkdir(exist_ok=False)
(R.TARGET/'tmp').mkdir()
R.freeze()
R.build('normal')
R.capture('profile')
inspect_assembly.inspect('baseline')
