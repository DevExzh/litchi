"""Serial baseline instruction-attribution capture; no candidate runtime."""
import run as R
from inspect_assembly import inspect
R.configure('baseline')
R.build('normal')
inspect()
R.capture('native')
R.capture('profile')
