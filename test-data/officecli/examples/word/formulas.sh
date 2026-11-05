#!/bin/bash
# Generate complex math/chemistry/physics formula test document
# Usage: ./gen_formulas.sh [officecli path]

CLI="${1:-officecli}"
OUT="$(dirname "$0")/formulas.docx"

rm -f "$OUT"
$CLI create "$OUT"
$CLI open "$OUT"

# ==================== Title ====================
$CLI add "$OUT" /body --type paragraph --prop text="Complex Math/Chemistry/Physics Formula Collection" --prop style=Heading1 --prop align=center

# ==================== I. Algebra ====================
$CLI add "$OUT" /body --type paragraph --prop text="I. Algebra" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="1. Quadratic Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=x = \frac{-b \pm \sqrt{b^{2} - 4ac}}{2a}'

$CLI add "$OUT" /body --type paragraph --prop text="2. Binomial Theorem:"
$CLI add "$OUT" /body --type equation --prop 'formula=(a+b)^{n} = \sum_{k=0}^{n} \binom{n}{k} a^{n-k} b^{k}'

$CLI add "$OUT" /body --type paragraph --prop text="3. Euler's Identity:"
$CLI add "$OUT" /body --type equation --prop 'formula=e^{i\pi} + 1 = 0'

# ==================== II. Calculus ====================
$CLI add "$OUT" /body --type paragraph --prop text="II. Calculus" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="4. Limit Definition of Derivative:"
$CLI add "$OUT" /body --type equation --prop 'formula=f^{\prime}(x) = \lim_{\Delta x \rightarrow 0} \frac{f(x + \Delta x) - f(x)}{\Delta x}'

$CLI add "$OUT" /body --type paragraph --prop text="5. Gaussian Integral:"
$CLI add "$OUT" /body --type equation --prop 'formula=\int_{-\infty}^{+\infty} e^{-x^{2}} dx = \sqrt{\pi}'

$CLI add "$OUT" /body --type paragraph --prop text="6. Taylor Series Expansion:"
$CLI add "$OUT" /body --type equation --prop 'formula=f(x) = \sum_{n=0}^{\infty} \frac{f^{(n)}(a)}{n!} (x-a)^{n}'

$CLI add "$OUT" /body --type paragraph --prop text="7. Newton-Leibniz Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=\int_{a}^{b} f(x) dx = F(b) - F(a)'

$CLI add "$OUT" /body --type paragraph --prop text="8. Triple Integral (Spherical Coordinates):"
$CLI add "$OUT" /body --type equation --prop 'formula=\iiint_{V} f(r, \theta, \phi) r^{2} \sin\theta \, dr \, d\theta \, d\phi'

$CLI add "$OUT" /body --type paragraph --prop text="9. Fourier Transform:"
$CLI add "$OUT" /body --type equation --prop 'formula=\hat{f}(\xi) = \int_{-\infty}^{+\infty} f(x) e^{-2\pi i x \xi} dx'

# ==================== III. Linear Algebra ====================
$CLI add "$OUT" /body --type paragraph --prop text="III. Linear Algebra" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="10. Matrix Characteristic Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\det(A - \lambda I) = 0'

# ==================== IV. Probability and Statistics ====================
$CLI add "$OUT" /body --type paragraph --prop text="IV. Probability and Statistics" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="11. Bayes' Theorem:"
$CLI add "$OUT" /body --type equation --prop 'formula=P(A|B) = \frac{P(B|A) \cdot P(A)}{P(B)}'

$CLI add "$OUT" /body --type paragraph --prop text="12. Normal Distribution PDF:"
$CLI add "$OUT" /body --type equation --prop 'formula=f(x) = \frac{1}{\sigma \sqrt{2\pi}} e^{-\frac{(x-\mu)^{2}}{2\sigma^{2}}}'

$CLI add "$OUT" /body --type paragraph --prop text="13. Variance Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=\sigma^{2} = \frac{1}{N} \sum_{i=1}^{N} (x_{i} - \mu)^{2}'

# ==================== V. Number Theory and Series ====================
$CLI add "$OUT" /body --type paragraph --prop text="V. Number Theory and Series" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="14. Riemann Zeta Function:"
$CLI add "$OUT" /body --type equation --prop 'formula=\zeta(s) = \sum_{n=1}^{\infty} \frac{1}{n^{s}}'

$CLI add "$OUT" /body --type paragraph --prop text="15. Stirling's Approximation:"
$CLI add "$OUT" /body --type equation --prop 'formula=n! \approx \sqrt{2\pi n} \left(\frac{n}{e}\right)^{n}'

# ==================== VI. Chemistry ====================
$CLI add "$OUT" /body --type paragraph --prop text="VI. Chemistry" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="16. Copper Sulfate Crystal Dissolution:"
$CLI add "$OUT" /body --type equation --prop 'formula=CuSO_{4} \cdot 5H_{2}O \rightarrow Cu^{2+} + SO_{4}^{2-} + 5H_{2}O'

$CLI add "$OUT" /body --type paragraph --prop text="17. Thermochemical Equation (Methane Combustion):"
$CLI add "$OUT" /body --type equation --prop 'formula=CH_{4}(g) + 2O_{2}(g) \rightarrow CO_{2}(g) + 2H_{2}O(l) \quad \Delta H = -890.3 \, kJ/mol'

$CLI add "$OUT" /body --type paragraph --prop text="18. Chemical Equilibrium Constant Expression:"
$CLI add "$OUT" /body --type equation --prop 'formula=K_{eq} = \frac{[C]^{c} [D]^{d}}{[A]^{a} [B]^{b}}'

$CLI add "$OUT" /body --type paragraph --prop text="19. Esterification Reaction (Reversible):"
$CLI add "$OUT" /body --type equation --prop 'formula=CH_{3}COOH + C_{2}H_{5}OH \rightleftharpoons CH_{3}COOC_{2}H_{5} + H_{2}O'

$CLI add "$OUT" /body --type paragraph --prop text="20. Henderson-Hasselbalch Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=pH = pK_{a} + \log \frac{[A^{-}]}{[HA]}'

$CLI add "$OUT" /body --type paragraph --prop text="21. Van der Waals Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\left(P + \frac{a n^{2}}{V^{2}}\right)(V - nb) = nRT'

$CLI add "$OUT" /body --type paragraph --prop text="22. Arrhenius Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=k = A e^{-\frac{E_{a}}{RT}}'

# ==================== VII. Physics ====================
$CLI add "$OUT" /body --type paragraph --prop text="VII. Physics" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="23. Maxwell's Equations (Differential Form):"
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \cdot E = \frac{\rho}{\epsilon_{0}}'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \cdot B = 0'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \times E = -\frac{\partial B}{\partial t}'
$CLI add "$OUT" /body --type equation --prop 'formula=\nabla \times B = \mu_{0} J + \mu_{0} \epsilon_{0} \frac{\partial E}{\partial t}'

$CLI add "$OUT" /body --type paragraph --prop text="24. Einstein Field Equations:"
$CLI add "$OUT" /body --type equation --prop 'formula=R_{\mu\nu} - \frac{1}{2} R g_{\mu\nu} + \Lambda g_{\mu\nu} = \frac{8\pi G}{c^{4}} T_{\mu\nu}'

$CLI add "$OUT" /body --type paragraph --prop text="25. Schrodinger Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=i\hbar \frac{\partial}{\partial t} \Psi(r, t) = \hat{H} \Psi(r, t)'

$CLI add "$OUT" /body --type paragraph --prop text="26. Dirac Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=(i\gamma^{\mu} \partial_{\mu} - m) \psi = 0'

$CLI add "$OUT" /body --type paragraph --prop text="27. Euler-Lagrange Equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\frac{d}{dt} \frac{\partial L}{\partial \dot{q}_{i}} - \frac{\partial L}{\partial q_{i}} = 0'

$CLI add "$OUT" /body --type paragraph --prop text="28. Heisenberg Uncertainty Principle:"
$CLI add "$OUT" /body --type equation --prop 'formula=\Delta x \cdot \Delta p \geq \frac{\hbar}{2}'

$CLI add "$OUT" /body --type paragraph --prop text="29. Planck's Black-Body Radiation Formula:"
$CLI add "$OUT" /body --type equation --prop 'formula=B(\nu, T) = \frac{2h\nu^{3}}{c^{2}} \cdot \frac{1}{e^{\frac{h\nu}{k_{B} T}} - 1}'

$CLI add "$OUT" /body --type paragraph --prop text="30. Lorentz Transformation:"
$CLI add "$OUT" /body --type equation --prop 'formula=t^{\prime} = \gamma \left(t - \frac{vx}{c^{2}}\right), \quad \gamma = \frac{1}{\sqrt{1 - \frac{v^{2}}{c^{2}}}}'

# ==================== VIII. Advanced Notation ====================
$CLI add "$OUT" /body --type paragraph --prop text="VIII. Advanced Notation" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="31. Matrix (pmatrix):"
$CLI add "$OUT" /body --type equation --prop 'formula=A = \begin{pmatrix} a_{11} & a_{12} & a_{13} \\ a_{21} & a_{22} & a_{23} \\ a_{31} & a_{32} & a_{33} \end{pmatrix}'

$CLI add "$OUT" /body --type paragraph --prop text="32. Determinant (vmatrix):"
$CLI add "$OUT" /body --type equation --prop 'formula=\det(A) = \begin{vmatrix} a & b \\ c & d \end{vmatrix} = ad - bc'

$CLI add "$OUT" /body --type paragraph --prop text="33. Bracketed Matrix (bmatrix):"
$CLI add "$OUT" /body --type equation --prop 'formula=I_{3} = \begin{bmatrix} 1 & 0 & 0 \\ 0 & 1 & 0 \\ 0 & 0 & 1 \end{bmatrix}'

$CLI add "$OUT" /body --type paragraph --prop text="34. Piecewise Function (cases):"
$CLI add "$OUT" /body --type equation --prop 'formula=|x| = \begin{cases} x, & x \geq 0 \\ -x, & x < 0 \end{cases}'

$CLI add "$OUT" /body --type paragraph --prop text="35. Auto-sized Delimiters (various brackets):"
$CLI add "$OUT" /body --type equation --prop 'formula=\left[ \frac{a}{b} \right] + \left\{ \frac{c}{d} \right\} + \left| \frac{e}{f} \right| + \left\langle \frac{g}{h} \right\rangle'

$CLI add "$OUT" /body --type paragraph --prop text="36. Floor and Ceiling:"
$CLI add "$OUT" /body --type equation --prop 'formula=\left\lfloor \frac{n}{2} \right\rfloor + \left\lceil \frac{n}{2} \right\rceil = n'

$CLI add "$OUT" /body --type paragraph --prop text="37. Underbrace and Overbrace:"
$CLI add "$OUT" /body --type equation --prop 'formula=\underbrace{1 + 2 + \cdots + n}_{n \text{ terms}} = \overbrace{\frac{n(n+1)}{2}}^{\text{closed form}}'

$CLI add "$OUT" /body --type paragraph --prop text="38. Overset (definition):"
$CLI add "$OUT" /body --type equation --prop 'formula=f(x) \overset{\text{def}}{=} \lim_{h \to 0} \frac{f(x+h) - f(x)}{h}'

$CLI add "$OUT" /body --type paragraph --prop text="39. Math Fonts (mathbb / mathcal / mathbf / mathrm):"
$CLI add "$OUT" /body --type equation --prop 'formula=\forall x \in \mathbb{R}, \exists \mathcal{L} : \mathbf{v} \mapsto \mathrm{d}\mathbf{v}'

$CLI add "$OUT" /body --type paragraph --prop text="40. Cancellation:"
$CLI add "$OUT" /body --type equation --prop 'formula=\frac{(x+1) \cancel{(x-1)}}{\cancel{(x-1)}} = x + 1'

$CLI add "$OUT" /body --type paragraph --prop text="41. Cancel-to (limit):"
$CLI add "$OUT" /body --type equation --prop 'formula=\lim_{x \to \infty} \cancelto{0}{\frac{1}{x}} + 1 = 1'

$CLI add "$OUT" /body --type paragraph --prop text="42. Boxed Result:"
$CLI add "$OUT" /body --type equation --prop 'formula=\boxed{E = mc^{2}}'

$CLI add "$OUT" /body --type paragraph --prop text="43. Accents (bar / vec / tilde / ddot):"
$CLI add "$OUT" /body --type equation --prop 'formula=\bar{x} = \frac{1}{n} \sum x_{i}, \quad \vec{F} = m\ddot{\vec{r}}, \quad \tilde{f}(\xi)'

$CLI add "$OUT" /body --type paragraph --prop text="44. Overline and Underline:"
$CLI add "$OUT" /body --type equation --prop 'formula=\overline{A \cup B} = \overline{A} \cap \overline{B}, \quad \underline{x} \leq x \leq \overline{x}'

$CLI add "$OUT" /body --type paragraph --prop text="45. Hyperbolic and Inverse Trig:"
$CLI add "$OUT" /body --type equation --prop 'formula=\arctan(x) = \int_{0}^{x} \frac{dt}{1+t^{2}}, \quad \cosh^{2}(x) - \sinh^{2}(x) = 1'

$CLI add "$OUT" /body --type paragraph --prop text="46. Custom Operator (operatorname):"
$CLI add "$OUT" /body --type equation --prop 'formula=\operatorname{lcm}(a, b) \cdot \gcd(a, b) = |ab|'

$CLI add "$OUT" /body --type paragraph --prop text="47. Modular Arithmetic:"
$CLI add "$OUT" /body --type equation --prop 'formula=a \equiv b \pmod{n} \iff n \mid (a - b), \quad 17 \bmod 5 = 2'

$CLI add "$OUT" /body --type paragraph --prop text="48. Double Integral with Text:"
$CLI add "$OUT" /body --type equation --prop 'formula=\iint_{D} f(x,y) \, dA \quad \text{where } D = \{(x,y) : x^{2}+y^{2} \leq 1\}'

$CLI add "$OUT" /body --type paragraph --prop text="49. Big Operators (bigcup / bigcap / coprod):"
$CLI add "$OUT" /body --type equation --prop 'formula=\bigcup_{i=1}^{n} A_{i} \supseteq \bigcap_{i=1}^{n} A_{i}, \quad \coprod_{i \in I} X_{i}'

$CLI add "$OUT" /body --type paragraph --prop text="50. Greek Letters (full uppercase set):"
$CLI add "$OUT" /body --type equation --prop 'formula=\Gamma, \Theta, \Xi, \Pi, \Phi, \Psi, \Omega \in \{\alpha, \beta, \gamma, \delta, \epsilon, \zeta, \eta, \theta\}'

$CLI add "$OUT" /body --type paragraph --prop text="51. Dots (ldots / cdots / vdots / ddots):"
$CLI add "$OUT" /body --type equation --prop 'formula=M = \begin{pmatrix} a_{11} & \cdots & a_{1n} \\ \vdots & \ddots & \vdots \\ a_{m1} & \cdots & a_{mn} \end{pmatrix}, \quad x_{1}, x_{2}, \ldots, x_{n}'

$CLI add "$OUT" /body --type paragraph --prop text="52. Spacing Control (quad / qquad / thinsp):"
$CLI add "$OUT" /body --type equation --prop 'formula=a + b \, c \; d \quad e \qquad f'

$CLI add "$OUT" /body --type paragraph --prop text="53. Colored Math (textcolor / color):"
$CLI add "$OUT" /body --type equation --prop 'formula=\textcolor{red}{x^{2}} + \textcolor{blue}{2xy} + \textcolor{green}{y^{2}} = \color{purple}{(x+y)^{2}}'

$CLI add "$OUT" /body --type paragraph --prop text="54. Set Theory:"
$CLI add "$OUT" /body --type equation --prop 'formula=A \subseteq B \iff \forall x \in A, x \in B; \quad A \setminus B = \{x : x \in A \land x \notin B\}; \quad \emptyset \subset A'

$CLI add "$OUT" /body --type paragraph --prop text="55. Norm and Inner Product:"
$CLI add "$OUT" /body --type equation --prop 'formula=\|x\|_{2} = \sqrt{\langle x, x \rangle} = \sqrt{\sum_{i=1}^{n} x_{i}^{2}}'

# ==================== IX. Equation Mode (display vs inline) ====================
$CLI add "$OUT" /body --type paragraph --prop text="IX. Equation Mode — display vs inline" --prop style=Heading2

# mode=display (default): equation gets its own block-level oMathPara element
$CLI add "$OUT" /body --type paragraph --prop text="56. Display mode (default) — centred block equation:"
$CLI add "$OUT" /body --type equation --prop 'formula=E = mc^{2}' --prop mode=display

# mode=inline: equation is appended to the parent paragraph as an oMath child
$CLI add "$OUT" /body --type paragraph --prop text="57. Inline mode — equation embedded mid-sentence:"
$CLI add "$OUT" /body --type equation --prop 'formula=A = \pi r^{2}' --prop mode=inline

# ==================== X. Coverage Completeness — Additional Supported Commands ====================
$CLI add "$OUT" /body --type paragraph --prop text="X. Coverage Completeness — Additional Supported Commands" --prop style=Heading2

$CLI add "$OUT" /body --type paragraph --prop text="58. N-ary Contour Integrals (oint / oiint / oiiint):"
$CLI add "$OUT" /body --type equation --prop 'formula=\oint_C \vec{F} \cdot d\vec{r} = \iint_S (\nabla \times \vec{F}) \cdot d\vec{S}, \quad \oiint_S \vec{E} \cdot d\vec{A} = \frac{Q}{\epsilon_0}, \quad \oiiint_V \rho \, dV'

$CLI add "$OUT" /body --type paragraph --prop text="59. Limit-style Operators with Under-limits (max / min / sup / inf):"
$CLI add "$OUT" /body --type equation --prop 'formula=\max_{1 \le i \le n} a_i \geq \min_{1 \le i \le n} a_i, \quad \sup_{x \in S} f(x) \geq \inf_{x \in S} f(x)'

$CLI add "$OUT" /body --type paragraph --prop text="60. More Limit Operators (limsup / liminf / argmax / argmin):"
$CLI add "$OUT" /body --type equation --prop 'formula=\limsup_{n \to \infty} x_n \geq \liminf_{n \to \infty} x_n, \quad \hat{\theta} = \argmax_{\theta} L(\theta) = \argmin_{\theta} (-L(\theta))'

$CLI add "$OUT" /body --type paragraph --prop text="61. Named Operators with Limits (det / gcd / Pr):"
$CLI add "$OUT" /body --type equation --prop 'formula=\det_{A \in M} A, \quad \gcd_{i} a_i, \quad \Pr_{x \sim D}[X = x]'

$CLI add "$OUT" /body --type paragraph --prop text="62. Limits Placement Control (\\limits / \\nolimits):"
$CLI add "$OUT" /body --type equation --prop 'formula=\lim\limits_{x \to 0} \frac{\sin x}{x} = 1, \quad \sum\nolimits_{i=1}^{n} i = \frac{n(n+1)}{2}'

$CLI add "$OUT" /body --type paragraph --prop text="63. N-ary Product (prod):"
$CLI add "$OUT" /body --type equation --prop 'formula=n! = \prod_{k=1}^{n} k, \quad \prod_{p \text{ prime}} \frac{1}{1 - p^{-s}} = \zeta(s)'

$CLI add "$OUT" /body --type paragraph --prop text="64. Binary Operators (div / ast / star / circ / oplus / ominus / otimes / odot / bullet):"
$CLI add "$OUT" /body --type equation --prop 'formula=a \div b, \quad f \ast g, \quad a \star b, \quad f \circ g, \quad a \oplus b \ominus c, \quad u \otimes v \odot w, \quad x \bullet y'

$CLI add "$OUT" /body --type paragraph --prop text="65. Arrows (leftarrow / uparrow / downarrow / leftrightarrow / Rightarrow / Leftarrow / Leftrightarrow / gets / implies):"
$CLI add "$OUT" /body --type equation --prop 'formula=a \leftarrow b \uparrow c \downarrow d \leftrightarrow e, \quad P \Rightarrow Q, \quad R \Leftarrow S, \quad X \Leftrightarrow Y, \quad n \gets n+1, \quad p \implies q'

$CLI add "$OUT" /body --type paragraph --prop text="66. Math Fonts (boldsymbol / mathit) and Over/Under-set:"
$CLI add "$OUT" /body --type equation --prop 'formula=\boldsymbol{\alpha} + \mathit{xyz}, \quad \overset{!}{=} \quad \underset{n \to \infty}{\lim} a_n'

$CLI add "$OUT" /body --type paragraph --prop text="67. Relations, Logic, Sets, Trig, and Legacy Fraction (\\neq / \\sim / \\subset / \\lor / \\neg / \\wedge / \\parallel / \\varnothing / \\complement / \\cos / \\tan / \\ln / {a \\over b}):"
$CLI add "$OUT" /body --type equation --prop 'formula=a \neq b \sim c, \quad A \subset B \supset C, \quad p \lor \neg q \wedge r, \quad u \vee v, \quad \ell_1 \parallel \ell_2, \quad \varnothing = \complement_U U, \quad \cos^2 x + \tan x - \ln x, \quad {a \over b}'

# ==================== XI. Full Symbol & Environment Coverage ====================
$CLI add "$OUT" /body --type paragraph --prop text="XI. Full Symbol & Environment Coverage" --prop style=Heading2

# Features: Greek variants — \chi \iota \kappa \omega \tau \upsilon \varepsilon \varphi \varpi \varrho \varsigma \vartheta \varkappa \digamma
$CLI add "$OUT" /body --type paragraph --prop text="68. Greek Variants and Extra Letters:"
$CLI add "$OUT" /body --type equation --prop 'formula=\chi, \iota, \kappa, \omega, \tau, \upsilon, \varepsilon, \varphi, \varpi, \varrho, \varsigma, \vartheta, \varkappa, \digamma'

# Features: relations — \cong \simeq \asymp \doteq \propto \prec \succ \preceq \succeq \ll \gg \models \vdash \dashv \perp \top \ni \sqsubset \sqsubseteq \sqsupset \sqsupseteq \subsetneq \supsetneq \Vdash
$CLI add "$OUT" /body --type paragraph --prop text="69. Relation Symbols:"
$CLI add "$OUT" /body --type equation --prop 'formula=a \cong b \simeq c \asymp d \doteq e, \quad f \propto g, \quad x \prec y \succ z, \quad p \preceq q \succeq r, \quad m \ll n \gg k, \quad \Gamma \models \phi \vdash \psi \dashv \chi \Vdash \omega, \quad u \perp v, \quad \top, \quad a \ni b, \quad S \sqsubset T \sqsubseteq U \sqsupset V \sqsupseteq W, \quad A \subsetneq B \supsetneq C'

# Features: negated relations — \nleq \ngeq \nmid \nparallel \nsubseteq \nsupseteq \nexists
$CLI add "$OUT" /body --type paragraph --prop text="70. Negated Relations:"
$CLI add "$OUT" /body --type equation --prop 'formula=a \nleq b, \quad c \ngeq d, \quad e \nmid f, \quad g \nparallel h, \quad A \nsubseteq B, \quad C \nsupseteq D, \quad \nexists x'

# Features: arrows — \longleftarrow \longrightarrow \longleftrightarrow \longmapsto \hookleftarrow \hookrightarrow \twoheadrightarrow \rightsquigarrow \leftharpoonup \leftharpoondown \rightharpoonup \rightharpoondown \nearrow \searrow \swarrow \nwarrow \curvearrowleft \curvearrowright \impliedby
$CLI add "$OUT" /body --type paragraph --prop text="71. Extended Arrows:"
$CLI add "$OUT" /body --type equation --prop 'formula=a \longleftarrow b \longrightarrow c \longleftrightarrow d, \quad x \longmapsto y, \quad e \hookleftarrow f \hookrightarrow g, \quad p \twoheadrightarrow q \rightsquigarrow r, \quad u \leftharpoonup v \leftharpoondown w \rightharpoonup s \rightharpoondown t, \quad \nearrow \searrow \swarrow \nwarrow, \quad \alpha \curvearrowleft \beta \curvearrowright \gamma, \quad P \impliedby Q'

# Features: misc symbols — \aleph \beth \gimel \daleth \wp \Re \Im \Sigma \angle \measuredangle \sphericalangle \triangle \triangleleft \triangleright \square \blacksquare \Diamond \diamond \diamondsuit \clubsuit \heartsuit \spadesuit \flat \sharp \natural \dagger \ddagger \bigstar \amalg \uplus \sqcap \sqcup \wr \bowtie \frown \smile \mp \bigtriangledown
$CLI add "$OUT" /body --type paragraph --prop text="72. Miscellaneous Symbols:"
$CLI add "$OUT" /body --type equation --prop 'formula=\aleph, \beth, \gimel, \daleth, \wp, \Re, \Im, \Sigma, \quad \angle, \measuredangle, \sphericalangle, \triangle, \triangleleft, \triangleright, \quad \square, \blacksquare, \Diamond, \diamond, \diamondsuit, \clubsuit, \heartsuit, \spadesuit, \quad \flat, \sharp, \natural, \dagger, \ddagger, \bigstar, \quad a \amalg b \uplus c \sqcap d \sqcup e \wr f, \quad x \bowtie y \frown z \smile w, \quad p \mp q, \quad \bigtriangledown'

# Features: math font families — \mathfrak \mathsf \mathtt \textbf \textit \textsf \texttt
$CLI add "$OUT" /body --type paragraph --prop text="73. Math Font Families (Fraktur / sans / mono / bold / italic text):"
$CLI add "$OUT" /body --type equation --prop 'formula=\mathfrak{ABCDabcd} \quad \mathsf{ABCDabcd} \quad \mathtt{ABCDabcd} \quad \textbf{ABCDabcd} \quad \textit{ABCDabcd} \quad \textsf{ABCDabcd} \quad \texttt{ABCDabcd}'

# Features: environments — Bmatrix Vmatrix (brace / double-bar matrices)
$CLI add "$OUT" /body --type paragraph --prop text="74. Environments — Bmatrix and Vmatrix:"
$CLI add "$OUT" /body --type equation --prop 'formula=\begin{Bmatrix} a & b \\ c & d \end{Bmatrix} \quad \begin{Vmatrix} a & b \\ c & d \end{Vmatrix}'

# Features: environments — smallmatrix array (with colspec)
$CLI add "$OUT" /body --type paragraph --prop text="75. Environments — smallmatrix and array (colspec):"
$CLI add "$OUT" /body --type equation --prop 'formula=\left(\begin{smallmatrix} 1 & 0 \\ 0 & 1 \end{smallmatrix}\right) \quad \begin{array}{cc} x & y \\ z & w \end{array}'

# Features: environments — aligned align (multi alignment points)
$CLI add "$OUT" /body --type paragraph --prop text="76. Environments — aligned and align (multi alignment points):"
$CLI add "$OUT" /body --type equation --prop 'formula=\begin{aligned} a &= b \\ c &= d \end{aligned} \qquad \begin{align} a &= b & c &= d \\ e &= f & g &= h \end{align}'

# Features: environments — gather split substack
$CLI add "$OUT" /body --type paragraph --prop text="77. Environments — gather, split, and substack:"
$CLI add "$OUT" /body --type equation --prop 'formula=\begin{gather} x = 1 \\ y = 2 \end{gather} \qquad \begin{split} a &= b + c \\ &= d \end{split} \qquad \sum_{\substack{i=1 \\ j=1}}^{n} a_{ij}'

$CLI close "$OUT"

$CLI validate "$OUT"
echo "Generated: $OUT"
