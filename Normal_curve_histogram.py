
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# 1) Ler e achatar o intervalo (pode vir 2D do Excel)
raw = xl('A1:A125')
vals = np.ravel(raw)

# 2) Tratar vírgula decimal e texto; converter para número
s = pd.Series(vals).astype(str).str.replace(',', '.', regex=False)
dados = pd.to_numeric(s, errors='coerce').dropna()

if dados.empty:
    raise ValueError('A1:A125 não contém números utilizáveis.')

# 3) Estatísticas
media = float(dados.mean())
desvio = float(dados.std(ddof=1))

# Evitar erro quando o desvio é zero (todos os valores iguais)
if not np.isfinite(desvio) or desvio <= 0:
    desvio = 1e-9  # bem pequeno só para desenhar a curva

# 4) Eixo x
xmin, xmax = float(dados.min()), float(dados.max())
if xmin == xmax:
    xmin -= 3
    xmax += 3
x = np.linspace(xmin, xmax, 200)

# 5) PDF normal sem SciPy
y = (1.0/(desvio*np.sqrt(2*np.pi))) * np.exp(-0.5*((x - media)/desvio)**2)

# 6) Plotar
bins = min(30, max(5, int(np.sqrt(len(dados)))))  # regra simples p/ bins
plt.hist(dados, bins=bins, density=True, alpha=0.6)
plt.plot(x, y, linewidth=2)
plt.title('Curva Normal Ajustada aos Dados (A1:A125)')
plt.xlabel('Valores'); plt.ylabel('Densidade')
plt.show()
