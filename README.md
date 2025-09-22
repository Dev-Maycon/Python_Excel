Análise de Distribuição Normal no Excel

Este projeto contém um script Python que permite analisar dados do Excel e gerar um histograma com uma curva normal ajustada, diretamente integrado com o Microsoft Excel.

📊 Funcionalidades

· Leitura direta do Excel: Lê dados de um intervalo específico (A1:A125)
· Pré-processamento automático: Converte vírgulas decimais para pontos e trata valores não numéricos
· Análise estatística: Calcula média e desvio padrão dos dados
· Visualização: Gera histograma com curva normal sobreposta
· Robusto: Trata casos especiais como desvio padrão zero ou dados constantes

🚀 Como Usar

Pré-requisitos

· Python 3.6 ou superior
· Microsoft Excel
· Biblioteca xlwings para integração com Excel

Uso no Excel

1. Coloque seus dados numéricos no intervalo A1:A125
2. No Excel, pressione Alt + F8
3. Execute a função Python que contém este código
4. O gráfico será exibido automaticamente

📁 Estrutura do Código

```python
# 1) Ler dados do Excel e achatar para 1D
# 2) Tratar formatação decimal e converter para numérico
# 3) Calcular estatísticas (média e desvio padrão)
# 4) Gerar eixo x para a curva normal
# 5) Calcular PDF normal sem dependência do SciPy
# 6) Plotar histograma e curva normal
```

⚙️ Personalização

Alterar o intervalo de dados

Modifique a linha:

```python
raw = xl('A1:A125')  # Altere para seu intervalo
```

Ajustar número de bins do histograma

Modifique a regra na linha:

```python
bins = min(30, max(5, int(np.sqrt(len(dados)))))
```

🐛 Solução de Problemas

Erro: "A1:A125 não contém números utilizáveis"

· Verifique se existem dados numéricos no intervalo especificado
· Confirme a formatação dos números (suporte a vírgula decimal)

Erro: Módulos não encontrados

· Execute pip install para todas as dependências
· Verifique se o Python path está correto no Excel

Gráfico não aparece

· Certifique-se de que o matplotlib está instalado corretamente
· Verifique se há dados válidos no intervalo

📈 Exemplo de Saída

O script gera um gráfico contendo:

· Histograma dos dados com densidade normalizada
· Curva normal teórica baseada na média e desvio padrão dos dados
· Título e labels automáticos baseados no intervalo de dados

🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

· Reportar issues
· Sugerir melhorias
· Enviar pull requests

📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo LICENSE para mais detalhes.

---

Nota: Este código foi desenvolvido para integração direta com Excel usando xlwings. Certifique-se de que as macros estão habilitadas no Excel para funcionamento correto.
