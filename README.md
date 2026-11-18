# Gerador de vinhetas clínicas

Transforma uma planilha de casos codificados (variáveis A a J) em vinhetas clínicas padronizadas, usando um dicionário de códigos. O texto é montado por regras fixas, sem IA: o mesmo código gera sempre o mesmo texto.

Feito para pesquisas que comparam hipóteses diagnósticas de ferramentas de IA a partir de casos odontológicos padronizados.

## Gerador web

**https://bbiamar.github.io/gerador-de-vinhetas/**

1. Baixe o modelo em branco na página.
2. Preencha a aba `dados_codificados`: uma linha por caso.
3. Envie a planilha na página. As vinhetas aparecem na hora, com botões para copiar e para baixar em `.txt` ou `.xlsx`.

A planilha é lida no seu navegador e não é enviada a nenhum servidor.

## Script em Python (alternativa)

Requer Python 3.

```bash
pip install openpyxl
python3 gerar_vinhetas.py minha_planilha.xlsx
```

Com só uma planilha na pasta do script, basta `python3 gerar_vinhetas.py`. O script cria `<planilha>_com_vinhetas.xlsx`, com as colunas **Vinheta** e **Status** preenchidas, e `<planilha>_com_vinhetas.txt`, com comando + vinheta de cada caso. A planilha original não é alterada.

## Arquivos

| Arquivo | O que é |
|---|---|
| `index.html` | Gerador web (página única, funciona no GitHub Pages) |
| `gerar_vinhetas.py` | O mesmo gerador em Python |
| `modelo_vinhetas.xlsx` | Modelo em branco: instruções, dados, dicionário de códigos e configuração |

## Preenchimento

| Variável | Conteúdo | Códigos |
|---|---|---|
| A | Sexo | 1 feminino, 2 masculino (um só) |
| B | Idade | anos, inteiro de 18 a 99 |
| C | Sintomatologia | vários, separados por `;` |
| D | Apresentação clínica | vários |
| E | Localização | vários |
| F | Fatores locais | vários |
| G | Fatores sistêmicos e medicamentos | vários |
| H | Radiografia | 1 não tem, 2 tem (um só) |
| I, J | Diagnóstico radiográfico e clínico | gabarito: conferidos, mas fora da vinheta |

Vários códigos na mesma célula vão separados por ponto e vírgula (`3;4`). Não use vírgula: o Excel em português transforma `3,4` em número decimal.

Cada caso recebe uma situação: **OK** (completo), **ATENÇÃO** (vinheta gerada, mas falta dado) ou **ERRO** (código inexistente ou mal digitado; sem vinheta).

Códigos novos podem ser incluídos na aba `dicionario_de_codigos`, desde que usem um dos subgrupos já existentes. A lista está na aba Instruções do modelo.

## Cuidados com os dados (LGPD)

- Dados de saúde são dados pessoais sensíveis (LGPD, art. 11). Use somente em pesquisa aprovada pelo CEP e conforme o que ele exigir.
- **Nunca envie planilhas preenchidas para este repositório.** O `.gitignore` bloqueia arquivos `.xlsx` além do modelo, mas confira antes de cada envio.
- No ID, use códigos como P01. Nunca nome, iniciais, prontuário ou data de nascimento.
- Radiografias: remova nome, data ou clínica gravados na imagem e os metadados do arquivo antes de enviar a qualquer serviço.
- Enviar vinhetas a uma IA (ChatGPT, Gemini etc.) é compartilhar dados com terceiros, muitas vezes fora do Brasil. Isso deve constar no projeto aprovado.
