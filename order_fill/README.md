# Atualização semanal do Order Fill

Script que atualiza o arquivo **Order Fill por Base** com os dados do mês corrente.

**Order Fill = Faturamento ÷ (Faturamento + Ruptura)**, em quantidade.

## Instalação (só na primeira vez)

1. Instale o Python: https://www.python.org/downloads/ (marque a opção **"Add Python to PATH"**).
2. Abra o Prompt de Comando e rode:

```bash
py -m pip install -r requirements.txt
```

## Toda semana

1. Coloque **na mesma pasta do script** (`order_fill`):
   - o arquivo Order Fill mais recente (o script sempre usa o mais novo da pasta);
   - o **de-para** de códigos (nome com "de-para");
   - o **faturamento** de cada marca exportado do Power BI: nomes com `Fat` e `BK`, `PLK` ou `SBX`
     (ex.: `Fat_Nov_BK.xlsx`). Abra cada um no Excel, clique em **Dados → Atualizar tudo** e salve antes;
   - a **ruptura** de cada marca: nomes com `Ruptura` e `BK`, `PLK` ou `SBX`.

   Faturamento e ruptura devem cobrir **do dia 01 do mês até a data de corte**.
2. Dê dois cliques em `atualizar_order_fill.bat`.
3. O resultado é salvo na mesma pasta como `Order_Fill_por_Base_<ano>_ate_<dd-mm>.xlsx`.
   O arquivo original não é alterado.

A data de corte é o último dia que tem **faturamento e ruptura** ao mesmo tempo.
Para escolher outra data:

```bash
py atualizar_order_fill.py --ate 10/11/2026
```

Para usar outra pasta:

```bash
py atualizar_order_fill.py --pasta "C:\caminho\da\pasta"
```

## O que o script faz

| Aba | Atualização |
|---|---|
| BK, PLK, SBX (mensal) | Cria ou atualiza a linha do mês atual |
| BK, PLK, SBX (semanal) | Mantém as semanas do mês anterior e refaz as do mês atual |
| Top Cortes BK / PLK / SBX | Mantém o mês anterior e refaz o top 10 de cortes das semanas do mês atual |
| Ofensores CD Z05 | Mesmo que Top Cortes, só com o CD Z05 |
| Todos Food | Refaz com o acumulado do mês dos itens "Food" do SBX |

## Regras usadas

- **Ativos e Linha**: no de-para, `Item ativo? = Ativo` e `Classificação = Linha`.
- **NBO**: itens de Ativos e Linha com `NBO = NBO`.
- **Bases**

  | CD | Base | Tipo |
  |---|---|---|
  | XK4 / XE4 | JDI | Próprio |
  | XKF / XE6 | R3 | Próprio |
  | XKA | ARA | Próprio |
  | Z14 | JDI | Compra e Venda |
  | Z17 | R3 | Compra e Venda |
  | Z10 | ARA | Compra e Venda |
  | Z05 | JP | Compra e Venda |

  Para mudar, edite `PROPRIO`, `COMPRA_VENDA` e `ABAS` no início do script.
- **Colunas JDI/R3/ARA gerais**: média entre Próprio e Compra e Venda (como já era na planilha).
  Se uma base não teve faturamento nem ruptura no período, aparece `n/d` e a média usa só a outra.
- **Semanas**: de segunda a domingo, cortadas na virada do mês. Se o mês começa num domingo,
  esse dia entra junto com a semana seguinte (ex.: 01/11 a 03/11/2026 = Semana 45).
- Itens que não estão no de-para ficam fora do cálculo; o script lista esses itens no final.
