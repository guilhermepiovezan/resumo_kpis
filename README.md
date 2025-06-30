# Resumo KPIs - Automação Google Sheets (Google Apps Script)

## Descrição

Este projeto consiste em um script Google Apps Script para automatizar a consolidação e o resumo de KPIs (Indicadores de Performance) em uma única aba dentro do Google Sheets. 

O script processa dados mensais de diferentes abas, calcula movimentos como desligamentos, férias, licenças, entradas, mudanças de HUB e retornos de afastamento, além de consolidar os dados nacionais e por regionais.

---

## Funcionalidades principais

- Limpeza automática da aba `Resumo_KPIs` para atualizar os dados sem resíduos antigos.
- Consolidação dos dados nacionais com cálculo de movimentos e contagem de colaboradores ativos por cargo/quadro.
- Consolidação dos dados por regionais com tratamento específico para variações de nomes.
- Mapeamento genérico de cargos para quadros (e.g., Executivo 1, Team Leader).
- Normalização de nomes de regionais para garantir consistência nos relatórios.
- Compatível com abas mensais nomeadas no formato `Dados_YYYYMM` (exemplo: Dados_202501 para Janeiro de 2025).

---

## Requisitos

- Google Sheets contendo abas mensais nomeadas conforme o padrão `Dados_YYYYMM`.
- Aba chamada `Resumo_KPIs` onde o resumo será gerado.
- Editor de scripts do Google Sheets para adicionar e executar o script.

---

## Como usar

1. Abra a planilha Google Sheets que contém os dados mensais e a aba de resumo.
2. Vá em **Extensões > Apps Script**.
3. Copie e cole o código do arquivo `resumo-kpis.gs` no editor do Apps Script.
4. Salve o projeto.
5. Execute a função principal `resumoBasicoERegional2025` para gerar o resumo dos KPIs.
6. Verifique a aba `Resumo_KPIs` para visualizar o resultado consolidado.

---

## Estrutura do Código

- **Função principal:** `resumoBasicoERegional2025()`  
  Controla o fluxo da limpeza da aba de resumo, coleta e consolida os dados nacionais e regionais.
  
- **Função auxiliar:** `normalizarRegional(nome, idxMes)`  
  Normaliza nomes regionais para nomes padronizados, considerando variações e regras específicas por mês.
  
- **Função auxiliar:** `listarRegionais()`  
  Retorna a lista de regionais padrão usadas no projeto.

- **Variáveis principais:**  
  - `meses`: Array com as informações das abas mensais e seus nomes.
  - `movimentos`: Lista dos tipos de movimentos monitorados (Desligamento, Férias, etc).
  - `cargosMapeados`: Objeto que mapeia cargos para quadros genéricos.
  - `quadros`: Lista dos quadros que serão considerados.

---

## Sobre o Autor

Guilherme Piovezan - 2025  
Este script foi desenvolvido como exemplo para automação de relatórios de KPIs em Google Sheets utilizando Google Apps Script.

---

## Licença

MIT License - Consulte o arquivo LICENSE para mais detalhes.

---

## Contato

Para dúvidas, sugestões ou colaborações, entre em contato pelo GitHub ou via email.

---

Obrigado por usar o projeto!

