/**
 * generate-manual.js — Generates the user manual in .docx format.
 * Run: node generate-manual.js
 */
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, ImageRun, TabStopPosition, TabStopType,
  PageBreak, Header, Footer, convertInchesToTwip
} = require('docx');
const fs = require('fs');

// ===== CEDISA Brand Palette =====
const NAVY = '00123C';   // CEDISA navy (primary dark)
const BLUE = '186BB8';   // CEDISA blue (primary accent)
const ORANGE = 'E85C04'; // CEDISA orange (warning / brand)
const RED = 'DC3002';    // CEDISA red (danger)
const GRAY_DARK = '595959';
const GRAY = '8B8B8B';
const DARK = '00123C';   // Body text anchor — reuses navy
// Functional tones harmonized with CEDISA
const INDIGO = BLUE;     // legacy alias — points to CEDISA blue
const GREEN = '0E8F5C';  // success (kept green, tuned slightly darker for print)
const AMBER = ORANGE;    // warning alias

function title(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 48, color: INDIGO, font: 'Segoe UI' })],
    spacing: { after: 200 },
    alignment: AlignmentType.CENTER,
  });
}

function subtitle(text) {
  return new Paragraph({
    children: [new TextRun({ text, size: 24, color: GRAY, font: 'Segoe UI' })],
    spacing: { after: 400 },
    alignment: AlignmentType.CENTER,
  });
}

function heading(text, level = HeadingLevel.HEADING_1) {
  const size = level === HeadingLevel.HEADING_1 ? 32 : level === HeadingLevel.HEADING_2 ? 26 : 22;
  const color = level === HeadingLevel.HEADING_1 ? INDIGO : DARK;
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size, color, font: 'Segoe UI' })],
    heading: level,
    spacing: { before: 300, after: 150 },
  });
}

function para(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({
      text,
      size: opts.size || 22,
      color: opts.color || DARK,
      font: opts.font || 'Segoe UI',
      bold: opts.bold || false,
      italics: opts.italics || false,
    })],
    spacing: { after: opts.after || 120 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
  });
}

function richPara(runs, opts = {}) {
  return new Paragraph({
    children: runs.map(r => new TextRun({
      text: r.text,
      size: r.size || 22,
      color: r.color || DARK,
      font: r.font || 'Segoe UI',
      bold: r.bold || false,
      italics: r.italics || false,
      underline: r.underline ? {} : undefined,
    })),
    spacing: { after: opts.after || 120 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    bullet: opts.bullet ? { level: opts.bulletLevel || 0 } : undefined,
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, color: DARK, font: 'Segoe UI' })],
    bullet: { level },
    spacing: { after: 80 },
  });
}

function bulletRich(runs, level = 0) {
  return new Paragraph({
    children: runs.map(r => new TextRun({
      text: r.text,
      size: r.size || 22,
      color: r.color || DARK,
      font: r.font || 'Segoe UI',
      bold: r.bold || false,
      italics: r.italics || false,
    })),
    bullet: { level },
    spacing: { after: 80 },
  });
}

function tip(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: '💡 Dica: ', bold: true, size: 22, color: GREEN, font: 'Segoe UI' }),
      new TextRun({ text, size: 22, color: DARK, font: 'Segoe UI', italics: true }),
    ],
    spacing: { before: 100, after: 150 },
    indent: { left: convertInchesToTwip(0.3) },
    border: {
      left: { style: BorderStyle.SINGLE, size: 6, color: GREEN, space: 10 },
    },
  });
}

function warning(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: '⚠️ Atenção: ', bold: true, size: 22, color: AMBER, font: 'Segoe UI' }),
      new TextRun({ text, size: 22, color: DARK, font: 'Segoe UI', italics: true }),
    ],
    spacing: { before: 100, after: 150 },
    indent: { left: convertInchesToTwip(0.3) },
    border: {
      left: { style: BorderStyle.SINGLE, size: 6, color: AMBER, space: 10 },
    },
  });
}

function spacer() {
  return new Paragraph({ children: [], spacing: { after: 200 } });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ===== Build the document =====
const doc = new Document({
  creator: 'CEDISA — Central de Aço',
  title: 'Manual do Usuário - Conciliação Bancária',
  description: 'Manual de uso do sistema de conciliação bancária CEDISA — Banco x Fusion (Oracle)',
  styles: {
    default: {
      document: {
        run: { font: 'Segoe UI', size: 22 },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(1),
          bottom: convertInchesToTwip(0.8),
          left: convertInchesToTwip(1.2),
          right: convertInchesToTwip(1.2),
        },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'Manual do Usuário — Conciliação Bancária', size: 16, color: GRAY, font: 'Segoe UI', italics: true }),
            ],
            alignment: AlignmentType.RIGHT,
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: 'CEDISA — Central de Aço · Documento Interno', size: 16, color: GRAY, font: 'Segoe UI' }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      }),
    },
    children: [
      // ===== CAPA =====
      spacer(), spacer(), spacer(), spacer(),
      title('Conciliação Bancária'),
      subtitle('Banco × Fusion (Oracle)'),
      spacer(),
      para('CEDISA — Central de Aço', { size: 26, align: AlignmentType.CENTER, color: ORANGE, bold: true }),
      para('Manual do Usuário', { size: 28, align: AlignmentType.CENTER, color: GRAY }),
      spacer(), spacer(),
      para('Este manual explica, passo a passo, como usar o sistema de conciliação bancária da CEDISA para comparar os pagamentos registrados no Banco com os lançamentos do sistema Fusion (Oracle).', { align: AlignmentType.CENTER }),
      spacer(), spacer(), spacer(),
      para(`Versão 2.0 — ${new Date().toLocaleDateString('pt-BR')}`, { align: AlignmentType.CENTER, color: GRAY, italics: true }),
      para('CEDISA — Central de Aço', { align: AlignmentType.CENTER, bold: true, size: 24, color: NAVY }),

      pageBreak(),

      // ===== SUMÁRIO =====
      heading('Sumário'),
      para('1. O que é a Conciliação Bancária?'),
      para('2. Como acessar o sistema'),
      para('3. Passo 1 — Upload dos arquivos'),
      para('4. Passo 2 — Padronização De-Para'),
      para('5. Passo 3 — Resultados da Conciliação'),
      para('6. Sugestões por Valor (validação manual)'),
      para('7. Dicionário De-Para (barra lateral)'),
      para('8. Exportar para Excel'),
      para('9. Como o sistema forma os pares (entender os níveis)'),
      para('10. Perguntas Frequentes'),

      pageBreak(),

      // ===== 1. O QUE É =====
      heading('1. O que é a Conciliação Bancária?'),
      para('A conciliação bancária é o processo de comparar dois arquivos:'),
      bulletRich([
        { text: 'Arquivo do Banco (DDA): ', bold: true },
        { text: 'Lista de pagamentos que saíram da conta bancária.' },
      ]),
      bulletRich([
        { text: 'Arquivo do Fusion (DIA/Oracle): ', bold: true },
        { text: 'Lista de pagamentos registrados no sistema ERP da empresa.' },
      ]),
      spacer(),
      para('O objetivo é verificar se todos os pagamentos feitos pelo banco estão corretamente registrados no Fusion, e vice-versa. O sistema faz essa comparação automaticamente e mostra:'),
      bulletRich([
        { text: 'Verde (Conciliado): ', bold: true, color: GREEN },
        { text: 'O valor bate entre banco e Fusion. Está tudo certo!' },
      ]),
      bulletRich([
        { text: 'Amarelo (Sugerido): ', bold: true, color: AMBER },
        { text: 'O valor é igual, mas os nomes dos fornecedores são diferentes. Precisa de uma conferência manual.' },
      ]),
      bulletRich([
        { text: 'Pendente: ', bold: true, color: RED },
        { text: 'Não foi possível encontrar correspondência automática. Precisa de análise.' },
      ]),

      pageBreak(),

      // ===== 2. COMO ACESSAR =====
      heading('2. Como acessar o sistema'),
      para('O sistema funciona pelo navegador de internet (Chrome, Edge, etc). Para acessar:'),
      bullet('Certifique-se de que o servidor está ligado (peça ao time de TI se necessário).'),
      bullet('Abra o navegador e digite o endereço: http://localhost:3000/'),
      bullet('A tela inicial de upload será exibida automaticamente.'),
      spacer(),
      warning('O sistema precisa de dois servidores rodando em segundo plano: o servidor de banco de dados (porta 3001) e o servidor web (porta 3000). Se a página não abrir, verifique com o time de TI.'),

      pageBreak(),

      // ===== 3. PASSO 1: UPLOAD =====
      heading('3. Passo 1 — Upload dos Arquivos'),
      para('Na tela inicial você verá dois quadrantes lado a lado:'),
      spacer(),

      heading('Quadrante Esquerdo — Arquivo do Banco', HeadingLevel.HEADING_2),
      bullet('Clique em "Selecionar Arquivo" ou arraste o arquivo do banco para dentro do quadrante.'),
      bullet('O arquivo deve ser do tipo Excel (.xlsx ou .xls).'),
      bullet('Geralmente é o arquivo DDA exportado do sistema do Banco Safra.'),
      spacer(),

      heading('Quadrante Direito — Arquivo do Fusion', HeadingLevel.HEADING_2),
      bullet('Clique em "Selecionar Arquivo" ou arraste o arquivo do Fusion para dentro do quadrante.'),
      bullet('O arquivo deve ser do tipo Excel (.xlsx ou .xls).'),
      bullet('Geralmente é o relatório DIA exportado do Oracle Fusion.'),
      spacer(),

      para('Após anexar os dois arquivos, o botão "Avançar para Padronização" ficará habilitado (roxo). Clique nele para ir ao próximo passo.'),
      spacer(),

      tip('Se você selecionou o arquivo errado, clique no botão "✕" vermelho ao lado do nome do arquivo para removê-lo e selecionar outro.'),

      pageBreak(),

      // ===== 4. PASSO 2: DE-PARA =====
      heading('4. Passo 2 — Padronização De-Para'),
      para('Esta é uma etapa muito importante! Os nomes dos fornecedores no extrato do Banco costumam ser diferentes (abreviados, sem acentos, com siglas) dos nomes cadastrados no Fusion. Por exemplo:'),
      spacer(),

      new Table({
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: 'Nome no Banco', bold: true, size: 22, font: 'Segoe UI', color: 'FFFFFF' })], alignment: AlignmentType.CENTER })],
                shading: { type: ShadingType.SOLID, fill: INDIGO, color: INDIGO },
                width: { size: 50, type: WidthType.PERCENTAGE },
              }),
              new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: 'Nome no Fusion (correto)', bold: true, size: 22, font: 'Segoe UI', color: 'FFFFFF' })], alignment: AlignmentType.CENTER })],
                shading: { type: ShadingType.SOLID, fill: GREEN, color: GREEN },
                width: { size: 50, type: WidthType.PERCENTAGE },
              }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [para('AIR LIQUIDE BRASIL L', { size: 20 })] }),
              new TableCell({ children: [para('AIR LIQUIDE BRASIL LTDA', { size: 20 })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [para('ALPE FUNDO DE INVESTIMENTO EM', { size: 20 })] }),
              new TableCell({ children: [para('ARCELORMITTAL BRASIL S A', { size: 20 })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [para('COMP ELE EST DA BA COELB', { size: 20 })] }),
              new TableCell({ children: [para('COMPANHIA DE ELETRICIDADE DO ESTADO DA BAHIA', { size: 20 })] }),
            ],
          }),
        ],
        width: { size: 100, type: WidthType.PERCENTAGE },
      }),

      spacer(),

      heading('O que você verá nesta tela:', HeadingLevel.HEADING_2),
      para('Uma tabela com 5 colunas:'),
      bulletRich([
        { text: 'Nome no Banco (abreviado): ', bold: true },
        { text: 'O nome do fornecedor como aparece no extrato bancário.' },
      ]),
      bulletRich([
        { text: 'Lanç. / Valor: ', bold: true },
        { text: 'Quantos lançamentos e o valor total desse fornecedor no banco.' },
      ]),
      bulletRich([
        { text: '→ Nome no Fusion (cadastrado): ', bold: true },
        { text: 'O nome correto do fornecedor no Oracle. Se estiver pendente, um seletor aparecerá para você escolher.' },
      ]),
      bulletRich([
        { text: 'Lanç. / Valor: ', bold: true },
        { text: 'Quantos lançamentos e o valor total desse fornecedor no Fusion.' },
      ]),
      bulletRich([
        { text: 'Status: ', bold: true },
        { text: 'Mostra se já foi mapeado (verde), auto (azul) ou pendente (amarelo).' },
      ]),

      spacer(),

      heading('Como fazer o mapeamento:', HeadingLevel.HEADING_2),
      bullet('Para fornecedores com status "Pendente", clique no seletor da coluna "Nome no Fusion".'),
      bullet('Procure e selecione o nome correto do fornecedor na lista.'),
      bullet('O sistema pode sugerir um nome automaticamente (aparece com ⭐). Verifique se a sugestão está correta.'),
      bullet('Ao selecionar, o mapeamento é salvo automaticamente no banco de dados. Na próxima vez que você usar o sistema, ele já saberá fazer essa correspondência.'),
      spacer(),

      tip('Use os filtros "Todos", "Não Mapeados" e "Mapeados" para facilitar a visualização. Você também pode pesquisar pelo nome do fornecedor na barra de busca.'),
      warning('Mesmo que haja fornecedores pendentes (sem mapeamento), você pode prosseguir para a conciliação. Eles aparecerão como não conciliados nos resultados.'),

      spacer(),
      para('Quando estiver pronto, clique em "Executar Conciliação" para ir ao próximo passo.'),

      pageBreak(),

      // ===== 5. PASSO 3: RESULTADOS =====
      heading('5. Passo 3 — Resultados da Conciliação'),
      para('Nesta tela você verá o resultado final da conciliação. No topo, chips coloridos mostram o resumo:'),
      bulletRich([
        { text: 'Conciliados (verde): ', bold: true, color: GREEN },
        { text: 'Fornecedores cujos valores batem perfeitamente entre banco e Fusion.' },
      ]),
      bulletRich([
        { text: 'Sugeridos (amarelo): ', bold: true, color: AMBER },
        { text: 'Possíveis correspondências encontradas por valor — precisam da sua aprovação.' },
      ]),
      bulletRich([
        { text: 'Pendentes (laranja): ', bold: true },
        { text: 'Fornecedores que não foram conciliados (falta mapeamento ou valores diferentes).' },
      ]),
      bulletRich([
        { text: 'Diferença Total: ', bold: true },
        { text: 'A soma de todas as diferenças encontradas.' },
      ]),

      spacer(),

      heading('Tabela de Resultados', HeadingLevel.HEADING_2),
      para('Cada linha mostra um fornecedor com:'),
      bullet('▶ Botão de expandir — clique para ver os lançamentos individuais do banco e do Fusion.'),
      bullet('Fornecedor — Nome do fornecedor.'),
      bullet('Valor Banco — Total pago no banco para esse fornecedor.'),
      bullet('Valor Fusion — Total registrado no Fusion para esse fornecedor.'),
      bullet('Status — Conciliado ou Pendente.'),
      bullet('Diferença — A diferença entre banco e Fusion (idealmente deve ser zero).'),

      spacer(),

      heading('Visualizando os lançamentos', HeadingLevel.HEADING_2),
      para('Ao clicar no botão ▶ ao lado de um fornecedor, uma área expansível mostra todos os lançamentos individuais, divididos em duas colunas:'),
      bullet('📄 Lançamentos Banco — Cada pagamento individual do banco, com documento e valor.'),
      bullet('🔗 Lançamentos Fusion — Cada lançamento do Fusion, com número da NF e valor.'),
      spacer(),
      para('Os lançamentos aparecem marcados em:'),
      bulletRich([
        { text: 'Verde (borda esquerda): ', bold: true, color: GREEN },
        { text: 'O lançamento encontrou correspondência no outro lado.' },
      ]),
      bulletRich([
        { text: 'Amarelo (borda esquerda): ', bold: true, color: AMBER },
        { text: 'O lançamento não encontrou correspondência exata.' },
      ]),

      pageBreak(),

      // ===== 6. SUGESTÕES POR VALOR =====
      heading('6. Sugestões por Valor (validação manual)'),
      para('Quando o sistema encontra fornecedores com valores iguais no banco e no Fusion, mas com nomes diferentes, ele cria uma seção especial chamada "Sugestões por Valor".'),
      spacer(),
      para('Por exemplo: no banco aparece "COMP ELE EST DA BA COELB" com R$ 946,00 e no Fusion aparece "COMPANHIA DE ELETRICIDADE DO ESTADO DA BAHIA" com R$ 946,00. Como os valores são iguais, o sistema sugere que podem ser o mesmo fornecedor.'),
      spacer(),

      heading('Tipos de sugestão', HeadingLevel.HEADING_2),
      bulletRich([
        { text: 'Por total do fornecedor: ', bold: true, color: BLUE },
        { text: 'a soma de todos os lançamentos do banco bate com a soma dos lançamentos do Fusion de outro fornecedor.' },
      ]),
      bulletRich([
        { text: 'Por valor de linha: ', bold: true, color: BLUE },
        { text: 'um ou mais lançamentos individuais têm valores coincidentes, mesmo que os totais não batam.' },
      ]),
      bulletRich([
        { text: 'Dentro de grupo parcial (novo): ', bold: true, color: ORANGE },
        { text: 'quando o valor do fornecedor "somente banco" bate com lançamentos que ficaram SEM par dentro de um grupo PENDENTE. A sugestão aparece com o aviso "⚠ Lançamentos sem par dentro deste grupo".' },
      ]),
      spacer(),
      para('Exemplo prático do último caso: no banco há "ALPE FUNDO DE INVESTIMENTO EM" com R$ 111.346,75 (fornecedor novo, sem mapeamento). No Fusion, dentro do grupo de ARCELORMITTAL BRASIL S A existe um lançamento de R$ 111.346,75 que não bateu com nenhum lançamento banco da ARCELORMITTAL. O sistema detecta esse "órfão" e sugere o par.'),
      spacer(),

      heading('Como usar:', HeadingLevel.HEADING_2),
      bullet('Analise cada sugestão: verifique se o fornecedor do banco realmente corresponde ao do Fusion.'),
      bulletRich([
        { text: 'Clique em "✓ Aprovar" ', bold: true, color: GREEN },
        { text: 'se a correspondência estiver correta. O mapeamento será salvo automaticamente para próximas análises.' },
      ]),
      bulletRich([
        { text: 'Clique em "✕ Recusar" ', bold: true, color: RED },
        { text: 'se os fornecedores não forem os mesmos, mesmo com valores iguais. A recusa também é registrada — essa sugestão não aparecerá mais.' },
      ]),
      spacer(),

      tip('Ao aprovar uma sugestão, o mapeamento é salvo no Dicionário De-Para. Isso significa que na próxima vez que você fizer uma conciliação, esse fornecedor será reconhecido automaticamente!'),
      tip('Para desfazer uma recusa, abra o Dicionário De-Para e vá na aba "Recusadas" — ali você pode restaurar sugestões descartadas por engano.'),

      pageBreak(),

      // ===== 7. DICIONÁRIO DE-PARA =====
      heading('7. Dicionário De-Para (barra lateral)'),
      para('O Dicionário De-Para é um banco de dados que guarda todas as associações entre nomes do banco e nomes do Fusion. Ele fica acessível a qualquer momento através do botão "Dicionário De-Para" no canto superior direito da tela.'),
      spacer(),

      heading('Abrindo a barra lateral:', HeadingLevel.HEADING_2),
      bullet('Clique no botão "📖 Dicionário De-Para" no cabeçalho.'),
      bullet('Uma barra lateral desliza da direita mostrando todos os mapeamentos salvos.'),
      bullet('Um pontinho verde no botão indica que já existem mapeamentos salvos.'),
      spacer(),

      heading('O que você pode fazer:', HeadingLevel.HEADING_2),
      bulletRich([
        { text: 'Adicionar novo mapeamento: ', bold: true },
        { text: 'Preencha os campos "Nome no Banco" e "Nome no Fusion" e clique em "Adicionar".' },
      ]),
      bulletRich([
        { text: 'Pesquisar: ', bold: true },
        { text: 'Use a barra de busca para encontrar mapeamentos existentes.' },
      ]),
      bulletRich([
        { text: 'Editar: ', bold: true },
        { text: 'Passe o mouse sobre um mapeamento e clique no ícone ✎ para alterar o nome do Fusion.' },
      ]),
      bulletRich([
        { text: 'Excluir: ', bold: true },
        { text: 'Passe o mouse sobre um mapeamento e clique no ícone ✕ para removê-lo.' },
      ]),
      spacer(),

      warning('Os mapeamentos são permanentes e ficam salvos no sistema. Uma vez criado, ele será aplicado automaticamente em todas as próximas conciliações. Exclua apenas se tiver certeza de que o mapeamento está errado.'),

      pageBreak(),

      // ===== 8. EXPORTAR =====
      heading('8. Exportar para Excel'),
      para('Na tela de resultados, clique no botão "Exportar Excel" para baixar um arquivo com todos os dados da conciliação. O arquivo agora é detalhado por cliente, com layout similar ao que você vê em tela.'),
      spacer(),

      heading('Abas do arquivo', HeadingLevel.HEADING_2),
      bulletRich([
        { text: 'Aba "Pendentes": ', bold: true, color: ORANGE },
        { text: 'contém apenas o que ficou SEM par. Se um fornecedor foi parcialmente conciliado, apenas os lançamentos órfãos aparecem aqui (ex.: o R$ 68.131,38 do banco e o R$ 13.745,44 do Fusion que não bateram com ninguém).' },
      ]),
      bulletRich([
        { text: 'Aba "Conciliados": ', bold: true, color: GREEN },
        { text: 'contém todos os pares formados — inclusive os pares 1:1, N:1 e 1:N encontrados dentro de grupos que, no geral, continuam pendentes. Esses blocos aparecem marcados como "CONCILIADO (parcial)".' },
      ]),
      bulletRich([
        { text: 'Aba "Sugestões": ', bold: true, color: BLUE },
        { text: 'lista as sugestões por valor e seu estado (SUGERIDO, APROVADO, RECUSADO).' },
      ]),
      spacer(),

      heading('Layout de cada bloco de cliente', HeadingLevel.HEADING_2),
      para('Cada fornecedor vira um bloco com:'),
      bullet('Linha cinza escuro "CLIENTE: <nome>" e badge colorido com o status (verde = conciliado, laranja = pendente).'),
      bullet('Linha de totais: Total Banco (fundo azul claro) e Total Fusion (fundo verde claro), mais a Diferença (fundo vermelho claro).'),
      bullet('Cabeçalho das colunas: BANCO em azul, FUSION em verde.'),
      bullet('Lançamentos lado a lado — à esquerda (azul) os lançamentos banco, à direita (verde) os lançamentos Fusion.'),
      bullet('Linha em branco separa um cliente do próximo.'),
      spacer(),

      heading('Formato dos valores', HeadingLevel.HEADING_2),
      bullet('Todas as células monetárias são NÚMERO com formato "R$ 1.234,56" — permite somar, filtrar e aplicar fórmulas no Excel.'),
      bullet('Valores negativos aparecem em vermelho automaticamente.'),
      bullet('As colunas de banco (A e B) têm fundo azul claro; as de Fusion (C e D) têm fundo verde claro — fica fácil diferenciar visualmente.'),
      spacer(),
      tip('O arquivo é salvo automaticamente na pasta de Downloads do seu computador com o nome "conciliacao_AAAA-MM-DD.xlsx".'),
      warning('Se um bloco aparece na aba "Conciliados" com o rótulo "(parcial)", significa que parte do fornecedor foi resolvida mas parte permanece pendente — procure o MESMO fornecedor na aba "Pendentes" para ver o que ficou em aberto.'),

      pageBreak(),

      // ===== 9. COMO O SISTEMA FORMA OS PARES =====
      heading('9. Como o sistema forma os pares (entender os níveis)'),
      para('Para conciliar, o sistema agrupa primeiro os lançamentos por fornecedor (aplicando o De-Para) e depois tenta formar pares em 4 níveis de rigor crescente:'),
      spacer(),

      heading('Nível 1 — Soma do grupo bate', HeadingLevel.HEADING_2),
      para('Se a soma de todos os lançamentos banco de um fornecedor é igual à soma dos lançamentos Fusion desse mesmo fornecedor (tolerância de R$ 0,015), o fornecedor é marcado como CONCILIADO imediatamente.'),
      spacer(),

      heading('Nível 2 — Emparelhamento por subconjuntos (N:1 e 1:N)', HeadingLevel.HEADING_2),
      para('Quando as somas não batem, o sistema procura combinações:'),
      bullet('Passo 1 — pares 1:1: um lançamento banco com o mesmo valor de um lançamento Fusion.'),
      bullet('Passo 2 — N:1: vários lançamentos banco que somados batem com um único lançamento Fusion.'),
      bullet('Passo 3 — 1:N: um lançamento banco que bate com a soma de vários lançamentos Fusion.'),
      para('Cada par encontrado aqui vai para a aba "Conciliados" do Excel, mesmo que o grupo, no geral, continue pendente.'),
      spacer(),

      heading('Nível 3 — Correspondência entre fornecedores diferentes', HeadingLevel.HEADING_2),
      para('Quando sobra fornecedor "somente no banco" ou "somente no Fusion", o sistema tenta parear pelo VALOR, mesmo com nomes diferentes. Geram-se sugestões nos seguintes casos:'),
      bullet('Valor total do banco = valor total do Fusion de outro fornecedor.'),
      bullet('Valor de linhas individuais coincidem.'),
      bulletRich([
        { text: 'NOVO: ', bold: true, color: ORANGE },
        { text: 'o valor do fornecedor "somente banco" bate com lançamentos ÓRFÃOS que ficaram sem par dentro de um grupo PENDENTE (ex.: ALPE FUNDO ↔ lançamento solto dentro de ARCELORMITTAL).' },
      ]),
      para('Todas essas sugestões aparecem na tela de resultados com os botões ✓ Aprovar / ✕ Recusar.'),
      spacer(),

      heading('Nível 4 — Sem correspondência', HeadingLevel.HEADING_2),
      para('O que não se enquadra nos níveis anteriores é marcado como PENDENTE. Aparece no Excel na aba "Pendentes" para análise manual.'),

      pageBreak(),

      // ===== 10. PERGUNTAS FREQUENTES =====
      heading('10. Perguntas Frequentes'),
      spacer(),

      richPara([{ text: 'P: Preciso mapear TODOS os fornecedores antes de conciliar?', bold: true }]),
      para('R: Não. Você pode conciliar a qualquer momento. Fornecedores que não foram mapeados aparecerão como "Pendentes" nos resultados. Você pode voltar, mapear mais fornecedores e conciliar novamente.'),
      spacer(),

      richPara([{ text: 'P: O mapeamento que eu fiz serve para todas as próximas análises?', bold: true }]),
      para('R: Sim! Uma vez que você mapeia "AIR LIQUIDE BRASIL L" → "AIR LIQUIDE BRASIL LTDA", todas as futuras conciliações já usarão esse mapeamento automaticamente.'),
      spacer(),

      richPara([{ text: 'P: Aprovei uma sugestão por valor errada. Como desfaço?', bold: true }]),
      para('R: Abra o Dicionário De-Para (botão no canto superior direito), procure o mapeamento e clique no ícone ✕ para excluí-lo.'),
      spacer(),

      richPara([{ text: 'P: O sistema está dizendo que não encontrou dados no arquivo. O que faço?', bold: true }]),
      para('R: Verifique se o arquivo está no formato correto (.xlsx). O sistema espera que o arquivo do banco tenha uma coluna "Favorecido / Beneficiário" e o do Fusion tenha "Fornecedor ou Parte". Se o layout do arquivo mudou, avise o time de TI.'),
      spacer(),

      richPara([{ text: 'P: Posso usar o sistema com arquivos de datas diferentes?', bold: true }]),
      para('R: Sim, basta selecionar os dois novos arquivos na tela de upload. Os mapeamentos do Dicionário De-Para continuam salvos independentemente dos arquivos usados.'),
      spacer(),

      richPara([{ text: 'P: O que significa a diferença total?', bold: true }]),
      para('R: É a soma de todas as diferenças entre banco e Fusion nos itens pendentes. Se for zero, significa que tudo está conciliado. Se for positivo, há mais saídas no banco do que no Fusion; se negativo, há mais registros no Fusion do que no banco.'),
      spacer(),

      richPara([{ text: 'P: A página não abre / mostra erro.', bold: true }]),
      para('R: Certifique-se de que o servidor está rodando. Acione o time de TI para verificar se os serviços estão ativos nas portas 3000 e 3001.'),
      spacer(),

      richPara([{ text: 'P: Um fornecedor aparece ao mesmo tempo em "Conciliados" (parcial) e em "Pendentes" no Excel. Por quê?', bold: true }]),
      para('R: Isso é esperado quando parte dos lançamentos casou e parte não. Na aba Conciliados você vê os pares formados; na aba Pendentes vê apenas o que ficou sem par. Some os dois blocos para ter a visão total do fornecedor.'),
      spacer(),

      richPara([{ text: 'P: Vi uma sugestão com o aviso "Lançamentos sem par dentro deste grupo". O que significa?', bold: true }]),
      para('R: Significa que o sistema encontrou, dentro de um grupo pendente (ex.: ARCELORMITTAL), um ou mais lançamentos que não bateram com nenhum pagamento banco daquele grupo, e cujo valor coincide com o fornecedor "somente banco" da sugestão. Ao aprovar, o De-Para é salvo e, na próxima rodada, o fornecedor será mesclado corretamente.'),
      spacer(),

      richPara([{ text: 'P: Existem dois lançamentos banco com o mesmo valor (ex.: R$ 111.346,75 na ARCELORMITTAL e R$ 111.346,75 na ALPE FUNDO). Como o sistema diferencia?', bold: true }]),
      para('R: Depois que o De-Para é aplicado, os dois lançamentos pertencem ao mesmo grupo. O sistema tenta emparelhar 1:1 com lançamentos Fusion de mesmo valor — se houver dois R$ 111.346,75 no Fusion, os dois casam; se houver só um, apenas um casa e o outro vai para Pendentes. Use a aba "Sugestões" para decidir casos ambíguos.'),
      spacer(),

      richPara([{ text: 'P: As cores do sistema mudaram. Por quê?', bold: true }]),
      para('R: O visual foi alinhado ao brand book da CEDISA (Central de Aço): azul marinho #00123C, azul #186BB8 e laranja #E85C04 passaram a ser as cores principais, e a tipografia foi ajustada para um padrão mais industrial.'),

      spacer(), spacer(),
      new Paragraph({
        children: [
          new TextRun({ text: '— Fim do Manual —', bold: true, size: 24, color: GRAY, font: 'Segoe UI', italics: true }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 400 },
      }),
    ],
  }],
});

// Generate
Packer.toBuffer(doc).then(buffer => {
  const path = 'Manual_Conciliacao_Bancaria.docx';
  fs.writeFileSync(path, buffer);
  console.log(`✅ Manual gerado com sucesso: ${path}`);
  console.log(`   Tamanho: ${(buffer.length / 1024).toFixed(1)} KB`);
});
