/* export.js — Excel (SheetJS) + PDF (jsPDF) export */
var Export = (function(){
  "use strict";
  var FC = {navy:'0B2545',green:'00B74A',red:'D32F2F',amb:'F57C00',blue:'1565C0',white:'FFFFFF',light:'F5F5F5',border:'D0D0D0'};
  var BRL = function(v){ return (v<0?'−':'')+'R$ '+Math.abs(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}); };
  var BRLi = function(v){ return (v<0?'−':'')+'R$ '+Math.abs(Math.round(v||0)).toLocaleString('pt-BR'); };
  var PCT = function(v){ return (v||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%'; };
  var NUM = function(v){ return (v||0).toLocaleString('pt-BR'); };

  function headerStyle(){return{font:{bold:true,color:{rgb:FC.white},sz:10,name:'Arial'},fill:{fgColor:{rgb:FC.navy}},alignment:{horizontal:'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:FC.border}}}}}
  function bodyStyle(align){return{font:{name:'Arial',sz:10},alignment:{horizontal:align||'left'},border:{bottom:{style:'thin',color:{rgb:'F0F0F0'}}}}}

  function addSheet(wb, name, headers, rows, colWidths){
    var ws = XLSX.utils.aoa_to_sheet([headers].concat(rows));
    if(colWidths) ws['!cols'] = colWidths.map(function(w){return{wch:w}});
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  function generateExcel(data, selected, processDate){
    var wb = XLSX.utils.book_new();
    if(selected.criticaResumo && data.critica){
      var c = data.critica;
      var rSum = [
        ['ACURACIDADE', PCT(c.acuracidade)],
        ['SKUs analisados', c.totalSKUs],
        ['SKUs sem divergência', c.okCount],
        ['SKUs com falta', c.faltaCount],
        ['SKUs com sobra', c.sobraCount],
        ['Valor das faltas', BRLi(c.totalFaltas)],
        ['Valor das sobras', BRLi(c.totalSobras)],
        ['Saldo líquido', BRLi(c.saldoLiquido)]
      ];
      var rCat = c.categorias.map(function(cat){
        return [cat.nome, cat.total, PCT(cat.acuracidade), BRLi(cat.faltaVal), BRLi(cat.sobraVal), BRLi(cat.saldo)];
      });
      var allRows = [['FORMULA CODE — CRÍTICA DO INVENTÁRIO'],['Processado em '+processDate],[]];
      allRows.push(['RESUMO GERAL']);
      rSum.forEach(function(r){allRows.push(r)});
      allRows.push([]);
      allRows.push(['RESULTADO POR CATEGORIA']);
      allRows.push(['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)']);
      rCat.forEach(function(r){allRows.push(r)});
      addSheet(wb, 'Crítica - Resumo', [], allRows, [22,14,14,16,16,16]);
    }
    if(selected.criticaDetalhe && data.critica){
      var hd = ['SKU','Descrição','Categoria','Qtd Sistema','Qtd Contada','Dif. Qtd','Custo Unit.','Dif. R$','Status'];
      var rd = data.critica.items.map(function(i){
        return [i.sku, i.descricao, i.categoria, i.qtdSistema, i.qtdContada, i.difQtd, i.custoUnit, i.difValor, i.status];
      });
      addSheet(wb, 'Crítica - Detalhado', hd, rd, [14,24,16,12,12,10,12,12,10]);
    }
    if(selected.ruptura && data.ruptura){
      var hr = ['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Qtd Loja','Venda Méd/Dia','Fat. Méd/Dia'];
      var rr = data.ruptura.items.map(function(i){
        return [i.sku,i.descricao,i.categoria,i.abc_valorVendido90||'C',i.abc_lucro90||'C',i.deposito,i.loja,Engine.round2(i.vendaMediaDia||0),BRL(i.fatMediaDia||0)];
      });
      addSheet(wb, 'Ruptura', hr, rr, [14,24,16,10,10,12,10,14,14]);
    }
    if(selected.dias && data.dias){
      var hde = ['SKU','Descrição','Categoria','Qtd Estoque','Venda Méd/Dia','Dias Estoque','Cobertura','Valor Estoque','ABC Fat.'];
      var rde = data.dias.items.map(function(i){
        return [i.sku,i.descricao,i.categoria,i.qtdEstoque,Engine.round2(i.vendaMediaDia||0),i.diasEstoque!==null?Engine.round2(i.diasEstoque):'—',i.faixa,BRL(i.valorEstoque),i.abcFat];
      });
      addSheet(wb, 'Dias de Estoque', hde, rde, [14,24,16,12,14,12,12,14,10]);
    }
    if(selected.abc && data.abc){
      var ha = ['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Estoque','Custo Unit.','Valor Investido','Fat. 90 dias','Lucro 90 dias'];
      var ra = data.abc.items.map(function(i){
        return [i.sku,i.descricao,i.categoria,i.abcFat,i.abcLucro,i.qtdEstoque,BRL(i.custoUnit),BRL(i.valorInvestido),BRL(i.fat90),BRL(i.lucro90)];
      });
      addSheet(wb, 'Investimento ABC', ha, ra, [14,24,16,10,10,12,12,14,14,14]);
    }
    if(selected.perda && data.perda){
      var hp = ['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Venda Méd/Dia','Fat. Méd/Dia','Lucro Méd/Dia','Perda Fat./Dia','Perda Lucro/Dia'];
      var rp = data.perda.items.map(function(i){
        return [i.sku,i.descricao,i.categoria,i.abcFat,i.abcLucro,i.qtdDeposito,i.vendaMediaDia,BRL(i.fatMediaDia),BRL(i.lucroMediaDia),BRL(i.perdaFatDia),BRL(i.perdaLucroDia)];
      });
      addSheet(wb, 'Projeção de Perda', hp, rp, [14,24,16,10,10,12,14,14,14,14,14]);
    }
    var out = XLSX.write(wb, {bookType:'xlsx',type:'array'});
    var blob = new Blob([out], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'auditoria_estoque_'+processDate.replace(/[\/ :]/g,'_')+'.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  }

  /* ========== PDF ========== */
  function generatePDF(reportType, data, processDate, logoDataUrl){
    var jsPDF = window.jspdf.jsPDF;
    var doc = new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
    var W = 210, H = 297, M = 15;
    var y = 0;
    function addHeader(){
      doc.setFillColor(11,37,69);
      doc.rect(0,0,W,18,('F'));
      if(logoDataUrl){try{doc.addImage(logoDataUrl,'PNG',M,3,36,12)}catch(e){}}
      doc.setFontSize(9); doc.setTextColor(255,255,255);
      doc.text('Auditoria de estoque', W-M, 8, {align:'right'});
      doc.setFontSize(7); doc.setTextColor(180,180,200);
      doc.text('Processado em '+processDate, W-M, 13, {align:'right'});
      y = 24;
    }
    function addFooter(pg){
      doc.setFontSize(7); doc.setTextColor(150,150,150);
      doc.text('Formula Code Tecnologia, Gestão e Automação', M, H-6);
      doc.text('Página '+pg, W-M, H-6, {align:'right'});
      doc.setDrawColor(200,200,200); doc.line(M, H-10, W-M, H-10);
    }
    function checkPage(need){
      if(y+need > H-18){doc.addPage();addHeader();addFooter(doc.getNumberOfPages());}
    }
    function title(txt){
      checkPage(12);
      doc.setFontSize(14); doc.setTextColor(11,37,69); doc.setFont(undefined,'bold');
      doc.text(txt, M, y); y+=6;
      doc.setFontSize(8); doc.setTextColor(150,150,150); doc.setFont(undefined,'normal');
      doc.text('Relatório gerado automaticamente pelo sistema Formula Code', M, y); y+=8;
    }
    function section(txt){
      checkPage(10);
      doc.setFontSize(11); doc.setTextColor(11,37,69); doc.setFont(undefined,'bold');
      doc.text(txt, M, y); y+=6; doc.setFont(undefined,'normal');
    }
    function autoT(head, body, opts){
      checkPage(20);
      doc.autoTable({startY:y,head:[head],body:body,margin:{left:M,right:M},
        headStyles:{fillColor:[11,37,69],fontSize:7,fontStyle:'bold',halign:'center'},
        bodyStyles:{fontSize:7},
        alternateRowStyles:{fillColor:[245,245,245]},
        styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},
        columnStyles:opts||{}
      });
      y = doc.lastAutoTable.finalY + 6;
    }
    function kpiRow(labels, values, colors){
      checkPage(18);
      var cw = (W-2*M)/labels.length;
      doc.setFillColor(245,245,245);
      doc.roundedRect(M, y-2, W-2*M, 16, 2, 2, 'F');
      for(var i=0;i<labels.length;i++){
        var x = M + i*cw + 4;
        doc.setFontSize(7); doc.setTextColor(150,150,150); doc.setFont(undefined,'bold');
        doc.text(labels[i], x, y+3);
        doc.setFontSize(12); doc.setFont(undefined,'bold');
        var c = colors[i]||[51,51,51];
        doc.setTextColor(c[0],c[1],c[2]);
        doc.text(values[i], x, y+10);
      }
      doc.setFont(undefined,'normal');
      y += 20;
    }

    addHeader(); addFooter(1);

    /* Generate based on reportType */
    if(reportType === 'critica'){
      var c = data.critica;
      title('Crítica do inventário — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(
        ['ACURACIDADE','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],
        [PCT(c.acuracidade), BRLi(c.totalFaltas), BRLi(c.totalSobras), BRLi(c.saldoLiquido)],
        [[0,183,74],[211,47,47],[245,124,0],[211,47,47]]
      );
      if(c.hasCategorias){
        section('Resultado por categoria');
        autoT(['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],
          c.categorias.map(function(cat){return [cat.nome, cat.total, PCT(cat.acuracidade), BRLi(cat.faltaVal), BRLi(cat.sobraVal), BRLi(cat.saldo)];}),
          {0:{halign:'left'},1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}}
        );
      }
      section('Maiores divergências (top 20)');
      var top20 = c.items.slice().sort(function(a,b){return a.difValor-b.difValor}).slice(0,20);
      autoT(['SKU','Descrição','Categoria','Dif. Qtd','Dif. R$'],
        top20.map(function(i){return [i.sku,i.descricao,i.categoria,i.difQtd,BRL(i.difValor)];}),
        {0:{halign:'left'},1:{halign:'left'},2:{halign:'left'},3:{halign:'right'},4:{halign:'right'}}
      );
    }
    else if(reportType === 'ruptura'){
      var r = data.ruptura;
      title('Ruptura — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(
        ['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CLASSE A (FAT.)','RUPTURA CLASSE A (LUCRO)'],
        [PCT(r.taxaRuptura), NUM(r.totalRupturas), PCT(r.taxaA), PCT(r.taxaALucro)],
        [[211,47,47],[51,51,51],[211,47,47],[211,47,47]]
      );
      section('Rupturas classe A (faturamento) — Top 30');
      var topA = r.items.filter(function(i){return i.abc_valorVendido90==='A'}).slice(0,30);
      autoT(['SKU','Descrição','ABC Fat.','ABC Lucro','Qtd Dep.','Venda Méd/Dia','Fat. Méd/Dia'],
        topA.map(function(i){return [i.sku,i.descricao,i.abc_valorVendido90,i.abc_lucro90,i.deposito,Engine.round2(i.vendaMediaDia||0),BRL(i.fatMediaDia||0)];}),
        {0:{halign:'left'},1:{halign:'left'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}}
      );
    }
    else if(reportType === 'dias'){
      var d = data.dias;
      title('Dias de estoque — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(
        ['COBERTURA MÉDIA','SKUS SEM GIRO','COBERTURA CLASSE A','OVERSTOCK (60+ DIAS)'],
        [d.mediaGeral+' dias', NUM(d.semGiro), d.mediaA+' dias', NUM(d.excessos)],
        [[51,51,51],[211,47,47],[51,51,51],[21,101,192]]
      );
      section('Distribuição');
      autoT(['Faixa','SKUs','% do Total'],
        [['Crítico (0-7 dias)',d.criticos,PCT(d.total?d.criticos/d.total*100:0)],
         ['Baixo (8-15 dias)',d.baixos,PCT(d.total?d.baixos/d.total*100:0)],
         ['Adequado (16-60 dias)',d.adequados,PCT(d.total?d.adequados/d.total*100:0)],
         ['Excesso (60+ dias)',d.excessos,PCT(d.total?d.excessos/d.total*100:0)],
         ['Sem giro',d.semGiro,PCT(d.total?d.semGiro/d.total*100:0)]],
        {0:{halign:'left'},1:{halign:'right'},2:{halign:'right'}}
      );
    }
    else if(reportType === 'abc'){
      var a = data.abc;
      title('Investimento em estoque por curva ABC — Resumo executivo');
      section('Curva ABC por faturamento');
      autoT(['Classe','Investimento (R$)','% Investimento','Faturamento (R$)','% Faturamento'],
        [['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],
         ['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],
         ['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}}
      );
      section('Curva ABC por lucro');
      autoT(['Classe','Investimento (R$)','% Investimento','Lucro (R$)','% Lucro'],
        [['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],
         ['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],
         ['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}}
      );
    }
    else if(reportType === 'perda'){
      var p = data.perda;
      title('Projeção de venda perdida por ruptura — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(
        ['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL (FAT.)','SKUS EM RUPTURA'],
        [BRLi(p.totalPerdaFat), BRLi(p.totalPerdaLucro), BRLi(p.perdaMensal), NUM(p.totalSKUs)],
        [[211,47,47],[211,47,47],[211,47,47],[51,51,51]]
      );
      section('Impacto por classe ABC');
      autoT(['Classe','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% da Perda Total','Perda Mensal Projetada'],
        [['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],
         ['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],
         ['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}}
      );
      section('Top 30 itens com maior perda projetada');
      var topP = p.items.slice().sort(function(a,b){return b.perdaFatDia-a.perdaFatDia}).slice(0,30);
      autoT(['SKU','Descrição','ABC Fat.','Perda Fat./Dia','Perda Lucro/Dia'],
        topP.map(function(i){return [i.sku,i.descricao,i.abcFat,BRL(i.perdaFatDia),BRL(i.perdaLucroDia)];}),
        {0:{halign:'left'},1:{halign:'left'},3:{halign:'right'},4:{halign:'right'}}
      );
    }
    // Note
    checkPage(12);
    doc.setFontSize(7); doc.setTextColor(150,150,150);
    doc.text('Nota: este relatório considera dados processados em '+processDate+'. Os valores são baseados nos arquivos fornecidos.', M, y);
    y+=4;
    doc.text('Premissa: a venda média dos últimos 90 dias representa a demanda normal. Valores projetados são estimativas.', M, y);

    doc.save('resumo_'+reportType+'_'+processDate.replace(/[\/ :]/g,'_')+'.pdf');
  }

  return { generateExcel:generateExcel, generatePDF:generatePDF };
})();
