/* export.js — v2.0 — Excel formatado (xlsx-js-style) + Dashboard + PDF completo */
var Export = (function(){
  "use strict";

  /* ===== CORES E FORMATADORES ===== */
  var C = {
    navy:'051323', navyL:'0B2545', green:'00B74A', greenD:'009B3E',
    red:'D32F2F', redL:'FDEAEA', amb:'F57C00', ambL:'FFF3E0',
    blue:'1565C0', blueL:'E3F2FD', white:'FFFFFF', light:'F5F5F5',
    lightG:'F0F0F0', border:'D0D0D0', text:'333333', muted:'888888'
  };
  var BRL=function(v){return(v<0?'−':'')+'R$ '+Math.abs(v||0).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2});};
  var BRLi=function(v){return(v<0?'−':'')+'R$ '+Math.abs(Math.round(v||0)).toLocaleString('pt-BR');};
  var PCT=function(v){return(v||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1})+'%';};
  var NUM=function(v){return(v||0).toLocaleString('pt-BR');};

  /* ===== ESTILOS REUTILIZÁVEIS ===== */
  function sHeader(){return{font:{bold:true,color:{rgb:C.white},sz:10,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:C.border}},top:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}
  function sBody(align,bold){return{font:{name:'Arial',sz:10,bold:!!bold,color:{rgb:C.text}},alignment:{horizontal:align||'left',vertical:'center'},border:{bottom:{style:'hair',color:{rgb:C.lightG}},left:{style:'hair',color:{rgb:C.lightG}},right:{style:'hair',color:{rgb:C.lightG}}}};}
  function sBodyAlt(align,bold){var s=sBody(align,bold);s.fill={fgColor:{rgb:C.light}};return s;}
  function sBrand(){return{font:{bold:true,color:{rgb:C.white},sz:14,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'left',vertical:'center'}};}
  function sSubBrand(){return{font:{color:{rgb:'B0C4DE'},sz:10,name:'Arial'},fill:{fgColor:{rgb:C.navy}},alignment:{horizontal:'left',vertical:'center'}};}
  function sNavyFill(){return{fill:{fgColor:{rgb:C.navy}}};}
  function sSectionTitle(){return{font:{bold:true,color:{rgb:C.navy},sz:11,name:'Arial'},border:{bottom:{style:'medium',color:{rgb:C.green}}}};}
  function sKpiLabel(){return{font:{bold:true,color:{rgb:C.muted},sz:8,name:'Arial'},fill:{fgColor:{rgb:C.light}},alignment:{horizontal:'center',vertical:'center'},border:{top:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}
  function sKpiValue(color){return{font:{bold:true,color:{rgb:color||C.text},sz:14,name:'Arial'},fill:{fgColor:{rgb:C.light}},alignment:{horizontal:'center',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:C.border}},left:{style:'thin',color:{rgb:C.border}},right:{style:'thin',color:{rgb:C.border}}}};}
  function sRedVal(){return{font:{name:'Arial',sz:10,color:{rgb:C.red}},alignment:{horizontal:'right',vertical:'center'},border:{bottom:{style:'hair',color:{rgb:C.lightG}}}};}
  function sGreenVal(){return{font:{name:'Arial',sz:10,color:{rgb:C.green}},alignment:{horizontal:'right',vertical:'center'},border:{bottom:{style:'hair',color:{rgb:C.lightG}}}};}

  /* ===== HELPERS PARA CONSTRUIR PLANILHAS ===== */
  function colLetter(n){var s='';while(n>=0){s=String.fromCharCode(65+(n%26))+s;n=Math.floor(n/26)-1;}return s;}
  function cellRef(r,c){return colLetter(c)+String(r+1);}

  function setCell(ws,r,c,val,style){
    var ref=cellRef(r,c);
    ws[ref]={v:val,t:typeof val==='number'?'n':'s',s:style||sBody()};
    /* Atualizar range */
    if(!ws['!ref']){ws['!ref']='A1:'+ref;}
    else{
      var range=XLSX.utils.decode_range(ws['!ref']);
      if(r>range.e.r)range.e.r=r;if(c>range.e.c)range.e.c=c;
      ws['!ref']=XLSX.utils.encode_range(range);
    }
  }

  function addBrandHeader(ws,row,info,processDate,totalCols){
    for(var i=0;i<totalCols;i++){setCell(ws,row,i,'',sNavyFill());}
    setCell(ws,row,0,'FORMULA CODE — AUDITORIA DE ESTOQUE',sBrand());
    ws['!merges']=ws['!merges']||[];
    ws['!merges'].push({s:{r:row,c:0},e:{r:row,c:Math.min(3,totalCols-1)}});
    for(var i=0;i<totalCols;i++){setCell(ws,row+1,i,'',sNavyFill());}
    setCell(ws,row+1,0,'Cliente: '+(info.cliente||'—')+' | Unidade: '+(info.unidade||'—')+' | Inventário: '+(info.dataInventario||'—')+' | Processado: '+processDate,sSubBrand());
    ws['!merges'].push({s:{r:row+1,c:0},e:{r:row+1,c:Math.min(5,totalCols-1)}});
    return row+3;
  }

  function addSectionTitle(ws,row,title,totalCols){
    setCell(ws,row,0,title,sSectionTitle());
    return row+1;
  }

  function addKpiRow(ws,row,labels,values,colors,startCol){
    startCol=startCol||0;
    for(var i=0;i<labels.length;i++){
      setCell(ws,row,startCol+i,labels[i],sKpiLabel());
      setCell(ws,row+1,startCol+i,values[i],sKpiValue(colors&&colors[i]?colors[i]:C.text));
    }
    return row+3;
  }

  function addDataTable(ws,row,headers,dataRows,colAligns,totalCols){
    /* Header */
    for(var i=0;i<headers.length;i++){
      setCell(ws,row,i,headers[i],sHeader());
    }
    row++;
    /* Body */
    for(var r=0;r<dataRows.length;r++){
      var isAlt=r%2===1;
      for(var c=0;c<dataRows[r].length;c++){
        var val=dataRows[r][c];
        var align=(colAligns&&colAligns[c])?colAligns[c]:'left';
        var style=isAlt?sBodyAlt(align):sBody(align);
        /* Colorir valores negativos de vermelho */
        if(typeof val==='string'&&val.indexOf('−')===0){
          style=JSON.parse(JSON.stringify(style));
          style.font.color={rgb:C.red};
        }
        setCell(ws,row+r,c,val,style);
      }
    }
    return row+dataRows.length+1;
  }

  /* ===== GERAR EXCEL COMPLETO ===== */
  function generateExcel(data,selected,processDate,info){
    info=info||{};
    var wb=XLSX.utils.book_new();
    var hasAny=false;

    /* --- ABA DASHBOARD --- */
    var ws={};
    var R=0;
    var DC=8; /* largura padrão do dashboard em colunas */

    R=addBrandHeader(ws,R,info,processDate,DC);

    /* Crítica no dashboard */
    if(data.critica){
      var c=data.critica;
      R=addSectionTitle(ws,R,'CRÍTICA DO INVENTÁRIO',DC);
      R=addKpiRow(ws,R,
        ['ACURACIDADE','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],
        [PCT(c.acuracidade),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],
        [C.green,C.red,C.amb,C.red]
      );
      if(c.hasCategorias){
        var catH=['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'];
        var catD=c.categorias.map(function(cat){return[cat.nome,cat.total,PCT(cat.acuracidade),BRLi(cat.faltaVal),BRLi(cat.sobraVal),BRLi(cat.saldo)];});
        R=addDataTable(ws,R,catH,catD,{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'},DC);
      }
    }

    /* Ruptura no dashboard */
    if(data.ruptura){
      var r=data.ruptura;
      R=addSectionTitle(ws,R,'RUPTURA',DC);
      R=addKpiRow(ws,R,
        ['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CLASSE A (FAT.)','RUPTURA CLASSE A (LUCRO)'],
        [PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],
        [C.red,C.text,C.red,C.red]
      );
      var rupH=['','Classe A','Classe B','Classe C','Total'];
      var rupD=[['SKUs em ruptura',NUM(r.rupturaA),NUM(r.rupturaB),NUM(r.rupturaC),NUM(r.totalRupturas)]];
      R=addDataTable(ws,R,rupH,rupD,{0:'left',1:'right',2:'right',3:'right',4:'right'},DC);
    }

    /* Dias de Estoque no dashboard */
    if(data.dias){
      var d=data.dias;
      R=addSectionTitle(ws,R,'DIAS DE ESTOQUE',DC);
      R=addKpiRow(ws,R,
        ['COBERTURA GERAL','CLASSE A','CLASSE B','CLASSE C','SEM GIRO'],
        [d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias',NUM(d.semGiro)],
        [C.text,C.text,C.text,C.text,C.red]
      );
      var diasH=['Faixa','SKUs','% do Total'];
      var diasD=[
        ['Ruptura (0-2 dias)',d.ruptura,PCT(d.total?d.ruptura/d.total*100:0)],
        ['Alto risco (3-5 dias)',d.altoRisco,PCT(d.total?d.altoRisco/d.total*100:0)],
        ['Médio risco (6-15 dias)',d.medioRisco,PCT(d.total?d.medioRisco/d.total*100:0)],
        ['Cobertura ideal (16-30 dias)',d.coberturaIdeal,PCT(d.total?d.coberturaIdeal/d.total*100:0)],
        ['Excesso de cobertura (31+ dias)',d.excessos,PCT(d.total?d.excessos/d.total*100:0)],
        ['Sem giro',d.semGiro,PCT(d.total?d.semGiro/d.total*100:0)]
      ];
      R=addDataTable(ws,R,diasH,diasD,{0:'left',1:'right',2:'right'},DC);
    }

    /* ABC no dashboard */
    if(data.abc){
      var a=data.abc;
      R=addSectionTitle(ws,R,'INVESTIMENTO ABC',DC);
      R=addKpiRow(ws,R,
        ['INVESTIMENTO TOTAL','FATURAMENTO 90D','LUCRO 90D','SKUS'],
        [BRLi(a.totalInvest),BRLi(a.totalFat),BRLi(a.totalLucro),NUM(a.items.length)],
        [C.text,C.green,C.green,C.text]
      );
      var abcH=['Classe','Investimento (R$)','% Investimento','Faturamento (R$)','% Faturamento'];
      var abcD=[
        ['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],
        ['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],
        ['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]
      ];
      R=addDataTable(ws,R,abcH,abcD,{0:'center',1:'right',2:'right',3:'right',4:'right'},DC);
    }

    /* Perda no dashboard */
    if(data.perda){
      var pe=data.perda;
      R=addSectionTitle(ws,R,'PROJEÇÃO DE PERDA POR RUPTURA',DC);
      R=addKpiRow(ws,R,
        ['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL (FAT.)','SKUS EM RUPTURA'],
        [BRLi(pe.totalPerdaFat),BRLi(pe.totalPerdaLucro),BRLi(pe.perdaMensal),NUM(pe.totalSKUs)],
        [C.red,C.red,C.red,C.text]
      );
      var perdaH=['Classe','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% da Perda','Perda Mensal'];
      var perdaD=[
        ['A',pe.classA.count,BRLi(pe.classA.perda),BRLi(pe.classA.lucro),PCT(pe.classA.pct),BRLi(pe.classA.perda*30)],
        ['B',pe.classB.count,BRLi(pe.classB.perda),BRLi(pe.classB.lucro),PCT(pe.classB.pct),BRLi(pe.classB.perda*30)],
        ['C',pe.classC.count,BRLi(pe.classC.perda),BRLi(pe.classC.lucro),PCT(pe.classC.pct),BRLi(pe.classC.perda*30)]
      ];
      R=addDataTable(ws,R,perdaH,perdaD,{0:'center',1:'right',2:'right',3:'right',4:'right',5:'right'},DC);
    }

    ws['!cols']=[{wch:28},{wch:18},{wch:16},{wch:18},{wch:16},{wch:18},{wch:16},{wch:16}];
    ws['!rows']=[{hpt:28},{hpt:20}];
    XLSX.utils.book_append_sheet(wb,ws,'Dashboard');

    /* --- ABAS DE DADOS --- */
    if(selected.criticaResumo&&data.critica){
      var wsC={};var rw=0;
      rw=addBrandHeader(wsC,rw,info,processDate,6);
      rw=addSectionTitle(wsC,rw,'RESUMO DA CRÍTICA',6);
      var c=data.critica;
      var sumLabels=['ACURACIDADE','SKUs analisados','SKUs sem divergência','SKUs com falta','SKUs com sobra','Valor das faltas','Valor das sobras','Saldo líquido'];
      var sumVals=[PCT(c.acuracidade),c.totalSKUs,c.okCount,c.faltaCount,c.sobraCount,BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)];
      for(var i=0;i<sumLabels.length;i++){
        setCell(wsC,rw+i,0,sumLabels[i],sBody('left',true));
        setCell(wsC,rw+i,1,sumVals[i],sBody('right'));
      }
      rw+=sumLabels.length+1;
      if(c.hasCategorias){
        rw=addSectionTitle(wsC,rw,'RESULTADO POR CATEGORIA',6);
        var catH2=['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'];
        var catD2=c.categorias.map(function(cat){return[cat.nome,cat.total,PCT(cat.acuracidade),BRLi(cat.faltaVal),BRLi(cat.sobraVal),BRLi(cat.saldo)];});
        rw=addDataTable(wsC,rw,catH2,catD2,{0:'left',1:'right',2:'right',3:'right',4:'right',5:'right'},6);
      }
      wsC['!cols']=[{wch:24},{wch:16},{wch:16},{wch:18},{wch:18},{wch:18}];
      wsC['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsC,'Crítica - Resumo');
    }

    if(selected.criticaDetalhe&&data.critica){
      var wsCD={};var rw=0;
      rw=addBrandHeader(wsCD,rw,info,processDate,9);
      var hd=['SKU','Descrição','Categoria','Qtd Sistema','Qtd Contada','Dif. Qtd','Custo Unit.','Dif. R$','Status'];
      var rd=data.critica.items.map(function(i){return[i.sku,i.descricao,i.categoria,i.qtdSistema,i.qtdContada,i.difQtd,BRL(i.custoUnit),BRL(i.difValor),i.status];});
      rw=addDataTable(wsCD,rw,hd,rd,{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'right',7:'right',8:'center'},9);
      wsCD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:12},{wch:10},{wch:14},{wch:14},{wch:10}];
      wsCD['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsCD,'Crítica - Detalhado');
    }

    if(selected.ruptura&&data.ruptura){
      var wsR={};var rw=0;
      rw=addBrandHeader(wsR,rw,info,processDate,9);
      var hr=['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Qtd Loja','Venda Méd/Dia','Fat. Méd/Dia'];
      var rr=data.ruptura.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90||'C',i.abc_lucro90||'C',i.deposito,i.loja,Engine.round2(i.vendaMediaDia||0),BRL(i.fatMediaDia||0)];});
      rw=addDataTable(wsR,rw,hr,rr,{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right'},9);
      wsR['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:14},{wch:10},{wch:14},{wch:14}];
      wsR['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsR,'Ruptura');
    }

    if(selected.dias&&data.dias){
      var wsD={};var rw=0;
      rw=addBrandHeader(wsD,rw,info,processDate,9);
      var hde=['SKU','Descrição','Categoria','Qtd Estoque','Venda Méd/Dia','Dias Estoque','Cobertura','Valor Estoque','ABC Fat.'];
      var rde=data.dias.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.qtdEstoque,Engine.round2(i.vendaMediaDia||0),i.diasEstoque!==null?Engine.round2(i.diasEstoque):'—',i.faixa,BRL(i.valorEstoque),i.abcFat];});
      rw=addDataTable(wsD,rw,hde,rde,{0:'left',1:'left',2:'left',3:'right',4:'right',5:'right',6:'center',7:'right',8:'center'},9);
      wsD['!cols']=[{wch:14},{wch:32},{wch:18},{wch:12},{wch:14},{wch:12},{wch:18},{wch:16},{wch:10}];
      wsD['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsD,'Dias de Estoque');
    }

    if(selected.abc&&data.abc){
      var wsA={};var rw=0;
      rw=addBrandHeader(wsA,rw,info,processDate,10);
      var ha=['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Estoque','Custo Unit.','Valor Investido','Fat. 90 dias','Lucro 90 dias'];
      var ra=data.abc.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro,i.qtdEstoque,BRL(i.custoUnit),BRL(i.valorInvestido),BRL(i.fat90),BRL(i.lucro90)];});
      rw=addDataTable(wsA,rw,ha,ra,{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right',9:'right'},10);
      wsA['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:16},{wch:16},{wch:16}];
      wsA['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsA,'Investimento ABC');
    }

    if(selected.perda&&data.perda){
      var wsP={};var rw=0;
      rw=addBrandHeader(wsP,rw,info,processDate,11);
      var hp=['SKU','Descrição','Categoria','ABC Fat.','ABC Lucro','Qtd Depósito','Venda Méd/Dia','Perda Fat./Dia','Perda Lucro/Dia','Perda Fat./Mês','Perda Lucro/Mês'];
      var rp=data.perda.items.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abcFat,i.abcLucro||'C',i.qtdDeposito,i.vendaMediaDia,BRL(i.perdaFatDia),BRL(i.perdaLucroDia),BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];});
      rw=addDataTable(wsP,rw,hp,rp,{0:'left',1:'left',2:'left',3:'center',4:'center',5:'right',6:'right',7:'right',8:'right',9:'right',10:'right'},11);
      wsP['!cols']=[{wch:14},{wch:32},{wch:18},{wch:10},{wch:10},{wch:12},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];
      wsP['!rows']=[{hpt:28},{hpt:20}];
      XLSX.utils.book_append_sheet(wb,wsP,'Projeção de Perda');
    }

    var out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    var blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var url=URL.createObjectURL(blob);
    var anc=document.createElement('a');
    anc.href=url;
    var fname='auditoria_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.xlsx';
    anc.download=fname;
    anc.click();
    URL.revokeObjectURL(url);
  }

  /* ========== PDF ========== */
  function generatePDF(reportType,data,processDate,logoDataUrl,info){
    info=info||{};
    var jsPDF=window.jspdf.jsPDF;
    var doc=new jsPDF({orientation:'portrait',unit:'mm',format:'a4'});
    var W=210,H=297,M=15;
    var y=0;
    function addHeader(){
      doc.setFillColor(5,19,35);doc.rect(0,0,W,22,'F');
      if(logoDataUrl){try{doc.addImage(logoDataUrl,'PNG',M,3,36,12);}catch(e){}}
      doc.setFontSize(9);doc.setTextColor(255,255,255);
      doc.text((info.cliente||'')+' — '+(info.unidade||''),W-M,7,{align:'right'});
      doc.setFontSize(7);doc.setTextColor(200,220,255);
      doc.text('Inventário: '+(info.dataInventario||'—'),W-M,12,{align:'right'});
      doc.setTextColor(180,180,200);
      doc.text('Processado em '+processDate,W-M,17,{align:'right'});
      y=28;
    }
    function addFooter(pg){
      doc.setFontSize(7);doc.setTextColor(150,150,150);
      doc.text('Formula Code Tecnologia, Gestão e Automação',M,H-6);
      doc.text('Página '+pg,W-M,H-6,{align:'right'});
      doc.setDrawColor(200,200,200);doc.line(M,H-10,W-M,H-10);
    }
    function checkPage(need){
      if(y+need>H-18){doc.addPage();addHeader();addFooter(doc.getNumberOfPages());}
    }
    function title(txt){
      checkPage(12);
      doc.setFontSize(14);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');
      doc.text(txt,M,y);y+=6;
      doc.setFontSize(8);doc.setTextColor(150,150,150);doc.setFont(undefined,'normal');
      doc.text('Relatório gerado automaticamente pelo sistema Formula Code',M,y);y+=8;
    }
    function section(txt){
      checkPage(10);
      doc.setFontSize(11);doc.setTextColor(5,19,35);doc.setFont(undefined,'bold');
      doc.text(txt,M,y);y+=6;doc.setFont(undefined,'normal');
    }
    function autoT(head,body,opts){
      checkPage(20);
      doc.autoTable({startY:y,head:[head],body:body,margin:{left:M,right:M},
        headStyles:{fillColor:[5,19,35],fontSize:7,fontStyle:'bold',halign:'center'},
        bodyStyles:{fontSize:7},
        alternateRowStyles:{fillColor:[245,245,245]},
        styles:{cellPadding:1.5,lineColor:[220,220,220],lineWidth:0.2},
        columnStyles:opts||{}
      });
      y=doc.lastAutoTable.finalY+6;
    }
    function kpiRow(labels,values,colors){
      checkPage(18);
      var cw=(W-2*M)/labels.length;
      doc.setFillColor(245,245,245);
      doc.roundedRect(M,y-2,W-2*M,16,2,2,'F');
      for(var i=0;i<labels.length;i++){
        var x=M+i*cw+4;
        doc.setFontSize(7);doc.setTextColor(150,150,150);doc.setFont(undefined,'bold');
        doc.text(labels[i],x,y+3);
        doc.setFontSize(12);doc.setFont(undefined,'bold');
        var cl=colors[i]||[51,51,51];
        doc.setTextColor(cl[0],cl[1],cl[2]);
        doc.text(values[i],x,y+10);
      }
      doc.setFont(undefined,'normal');
      y+=20;
    }

    addHeader();addFooter(1);

    /* ===== GERAR POR TIPO ===== */
    if(reportType==='critica'){
      var c=data.critica;
      title('Crítica do inventário — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(['ACURACIDADE','VALOR DAS FALTAS','VALOR DAS SOBRAS','SALDO LÍQUIDO'],
        [PCT(c.acuracidade),BRLi(c.totalFaltas),BRLi(c.totalSobras),BRLi(c.saldoLiquido)],
        [[0,183,74],[211,47,47],[245,124,0],[211,47,47]]);
      if(c.hasCategorias){
        section('Resultado por categoria');
        autoT(['Categoria','SKUs','Acuracidade','Faltas (R$)','Sobras (R$)','Saldo (R$)'],
          c.categorias.map(function(cat){return[cat.nome,cat.total,PCT(cat.acuracidade),BRLi(cat.faltaVal),BRLi(cat.sobraVal),BRLi(cat.saldo)];}),
          {0:{halign:'left'},1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
      }
      section('Maiores divergências (top 20)');
      var top20=c.items.slice().sort(function(a,b){return a.difValor-b.difValor;}).slice(0,20);
      autoT(['SKU','Descrição','Categoria','Dif. Qtd','Dif. R$'],
        top20.map(function(i){return[i.sku,i.descricao,i.categoria,i.difQtd,BRL(i.difValor)];}),
        {0:{halign:'left'},1:{halign:'left'},2:{halign:'left'},3:{halign:'right'},4:{halign:'right'}});
    }
    else if(reportType==='ruptura'){
      var r=data.ruptura;
      title('Ruptura — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(['TAXA DE RUPTURA','SKUS EM RUPTURA','RUPTURA CLASSE A (FAT.)','RUPTURA CLASSE A (LUCRO)'],
        [PCT(r.taxaRuptura),NUM(r.totalRupturas),PCT(r.taxaA),PCT(r.taxaALucro)],
        [[211,47,47],[51,51,51],[211,47,47],[211,47,47]]);
      section('Rupturas classe A (faturamento) — Top 30');
      var topA=r.items.filter(function(i){return i.abc_valorVendido90==='A';}).slice(0,30);
      autoT(['SKU','Descrição','Categoria','ABC Fat.','Qtd Dep.','Venda Méd/Dia','Fat. Méd/Dia'],
        topA.map(function(i){return[i.sku,i.descricao,i.categoria||'',i.abc_valorVendido90,i.deposito,Engine.round2(i.vendaMediaDia||0),BRL(i.fatMediaDia||0)];}),
        {0:{halign:'left'},1:{halign:'left'},2:{halign:'left'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}});
    }
    else if(reportType==='dias'){
      var d=data.dias;
      title('Dias de estoque — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(['COBERTURA GERAL','COBERTURA CLASSE A','COBERTURA CLASSE B','COBERTURA CLASSE C'],
        [d.coberturaGeral+' dias',d.coberturaA+' dias',d.coberturaB+' dias',d.coberturaC+' dias'],
        [[51,51,51],[211,47,47],[245,124,0],[136,136,136]]);

      section('Distribuição por faixa');
      autoT(['Faixa','SKUs','% do Total','Valor estoque (R$)'],
        [
          ['Ruptura (0-2 dias)',d.ruptura,PCT(d.total?d.ruptura/d.total*100:0),'—'],
          ['Alto risco (3-5 dias)',d.altoRisco,PCT(d.total?d.altoRisco/d.total*100:0),'—'],
          ['Médio risco (6-15 dias)',d.medioRisco,PCT(d.total?d.medioRisco/d.total*100:0),'—'],
          ['Cobertura ideal (16-30 dias)',d.coberturaIdeal,PCT(d.total?d.coberturaIdeal/d.total*100:0),'—'],
          ['Excesso de cobertura (31+ dias)',d.excessos,PCT(d.total?d.excessos/d.total*100:0),BRLi(d.valorExcesso||0)],
          ['Sem giro',d.semGiro,PCT(d.total?d.semGiro/d.total*100:0),'—']
        ],
        {0:{halign:'left'},1:{halign:'right'},2:{halign:'right'},3:{halign:'right'}});

      if(d.hasCategorias){
        section('Cobertura por categoria');
        autoT(['Categoria','SKUs','Cobertura média','Ruptura+Alto risco','Sem giro','Excessos','Val. estoque (R$)'],
          d.categorias.map(function(cat){return[cat.nome,cat.total,cat.mediaCobertura+' dias',cat.criticos,cat.semGiro,cat.excessos,BRLi(cat.valorEstoque)];}),
          {0:{halign:'left'},1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'},6:{halign:'right'}});
      }

      /* Listar produtos por faixa */
      var faixas=['Ruptura','Alto risco','Médio risco','Cobertura ideal','Excesso de cobertura','Sem giro'];
      var pHead=['SKU','Descrição','Categoria','Dias est.','ABC Fat.','Val. estoque'];
      var pOpts={0:{halign:'left'},1:{halign:'left'},2:{halign:'left'},3:{halign:'right'},4:{halign:'center'},5:{halign:'right'}};
      faixas.forEach(function(fx){
        var itens=d.items.filter(function(i){return i.faixa===fx;});
        if(itens.length){
          section(fx+' — '+itens.length+' itens');
          autoT(pHead,
            itens.slice(0,50).map(function(i){return[i.sku,i.descricao,i.categoria||'',i.diasEstoque!==null?Engine.round2(i.diasEstoque):'—',i.abcFat,BRL(i.valorEstoque)];}),
            pOpts);
          if(itens.length>50){
            doc.setFontSize(7);doc.setTextColor(150,150,150);
            doc.text('... e mais '+(itens.length-50)+' itens (ver Excel para lista completa)',M,y);y+=4;
          }
        }
      });
    }
    else if(reportType==='abc'){
      var a=data.abc;
      title('Investimento em estoque por curva ABC — Resumo executivo');
      section('Curva ABC por faturamento');
      autoT(['Classe','Investimento (R$)','% Investimento','Faturamento (R$)','% Faturamento'],
        [['A',BRLi(a.fatA.invest),PCT(a.fatA.pctInvest),BRLi(a.fatA.fat),PCT(a.fatA.pctFat)],
         ['B',BRLi(a.fatB.invest),PCT(a.fatB.pctInvest),BRLi(a.fatB.fat),PCT(a.fatB.pctFat)],
         ['C',BRLi(a.fatC.invest),PCT(a.fatC.pctInvest),BRLi(a.fatC.fat),PCT(a.fatC.pctFat)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
      section('Curva ABC por lucro');
      autoT(['Classe','Investimento (R$)','% Investimento','Lucro (R$)','% Lucro'],
        [['A',BRLi(a.lucA.invest),PCT(a.lucA.pctInvest),BRLi(a.lucA.luc),PCT(a.lucA.pctLuc)],
         ['B',BRLi(a.lucB.invest),PCT(a.lucB.pctInvest),BRLi(a.lucB.luc),PCT(a.lucB.pctLuc)],
         ['C',BRLi(a.lucC.invest),PCT(a.lucC.pctInvest),BRLi(a.lucC.luc),PCT(a.lucC.pctLuc)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'}});
    }
    else if(reportType==='perda'){
      var p=data.perda;
      title('Projeção de venda perdida por ruptura — Resumo executivo');
      section('Indicadores gerais');
      kpiRow(['PERDA FAT./DIA','PERDA LUCRO/DIA','PERDA MENSAL (FAT.)','SKUS EM RUPTURA'],
        [BRLi(p.totalPerdaFat),BRLi(p.totalPerdaLucro),BRLi(p.perdaMensal),NUM(p.totalSKUs)],
        [[211,47,47],[211,47,47],[211,47,47],[51,51,51]]);
      section('Impacto por classe ABC');
      autoT(['Classe','SKUs','Perda Fat./Dia','Perda Lucro/Dia','% da Perda Total','Perda Mensal Projetada'],
        [['A',p.classA.count,BRLi(p.classA.perda),BRLi(p.classA.lucro),PCT(p.classA.pct),BRLi(p.classA.perda*30)],
         ['B',p.classB.count,BRLi(p.classB.perda),BRLi(p.classB.lucro),PCT(p.classB.pct),BRLi(p.classB.perda*30)],
         ['C',p.classC.count,BRLi(p.classC.perda),BRLi(p.classC.lucro),PCT(p.classC.pct),BRLi(p.classC.perda*30)]],
        {1:{halign:'right'},2:{halign:'right'},3:{halign:'right'},4:{halign:'right'},5:{halign:'right'}});
      var perdaHead=['SKU','Descrição','Categoria','Perda Fat./Mês','Perda Lucro/Mês'];
      var perdaCols={0:{halign:'left'},1:{halign:'left'},2:{halign:'left'},3:{halign:'right'},4:{halign:'right'}};
      var itemsA=p.items.filter(function(i){return i.abcFat==='A';}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});
      if(itemsA.length){section('Classe A — '+itemsA.length+' itens');autoT(perdaHead,itemsA.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),perdaCols);}
      var itemsB=p.items.filter(function(i){return i.abcFat==='B';}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});
      if(itemsB.length){section('Classe B — '+itemsB.length+' itens');autoT(perdaHead,itemsB.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),perdaCols);}
      var itemsC=p.items.filter(function(i){return i.abcFat==='C';}).sort(function(a,b){return b.perdaFatMes-a.perdaFatMes;});
      if(itemsC.length){section('Classe C — '+itemsC.length+' itens');autoT(perdaHead,itemsC.map(function(i){return[i.sku,i.descricao,i.categoria||'',BRL(i.perdaFatMes),BRL(i.perdaLucroMes)];}),perdaCols);}
    }

    checkPage(12);
    doc.setFontSize(7);doc.setTextColor(150,150,150);
    doc.text('Nota: este relatório considera dados processados em '+processDate+'. Os valores são baseados nos arquivos fornecidos.',M,y);y+=4;
    doc.text('Premissa: a venda média dos últimos 90 dias representa a demanda normal. Valores projetados são estimativas.',M,y);

    var pdfName='resumo_'+reportType+'_'+(info.cliente||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.unidade||'').replace(/[^a-zA-Z0-9]/g,'_')+'_'+(info.dataInventario||'').replace(/\//g,'-')+'.pdf';
    doc.save(pdfName);
  }

  return{generateExcel:generateExcel,generatePDF:generatePDF};
})();
