/* engine.js — Cálculos dos 5 relatórios com visão por categoria */
var Engine = (function(){
  "use strict";

  function round2(n){ return Math.round((n||0)*100)/100; }
  function roundInt(n){ return Math.round(n||0); }
  function safeDiv(a,b){ return b ? a/b : null; }

  /* Helper: agrupa itens por categoria e retorna lista ordenada */
  function groupByCategoria(items, aggFn){
    var cats = {};
    items.forEach(function(it){
      var cat = it.categoria || it.descricao_categoria || 'Sem categoria';
      if(!cats[cat]) cats[cat] = {nome:cat, items:[]};
      cats[cat].items.push(it);
    });
    var list = Object.keys(cats).map(function(k){ return aggFn(cats[k]); });
    return list.sort(function(a,b){ return (b.destaque||0) - (a.destaque||0); });
  }

  function hasRealCategorias(catList){
    return catList.length > 1 || (catList.length===1 && catList[0].nome !== 'Sem categoria');
  }

  /* ========== 1. CRITICA ========== */
  function calcCritica(estoque, contagem, cadastro){
    var skuMap = {};
    estoque.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdSistema:Number(row.qtdSistema)||0, custoUnit:Number(row.custoUnit)||0};
    });
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuMap[sku]) skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdSistema:0, custoUnit:Number(row.custoUnit)||0};
      var item = skuMap[sku];
      item.qtdContada = (item.qtdContada||0) + (Number(row.qtdContada)||0);
      if(row.descricao && !item.descricao) item.descricao = row.descricao;
      if(row.categoria && !item.categoria) item.categoria = row.categoria;
      if(row.custoUnit && !item.custoUnit) item.custoUnit = Number(row.custoUnit)||0;
      var local = String(row.local||'').toLowerCase().trim();
      if(!item.locais) item.locais = {};
      item.locais[local] = (item.locais[local]||0) + (Number(row.qtdContada)||0);
    });
    if(cadastro && cadastro.length){
      var cadMap = {};
      cadastro.forEach(function(r){ cadMap[String(r.sku||'').trim()] = r; });
      Object.keys(skuMap).forEach(function(sku){
        var c = cadMap[sku];
        if(c){
          if(c.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = c.descricao;
          if(c.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = c.categoria;
          if(c.custoUnit && !skuMap[sku].custoUnit) skuMap[sku].custoUnit = Number(c.custoUnit)||0;
        }
      });
    }
    var items = [];
    Object.keys(skuMap).forEach(function(sku){
      var it = skuMap[sku];
      it.qtdContada = it.qtdContada || 0;
      it.difQtd = it.qtdContada - it.qtdSistema;
      it.difValor = round2(it.difQtd * it.custoUnit);
      it.status = it.difQtd < 0 ? 'Falta' : (it.difQtd > 0 ? 'Sobra' : 'OK');
      items.push(it);
    });
    var totalSKUs = items.length;
    var okCount = items.filter(function(i){return i.status==='OK'}).length;
    var faltaItems = items.filter(function(i){return i.status==='Falta'});
    var sobraItems = items.filter(function(i){return i.status==='Sobra'});
    var totalFaltas = faltaItems.reduce(function(s,i){return s+i.difValor},0);
    var totalSobras = sobraItems.reduce(function(s,i){return s+i.difValor},0);
    var catList = groupByCategoria(items, function(g){
      var ok = g.items.filter(function(i){return i.status==='OK'}).length;
      var fv = g.items.filter(function(i){return i.status==='Falta'}).reduce(function(s,i){return s+i.difValor},0);
      var sv = g.items.filter(function(i){return i.status==='Sobra'}).reduce(function(s,i){return s+i.difValor},0);
      return {nome:g.nome, total:g.items.length, ok:ok, faltas:g.items.filter(function(i){return i.status==='Falta'}).length, sobras:g.items.filter(function(i){return i.status==='Sobra'}).length, faltaVal:round2(fv), sobraVal:round2(sv), acuracidade:g.items.length?round2(ok/g.items.length*100):0, saldo:round2(fv+sv), destaque:Math.abs(fv)};
    });
    return {items:items, totalSKUs:totalSKUs, okCount:okCount, acuracidade:totalSKUs?round2(okCount/totalSKUs*100):0, faltaCount:faltaItems.length, sobraCount:sobraItems.length, totalFaltas:round2(totalFaltas), totalSobras:round2(totalSobras), saldoLiquido:round2(totalFaltas+totalSobras), categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== ABC helper ========== */
  function calcABC(items, valueField){
    var sorted = items.slice().filter(function(i){return (i[valueField]||0)>0;});
    sorted.sort(function(a,b){return (b[valueField]||0)-(a[valueField]||0);});
    var total = sorted.reduce(function(s,i){return s+(i[valueField]||0);},0);
    var cum = 0;
    sorted.forEach(function(it){
      cum += (it[valueField]||0);
      var pct = total?cum/total*100:0;
      it['abc_'+valueField] = pct<=80?'A':(pct<=95?'B':'C');
    });
    items.forEach(function(it){if(!it['abc_'+valueField]) it['abc_'+valueField]='C';});
    return items;
  }

  function isDeposito(local){
    var d = ['deposito','depósito','dep','dep.','retaguarda','cd','estoque','armazem','armazém','back','reserva'];
    for(var i=0;i<d.length;i++) if(local.indexOf(d[i])>=0) return true;
    return false;
  }

  /* ========== 2. RUPTURA ========== */
  function calcRuptura(contagem, vendas90){
    var skuLocais = {};
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuLocais[sku]) skuLocais[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', deposito:0, loja:0, custoUnit:Number(row.custoUnit)||0};
      var local = String(row.local||'').toLowerCase().trim();
      var qty = Number(row.qtdContada)||0;
      if(isDeposito(local)) skuLocais[sku].deposito += qty;
      else skuLocais[sku].loja += qty;
      if(row.descricao) skuLocais[sku].descricao = row.descricao;
      if(row.categoria) skuLocais[sku].categoria = row.categoria;
    });
    var vendasMap = {};
    if(vendas90 && vendas90.length){
      vendas90.forEach(function(r){
        var sku = String(r.sku||'').trim();
        vendasMap[sku] = {qtdVendida:Number(r.qtdVendida)||0, valorVendido:Number(r.valorVendido)||0, custoVendido:Number(r.custoVendido)||0, lucro:Number(r.lucro)||0};
      });
    }
    var comDeposito = [], rupturas = [];
    Object.keys(skuLocais).forEach(function(sku){
      var it = skuLocais[sku];
      if(it.deposito > 0){
        var v = vendasMap[sku]||{};
        it.vendaMediaDia = safeDiv(v.qtdVendida,90);
        it.fatMediaDia = safeDiv(v.valorVendido,90);
        it.lucroMediaDia = safeDiv(v.lucro||(v.valorVendido-v.custoVendido),90);
        it.valorVendido90 = v.valorVendido||0;
        it.lucro90 = v.lucro||((v.valorVendido||0)-(v.custoVendido||0));
        comDeposito.push(it);
        if(it.loja===0) rupturas.push(it);
      }
    });
    calcABC(comDeposito,'valorVendido90');
    calcABC(comDeposito,'lucro90');
    var order = {A:0,B:1,C:2};
    rupturas.sort(function(a,b){
      var da=order[a.abc_valorVendido90]||2, db=order[b.abc_valorVendido90]||2;
      if(da!==db) return da-db;
      return (b.fatMediaDia||0)-(a.fatMediaDia||0);
    });
    var ruptA=rupturas.filter(function(i){return i.abc_valorVendido90==='A'});
    var comDepA=comDeposito.filter(function(i){return i.abc_valorVendido90==='A'});
    // Categorias
    var catList = groupByCategoria(rupturas, function(g){
      var dep = comDeposito.filter(function(i){return (i.categoria||'Sem categoria')===g.nome});
      var rA = g.items.filter(function(i){return i.abc_valorVendido90==='A'}).length;
      var taxa = dep.length ? round2(g.items.length/dep.length*100) : 0;
      var perdaDia = g.items.reduce(function(s,i){return s+(i.fatMediaDia||0);},0);
      return {nome:g.nome, totalRupturas:g.items.length, totalDeposito:dep.length, taxa:taxa, rupturaA:rA, perdaDia:round2(perdaDia), destaque:g.items.length};
    });
    return {items:rupturas, allDeposito:comDeposito, totalComDeposito:comDeposito.length, totalRupturas:rupturas.length,
      taxaRuptura:comDeposito.length?round2(rupturas.length/comDeposito.length*100):0,
      rupturaA:ruptA.length, rupturaB:rupturas.filter(function(i){return i.abc_valorVendido90==='B'}).length,
      rupturaC:rupturas.filter(function(i){return i.abc_valorVendido90==='C'}).length,
      taxaA:comDepA.length?round2(ruptA.length/comDepA.length*100):0,
      taxaALucro:(function(){var cA=comDeposito.filter(function(i){return i.abc_lucro90==='A'});var rA2=rupturas.filter(function(i){return i.abc_lucro90==='A'});return cA.length?round2(rA2.length/cA.length*100):0;})(),
      comDepA:comDepA.length, categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 3. DIAS DE ESTOQUE ========== */
  function calcDiasEstoque(critica, vendas90){
    var vendasMap = {};
    if(vendas90 && vendas90.length){
      vendas90.forEach(function(r){
        var sku=String(r.sku||'').trim();
        var qtdV=Number(r.qtdVendida)||0;
        var valV=Number(r.valorVendido)||0;
        var custoV=Number(r.custoVendido)||0;
        // Derivar custo unitário do CMV quando disponível
        var custoUnitDerivado = qtdV>0 ? custoV/qtdV : 0;
        vendasMap[sku]={qtdVendida:qtdV, valorVendido:valV, custoVendido:custoV, custoUnitDerivado:custoUnitDerivado};
      });
    }
    var items = critica.items.map(function(it){
      var v = vendasMap[it.sku]||{};
      var vendaMediaDia = safeDiv(v.qtdVendida,90);
      var dias = vendaMediaDia?round2(it.qtdContada/vendaMediaDia):null;
      // Custo: usa custoUnit do item, senão deriva do CMV/Qtde das vendas
      var custo = it.custoUnit || v.custoUnitDerivado || 0;
      var faixa = dias===null?'Sem giro':(dias<=2?'Ruptura':(dias<=5?'Alto risco':(dias<=15?'Médio risco':(dias<=30?'Cobertura ideal':'Excesso de cobertura'))));
      return {sku:it.sku, descricao:it.descricao, categoria:it.categoria, qtdEstoque:it.qtdContada, vendaMediaDia:vendaMediaDia, diasEstoque:dias, faixa:faixa, valorEstoque:round2(it.qtdContada*custo), custoUnit:custo, valorVendido90:v.valorVendido||0};
    });
    items = calcABC(items,'valorVendido90');
    items.forEach(function(it){it.abcFat=it.abc_valorVendido90||'C';});
    var semGiro=items.filter(function(i){return i.faixa==='Sem giro'}).length;
    var ruptura=items.filter(function(i){return i.faixa==='Ruptura'}).length;
    var altoRisco=items.filter(function(i){return i.faixa==='Alto risco'}).length;
    var medioRisco=items.filter(function(i){return i.faixa==='Médio risco'}).length;
    var coberturaIdeal=items.filter(function(i){return i.faixa==='Cobertura ideal'}).length;
    var excessos=items.filter(function(i){return i.faixa==='Excesso de cobertura'}).length;
    var comVenda=items.filter(function(i){return i.diasEstoque!==null&&i.diasEstoque>0});
    // Cobertura = soma estoque / soma venda média diária
    var somaEstoque=comVenda.reduce(function(s,i){return s+i.qtdEstoque},0);
    var somaVendaDia=comVenda.reduce(function(s,i){return s+(i.vendaMediaDia||0)},0);
    var coberturaGeral=somaVendaDia?round2(somaEstoque/somaVendaDia):0;
    function coberturaClasse(cls){
      var fi=items.filter(function(i){return i.abcFat===cls&&i.diasEstoque!==null&&i.vendaMediaDia>0});
      var sE=fi.reduce(function(s,i){return s+i.qtdEstoque},0);
      var sV=fi.reduce(function(s,i){return s+(i.vendaMediaDia||0)},0);
      return sV?round2(sE/sV):0;
    }
    var valExcesso=items.filter(function(i){return i.faixa==='Excesso de cobertura'}).reduce(function(s,i){return s+i.valorEstoque},0);
    var catList = groupByCategoria(items, function(g){
      var cv=g.items.filter(function(i){return i.diasEstoque!==null&&i.vendaMediaDia>0});
      var sE=cv.reduce(function(s,i){return s+i.qtdEstoque},0);
      var sV=cv.reduce(function(s,i){return s+(i.vendaMediaDia||0)},0);
      var media=sV?round2(sE/sV):0;
      var sg=g.items.filter(function(i){return i.faixa==='Sem giro'}).length;
      var cr=g.items.filter(function(i){return i.faixa==='Ruptura'||i.faixa==='Alto risco'}).length;
      var ex=g.items.filter(function(i){return i.faixa==='Excesso de cobertura'}).length;
      var valEst=g.items.reduce(function(s,i){return s+i.valorEstoque},0);
      var valEx=g.items.filter(function(i){return i.faixa==='Excesso de cobertura'}).reduce(function(s,i){return s+i.valorEstoque},0);
      return {nome:g.nome, total:g.items.length, mediaCobertura:media, semGiro:sg, criticos:cr, excessos:ex, valorEstoque:round2(valEst), valorExcesso:round2(valEx), destaque:cr+sg};
    });
    return {items:items, coberturaGeral:coberturaGeral, coberturaA:coberturaClasse('A'), coberturaB:coberturaClasse('B'), coberturaC:coberturaClasse('C'), semGiro:semGiro, ruptura:ruptura, altoRisco:altoRisco, medioRisco:medioRisco, coberturaIdeal:coberturaIdeal, excessos:excessos, valorExcesso:round2(valExcesso), total:items.length, categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 4. INVESTIMENTO ABC ========== */
  function calcInvestimentoABC(critica, vendas90){
    var vendasMap = {};
    vendas90.forEach(function(r){
      var sku=String(r.sku||'').trim();
      vendasMap[sku]={valorVendido:Number(r.valorVendido)||0, custoVendido:Number(r.custoVendido)||0, lucro:Number(r.lucro)||0, qtdVendida:Number(r.qtdVendida)||0};
    });
    var items = critica.items.map(function(it){
      var v=vendasMap[it.sku]||{};
      var lucro=v.lucro||((v.valorVendido||0)-(v.custoVendido||0));
      if(!lucro&&v.valorVendido&&it.custoUnit) lucro=v.valorVendido-(v.qtdVendida*it.custoUnit);
      return {sku:it.sku, descricao:it.descricao, categoria:it.categoria, qtdEstoque:it.qtdContada, custoUnit:it.custoUnit, valorInvestido:round2(it.qtdContada*it.custoUnit), fat90:v.valorVendido||0, lucro90:round2(lucro), qtdVendida90:v.qtdVendida||0};
    });
    items=calcABC(items,'fat90');
    items=calcABC(items,'lucro90');
    items.forEach(function(it){it.abcFat=it.abc_fat90||'C';it.abcLucro=it.abc_lucro90||'C';});
    function agg(cls,f){return items.filter(function(i){return i['abc_'+f]===cls}).reduce(function(s,i){return s+i.valorInvestido},0);}
    function aggF(cls,f,sf){return items.filter(function(i){return i['abc_'+f]===cls}).reduce(function(s,i){return s+i[sf]},0);}
    var totalInvest=items.reduce(function(s,i){return s+i.valorInvestido},0);
    var totalFat=items.reduce(function(s,i){return s+i.fat90},0);
    var totalLucro=items.reduce(function(s,i){return s+i.lucro90},0);
    // Categorias
    var catList = groupByCategoria(items, function(g){
      var inv=g.items.reduce(function(s,i){return s+i.valorInvestido},0);
      var fat=g.items.reduce(function(s,i){return s+i.fat90},0);
      var luc=g.items.reduce(function(s,i){return s+i.lucro90},0);
      var countA=g.items.filter(function(i){return i.abcFat==='A'}).length;
      return {nome:g.nome, total:g.items.length, investimento:round2(inv), faturamento:round2(fat), lucro:round2(luc), countA:countA, pctInvest:totalInvest?round2(inv/totalInvest*100):0, destaque:inv};
    });
    return {items:items, totalInvest:round2(totalInvest), totalFat:round2(totalFat), totalLucro:round2(totalLucro),
      fatA:{invest:round2(agg('A','fat90')),fat:round2(aggF('A','fat90','fat90')),pctInvest:totalInvest?round2(agg('A','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('A','fat90','fat90')/totalFat*100):0},
      fatB:{invest:round2(agg('B','fat90')),fat:round2(aggF('B','fat90','fat90')),pctInvest:totalInvest?round2(agg('B','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('B','fat90','fat90')/totalFat*100):0},
      fatC:{invest:round2(agg('C','fat90')),fat:round2(aggF('C','fat90','fat90')),pctInvest:totalInvest?round2(agg('C','fat90')/totalInvest*100):0,pctFat:totalFat?round2(aggF('C','fat90','fat90')/totalFat*100):0},
      lucA:{invest:round2(agg('A','lucro90')),luc:round2(aggF('A','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('A','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('A','lucro90','lucro90')/totalLucro*100):0},
      lucB:{invest:round2(agg('B','lucro90')),luc:round2(aggF('B','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('B','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('B','lucro90','lucro90')/totalLucro*100):0},
      lucC:{invest:round2(agg('C','lucro90')),luc:round2(aggF('C','lucro90','lucro90')),pctInvest:totalInvest?round2(agg('C','lucro90')/totalInvest*100):0,pctLuc:totalLucro?round2(aggF('C','lucro90','lucro90')/totalLucro*100):0},
      categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* ========== 5. PROJECAO DE PERDA ========== */
  function calcProjecaoPerda(ruptura){
    var items = ruptura.items.map(function(it){
      return {sku:it.sku, descricao:it.descricao, categoria:it.categoria,
        abcFat:it.abc_valorVendido90||'C', abcLucro:it.abc_lucro90||'C',
        qtdDeposito:it.deposito, vendaMediaDia:round2(it.vendaMediaDia||0),
        fatMediaDia:round2(it.fatMediaDia||0), lucroMediaDia:round2(it.lucroMediaDia||0),
        perdaFatDia:round2(it.fatMediaDia||0), perdaLucroDia:round2(it.lucroMediaDia||0)};
    });
    var totalPerdaFat=round2(items.reduce(function(s,i){return s+i.perdaFatDia},0));
    var totalPerdaLucro=round2(items.reduce(function(s,i){return s+i.perdaLucroDia},0));
    function byAbc(cls,f){return round2(items.filter(function(i){return i.abcFat===cls}).reduce(function(s,i){return s+i[f]},0));}
    function cntAbc(cls){return items.filter(function(i){return i.abcFat===cls}).length;}
    // Categorias
    var catList = groupByCategoria(items, function(g){
      var pf=g.items.reduce(function(s,i){return s+i.perdaFatDia},0);
      var pl=g.items.reduce(function(s,i){return s+i.perdaLucroDia},0);
      var rA=g.items.filter(function(i){return i.abcFat==='A'}).length;
      return {nome:g.nome, totalRupturas:g.items.length, perdaFatDia:round2(pf), perdaLucroDia:round2(pl), perdaMensal:round2(pf*30), rupturaA:rA, destaque:pf};
    });
    return {items:items, totalPerdaFat:totalPerdaFat, totalPerdaLucro:totalPerdaLucro, perdaMensal:round2(totalPerdaFat*30), totalSKUs:items.length,
      classA:{perda:byAbc('A','perdaFatDia'),lucro:byAbc('A','perdaLucroDia'),count:cntAbc('A'),pct:totalPerdaFat?round2(byAbc('A','perdaFatDia')/totalPerdaFat*100):0},
      classB:{perda:byAbc('B','perdaFatDia'),lucro:byAbc('B','perdaLucroDia'),count:cntAbc('B'),pct:totalPerdaFat?round2(byAbc('B','perdaFatDia')/totalPerdaFat*100):0},
      classC:{perda:byAbc('C','perdaFatDia'),lucro:byAbc('C','perdaLucroDia'),count:cntAbc('C'),pct:totalPerdaFat?round2(byAbc('C','perdaFatDia')/totalPerdaFat*100):0},
      categorias:catList, hasCategorias:hasRealCategorias(catList)};
  }

  /* Helper: constrói lista de itens a partir da contagem quando não há estoque sistema */
  function buildItemsFromContagem(contagem, cadastro){
    var skuMap = {};
    contagem.forEach(function(row){
      var sku = String(row.sku||'').trim();
      if(!sku) return;
      if(!skuMap[sku]) skuMap[sku] = {sku:sku, descricao:row.descricao||'', categoria:row.categoria||'', qtdContada:0, custoUnit:Number(row.custoUnit)||0, qtdSistema:0, difQtd:0, difValor:0, status:'—'};
      skuMap[sku].qtdContada += (Number(row.qtdContada)||0);
      if(row.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = row.descricao;
      if(row.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = row.categoria;
      if(row.custoUnit && !skuMap[sku].custoUnit) skuMap[sku].custoUnit = Number(row.custoUnit)||0;
    });
    if(cadastro && cadastro.length){
      var cadMap = {};
      cadastro.forEach(function(r){ cadMap[String(r.sku||'').trim()] = r; });
      Object.keys(skuMap).forEach(function(sku){
        var c = cadMap[sku];
        if(c){
          if(c.descricao && !skuMap[sku].descricao) skuMap[sku].descricao = c.descricao;
          if(c.categoria && !skuMap[sku].categoria) skuMap[sku].categoria = c.categoria;
          if(c.custoUnit && !skuMap[sku].custoUnit) skuMap[sku].custoUnit = Number(c.custoUnit)||0;
        }
      });
    }
    return Object.keys(skuMap).map(function(k){ return skuMap[k]; });
  }

  return {calcCritica:calcCritica, calcRuptura:calcRuptura, calcDiasEstoque:calcDiasEstoque, calcInvestimentoABC:calcInvestimentoABC, calcProjecaoPerda:calcProjecaoPerda, calcABC:calcABC, buildItemsFromContagem:buildItemsFromContagem, round2:round2, roundInt:roundInt};
})();
