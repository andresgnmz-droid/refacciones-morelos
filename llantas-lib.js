/* Cálculos del módulo Control de llantas (sin pantalla, para poder probarlos aparte).
   Posiciones: 1-2 delanteras, 3-4 dual trasero lado conductor, 5-6 dual trasero lado ayudante.
   Isuzu: 4 llantas sencillas (1-2 delanteras, 3-4 traseras).
   Si cambias este archivo, sube el número de VERSION en sw.js. */
(function(root){
  "use strict";

  var CEDIS = { MCV:"Cuernavaca", MCU:"Cuautla", MJO:"Jojutla", MPU:"Puente de Ixtla", MAX:"Axochiapan", MCC:"Cocoyoc" };
  var ORDEN = ["MCV", "MCU", "MJO", "MPU", "MAX", "MCC"];
  var HOJA_A_CEDIS = { "cuernavaca":"MCV", "cuautla":"MCU", "jojutla":"MJO", "puente de ixtla":"MPU", "axochiapan":"MAX", "cocoyoc":"MCC" };
  var PREFIJO_DIGITO = { "1":"MCV", "2":"MCU", "3":"MJO", "4":"MPU", "5":"MAX", "6":"MCC" };
  var MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];

  // Reglas confirmadas por Andrés
  var UMBRAL_CAMBIO = 4;   // norma de la empresa: cambio en 4 mm o menos
  var UMBRAL_LEGAL = 3;    // máximo legal: 3 mm o menos no se puede operar
  var UMBRAL_VIGILAR = 6;  // de 5 a 6 mm se vigila
  var BUENA = 7;           // llanta "buena" recuperable para emparejar: más de 7 mm
  var TASA = 1;            // desgaste promedio de reparto: 1 mm por mes
  var TOLERANCIA = 2;      // un valor más de 2 mm sobre la llanta nueva se considera error de captura

  var TIPOS = {
    M2:    { rin:"22.5", nombre:"Freightliner M2" },
    FL360: { rin:"17.5", nombre:"Freightliner FL360" },
    Isuzu: { rin:"16",   nombre:"Isuzu" }
  };

  function norm(s){ return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/\s+/g, " ").trim(); }

  function tipoDeMarca(marca){
    var m = norm(marca);
    if(!m) return null;
    if(m.indexOf("isuzu") >= 0) return "Isuzu";
    if(/fl-? ?360/.test(m)) return "FL360";
    if(m.indexOf("freightliner") >= 0) return "M2";
    return null;
  }
  function nueva(tipo, pos){
    if(tipo === "M2") return pos <= 2 ? 20 : 18;
    if(tipo === "FL360") return 11;
    if(tipo === "Isuzu") return 10;
    return 20;
  }
  function posActivas(tipo){ return tipo === "Isuzu" ? [1,2,3,4] : [1,2,3,4,5,6]; }
  function eje(pos){ return pos <= 2 ? "Delantero" : "Trasero"; }
  function esDual(tipo){ return tipo === "M2" || tipo === "FL360"; }

  function estado(mm){
    if(mm == null) return "sin";
    if(mm <= UMBRAL_LEGAL) return "ilegal";
    if(mm <= UMBRAL_CAMBIO) return "cambio";
    if(mm <= UMBRAL_VIGILAR) return "vigilar";
    return "ok";
  }

  function periodoDeTexto(t){
    var m = /(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\s*(?:de\s*)?(20\d\d)/.exec(norm(t));
    if(!m) return null;
    var i = MESES.indexOf(m[1] === "setiembre" ? "septiembre" : m[1]);
    return m[2] + "-" + (i < 9 ? "0" : "") + (i + 1) + "-01";
  }
  function etiquetaPeriodo(p){
    if(!p) return "";
    var a = p.split("-"); return MESES[+a[1] - 1] + " " + a[0];
  }
  function sumaMeses(p, n){
    var a = p.split("-"), d = new Date(Date.UTC(+a[0], +a[1] - 1 + n, 1));
    return d.getUTCFullYear() + "-" + (d.getUTCMonth() < 9 ? "0" : "") + (d.getUTCMonth() + 1) + "-01";
  }
  function mesesEntre(p1, p2){
    var a = p1.split("-"), b = p2.split("-");
    return (+b[0] - +a[0]) * 12 + (+b[1] - +a[1]);
  }

  /* ---------- lectura del libro de Excel ----------
     cat: { MCV1000: { marca:"...", tipo:"M2" }, ... } (catálogo maestro de unidades) o null. */
  function parsearLibro(XLSX, wb, cat){
    var out = { periodo:null, filas:[], avisos:[], hojas:[], ignoradas:[] };
    var periodos = {}, vistos = {};
    function aviso(nivel, hoja, eco, texto){ out.avisos.push({ nivel:nivel, hoja:hoja || "", eco:eco || "", texto:texto }); }

    wb.SheetNames.forEach(function(nombre){
      var cedis = HOJA_A_CEDIS[norm(nombre)];
      if(!cedis){ out.ignoradas.push(nombre); return; }
      var ws = wb.Sheets[nombre];
      var rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:null, raw:true });
      var h = -1, i;
      for(i = 0; i < Math.min(rows.length, 10); i++){
        var r = rows[i] || [];
        for(var c = 0; c < r.length; c++){ if(norm(r[c]) === "economico"){ h = i; break; } }
        if(h >= 0) break;
      }
      if(h < 0){ aviso("error", nombre, "", "No encontré el encabezado ECONÓMICO; la pestaña no se cargó."); return; }
      var cab = rows[h], ecoCol = -1, posCol = {}, camCol = -1, alaCol = -1;
      cab.forEach(function(v, c){
        var t = norm(v), m;
        if(t === "economico") ecoCol = c;
        else if((m = /^(?:pos\.? ?)?([1-6])$/.exec(t))) posCol[+m[1]] = c;
        else if(t.indexOf("camara") === 0) camCol = c;
        else if(t.indexOf("alarma") === 0) alaCol = c;
      });
      if(Object.keys(posCol).length < 6){ aviso("error", nombre, "", "El encabezado no trae las 6 columnas de posición (1 a 6); la pestaña no se cargó."); return; }
      for(i = 0; i <= h; i++){
        (rows[i] || []).forEach(function(v){ var p = typeof v === "string" ? periodoDeTexto(v) : null; if(p) periodos[p] = (periodos[p] || 0) + 1; });
      }
      var n = 0;
      for(i = h + 1; i < rows.length; i++){
        var row = rows[i] || [], ecoRaw = row[ecoCol], fila = i + 1;
        var tieneDato = Object.keys(posCol).some(function(k){ var v = row[posCol[k]]; return v != null && v !== ""; }) ||
                        (camCol >= 0 && row[camCol] != null) || (alaCol >= 0 && row[alaCol] != null);
        if(ecoRaw == null || String(ecoRaw).trim() === ""){
          if(tieneDato) aviso("dato", nombre, "", "Fila " + fila + ": hay un valor suelto sin económico (" + Object.keys(posCol).map(function(k){ return row[posCol[k]]; }).filter(function(v){ return v != null && v !== ""; }).join(", ") + "). Se ignoró.");
          continue;
        }
        var s = String(ecoRaw).trim().toUpperCase(), eco;
        if(/^\d{4}$/.test(s)) eco = cedis + s;
        else if(/^M[A-Z]{2}\d{4}$/.test(s)) eco = s;
        else { aviso("dato", nombre, s, "Fila " + fila + ": el económico \"" + s + "\" no tiene el formato esperado (4 dígitos). Se ignoró."); continue; }
        if(eco.slice(0, 3) !== cedis) aviso("dato", nombre, eco, "El económico no corresponde a la pestaña " + nombre + ". Se guardó con el CEDIS de la pestaña.");
        else if(/^\d{4}$/.test(s) && PREFIJO_DIGITO[s.charAt(0)] !== cedis) aviso("dato", nombre, eco, "El primer dígito del económico no corresponde al CEDIS de la pestaña.");
        if(vistos[eco]){ aviso("dato", nombre, eco, "Económico repetido (también en " + vistos[eco] + "). Se conservó la primera captura."); continue; }
        vistos[eco] = nombre;

        var info = cat ? cat[eco] : null, tipo = info ? info.tipo : null;
        if(cat && !info) aviso("info", nombre, eco, "No está en el catálogo de unidades activo: no se conoce su tipo, así que no entra a emparejamiento de duales.");
        else if(cat && !tipo) aviso("info", nombre, eco, "Su marca en el catálogo no es Freightliner M2, FL360 ni Isuzu: no entra a emparejamiento de duales.");
        var p = [], dud = [], faltan = [], activas = tipo ? posActivas(tipo) : [1,2,3,4,5,6];
        for(var k = 1; k <= 6; k++){
          var v = row[posCol[k]];
          if(v == null || v === ""){ p.push(null); if(activas.indexOf(k) >= 0) faltan.push(k); continue; }
          var num = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
          if(isNaN(num)){ p.push(null); dud.push(k); aviso("dato", nombre, eco, "Posición " + k + ": el valor \"" + v + "\" no es un número. Se excluyó del cálculo."); continue; }
          p.push(num);
          if(activas.indexOf(k) < 0){ dud.push(k); aviso("dato", nombre, eco, "Posición " + k + " trae " + num + " mm pero esta unidad (" + tipo + ") solo usa 4 llantas. Se excluyó del cálculo."); continue; }
          var tope = tipo ? nueva(tipo, k) : 20;
          if(num < 0 || num > tope + TOLERANCIA){
            dud.push(k);
            aviso("dato", nombre, eco, "Posición " + k + ": " + num + " mm no es posible para llanta " + (tipo || "") + " (nueva trae " + tope + " mm). Se marcó como dudosa y se excluyó del cálculo.");
          } else if(num > tope){
            aviso("info", nombre, eco, "Posición " + k + ": " + num + " mm supera por " + (num - tope) + " mm a la llanta nueva (" + tope + " mm). Se incluyó en el cálculo.");
          }
        }
        var vacia = p.every(function(x){ return x == null; });
        if(vacia) aviso("dato", nombre, eco, "Sin ninguna captura de profundidad este mes.");
        else if(faltan.length) aviso("dato", nombre, eco, "Faltan las posiciones " + faltan.join(", ") + ".");
        var cam = camCol >= 0 ? row[camCol] : null, ala = alaCol >= 0 ? row[alaCol] : null;
        [["cámara", cam], ["alarma", ala]].forEach(function(x){ if(x[1] != null && x[1] !== 1 && x[1] !== 2) aviso("dato", nombre, eco, "El código de " + x[0] + " (" + x[1] + ") no es 1 ni 2. Se ignoró."); });
        out.filas.push({
          economico:eco, cedis:cedis, p:p, dudosas:dud,
          camara:(cam === 1 || cam === 2) ? cam : null, alarma:(ala === 1 || ala === 2) ? ala : null,
          sin_captura:vacia
        });
        n++;
      }
      out.hojas.push({ hoja:nombre, cedis:cedis, unidades:n, camara:camCol >= 0, alarma:alaCol >= 0 });
    });

    var ps = Object.keys(periodos);
    if(ps.length) out.periodo = ps.sort(function(a, b){ return periodos[b] - periodos[a]; })[0];
    if(ps.length > 1) aviso("dato", "", "", "Las pestañas no dicen el mismo mes (" + ps.map(etiquetaPeriodo).join(", ") + "). Verifica el mes antes de guardar.");
    if(!out.periodo) aviso("dato", "", "", "No pude leer el mes del encabezado; elígelo antes de guardar.");
    var sinCamara = out.hojas.filter(function(x){ return !x.camara && !x.alarma; }).map(function(x){ return x.hoja; });
    if(sinCamara.length) aviso("info", "", "", "Sin columnas de cámara/alarma en: " + sinCamara.join(", ") + ".");
    if(cat){
      var faltantes = Object.keys(cat).filter(function(e){ return cat[e].tipo && !vistos[e]; });
      if(faltantes.length) aviso("dato", "", "", faltantes.length + " unidad(es) activa(s) del catálogo no están en el archivo: " + faltantes.sort().join(", ") + ".");
    }
    return out;
  }

  /* ---------- filas guardadas ---------- */
  function deBase(r){
    return { periodo:r.periodo, economico:r.economico, cedis:r.cedis,
      p:[r.p1, r.p2, r.p3, r.p4, r.p5, r.p6].map(function(x){ return x == null ? null : Number(x); }),
      dudosas:r.dudosas || [], camara:r.camara, alarma:r.alarma, sin_captura:!!r.sin_captura };
  }
  function aBase(f, periodo){
    return { periodo:periodo, economico:f.economico, cedis:f.cedis,
      p1:f.p[0], p2:f.p[1], p3:f.p[2], p4:f.p[3], p5:f.p[4], p6:f.p[5],
      dudosas:f.dudosas, camara:f.camara, alarma:f.alarma, sin_captura:f.sin_captura };
  }
  // Profundidades que sí entran al cálculo: con dato, no dudosas y de una posición que la unidad usa.
  function validas(f, tipo){
    var act = tipo ? posActivas(tipo) : [1,2,3,4,5,6], out = [];
    for(var k = 1; k <= 6; k++){
      var v = f.p[k - 1];
      if(v == null || f.dudosas.indexOf(k) >= 0 || act.indexOf(k) < 0) continue;
      out.push({ pos:k, mm:v });
    }
    return out;
  }

  /* ---------- resumen y alertas ---------- */
  function resumen(filas, tipos){
    var r = { unidades:0, sinCaptura:0, posiciones:0, cambio:0, ilegal:0, vigilar:0, ok:0, delCambio:0, unidadesCambio:0, dudosas:0,
              camFalla:0, alaFalla:0, camCap:0, alaCap:0 };
    filas.forEach(function(f){
      var t = tipos[f.economico], v = validas(f, t);
      r.unidades++; r.dudosas += f.dudosas.length;
      if(f.camara != null){ r.camCap++; if(f.camara === 1) r.camFalla++; }
      if(f.alarma != null){ r.alaCap++; if(f.alarma === 1) r.alaFalla++; }
      if(!v.length){ r.sinCaptura++; return; }
      var hay = false;
      v.forEach(function(x){
        r.posiciones++;
        var e = estado(x.mm);
        if(e === "ilegal"){ r.ilegal++; r.cambio++; hay = true; if(x.pos <= 2) r.delCambio++; }
        else if(e === "cambio"){ r.cambio++; hay = true; if(x.pos <= 2) r.delCambio++; }
        else if(e === "vigilar") r.vigilar++;
        else r.ok++;
      });
      if(hay) r.unidadesCambio++;
    });
    return r;
  }
  function alertas(filas, tipos){
    var out = [];
    filas.forEach(function(f){
      var t = tipos[f.economico], v = validas(f, t).filter(function(x){ return x.mm <= UMBRAL_CAMBIO; });
      if(!v.length) return;
      var min = Math.min.apply(null, v.map(function(x){ return x.mm; }));
      out.push({ fila:f, tipo:t, posiciones:v, min:min, nivel:min <= UMBRAL_LEGAL ? "ilegal" : "cambio" });
    });
    out.sort(function(a, b){ return (a.min - b.min) || a.fila.economico.localeCompare(b.fila.economico); });
    return out;
  }

  /* ---------- proyección ---------- */
  // Mediana de la tasa real (mm/mes) de una unidad entre sus dos últimas inspecciones. Ignora posiciones que subieron (llanta cambiada).
  function tasaReal(actual, previa, periodoA, periodoP, tipo){
    if(!actual || !previa) return null;
    var n = mesesEntre(periodoP, periodoA);
    if(n <= 0) return null;
    var tasas = [];
    validas(actual, tipo).forEach(function(x){
      var y = previa.p[x.pos - 1];
      if(y == null || previa.dudosas.indexOf(x.pos) >= 0) return;
      var d = y - x.mm;
      if(d > 0 && d / n <= 3) tasas.push(d / n);
    });
    if(tasas.length < 2) return null;
    tasas.sort(function(a, b){ return a - b; });
    return tasas[Math.floor(tasas.length / 2)];
  }
  // Meses que le quedan a cada posición antes de llegar a umbral (0 si ya está en o por debajo).
  function mesesRestantes(mm, tasa, umbral){ return mm <= umbral ? 0 : Math.ceil((mm - umbral) / tasa - 1e-9); }
  function proyeccion(filas, tipos, periodo, tasaDe){
    return filas.map(function(f){
      var t = tipos[f.economico], v = validas(f, t), tasa = tasaDe(f.economico);
      if(!v.length) return { fila:f, tipo:t, tasa:tasa, sin:true };
      var lim = v.reduce(function(a, b){ return b.mm < a.mm ? b : a; });
      var m4 = mesesRestantes(lim.mm, tasa, UMBRAL_CAMBIO), m3 = mesesRestantes(lim.mm, tasa, UMBRAL_LEGAL);
      return { fila:f, tipo:t, tasa:tasa, limite:lim, meses:m4, mesesLegal:m3, fecha:sumaMeses(periodo, m4), sin:false };
    });
  }

  /* ---------- pedido y emparejamiento de duales ----------
     h = meses hacia adelante (0 = hoy). tasaDe(eco) = mm/mes de esa unidad.
     Un dual con cualquier llanta en o bajo 4 mm se cambia completo (2 llantas, nunca una sola).
     Emparejar: dos unidades del mismo CEDIS y misma medida, cada una con un dual trasero con una llanta mala (≤4 mm)
     y otra buena (>7 mm). La llanta buena de la "donante" pasa al dual de la "receptora" y la donante recibe un dual nuevo:
     se resuelven 2 duales con 2 llantas nuevas en lugar de 4. Se prioriza la menor diferencia de mm en el dual resultante. */
  function plan(filas, tipos, h, tasaDe){
    var singles = [], duales = [], cand = {};
    filas.forEach(function(f){
      var t = tipos[f.economico], tasa = tasaDe(f.economico);
      if(!t) return;
      var val = {};
      validas(f, t).forEach(function(x){ val[x.pos] = x.mm - tasa * h; });
      var lista = (t === "Isuzu") ? [[1],[2],[3],[4]] : [[1],[2],[3,4],[5,6]];
      lista.forEach(function(g){
        if(g.length === 1){
          var v = val[g[0]];
          if(v != null && v <= UMBRAL_CAMBIO) singles.push({ cedis:f.cedis, tipo:t, eje:eje(g[0]), eco:f.economico, pos:g[0], mm:v });
          return;
        }
        var a = val[g[0]], b = val[g[1]];
        var malA = a != null && a <= UMBRAL_CAMBIO, malB = b != null && b <= UMBRAL_CAMBIO;
        if(!malA && !malB) return;
        var d = { cedis:f.cedis, tipo:t, eco:f.economico, dual:g.join("-"), pos:g, a:a, b:b };
        duales.push(d);
        if(a != null && b != null && (malA !== malB)){
          var buenaPos = malA ? g[1] : g[0], buenaMm = malA ? b : a, malaPos = malA ? g[0] : g[1], malaMm = malA ? a : b;
          if(buenaMm > BUENA){
            d.buena = { pos:buenaPos, mm:buenaMm }; d.mala = { pos:malaPos, mm:malaMm };
            var k = f.cedis + "|" + t;
            (cand[k] = cand[k] || []).push(d);
          }
        }
      });
    });
    var pares = [];
    Object.keys(cand).forEach(function(k){
      var L = cand[k], opciones = [], i, j;
      for(i = 0; i < L.length; i++) for(j = i + 1; j < L.length; j++){
        if(L[i].eco === L[j].eco) continue;
        opciones.push({ a:L[i], b:L[j], dif:Math.abs(L[i].buena.mm - L[j].buena.mm) });
      }
      opciones.sort(function(x, y){ return (x.dif - y.dif) || x.a.eco.localeCompare(y.a.eco) || x.b.eco.localeCompare(y.b.eco); });
      var usados = {};
      opciones.forEach(function(o){
        var ka = o.a.eco + o.a.dual, kb = o.b.eco + o.b.dual;
        if(usados[ka] || usados[kb]) return;
        usados[ka] = usados[kb] = true;
        var don = o.a.buena.mm <= o.b.buena.mm ? o.a : o.b, rec = don === o.a ? o.b : o.a;
        don.par = rec.par = true;
        pares.push({ cedis:o.a.cedis, tipo:o.a.tipo, donante:don, receptora:rec, dif:o.dif,
                     resultado:[rec.buena.mm, don.buena.mm] });
      });
    });
    // pedido por CEDIS · medida · eje
    var mapa = {};
    function g(cedis, tipo, ej){
      var k = cedis + "|" + tipo + "|" + ej;
      return mapa[k] = mapa[k] || { cedis:cedis, tipo:tipo, rin:TIPOS[tipo].rin, eje:ej, llantas:0, ahorro:0, pedir:0, unidades:{} };
    }
    singles.forEach(function(s){ var x = g(s.cedis, s.tipo, s.eje); x.llantas++; x.unidades[s.eco] = 1; });
    duales.forEach(function(d){ var x = g(d.cedis, d.tipo, "Trasero"); x.llantas += 2; x.unidades[d.eco] = 1; });
    pares.forEach(function(p){ g(p.cedis, p.tipo, "Trasero").ahorro += 2; });
    var pedido = Object.keys(mapa).map(function(k){ var x = mapa[k]; x.pedir = x.llantas - x.ahorro; x.nUnidades = Object.keys(x.unidades).length; return x; });
    pedido.sort(function(a, b){ return ORDEN.indexOf(a.cedis) - ORDEN.indexOf(b.cedis) || a.tipo.localeCompare(b.tipo) || a.eje.localeCompare(b.eje); });
    return { singles:singles, duales:duales, pares:pares, pedido:pedido };
  }

  /* ---------- propuesta de pedido y gasto contra presupuesto ----------
     Cuentas SAP: 6018302 llantas nuevas (posiciones 1-2 de M2/FL360 y todas las del Isuzu)
                  6018303 llantas vitalizadas (posiciones 3-6 de M2/FL360).
     El precio de cada llanta sale del catálogo de refacciones (Stock: valor libre ÷ piezas libres).
     Si el material no tiene existencia en ningún CEDIS no hay precio en stock y se usa un estimado editable. */
  var CUENTAS = { nuevas:{ clase:"6018302", nombre:"Llantas nuevas" }, vitalizadas:{ clase:"6018303", nombre:"Llantas vitalizadas" } };
  var CATEGORIAS = {
    NUEVA_M2:    { cuenta:"nuevas",      nombre:"Nueva M2 rin 22.5 (delanteras)",      mats:["535391","545421"], precioMat:"535391", respaldo:3235 },
    NUEVA_FL360: { cuenta:"nuevas",      nombre:"Nueva FL360 rin 17.5 (delanteras)",   mats:["518602"],          precioMat:"518602", respaldo:2200 },
    NUEVA_ISUZU: { cuenta:"nuevas",      nombre:"Nueva Isuzu rin 16 (las 4 llantas)",  mats:["503228"],          precioMat:"503228", respaldo:2800 },
    VIT_M2:      { cuenta:"vitalizadas", nombre:"Vitalizada M2 rin 22.5 (traseras)",   mats:[],                  precioMat:null,     respaldo:7100 },
    VIT_FL360:   { cuenta:"vitalizadas", nombre:"Vitalizada FL360 rin 17.5 (traseras)",mats:["519611"],          precioMat:"519611", respaldo:1800 }
  };
  var ORDEN_CAT = ["NUEVA_M2", "NUEVA_FL360", "NUEVA_ISUZU", "VIT_M2", "VIT_FL360"];
  function categoriaDe(x){
    if(x.tipo === "Isuzu") return "NUEVA_ISUZU";
    if(x.eje === "Delantero") return x.tipo === "FL360" ? "NUEVA_FL360" : "NUEVA_M2";
    return x.tipo === "FL360" ? "VIT_FL360" : "VIT_M2";
  }
  function cecoDe(cedis){ return cedis + "M0002"; }
  function indicesCat(d){ var c = d.cols, o = {}; ["mat","ce","libre","vlibre"].forEach(function(k){ o[k] = c.indexOf(k); }); return o; }
  // Precio por llanta de un material: valor libre ÷ piezas libres, sumando todos los CEDIS con existencia.
  function precioStock(d, mat){
    if(!d || !d.rows) return null;
    var i = indicesCat(d), v = 0, q = 0;
    d.rows.forEach(function(r){ if(r[i.mat] === mat && +r[i.libre] > 0 && +r[i.vlibre] > 0){ v += +r[i.vlibre]; q += +r[i.libre]; } });
    return q > 0 ? Math.round(v / q * 100) / 100 : null;
  }
  function existenciaStock(d, mats, cedis){
    if(!d || !d.rows) return 0;
    var i = indicesCat(d), n = 0;
    d.rows.forEach(function(r){ if(r[i.ce] === "M" + cedis && mats.indexOf(r[i.mat]) >= 0 && +r[i.libre] > 0) n += +r[i.libre]; });
    return n;
  }
  // Precio de cada categoría y de dónde sale. editados = { CATEGORIA: precio } puesto a mano.
  function preciosDe(d, editados){
    var out = {};
    ORDEN_CAT.forEach(function(k){
      var c = CATEGORIAS[k], p = c.precioMat ? precioStock(d, c.precioMat) : null;
      if(editados && editados[k] > 0) out[k] = { precio:editados[k], origen:"manual" };
      else if(p != null) out[k] = { precio:p, origen:"stock" };
      else out[k] = { precio:c.respaldo, origen:"estimado" };
    });
    return out;
  }
  // pedido = plan().pedido. Devuelve una fila por CEDIS y categoría con consumo (gasto) y compra (descontando stock).
  function propuesta(pedido, precios, d){
    var m = {};
    pedido.forEach(function(x){
      var cat = categoriaDe(x), k = x.cedis + "|" + cat;
      var r = m[k] = m[k] || { cedis:x.cedis, cat:cat, cuenta:CATEGORIAS[cat].cuenta, llantas:0, ahorro:0, pedir:0 };
      r.llantas += x.llantas; r.ahorro += x.ahorro; r.pedir += x.pedir;
    });
    var filas = Object.keys(m).map(function(k){
      var r = m[k];
      r.precio = precios[r.cat].precio; r.origen = precios[r.cat].origen;
      r.stock = existenciaStock(d, CATEGORIAS[r.cat].mats, r.cedis);
      r.comprar = Math.max(r.pedir - r.stock, 0);
      r.gasto = Math.round(r.pedir * r.precio);
      r.compra = Math.round(r.comprar * r.precio);
      return r;
    });
    filas.sort(function(a, b){ return ORDEN.indexOf(a.cedis) - ORDEN.indexOf(b.cedis) || ORDEN_CAT.indexOf(a.cat) - ORDEN_CAT.indexOf(b.cat); });
    return filas;
  }
  // presupuesto = { MCV:{ nuevas:{ plan, planMes, real }, vitalizadas:{...} } }. Una fila por CEDIS y cuenta.
  function contraPresupuesto(filas, presupuesto){
    var out = [];
    ORDEN.forEach(function(c){
      ["nuevas", "vitalizadas"].forEach(function(cu){
        var mias = filas.filter(function(f){ return f.cedis === c && f.cuenta === cu; });
        var p = presupuesto && presupuesto[c] && presupuesto[c][cu];
        var r = { cedis:c, cuenta:cu, pedir:0, gasto:0, comprar:0, compra:0, tiene:!!p, plan:p ? p.plan : null, planMes:p ? p.planMes : null, real:p ? p.real : null };
        mias.forEach(function(f){ r.pedir += f.pedir; r.gasto += f.gasto; r.comprar += f.comprar; r.compra += f.compra; });
        r.saldo = p ? p.plan - p.real : null;
        r.resultado = p ? r.saldo - r.gasto : null;      // positivo = sobra, negativo = se pasa
        r.estado = !p ? "sin" : r.resultado >= 0 ? "dentro" : "pasa";
        out.push(r);
      });
    });
    return out;
  }

  var API = {
    CUENTAS:CUENTAS, CATEGORIAS:CATEGORIAS, ORDEN_CAT:ORDEN_CAT, categoriaDe:categoriaDe, cecoDe:cecoDe,
    precioStock:precioStock, existenciaStock:existenciaStock, preciosDe:preciosDe, propuesta:propuesta, contraPresupuesto:contraPresupuesto,
    CEDIS:CEDIS, ORDEN:ORDEN, MESES:MESES, TIPOS:TIPOS,
    UMBRAL_CAMBIO:UMBRAL_CAMBIO, UMBRAL_LEGAL:UMBRAL_LEGAL, UMBRAL_VIGILAR:UMBRAL_VIGILAR, BUENA:BUENA, TASA:TASA, TOLERANCIA:TOLERANCIA,
    norm:norm, tipoDeMarca:tipoDeMarca, nueva:nueva, posActivas:posActivas, eje:eje, esDual:esDual, estado:estado,
    periodoDeTexto:periodoDeTexto, etiquetaPeriodo:etiquetaPeriodo, sumaMeses:sumaMeses, mesesEntre:mesesEntre,
    parsearLibro:parsearLibro, deBase:deBase, aBase:aBase, validas:validas,
    resumen:resumen, alertas:alertas, tasaReal:tasaReal, mesesRestantes:mesesRestantes, proyeccion:proyeccion, plan:plan
  };
  if(typeof module !== "undefined" && module.exports) module.exports = API;
  root.LLA = API;
})(typeof window !== "undefined" ? window : this);
