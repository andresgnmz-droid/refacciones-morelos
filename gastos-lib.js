/* Lectura de los archivos de SAP del módulo Gastos.
   Acepta lo que SAP baja con Lista > Exportar > Fichero local > "Texto sin tabuladores" (columnas separadas por tabulador):
   - KSB1  Partidas individuales de costes reales por centros de coste  → gasto real (una fila por partida)
   - KSBP  Partidas individuales de plan por centros de coste           → presupuesto (12 columnas "Val./MI", una por mes)
   Se reconoce cada archivo por sus encabezados, así que el orden de las columnas del layout no importa.
   Si cambias este archivo, sube el número de VERSION en sw.js. */
(function(root){
  "use strict";

  function quitaAcentos(s){ return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, ""); }
  function clave(s){ return quitaAcentos(s).replace(/\s+/g, " ").trim().toLowerCase(); }

  // Texto del archivo: SAP lo guarda en UTF-16 (con marca al inicio); si no, UTF-8 o Windows-1252.
  function decodificar(buf){
    var b = new Uint8Array(buf);
    if(b.length >= 2 && b[0] === 0x50 && b[1] === 0x4B) throw new Error("Este archivo es un Excel (.xlsx). Vuelve a descargarlo de SAP con Lista > Exportar > Fichero local > Texto sin tabuladores.");
    if(b.length >= 2 && b[0] === 0xD0 && b[1] === 0xCF) throw new Error("Este archivo es un Excel (.xls). Vuelve a descargarlo de SAP con Lista > Exportar > Fichero local > Texto sin tabuladores.");
    if(b.length >= 2 && b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder("utf-16le").decode(b.subarray(2));
    if(b.length >= 2 && b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder("utf-16be").decode(b.subarray(2));
    try{ return new TextDecoder("utf-8", { fatal: true }).decode(b); }
    catch(e){ return new TextDecoder("windows-1252").decode(b); }
  }

  // "1,234.56"  "-1,234.56"  "1,234.56-"  → número
  function numero(s){
    s = String(s == null ? "" : s).trim().replace(/,/g, "");
    if(s === "") return 0;
    var neg = false;
    if(s.charAt(s.length - 1) === "-"){ neg = true; s = s.slice(0, -1); }
    var n = Number(s);
    if(isNaN(n)) return NaN;
    return neg ? -n : n;
  }
  function redondea(n){ return Math.round(n * 100) / 100; }

  // "30.06.2026" → "2026-06-30"
  function fechaISO(s){
    var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(String(s || "").trim());
    if(!m) return null;
    var d = +m[1], mo = +m[2], y = +m[3];
    if(mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    return y + "-" + (mo < 10 ? "0" : "") + mo + "-" + (d < 10 ? "0" : "") + d;
  }

  var ALIAS = {
    clase:     ["cl.coste", "clase de coste"],
    clase_den: ["denom.cl.coste", "denom.clase de coste"],
    area:      ["denom.obj.", "denominacion objeto", "denominacion del objeto"],
    orden:     ["ordpartner", "orden partner"],
    orden_den: ["denominacion del objeto del interlocutor"],
    valor:     ["val./mi", "valor/mon.inf."],
    ceco:      ["ce.coste", "centro coste", "centro de coste"],
    fecha:     ["fe.contab.", "fecha contab.", "fe.contabilizacion", "fecha de contabilizacion"],
    reg:       ["registrado", "fecha entrada"],
    usuario:   ["usuario"],
    anio:      ["ano", "ejercicio"]
  };
  function buscaCol(cabecera, campo){
    var a = ALIAS[campo];
    for(var i = 0; i < cabecera.length; i++) if(a.indexOf(cabecera[i]) >= 0) return i;
    return -1;
  }

  // Devuelve { tipo:"real"|"plan", filas:[...], resumen:{...}, avisos:[...] }  (lanza Error si no es un archivo reconocible)
  function leer(texto){
    var lineas = String(texto).split(/\r?\n/);
    var iCab = -1, i;
    for(i = 0; i < lineas.length; i++){
      var cel = lineas[i].split("\t").map(clave);
      if(cel.indexOf("cl.coste") >= 0 || cel.indexOf("clase de coste") >= 0){ iCab = i; break; }
    }
    if(iCab < 0) throw new Error("No reconozco este archivo: no encuentro la columna de clase de coste (Cl.coste). ¿Es de KSB1 o KSBP?");
    var cab = lineas[iCab].split("\t").map(clave);
    var iClase = buscaCol(cab, "clase"), iCeco = buscaCol(cab, "ceco");
    if(iCeco < 0) throw new Error("Falta la columna del centro de coste (Ce.coste). Agrégala al layout en SAP y vuelve a bajar el archivo.");

    var colsMes = [];
    cab.forEach(function(c, k){ if(c === "val./mi") colsMes.push(k); });
    var esPlan = colsMes.length === 12;
    var avisos = [];

    // filas de datos: la clase de coste es un número de 6 o más cifras (se saltan los encabezados repetidos de cada página)
    var datos = [];
    for(i = iCab + 1; i < lineas.length; i++){
      var p = lineas[i].split("\t");
      if(p.length < cab.length - 1) continue;
      if(!/^\d{6,}$/.test((p[iClase] || "").trim())) continue;
      datos.push(p);
    }
    if(!datos.length) throw new Error("El archivo no trae renglones de datos.");
    function celda(p, k){ return k >= 0 ? (p[k] || "").trim() : ""; }

    var salida;
    if(esPlan){
      var iDen = buscaCol(cab, "clase_den"), iAnio = buscaCol(cab, "anio");
      var mapa = {}, orden = [], anios = {}, malos = 0, rangoRaro = 0;
      var iDe = cab.indexOf("de"), iA = cab.indexOf("a");
      datos.forEach(function(p){
        var m = colsMes.map(function(k){ return numero(p[k]); });
        if(m.some(isNaN)){ malos++; return; }
        var c = celda(p, iCeco), k = celda(p, iClase), id = c + "|" + k;
        if(iAnio >= 0 && celda(p, iAnio)) anios[celda(p, iAnio)] = 1;
        if(iDe >= 0 && iA >= 0 && (celda(p, iDe) !== "1" || celda(p, iA) !== "12")) rangoRaro++;
        if(mapa[id]){
          mapa[id].m = mapa[id].m.map(function(v, j){ return redondea(v + m[j]); });
          avisos.push("El renglón " + c + " · " + k + " viene repetido en el archivo; se sumaron sus valores.");
        } else {
          mapa[id] = { c: c, k: k, kd: celda(p, iDen), m: m.map(redondea) };
          orden.push(id);
        }
      });
      if(malos) avisos.push(malos + " renglones con importes que no se pudieron leer se omitieron.");
      if(rangoRaro) avisos.push(rangoRaro + " renglones traen un rango de periodos distinto de 1 a 12. El ERP usa el valor de cada mes tal como viene.");
      salida = orden.map(function(id){ return mapa[id]; });
      var cecos = {}, mes = new Array(12).fill(0);
      salida.forEach(function(r){ cecos[r.c] = (cecos[r.c] || 0) + 1; r.m.forEach(function(v, j){ mes[j] += v; }); });
      mes = mes.map(redondea);
      if(mes[10] === 0 && mes[11] === 0 && mes[9] !== 0) avisos.push("El plan no trae noviembre ni diciembre. En SAP corre KSBP con periodos 1 a 12.");
      if(datos.length === 50) avisos.push("El archivo trae exactamente 50 renglones: SAP corta la lista en 50 por defecto. Sube el máximo de resultados en Otras opciones y vuelve a correr.");
      var anioPlan = Object.keys(anios);
      return {
        tipo: "plan", filas: salida, avisos: avisos,
        resumen: { n: salida.length, cecos: cecos, porMes: mes, totalAnual: redondea(mes.reduce(function(a, b){ return a + b; }, 0)), anios: anioPlan }
      };
    }

    // gasto real (KSB1)
    var iValor = buscaCol(cab, "valor"), iFecha = buscaCol(cab, "fecha");
    if(iValor < 0 || iFecha < 0) throw new Error("Faltan columnas: se necesitan el valor (Val./MI) y la fecha de contabilización (Fe.contab.). Revisa el layout en SAP.");
    var iDenR = buscaCol(cab, "clase_den"), iArea = buscaCol(cab, "area"), iOrd = buscaCol(cab, "orden"), iOrdDen = buscaCol(cab, "orden_den"),
        iReg = buscaCol(cab, "reg"), iUsu = buscaCol(cab, "usuario");
    var filas = [], sinFecha = 0, sinValor = 0, cecosR = {}, porMes = {}, anioC = {}, total = 0, fmin = null, fmax = null;
    datos.forEach(function(p){
      var v = numero(p[iValor]), f = fechaISO(p[iFecha]);
      if(isNaN(v)){ sinValor++; return; }
      if(!f){ sinFecha++; return; }
      var c = celda(p, iCeco);
      filas.push({ c: c, k: celda(p, iClase), kd: celda(p, iDenR), a: celda(p, iArea), o: celda(p, iOrd), od: celda(p, iOrdDen),
                   f: f, v: redondea(v), u: celda(p, iUsu), r: fechaISO(celda(p, iReg)) || "" });
      cecosR[c] = (cecosR[c] || 0) + 1;
      var mm = +f.slice(5, 7); porMes[mm] = redondea((porMes[mm] || 0) + v);
      anioC[f.slice(0, 4)] = (anioC[f.slice(0, 4)] || 0) + 1;
      total += v;
      if(!fmin || f < fmin) fmin = f;
      if(!fmax || f > fmax) fmax = f;
    });
    if(sinValor) avisos.push(sinValor + " renglones con importe ilegible se omitieron.");
    if(sinFecha) avisos.push(sinFecha + " renglones sin fecha de contabilización válida se omitieron.");
    if(datos.length === 5000) avisos.push("El archivo trae exactamente 5,000 renglones: SAP corta la lista en 5,000 por defecto, así que probablemente faltan partidas. Sube el máximo de resultados en Otras opciones de KSB1 y vuelve a correr.");
    var anioMayor = Object.keys(anioC).sort(function(a, b){ return anioC[b] - anioC[a]; })[0];
    if(Object.keys(anioC).length > 1) avisos.push("El archivo mezcla fechas de más de un año (" + Object.keys(anioC).join(", ") + "). Corre KSB1 con un solo año.");
    return {
      tipo: "real", filas: filas, avisos: avisos,
      resumen: { n: filas.length, total: redondea(total), fmin: fmin, fmax: fmax, cecos: cecosR, porMes: porMes, anio: anioMayor ? +anioMayor : null }
    };
  }

  root.GASTOS_LIB = { decodificar: decodificar, leer: leer, numero: numero, fechaISO: fechaISO, redondea: redondea };
  if(typeof module !== "undefined" && module.exports) module.exports = root.GASTOS_LIB;
})(typeof window !== "undefined" ? window : globalThis);
