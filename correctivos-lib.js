/* Lectura de las órdenes de mantenimiento desde correctivos_datos (el mismo archivo de Carga de trabajo).
   La usan correctivos.html y la portada de Mantenimiento en erp.html. Si cambias este archivo, sube VERSION en sw.js.

   Regla de fecha: el gasto se asigna al mes de la "Fecha de documento" de cada movimiento de material (pestaña OM01OM02MOV),
   no a la "Fecha de referencia" de la orden. Así cuadra con el control manual y una orden que cruza de mes reparte su gasto.
   Clases que se leen: OM01 (correctivo), OM02 (preventivo) y OM12 (proyectos especiales, para Hojalatería). OM03 no se usa. */
(function(){
  "use strict";
  var CEDIS = { MMCV:"Cuernavaca", MMCU:"Cuautla", MMJO:"Jojutla", MMPU:"Puente de Ixtla", MMAX:"Axochiapan", MMCC:"Cocoyoc" };
  var ORDEN_CEDIS = ["Cuernavaca","Cuautla","Jojutla","Puente de Ixtla","Axochiapan","Cocoyoc"];
  var CLASES = { OM01:1, OM02:1, OM12:1 };

  function serie(v){
    if(typeof v === "number" && v > 20000){ return new Date(Math.round((v - 25569) * 86400000)); }
    if(typeof v === "string" && v){ var d = new Date(v); return isNaN(d) ? null : d; }
    return null;
  }
  function mesDe(f){ return f ? f.getUTCFullYear() + "-" + ("0" + (f.getUTCMonth() + 1)).slice(-2) : ""; }
  function filaEnc(rows, nombre){
    for(var i = 0; i < rows.length; i++){
      if(rows[i] && rows[i].some(function(c){ return c != null && String(c).trim() === nombre; })) return i;
    }
    return -1;
  }
  function mapa(row){ var m = {}; row.forEach(function(h, i){ if(h != null) m[String(h).trim()] = i; }); return m; }
  function hoja(datos, nombre){
    var k = Object.keys(datos || {}).filter(function(n){ return n.toLowerCase() === nombre.toLowerCase(); })[0];
    return k ? datos[k] : null;
  }

  // Devuelve { ordenes:{orden:{...}}, lineas:[{orden,clase,cedis,equipo,texto,mes,fecha,importe,material,...}], movs:{orden:[...]}, meses:[...] }
  function leer(datos){
    var out = { ordenes: {}, lineas: [], movs: {}, meses: [] };
    var om = hoja(datos, "OM01OM02");
    if(om){
      var hi = filaEnc(om, "Clase de orden");
      if(hi >= 0){
        var c = mapa(om[hi]);
        for(var i = hi + 1; i < om.length; i++){
          var r = om[i]; if(!r) continue;
          var clase = r[c["Clase de orden"]]; clase = clase ? String(clase).trim() : "";
          if(!CLASES[clase]) continue;
          var orden = r[c["Orden"]]; if(orden == null || orden === "") continue;
          orden = String(orden).trim();
          var costo = r[c["Costes tot.reales"]];
          out.ordenes[orden] = {
            orden: orden, clase: clase,
            texto: r[c["Texto breve"]] || "",
            equipo: r[c["Equipo"]] || "",
            cedis: CEDIS[r[c["Centro"]]] || String(r[c["Centro"]] || "Sin CEDIS"),
            costo: typeof costo === "number" ? costo : 0,
            ref: serie(r[c["Fecha de referencia"]])
          };
        }
      }
    }
    var mv = hoja(datos, "OM01OM02MOV");
    if(mv){
      var hm = filaEnc(mv, "Material");
      if(hm >= 0){
        var cm = mapa(mv[hm]), ms = {};
        for(var j = hm + 1; j < mv.length; j++){
          var m = mv[j]; if(!m) continue;
          var oo = m[cm["Orden"]]; if(oo == null) continue; oo = String(oo).trim();
          var o = out.ordenes[oo]; if(!o) continue;
          var f = serie(m[cm["Fecha de documento"]]);
          var mov = {
            orden: oo, clase: o.clase, cedis: o.cedis, equipo: o.equipo, texto0: o.texto,
            material: m[cm["Material"]],
            texto: m[cm["Texto breve de material"]] || "(sin descripción)",
            cantidad: typeof m[cm["Cantidad"]] === "number" ? m[cm["Cantidad"]] : 0,
            importe: typeof m[cm["Importe ML"]] === "number" ? m[cm["Importe ML"]] : 0,
            um: m[cm["Un.medida de entrada"]] || "",
            fecha: f, mes: mesDe(f)
          };
          out.lineas.push(mov);
          (out.movs[oo] = out.movs[oo] || []).push(mov);
          if(mov.mes) ms[mov.mes] = 1;
        }
        out.meses = Object.keys(ms).sort();
      }
    }
    return out;
  }

  // Gasto de una clase de orden (o varias) en un mes ("" = todo), opcionalmente solo de ciertas órdenes.
  function suma(D, clases, mes, soloOrdenes){
    var cl = [].concat(clases), t = 0;
    D.lineas.forEach(function(l){
      if(cl.indexOf(l.clase) < 0) return;
      if(mes && l.mes !== mes) return;
      if(soloOrdenes && !soloOrdenes[l.orden]) return;
      t += l.importe;
    });
    return t;
  }

  window.CORR = { leer: leer, suma: suma, CEDIS: ORDEN_CEDIS };
})();
