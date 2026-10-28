/* Lectura de las órdenes OM01 (correctivos) desde correctivos_datos (el mismo archivo de Carga de trabajo).
   La usan correctivos.html y la portada de Mantenimiento en erp.html. Si cambias este archivo, sube VERSION en sw.js. */
(function(){
  "use strict";
  var CEDIS = { MMCV:"Cuernavaca", MMCU:"Cuautla", MMJO:"Jojutla", MMPU:"Puente de Ixtla", MMAX:"Axochiapan", MMCC:"Cocoyoc" };
  var ORDEN_CEDIS = ["Cuernavaca","Cuautla","Jojutla","Puente de Ixtla","Axochiapan","Cocoyoc"];

  function serie(v){
    if(typeof v === "number" && v > 20000){ return new Date(Math.round((v - 25569) * 86400000)); }
    if(typeof v === "string" && v){ var d = new Date(v); return isNaN(d) ? null : d; }
    return null;
  }
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

  // Devuelve { ordenes:[{orden,texto,equipo,cedis,costo,fecha,mes,status}], movs:{orden:[...]}, meses:[...] }
  function leer(datos){
    var out = { ordenes: [], movs: {}, meses: [] };
    var om = hoja(datos, "OM01OM02");
    if(om){
      var hi = filaEnc(om, "Clase de orden");
      if(hi >= 0){
        var c = mapa(om[hi]);
        for(var i = hi + 1; i < om.length; i++){
          var r = om[i]; if(!r) continue;
          var clase = r[c["Clase de orden"]]; if(!clase || String(clase).trim() !== "OM01") continue;
          var orden = r[c["Orden"]]; if(orden == null || orden === "") continue;
          var f = serie(r[c["Fecha de referencia"]]);
          var costo = r[c["Costes tot.reales"]];
          out.ordenes.push({
            orden: String(orden).trim(),
            texto: r[c["Texto breve"]] || "",
            equipo: r[c["Equipo"]] || "",
            cedis: CEDIS[r[c["Centro"]]] || String(r[c["Centro"]] || "Sin CEDIS"),
            costo: typeof costo === "number" ? costo : 0,
            fecha: f,
            mes: f ? f.getUTCFullYear() + "-" + ("0" + (f.getUTCMonth() + 1)).slice(-2) : "",
            status: r[c["Status del sistema"]] || ""
          });
        }
      }
    }
    var mv = hoja(datos, "OM01OM02MOV");
    if(mv){
      var hm = filaEnc(mv, "Material");
      if(hm >= 0){
        var cm = mapa(mv[hm]);
        var propias = {}; out.ordenes.forEach(function(o){ propias[o.orden] = 1; });
        for(var j = hm + 1; j < mv.length; j++){
          var m = mv[j]; if(!m) continue;
          var oo = m[cm["Orden"]]; if(oo == null) continue; oo = String(oo).trim();
          if(!propias[oo]) continue;
          (out.movs[oo] = out.movs[oo] || []).push({
            material: m[cm["Material"]],
            texto: m[cm["Texto breve de material"]] || "(sin descripción)",
            cantidad: typeof m[cm["Cantidad"]] === "number" ? m[cm["Cantidad"]] : 0,
            importe: typeof m[cm["Importe ML"]] === "number" ? m[cm["Importe ML"]] : 0,
            um: m[cm["Un.medida de entrada"]] || "",
            fecha: serie(m[cm["Fecha de documento"]])
          });
        }
      }
    }
    var ms = {}; out.ordenes.forEach(function(o){ if(o.mes) ms[o.mes] = 1; });
    out.meses = Object.keys(ms).sort();
    return out;
  }

  window.CORR = { leer: leer, CEDIS: ORDEN_CEDIS };
})();
