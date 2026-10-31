/* Modo sin conexión del ERP Automotriz Morelos.
   Guarda en el equipo las páginas, las librerías y la última respuesta de Supabase,
   para poder consultar la información ya cargada cuando no hay internet.
   Si cambias este archivo, sube el número de VERSION. */
var VERSION = "v16";
var SHELL = "erp-paginas-" + VERSION;
var LIBS = "erp-librerias-" + VERSION;
var DATOS = "erp-datos-" + VERSION;

var API_HOSTS = ["aoutudbxwawlnbwvypir.supabase.co"];
var PAGINAS = ["./", "index.html", "preventivos.html", "catalogo-maestro.html", "hojalateria.html", "carga-mecanicos.html", "rendimientos.html", "erp.html", "historial.html", "unidades-piso.html", "faltantes.html", "catalogo-refacciones.html", "carta-peticion.html", "offline.js", "erp-auth.js", "correctivos.html", "correctivos-lib.js"];
var LIBRERIAS = [
  "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
  "https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js",
  "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"
];

self.addEventListener("install", function(e){
  e.waitUntil(Promise.all([
    caches.open(SHELL).then(function(c){
      return Promise.all(PAGINAS.map(function(u){ return c.add(new Request(u, { cache: "reload" })).catch(function(){}); }));
    }),
    caches.open(LIBS).then(function(c){
      return Promise.all(LIBRERIAS.map(function(u){
        return fetch(new Request(u, { mode: "cors", credentials: "omit" })).then(function(r){ if(r.ok) return c.put(u, r); }).catch(function(){});
      }));
    })
  ]).then(function(){ return self.skipWaiting(); }));
});

self.addEventListener("activate", function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k.indexOf("erp-") === 0 && [SHELL, LIBS, DATOS].indexOf(k) < 0; }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

function conTiempo(p, ms){
  return new Promise(function(ok, fail){
    var t = setTimeout(function(){ fail(new Error("tiempo agotado")); }, ms);
    p.then(function(r){ clearTimeout(t); ok(r); }, function(er){ clearTimeout(t); fail(er); });
  });
}

// Guarda la respuesta con la hora en que se guardó (los encabezados de fecha del servidor no son visibles para la página).
function guardar(cache, req, res){
  var h = new Headers(res.headers);
  h.set("x-guardado", new Date().toISOString());
  return res.clone().blob().then(function(b){
    return cache.put(req, new Response(b, { status: res.status, statusText: res.statusText, headers: h }));
  }).catch(function(){});
}

function avisarGuardado(hit){
  var fecha = hit.headers.get("x-guardado");
  self.clients.matchAll({ includeUncontrolled: true }).then(function(cs){
    cs.forEach(function(c){ c.postMessage({ type: "datos-guardados", fecha: fecha }); });
  });
}

// Primero la red; si falla o tarda, la copia guardada.
function redPrimero(req, nombre, espera, ignorarQuery, alUsarCopia, revalidar){
  return caches.open(nombre).then(function(cache){
    /* Las páginas se revalidan siempre contra el servidor (si no cambiaron, responde 304 y es rápido): así un cambio publicado se ve de inmediato y no hasta 10 minutos después. */
    var red = fetch(revalidar ? new Request(req, { cache: "no-cache" }) : req).then(function(res){
      if(res.status >= 500) throw new Error("servidor");
      if(res.status === 200) guardar(cache, req, res);
      return res;
    });
    return conTiempo(red, espera).catch(function(){
      return cache.match(req, { ignoreSearch: !!ignorarQuery, ignoreVary: true }).then(function(hit){
        if(hit){ if(alUsarCopia) alUsarCopia(hit); return hit; }
        return red;
      });
    });
  });
}

function buscarLibreria(req){
  var u = new URL(req.url);
  if(u.origin === self.location.origin) return fetch(req);
  return fetch(new Request(req.url, { mode: "cors", credentials: "omit" })).catch(function(){ return fetch(req); });
}

// Usa la copia guardada y la actualiza en segundo plano.
function copiaYActualiza(e, nombre){
  var req = e.request;
  return caches.open(nombre).then(function(cache){
    return cache.match(req.url, { ignoreVary: true }).then(function(hit){
      var red = buscarLibreria(req).then(function(res){
        if(res && (res.ok || res.type === "opaque")) cache.put(req.url, res.clone());
        return res;
      });
      if(hit){ e.waitUntil(red.catch(function(){})); return hit; }
      return red;
    });
  });
}

self.addEventListener("fetch", function(e){
  var req = e.request;
  if(req.method !== "GET" || req.headers.has("range")) return;
  var url = new URL(req.url);
  if(url.protocol !== "http:" && url.protocol !== "https:") return;

  if(API_HOSTS.indexOf(url.host) >= 0){
    e.respondWith(redPrimero(req, DATOS, 8000, false, avisarGuardado));
    return;
  }
  if(req.mode === "navigate" || (url.origin === self.location.origin && /\.html?$/.test(url.pathname))){
    e.respondWith(redPrimero(req, SHELL, 5000, true, null, true));
    return;
  }
  e.respondWith(copiaYActualiza(e, url.origin === self.location.origin ? SHELL : LIBS));
});
