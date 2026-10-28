/* Sesión compartida del ERP.
   erp.html inicia sesión con Supabase Auth; este archivo deja que cada módulo (que corre dentro del ERP
   o abierto por separado en el mismo sitio) use esa misma sesión para leer y escribir en la base,
   y sepa qué puede editar el usuario. Si cambias este archivo, sube el número de VERSION en sw.js. */
(function(){
  "use strict";
  var URL = "https://aoutudbxwawlnbwvypir.supabase.co";
  var KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFvdXR1ZGJ4d2F3bG5id3Z5cGlyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODU4NzQzNTMsImV4cCI6MjEwMTQ1MDM1M30.cfdjH9rwQDYyxrrKZIAw7JpcQKYcFWtiEYXEDF8fa_E";
  var TK = "sb-aoutudbxwawlnbwvypir-auth-token";
  var PK = "erp_perfil";

  // false = los módulos sueltos siguen abriendo sin sesión (fase de transición).
  // true  = un módulo abierto sin sesión muestra un aviso para entrar por erp.html.
  var EXIGIR_SESION = false;

  function leer(k){ try{ return JSON.parse(localStorage.getItem(k) || "null"); }catch(e){ return null; } }

  function sesion(){
    var j = leer(TK);
    if(!j) return null;
    var s = j.currentSession || j;
    if(!s || !s.access_token) return null;
    if(s.expires_at && s.expires_at * 1000 < Date.now() + 5000) return null;
    return s;
  }
  function token(){ var s = sesion(); return s ? s.access_token : null; }
  function headers(extra){
    var h = { apikey: KEY, Authorization: "Bearer " + (token() || KEY), "Content-Type": "application/json" };
    if(extra) for(var k in extra) h[k] = extra[k];
    return h;
  }
  function perfil(){ return token() ? leer(PK) : null; }
  function nivel(modulo){
    var p = perfil();
    return (p && p.permisos && p.permisos[modulo]) || "oculto";
  }
  function puede(modulo){ return nivel(modulo) === "editar"; }
  function esAdmin(){ var p = perfil(); return !!(p && p.rol === "admin"); }

  // Las reglas de cálculo y los textos técnicos son solo para el administrador.
  // Marca un elemento con class="solo-admin" para que lo vea únicamente él;
  // con data-publico="texto" los demás ven ese texto (qué informa la pestaña) en lugar del original.
  document.addEventListener("DOMContentLoaded", function(){
    if(esAdmin()) return;
    var i, a = document.querySelectorAll(".solo-admin");
    for(i = 0; i < a.length; i++) a[i].parentNode.removeChild(a[i]);
    var b = document.querySelectorAll("[data-publico]");
    for(i = 0; i < b.length; i++) b[i].textContent = b[i].getAttribute("data-publico");
  });

  // Llamada REST con la sesión. opts: { method, body, prefer, count }
  function rest(path, opts){
    opts = opts || {};
    var h = headers();
    if(opts.prefer) h.Prefer = opts.prefer;
    return fetch(URL + "/rest/v1/" + path, {
      method: opts.method || "GET",
      headers: h,
      body: opts.body == null ? undefined : JSON.stringify(opts.body)
    }).then(function(r){
      if(!r.ok) return r.text().then(function(t){ throw new Error("HTTP " + r.status + (t ? " · " + t.slice(0, 160) : "")); });
      return r.status === 204 ? null : r.text().then(function(t){ return t ? JSON.parse(t) : null; });
    });
  }

  window.ERP = { URL: URL, KEY: KEY, token: token, headers: headers, perfil: perfil, nivel: nivel, puede: puede, esAdmin: esAdmin, rest: rest, exigir: EXIGIR_SESION };

  if(EXIGIR_SESION && !/(^|\/)erp\.html$/.test(location.pathname) && !token()){
    document.addEventListener("DOMContentLoaded", function(){
      var d = document.createElement("div");
      d.style.cssText = "position:fixed;inset:0;z-index:99999;background:#0F1620;color:#E7EDF2;display:flex;align-items:center;justify-content:center;font:15px system-ui,sans-serif;text-align:center;padding:24px";
      d.innerHTML = "<div><p style='margin:0 0 14px'>Para ver esta información inicia sesión en el ERP.</p><a href='erp.html' style='color:#35B6C2;font-weight:600'>Ir al inicio de sesión</a></div>";
      document.body.appendChild(d);
    });
  }
})();
