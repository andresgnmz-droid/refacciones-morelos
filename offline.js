/* Aviso de modo sin conexión y registro del service worker (sw.js). */
(function(){
  if(!("serviceWorker" in navigator)) return;
  var KEY = "erp_ultima_conexion";
  var fechaGuardada = null;
  var barra = null;

  function guardarMarca(){ try{ localStorage.setItem(KEY, new Date().toISOString()); }catch(e){} }
  function leerMarca(){ try{ return localStorage.getItem(KEY); }catch(e){ return null; } }
  function formato(iso){
    try{ return new Date(iso).toLocaleString("es-MX", { day:"2-digit", month:"short", hour:"2-digit", minute:"2-digit" }); }
    catch(e){ return ""; }
  }

  function crearBarra(){
    if(barra) return barra;
    var st = document.createElement("style");
    st.textContent =
      "#erpSinConexion{position:fixed;z-index:99999;left:max(12px,env(safe-area-inset-left,0px));bottom:max(12px,env(safe-area-inset-bottom,0px));max-width:min(560px,calc(100vw - 24px));display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:10px 14px;border-radius:10px;background:#2b2a35;color:#fff;font:13px/1.4 system-ui,-apple-system,Segoe UI,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.35)}" +
      "#erpSinConexion[hidden]{display:none}" +
      "#erpSinConexion button{font:inherit;font-weight:600;color:#2b2a35;background:#fff;border:0;border-radius:7px;padding:5px 10px;cursor:pointer}";
    document.head.appendChild(st);
    barra = document.createElement("div");
    barra.id = "erpSinConexion";
    barra.setAttribute("role", "status");
    barra.hidden = true;
    document.body.appendChild(barra);
    return barra;
  }
  function mostrar(texto, boton){
    var b = crearBarra();
    b.textContent = "";
    var s = document.createElement("span");
    s.textContent = texto;
    b.appendChild(s);
    if(boton){
      var btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = boton;
      btn.addEventListener("click", function(){ location.reload(); });
      b.appendChild(btn);
    }
    b.hidden = false;
  }
  function ocultar(){ if(barra) barra.hidden = true; }

  function actualizar(){
    if(!navigator.onLine){
      var d = fechaGuardada || leerMarca();
      mostrar("Sin conexión. Estás viendo la información guardada" + (d ? " el " + formato(d) : "") + ".");
    } else if(fechaGuardada){
      mostrar("No se pudo actualizar. Estás viendo la información guardada el " + formato(fechaGuardada) + ".", "Reintentar");
    } else {
      ocultar();
    }
  }

  navigator.serviceWorker.addEventListener("message", function(e){
    if(!e.data || e.data.type !== "datos-guardados") return;
    var f = e.data.fecha;
    if(f && (!fechaGuardada || f < fechaGuardada)) fechaGuardada = f;
    actualizar();
  });
  window.addEventListener("offline", actualizar);
  window.addEventListener("online", function(){
    if(fechaGuardada || barra && !barra.hidden) mostrar("Conexión restablecida.", "Actualizar");
  });

  window.addEventListener("load", function(){
    navigator.serviceWorker.register("sw.js").catch(function(){});
    if(!navigator.onLine) actualizar();
    setTimeout(function(){ if(navigator.onLine && !fechaGuardada) guardarMarca(); }, 4000);
  });
})();
