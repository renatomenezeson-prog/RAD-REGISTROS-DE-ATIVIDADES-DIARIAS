<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>RAD Digital</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js" defer></script>
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);--bg:#f2f5f4;--card:#fff;--ink:#14211e;--mut:#5d6e69;--line:#d8e0dd;--brand:#0f5c4d;--on:#fff;--ok:#2e8b57;--warn:#c62828;--run:#e0a100;--sbw:260px}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0f1715;--card:#18231f;--ink:#e8efec;--mut:#9bb0a9;--line:#2b3a35;--brand:#4fc3a6;--on:#0f1715}}
:root[data-theme="dark"]{--bg:#0f1715;--card:#18231f;--ink:#e8efec;--mut:#9bb0a9;--line:#2b3a35;--brand:#4fc3a6;--on:#0f1715}
*,*:before,*:after{box-sizing:inherit}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
body{margin:0;background:var(--bg);color:var(--ink);font:17px/1.4 "Segoe UI",system-ui,sans-serif}
main{max-width:720px;margin:0 auto;padding:12px}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px;margin-bottom:12px}
input,select,textarea{width:100%;padding:13px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin-bottom:8px}
.btn{display:block;width:100%;padding:15px;border:0;border-radius:10px;background:var(--brand);color:var(--on);font:inherit;font-weight:700;margin-top:6px;cursor:pointer}
.btn.alt{background:none;color:var(--brand);border:2px solid var(--brand)}.btn.red{background:var(--warn);color:#fff}
.btn.nfc{background:linear-gradient(135deg,#0f5c4d 0%,#1a8a72 100%);color:#fff}
.btn.nfc:before{content:"📶 "}
.row{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:13px 4px;border-top:1px solid var(--line)}
.row b{font-weight:600}.s{font-size:14px;color:var(--mut)}
.pill{padding:3px 10px;border-radius:99px;font-size:14px;font-weight:700;color:#fff;white-space:nowrap;cursor:pointer}
.ok{background:var(--ok)}.run{background:var(--run);color:#222}.no{background:var(--warn)}.pen{background:var(--mut)}
.run-card{border:2px solid var(--run)}
.chips{display:flex;gap:6px;margin:8px 0;flex-wrap:wrap}.chips button{flex:1 1 40%;padding:12px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;font-size:14px;cursor:pointer}
.chips button.sel{background:var(--brand);color:var(--on);border-color:var(--brand)}
.bar{height:10px;background:var(--line);border-radius:6px;overflow:hidden;margin:6px 0}.bar i{display:block;height:100%;background:var(--ok)}
.tbl{overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:14px}th,td{text-align:left;padding:7px 6px;border-bottom:1px solid var(--line);white-space:nowrap}
video{width:100%;border-radius:8px;background:#000;margin-top:8px}
.sl{position:relative;background:#fff;color:#14213d;border:1.5px solid #1f4e8c;padding:16px 14px 20px;margin-bottom:14px;overflow:hidden}
.sl:before{content:"";position:absolute;top:0;right:0;width:64px;height:64px;background:linear-gradient(225deg,#36b3e8 0 30%,#3a709f 30% 62%,#6699c1 62% 80%,transparent 80%)}
.sl:after{content:"";position:absolute;bottom:0;left:0;width:48px;height:48px;background:linear-gradient(45deg,#36b3e8 0 30%,#3a709f 30% 62%,#6699c1 62% 80%,transparent 80%)}
.sl h2{color:#1f3864;text-align:center;text-transform:uppercase;font-size:17px;margin:10px 0}
.sl table th,.sl table td{border:1px solid #14213d;padding:6px 8px;color:#14213d;white-space:normal}
.sl ul{line-height:1.8;padding-left:22px;margin:6px 0}.sl .dt{font-size:12px;text-align:center}
@media print{header,nav,.noprint,#sidebar,#topbar{display:none!important}body{background:#fff}#main{margin-left:0!important}.sl{break-inside:avoid;page-break-after:always;margin:0}}
.hide{display:none!important}
mark{background:#ffe08a;color:#14211e;border-radius:3px;padding:0 1px}
.ativ-ed{border:1px solid var(--line);border-radius:8px;padding:8px;margin-bottom:6px;background:var(--bg)}
.ativ-ed input,.ativ-ed select{margin-bottom:4px}
.qrs{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:10px}.qrs div{text-align:center;font-size:13px}
#nfclive{display:none;padding:10px;border-radius:8px;background:rgba(15,92,77,.12);border:1px dashed var(--brand);color:var(--brand);font-weight:600;text-align:center;margin-top:8px}
#nfclive.on{display:block}
#nfclive.err{background:rgba(198,40,40,.12);border-color:var(--warn);color:var(--warn)}
.tag-row{display:flex;gap:6px;align-items:center;padding:6px 0;border-top:1px solid var(--line);font-size:13px}
.tag-row code{flex:1;background:var(--bg);padding:4px 6px;border-radius:6px;font-size:11px;word-break:break-all;overflow:hidden}
#login{position:fixed;inset:0;background:var(--bg);display:flex;align-items:flex-start;justify-content:center;padding:20px;z-index:100;overflow-y:auto}
.login-card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:22px;max-width:420px;width:100%;box-shadow:0 10px 40px rgba(0,0,0,.12);margin:auto}
.login-card h1{margin:0;font-size:22px;color:var(--brand);text-align:center;letter-spacing:.5px}
.login-card .sub{text-align:center;color:var(--mut);font-size:13px;margin:6px 0 18px}
.user-list{max-height:min(50vh,340px);overflow-y:auto;border:1px solid var(--line);border-radius:10px;padding:4px;margin-top:6px}
.user-list button{width:100%;text-align:left;padding:12px 14px;border:0;background:none;color:var(--ink);font:inherit;border-radius:8px;cursor:pointer;display:flex;align-items:center;gap:10px}
.user-list button:hover{background:var(--bg)}
.uav{width:34px;height:34px;border-radius:50%;background:var(--brand);color:var(--on);display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;flex:0 0 auto}
.un{flex:1;overflow:hidden}.un b{display:block;font-size:15px}.un span{font-size:12px;color:var(--mut)}
.keypad{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin:14px 0}
.keypad button{padding:18px 0;font-size:22px;font-weight:600;border:1px solid var(--line);border-radius:12px;background:var(--bg);color:var(--ink);cursor:pointer}
.keypad button.act{background:var(--brand);color:var(--on)}
#pinDots{display:flex;justify-content:center;gap:14px;margin:18px 0 4px;min-height:22px}
#pinDots i{width:16px;height:16px;border-radius:50%;border:2px solid var(--brand);display:inline-block}
#pinDots i.f{background:var(--brand)}
.pin-msg{text-align:center;font-weight:600;font-size:14px;margin-top:10px;min-height:20px;color:var(--mut)}
.pin-msg.err{color:var(--warn)}.pin-msg.ok{color:var(--ok)}
.login-back{margin-top:14px;text-align:center}
.login-back a{color:var(--mut);font-size:13px;cursor:pointer}
.empty-users{text-align:center;padding:20px;color:var(--mut);font-size:14px}
#langlogin{display:flex;justify-content:center;gap:6px;margin-bottom:14px;flex-wrap:wrap}
#langlogin button{padding:6px 12px;border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:99px;font:inherit;font-size:13px;cursor:pointer}
#langlogin button.on{background:var(--brand);color:var(--on);border-color:var(--brand)}
#shell{min-height:100vh}
#topbar{display:flex;align-items:center;gap:12px;padding:12px 16px;background:var(--brand);color:var(--on);position:sticky;top:0;z-index:11;font-weight:700}
#topbar #menuBtn{background:none;border:0;color:var(--on);font-size:24px;cursor:pointer}
#topbar .ttl{flex:1}#topbar .who{font-size:12px;opacity:.9}
#sidebar{position:fixed;top:0;left:0;bottom:0;width:var(--sbw);background:var(--card);border-right:1px solid var(--line);transform:translateX(-100%);transition:transform .22s ease;z-index:20;display:flex;flex-direction:column;overflow:hidden}
#sidebar.open{transform:translateX(0)}
.sidebar-head{padding:16px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;color:var(--brand);font-weight:700}
#sbClose{background:none;border:0;color:var(--mut);font-size:22px;cursor:pointer}
#sbNav{display:flex;flex-direction:column;padding:8px;gap:2px;overflow-y:auto;flex:1}
#sbNav button{width:100%;text-align:left;padding:12px 14px;border:0;background:none;color:var(--ink);font:inherit;font-size:15px;font-weight:600;border-radius:8px;cursor:pointer;display:flex;align-items:center;gap:10px}
#sbNav .navg{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--mut,#7a8a86);padding:12px 14px 4px;opacity:.8}
#sbNav button.on{background:var(--brand);color:var(--on)}
#sbNav button .ico{font-size:18px;width:22px;text-align:center}
.sidebar-foot{padding:12px;border-top:1px solid var(--line);font-size:13px}
.sidebar-foot .who{font-weight:700;margin-bottom:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sidebar-foot .perfil{color:var(--mut);font-size:12px;margin-bottom:10px}
.sidebar-foot .btn{padding:10px;font-size:14px;margin-top:6px}
#langbar{display:flex;gap:4px;flex-wrap:wrap;margin-bottom:8px}
#langbar button{padding:4px 8px;border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:99px;font:inherit;font-size:11px;cursor:pointer}
#langbar button.on{background:var(--brand);color:var(--on);border-color:var(--brand)}
#overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:15;display:none}
#overlay.on{display:block}
#main{padding:12px;padding-bottom:60px}
@media(min-width:820px){#topbar{display:none}#sidebar{transform:translateX(0)}#overlay{display:none!important}#main{padding:20px;max-width:900px;margin-left:var(--sbw);margin-right:auto}}
.acoes{background:linear-gradient(135deg,var(--card) 0%,var(--bg) 100%);border-left:4px solid var(--brand)}
.grid-acoes{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}
.grid-acoes .btn{margin:0;padding:13px 8px;font-size:14px}
@media(max-width:420px){.grid-acoes{grid-template-columns:1fr}}
.esp-linha{display:flex;gap:4px;align-items:center;margin-bottom:6px}
.esp-linha.off{opacity:.55}
.esp-linha input{flex:1;margin:0}
.esp-linha .mini-btn{padding:8px 10px;border:0;cursor:pointer}
.trn-card{border:1px solid var(--line);border-radius:10px;padding:12px;margin-bottom:10px;background:var(--bg)}
.trn-card h4{margin:0 0 4px 0;font-size:15px;color:var(--brand)}
.trn-meta{font-size:12px;color:var(--mut);margin-bottom:6px}
.trn-desc{font-size:14px;line-height:1.5;margin:8px 0}
.trn-part{font-size:12px;color:var(--mut);margin-top:6px}
.foto{cursor:pointer;border:2px dashed var(--line);border-radius:50%;width:72px;height:72px;display:flex;align-items:center;justify-content:center;overflow:hidden;flex:0 0 auto;background:var(--bg)}
.foto img{width:100%;height:100%;object-fit:cover}
.nr-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(90px,1fr));gap:6px;margin-top:6px}
.nr-grid button{padding:8px 4px;font-size:12px;border-radius:6px;border:1px solid var(--line);background:var(--bg);color:var(--ink);cursor:pointer}
.nr-grid button.on{background:var(--ok);color:#fff;border-color:var(--ok)}
.qrp{display:grid;place-items:center;padding:8px;background:#fff;border-radius:8px;margin:6px 0;min-height:160px}
</style>
</head>
<body>
<div id="login"><div class="login-card">
  <div id="langlogin"></div>
  <h1 id="loginTitle">RAD Digital</h1>
  <div class="sub" id="loginSub">Registro de Atividades Diárias</div>
  <div id="loginBody"></div>
</div></div>
<div id="shell" class="hide">
<header id="topbar"><button id="menuBtn" onclick="openMenu()">☰</button><span class="ttl">RAD Digital</span><span class="who" id="topWho"></span></header>
<aside id="sidebar">
<div class="sidebar-head"><span>RAD Digital</span><button id="sbClose" onclick="closeMenu()">×</button></div>
<nav id="sbNav"></nav>
<div class="sidebar-foot">
<div class="who" id="sbUser"></div>
<div class="perfil" id="sbPerfil"></div>
<div id="langbar"></div>
<div id="sbSim"></div>
<div id="sbBtns">
<button class="btn alt" onclick="telaAlterarPin()" id="btnAlterPin">🔑 Alterar meu PIN</button>
<button class="btn red" onclick="logout()" id="btnLogout">Sair</button>
</div>
</div>
</aside>
<div id="overlay" onclick="closeMenu()"></div>
<main id="main"><div id="app"></div></main>
</div>
<script>
/* ============ UTIL ============ */
function ld(k,d){try{var v=localStorage.getItem(k);return v?JSON.parse(v):d}catch(e){return d}}
function sv(k,v){try{localStorage.setItem(k,JSON.stringify(v))}catch(e){alert("Não foi possível salvar (armazenamento cheio?). "+(e.message||""))}}
function esc(s){return String(s==null?"":s).replace(/[&<>"']/g,function(c){return{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]})}
function $(s){return document.querySelector(s)}
function $all(s){return [].slice.call(document.querySelectorAll(s))}
function p2(n){return (n<10?"0":"")+n}
function dia(t){return new Date(t).toLocaleDateString("en-CA")}
function hm(t){return new Date(t).toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
function mk(t){return new Date(t).toLocaleDateString("en-CA").slice(0,7)}
function f1(n){return (Math.round(n*100)/100).toString().replace(".",",")}
function pct(a,b){return b?Math.round(a/b*100):0}
function pctStr(a,c){return a?(Math.round(c/a*1000)/10).toString().replace(".",",")+"%":"—"}
/* O sistema não informa férias: essa situação é exibida como vazia e contada como ativo */
function ehFerias(s){return /f[eé]ria/i.test(s||"")}
function sitExib(s){return ehFerias(s)?"":(s||"")}
function ehAfast(s){return /afast|licen|inss|doen/i.test(s||"")}
function ehDeslig(s){return /demit|deslig|rescis/i.test(s||"")}
function ehAtivo(s){return !ehAfast(s)&&!ehDeslig(s)}
function sitCor(s){if(!s)return"var(--mut)";var t=s.toLowerCase();if(/ativo|trabalhando/.test(t))return"var(--ok)";if(/afast|licen|inss|doen/.test(t))return"var(--warn)";if(/demit|deslig|rescis/.test(t))return"var(--mut)";return"var(--brand)"}

/* ============ BUSCA ============ */
var BUSCA={};
function normT(s){s=String(s==null?"":s).toLowerCase();try{s=s.normalize("NFD").replace(/[\u0300-\u036f]/g,"")}catch(e){}return s}
function termo(k){return normT(BUSCA[k]||"").trim()}
function bate(k){var t=termo(k);if(!t)return true;for(var i=1;i<arguments.length;i++){if(normT(arguments[i]).indexOf(t)>=0)return true}return false}
function campoBusca(k,ph){return '<div style="position:relative"><input id="busca_'+k+'" type="search" placeholder="🔍 '+esc(ph)+'" value="'+esc(BUSCA[k]||"")+'" oninput="buscar(\''+k+'\',this)" autocomplete="off" style="padding-right:44px">'+(BUSCA[k]?'<button onclick="BUSCA[\''+k+'\']=\'\';draw()" style="position:absolute;right:6px;top:6px;border:0;background:none;color:var(--mut);font-size:20px;cursor:pointer" title="Limpar">×</button>':'')+'</div>'}
function buscar(k,el){BUSCA[k]=el.value;var p=el.selectionStart;var y=window.scrollY;draw();window.scrollTo(0,y);var e=document.getElementById("busca_"+k);if(e){e.focus();try{e.setSelectionRange(p,p)}catch(x){}}}
function marca(txt,k){var t=termo(k),s=String(txt==null?"":txt);if(!t)return esc(s);var i=normT(s).indexOf(t);if(i<0)return esc(s);return esc(s.slice(0,i))+'<mark>'+esc(s.slice(i,i+t.length))+'</mark>'+esc(s.slice(i+t.length))}

/* ============ i18n ============ */
var IDIOMAS_UI=[
  {id:"pt-BR",nome:"🇧🇷 Português"},
  {id:"en",nome:"🇺🇸 English"},
  {id:"es",nome:"🇪🇸 Español"},
  {id:"fr",nome:"🇫🇷 Français"}
];
var IDIOMAS_COMUNICACAO=[
  {id:"pt-BR",nome:"Português brasileiro",ico:"🇧🇷"},
  {id:"en",nome:"Inglês",ico:"🇺🇸"},
  {id:"es",nome:"Espanhol",ico:"🇪🇸"},
  {id:"fr",nome:"Francês",ico:"🇫🇷"},
  {id:"ind",nome:"Língua indígena brasileira",ico:"🌿",extra:true},
  {id:"libras",nome:"Libras — Língua Brasileira de Sinais",ico:"🤟"},
  {id:"outro",nome:"Outro idioma ou língua de sinais",ico:"🌐"},
  {id:"nao",nome:"Não informado",ico:"❓"}
];
var TRAD={
  "pt-BR":{app_title:"RAD Digital",app_sub:"Registro de Atividades Diárias",login_tap:"Toque no seu nome para entrar",login_search:"🔍 Buscar pelo nome ou matrícula",login_nousers:"Nenhum usuário cadastrado.",login_nofound:"Nenhum usuário encontrado.",login_back:"← escolher outro usuário",login_welcome:"✓ Bem-vindo",login_wrong:"PIN incorreto",login_blocked:"Bloqueado por",login_min:"min.",sb_changePin:"🔑 Alterar meu PIN",sb_logout:"Sair",sb_simulate:"👁️ Ver como outro perfil",sb_back:"↩️ Voltar a",menu_rotina:"Rotina",menu_giro:"Giro",menu_hOP:"Chamados Operacionais",menu_hAD:"Chamados Administrativos",menu_cli:"Chamados do Cliente",menu_imp:"Importar Efetivo",menu_eq:"Equipe",menu_ct:"Contratos",menu_pti:"PTI",menu_plano:"Plano Geral (PTG)",menu_painel:"Painel",menu_trein:"Treinamentos",menu_aval:"Avaliação",menu_rel:"Relatório",menu_cfg:"Configuração do Sistema",menu_ind:"Indicadores",menu_qr:"QR Codes",menu_estoque:"Estoque",menu_req:"Requisições",menu_set:"Setores",menu_cont:"Contatos e Links",menu_perfil:"Meu Perfil",btn_save:"Salvar",btn_cancel:"Cancelar",btn_send:"Enviar",btn_close:"Fechar",btn_open:"Abrir",btn_edit:"Editar",btn_delete:"Excluir",btn_add:"Adicionar",btn_back:"Voltar",btn_next:"Avançar",btn_print:"Imprimir",btn_approve:"Aprovar",btn_reject:"Rejeitar",btn_reset:"Resetar",lang_select:"Idioma",lang_comunic:"Idioma preferencial de comunicação",lang_indigenous:"Qual língua indígena? (opcional)"},
  "en":{app_title:"RAD Digital",app_sub:"Daily Activities Log",login_tap:"Tap your name to enter",login_search:"🔍 Search by name or ID",login_nousers:"No users registered.",login_nofound:"No users found.",login_back:"← choose another user",login_welcome:"✓ Welcome",login_wrong:"Wrong PIN",login_blocked:"Blocked for",login_min:"min.",sb_changePin:"🔑 Change my PIN",sb_logout:"Log out",sb_simulate:"👁️ View as another profile",sb_back:"↩️ Back to",menu_rotina:"Routine",menu_giro:"Room Turnover",menu_hOP:"Operational Tickets",menu_hAD:"Administrative Tickets",menu_cli:"Client Tickets",menu_imp:"Import Staff",menu_eq:"Team",menu_ct:"Contracts",menu_pti:"IWP",menu_plano:"General Plan (PTG)",menu_painel:"Dashboard",menu_trein:"Training",menu_aval:"Evaluation",menu_rel:"Report",menu_cfg:"Settings",menu_ind:"Indicators",menu_qr:"QR Codes",menu_estoque:"Inventory",menu_req:"Requests",menu_set:"Departments",menu_cont:"Contacts & Links",menu_perfil:"My Profile",btn_save:"Save",btn_cancel:"Cancel",btn_send:"Send",btn_close:"Close",btn_open:"Open",btn_edit:"Edit",btn_delete:"Delete",btn_add:"Add",btn_back:"Back",btn_next:"Next",btn_print:"Print",btn_approve:"Approve",btn_reject:"Reject",btn_reset:"Reset",lang_select:"Language",lang_comunic:"Preferred communication language",lang_indigenous:"Which indigenous language? (optional)"},
  "es":{app_title:"RAD Digital",app_sub:"Registro de Actividades Diarias",login_tap:"Toque su nombre para entrar",login_search:"🔍 Buscar por nombre o matrícula",login_nousers:"Ningún usuario registrado.",login_nofound:"Ningún usuario encontrado.",login_back:"← elegir otro usuario",login_welcome:"✓ Bienvenido",login_wrong:"PIN incorrecto",login_blocked:"Bloqueado por",login_min:"min.",sb_changePin:"🔑 Cambiar mi PIN",sb_logout:"Salir",sb_simulate:"👁️ Ver como otro perfil",sb_back:"↩️ Volver a",menu_rotina:"Rutina",menu_giro:"Rotación de sala",menu_hOP:"Llamados Operativos",menu_hAD:"Llamados Administrativos",menu_cli:"Llamados del Cliente",menu_imp:"Importar Personal",menu_eq:"Equipo",menu_ct:"Contratos",menu_pti:"PTI",menu_plano:"Plan General (PTG)",menu_painel:"Panel",menu_trein:"Capacitación",menu_aval:"Evaluación",menu_rel:"Informe",menu_cfg:"Configuración",menu_ind:"Indicadores",menu_qr:"Códigos QR",menu_estoque:"Inventario",menu_req:"Solicitudes",menu_set:"Sectores",menu_cont:"Contactos y Enlaces",menu_perfil:"Mi Perfil",btn_save:"Guardar",btn_cancel:"Cancelar",btn_send:"Enviar",btn_close:"Cerrar",btn_open:"Abrir",btn_edit:"Editar",btn_delete:"Eliminar",btn_add:"Agregar",btn_back:"Volver",btn_next:"Siguiente",btn_print:"Imprimir",btn_approve:"Aprobar",btn_reject:"Rechazar",btn_reset:"Restablecer",lang_select:"Idioma",lang_comunic:"Idioma preferido de comunicación",lang_indigenous:"¿Cuál lengua indígena? (opcional)"},
  "fr":{app_title:"RAD Digital",app_sub:"Registre des Activités Quotidiennes",login_tap:"Touchez votre nom pour entrer",login_search:"🔍 Rechercher par nom ou matricule",login_nousers:"Aucun utilisateur enregistré.",login_nofound:"Aucun utilisateur trouvé.",login_back:"← choisir un autre utilisateur",login_welcome:"✓ Bienvenue",login_wrong:"PIN incorrect",login_blocked:"Bloqué pendant",login_min:"min.",sb_changePin:"🔑 Changer mon PIN",sb_logout:"Déconnexion",sb_simulate:"👁️ Voir comme autre profil",sb_back:"↩️ Retour à",menu_rotina:"Routine",menu_giro:"Rotation de salle",menu_hOP:"Tickets Opérationnels",menu_hAD:"Tickets Administratifs",menu_cli:"Tickets Client",menu_imp:"Importer le Personnel",menu_eq:"Équipe",menu_ct:"Contrats",menu_pti:"PTI",menu_plano:"Plan Général (PTG)",menu_painel:"Tableau de bord",menu_trein:"Formation",menu_aval:"Évaluation",menu_rel:"Rapport",menu_cfg:"Paramètres",menu_ind:"Indicateurs",menu_qr:"QR Codes",menu_estoque:"Inventaire",menu_req:"Demandes",menu_set:"Secteurs",menu_cont:"Contacts et Liens",menu_perfil:"Mon Profil",btn_save:"Enregistrer",btn_cancel:"Annuler",btn_send:"Envoyer",btn_close:"Fermer",btn_open:"Ouvrir",btn_edit:"Modifier",btn_delete:"Supprimer",btn_add:"Ajouter",btn_back:"Retour",btn_next:"Suivant",btn_print:"Imprimer",btn_approve:"Approuver",btn_reject:"Rejeter",btn_reset:"Réinitialiser",lang_select:"Langue",lang_comunic:"Langue de communication préférée",lang_indigenous:"Quelle langue autochtone ? (optionnel)"}
};
TRAD["pt-BR"].menu_ag="Agenda e Avisos";TRAD.en.menu_ag="Calendar & Notices";TRAD.es.menu_ag="Agenda y Avisos";TRAD.fr.menu_ag="Agenda et Avis";
TRAD["pt-BR"].menu_ar="Áreas do Contrato";TRAD.en.menu_ar="Contract Areas";TRAD.es.menu_ar="Áreas del Contrato";TRAD.fr.menu_ar="Zones du Contrat";
var IDIOMA_UI=ld("rad_idioma_ui","pt-BR");
function T(k){var d=TRAD[IDIOMA_UI]||TRAD["pt-BR"];return d[k]||TRAD["pt-BR"][k]||k}
function setIdiomaUI(v){
  IDIOMA_UI=v;sv("rad_idioma_ui",v);
  document.documentElement.lang=v;
  document.getElementById("loginTitle").textContent=T("app_title");
  document.getElementById("loginSub").textContent=T("app_sub");
  var bp=document.getElementById("btnAlterPin");if(bp)bp.textContent=T("sb_changePin");
  var bl=document.getElementById("btnLogout");if(bl)bl.textContent=T("sb_logout");
  renderLangLogin();renderLangBar();try{aplicarTema()}catch(e){}
  if(session){buildNav();draw()}else renderLogin();
}
function renderLangLogin(){
  var el=document.getElementById("langlogin");if(!el)return;
  el.innerHTML=IDIOMAS_UI.map(function(i){
    return '<button class="'+(i.id==IDIOMA_UI?"on":"")+'" onclick="setIdiomaUI(\''+i.id+'\')">'+i.nome+'</button>';
  }).join("");
}
function renderLangBar(){
  var el=document.getElementById("langbar");if(!el)return;
  el.innerHTML='<div class="s" style="width:100%;margin-bottom:2px">'+T("lang_select")+'</div>'
    +IDIOMAS_UI.map(function(i){
      return '<button class="'+(i.id==IDIOMA_UI?"on":"")+'" onclick="setIdiomaUI(\''+i.id+'\')">'+i.nome.split(" ")[0]+'</button>';
    }).join("");
}
function idiomaComunicNome(id){
  var x=IDIOMAS_COMUNICACAO.filter(function(i){return i.id==id})[0];
  return x?x.ico+" "+x.nome:"❓ Não informado";
}

/* ============ CARGOS ============ */
var ACOES_LIST=[
  {id:"ver_todos_contratos",nome:"Ver todos os contratos"},
  {id:"editar_areas",nome:"Editar áreas e espaços"},
  {id:"editar_colaboradores",nome:"Editar dados de colaboradores"},
  {id:"excluir_colaborador",nome:"Excluir colaboradores"},
  {id:"aprovar_pti_sup",nome:"Aprovar PTI (etapa Supervisor)"},
  {id:"aprovar_pti_coord",nome:"Aprovar PTI (etapa Coordenação)"},
  {id:"aprovar_pti_ger",nome:"Aprovar PTI (etapa Gerência)"},
  {id:"aprovar_pti_admin",nome:"Aprovar PTI (etapa Administrador)"},
  {id:"gerenciar_cargos",nome:"Gerenciar cargos e permissões"},
  {id:"gerenciar_usuarios",nome:"Gerenciar usuários e acessos"},
  {id:"gerenciar_contratos",nome:"Gerenciar contratos"},
  {id:"importar_efetivo",nome:"Importar efetivo"},
  {id:"ver_auditoria",nome:"Ver log de auditoria"},
  {id:"criar_avisos",nome:"Criar recados, compromissos e reuniões"},
  {id:"registrar_indicadores",nome:"Registrar indicadores de meses anteriores"}
];
var CARGOS=ld("rad_cargos",null);
if(!CARGOS){
  CARGOS=[
    {id:"admin",nome:"Administrador",cor:"#c62828",fixo:true,acessos:["*"],acoes:ACOES_LIST.map(function(a){return a.id})},
    {id:"ger",nome:"Gerência",cor:"#0f5c4d",fixo:false,acessos:["perfil","p","t","ct","ar","pti","plano","r","i","eq","CLI"],acoes:["ver_todos_contratos","aprovar_pti_ger","editar_areas"]},
    {id:"coord",nome:"Coordenação",cor:"#0f5c4d",fixo:false,acessos:["perfil","p","t","ct","ar","pti","plano","r","i","eq","CLI"],acoes:["ver_todos_contratos","aprovar_pti_coord","editar_areas"]},
    {id:"sup",nome:"Supervisor",cor:"#e0a100",fixo:false,acessos:["perfil","ar","c","p","t","ct","pti","plano","k","l","q","hOP","hAD","CLI","imp","eq","st","r","es","rq"],acoes:["aprovar_pti_sup","editar_areas","editar_colaboradores","importar_efetivo","gerenciar_contratos"]},
    {id:"enc",nome:"Encarregado",cor:"#e0a100",fixo:false,acessos:["perfil","l","k","hOP","hAD","p","t","rq"],acoes:[]},
    {id:"colab",nome:"Colaborador",cor:"#5d6e69",fixo:false,acessos:["perfil","c","g","hOP","hAD","rq"],acoes:[]},
    {id:"cli",nome:"Cliente",cor:"#5d6e69",fixo:false,acessos:["perfil","i","r","CLI"],acoes:[]}
  ];
  sv("rad_cargos",CARGOS);
  sv("rad_mig_cli",true);sv("rad_mig_ar",true);
}
/* Migração: cargos já salvos não tinham o menu de chamados do Cliente */
if(!ld("rad_mig_cli",false)){
  CARGOS.forEach(function(c){if(["sup","coord","ger","cli"].indexOf(c.id)>=0&&c.acessos.indexOf("*")<0&&c.acessos.indexOf("CLI")<0)c.acessos.push("CLI")});
  sv("rad_cargos",CARGOS);sv("rad_mig_cli",true);
}
if(!ld("rad_mig_ind",false)){
  CARGOS.forEach(function(c){c.acoes=c.acoes||[];if(["sup","coord","ger","admin"].indexOf(c.id)>=0&&c.acoes.indexOf("registrar_indicadores")<0)c.acoes.push("registrar_indicadores")});
  sv("rad_cargos",CARGOS);sv("rad_mig_ind",true);
}
if(!ld("rad_mig_es",false)){
  CARGOS.forEach(function(c){if(["enc","coord","ger","sup"].indexOf(c.id)>=0&&c.acessos.indexOf("*")<0&&c.acessos.indexOf("es")<0)c.acessos.push("es")});
  sv("rad_cargos",CARGOS);sv("rad_mig_es",true);
}
/* Migração: Agenda e Avisos para todos; criação para Supervisor/Coordenação/Gerência */
if(!ld("rad_mig_ag",false)){
  CARGOS.forEach(function(c){
    if(c.acessos.indexOf("*")<0&&c.acessos.indexOf("ag")<0)c.acessos.push("ag");
    c.acoes=c.acoes||[];
    if(["sup","coord","ger","admin"].indexOf(c.id)>=0&&c.acoes.indexOf("criar_avisos")<0)c.acoes.push("criar_avisos");
  });
  sv("rad_cargos",CARGOS);sv("rad_mig_ag",true);
}
/* Migração: menu "Áreas do Contrato" para Supervisor/Coordenação/Gerência */
if(!ld("rad_mig_ar",false)){
  CARGOS.forEach(function(c){
    if(["sup","coord","ger"].indexOf(c.id)<0||c.acessos.indexOf("*")>=0)return;
    ["ar"].concat(c.id=="sup"?["c","q"]:[]).forEach(function(k){if(c.acessos.indexOf(k)<0)c.acessos.push(k)});
    c.acoes=c.acoes||[];if(c.acoes.indexOf("editar_areas")<0)c.acoes.push("editar_areas");
  });
  sv("rad_cargos",CARGOS);sv("rad_mig_ar",true);
}
function svCargos(){sv("rad_cargos",CARGOS)}
function cargoById(id){return CARGOS.filter(function(c){return c.id==id})[0]||null}
function cargoDoUsuario(u){
  if(!u)return null;
  /* Simulação de perfil (somente para o usuário logado) */
  if(SIM&&session&&u.id==session.userId){var s=cargoById(SIM);if(s)return s}
  return cargoById(u.cargoId)||cargoById("colab");
}
function cargoAtual(){return cargoDoUsuario(session?userById(session.userId):null)}
function temAcao(id){
  var c=cargoAtual();
  if(!c)return false;
  return c.fixo||c.id=="admin"||(c.acoes||[]).indexOf(id)>=0;
}
function acessosDoUsuario(){
  var c=cargoAtual();
  if(!c)return ["perfil"];
  if(c.acessos.indexOf("*")>=0)return Object.keys(NAV).filter(function(k){return k!=="perfil"});
  var l=c.acessos.slice();if((c.acoes||[]).indexOf("gerenciar_usuarios")>=0&&l.indexOf("ac")<0)l.push("ac");
  if((l.indexOf("ct")>=0||l.indexOf("ar")>=0)&&l.indexOf("fc")<0)l.push("fc");
  return l;
}

/* ============ LOG ============ */
var LOG=ld("rad_log",[]);
function svLog(){sv("rad_log",LOG)}
function logAcao(acao,alvo,extra){
  try{
    LOG.push({t:Date.now(),uid:session?session.userId:null,user:user||"—",acao:acao,alvo:alvo||"",extra:extra||""});
    if(LOG.length>500)LOG=LOG.slice(-500);
    svLog();
  }catch(e){}
}

/* ============ USUÁRIOS ============ */
var USERS=ld("rad_users",[]);
function saveUsers(){sv("rad_users",USERS)}
function userById(id){for(var i=0;i<USERS.length;i++)if(USERS[i].id==id)return USERS[i];return null}
function userByName(n){n=(n||"").trim().toLowerCase();for(var i=0;i<USERS.length;i++)if((USERS[i].nome||"").toLowerCase()===n)return USERS[i];return null}
function iniciais(n){var p=(n||"").trim().split(/\s+/);return((p[0]||"?")[0]+((p[1]||"")[0]||"")).toUpperCase()}
if(!USERS.length){USERS.push({id:1,nome:"Administrador",mat:"",cargoId:"admin",pin:"0000",idiomaComun:"pt-BR",perfilCompleto:true,criado:Date.now()});saveUsers()}
var MAPA_PERFIL_CARGO={"Administrador":"admin","Gerência":"ger","Coordenação":"coord","Supervisor":"sup","Encarregado":"enc","Colaborador":"colab","Cliente":"cli"};
USERS.forEach(function(u){
  if(!u.cargoId)u.cargoId=MAPA_PERFIL_CARGO[u.perfil]||"colab";
  if(u.perfilCompleto===undefined)u.perfilCompleto=(u.cargoId!=="colab");
});
saveUsers();
function delUser(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var u=userById(id);if(!u)return;
  if(u.cargoId=="admin"){
    var admins=USERS.filter(function(x){return x.cargoId=="admin"});
    if(admins.length<=1){alert("É preciso manter pelo menos um Administrador.");return}
  }
  if(!confirm("Excluir o acesso de "+u.nome+"?"))return;
  USERS=USERS.filter(function(x){return x.id!=id});saveUsers();
  logAcao("Excluiu usuário",u.nome);
  if(session&&session.userId==id)logout(true);else draw();
}
function editarIdiomaComunicacao(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var u=userById(id);if(!u)return;
  var lista=IDIOMAS_COMUNICACAO.map(function(i,n){return (n+1)+" - "+i.ico+" "+i.nome}).join("\n");
  var atualIdx=IDIOMAS_COMUNICACAO.findIndex(function(i){return i.id==(u.idiomaComun||"nao")});
  var resp=prompt("Idioma preferencial de comunicação de "+u.nome+":\n\n"+lista+"\n\nDigite o número:",(atualIdx>=0?atualIdx+1:8));
  if(resp===null)return;
  var idx=+resp-1;
  if(!IDIOMAS_COMUNICACAO[idx]){alert("Opção inválida.");return}
  u.idiomaComun=IDIOMAS_COMUNICACAO[idx].id;
  if(u.idiomaComun==="ind"){var q=prompt("Qual língua indígena?","");u.idiomaIndigena=(q||"").trim()}
  else u.idiomaIndigena="";
  saveUsers();logAcao("Editou idioma do usuário",u.nome);draw();
}
function resetPinUser(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var u=userById(id);if(!u)return;
  var np=prompt("Novo PIN (4 dígitos) para "+u.nome+":");
  if(np===null)return;
  if(!/^\d{4}$/.test(np)){alert("O PIN deve ter 4 dígitos.");return}
  u.pin=np;u.pinTrocar=true;saveUsers();PINLOCK[id]={tentativas:0,ate:0};sv("rad_pinlock",PINLOCK);
  logAcao("Resetou PIN",u.nome);draw();if(confirm("PIN atualizado. Enviar a nova senha para "+u.nome+" por WhatsApp ou e-mail?"))enviarAcesso(u.id,np);
}
function editarCargoUsuario(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var u=userById(id);if(!u)return;
  var lista=CARGOS.map(function(c,i){return (i+1)+" - "+c.nome}).join("\n");
  var atual=CARGOS.findIndex(function(c){return c.id==u.cargoId});
  var resp=prompt("Cargo de "+u.nome+":\n\n"+lista+"\n\nDigite o número:",(atual>=0?atual+1:1));
  if(resp===null)return;
  var idx=+resp-1;
  if(!CARGOS[idx]){alert("Opção inválida.");return}
  if(u.cargoId=="admin"&&CARGOS[idx].id!="admin"&&USERS.filter(function(x){return x.cargoId=="admin"}).length<=1){alert("É preciso manter pelo menos um Administrador.");return}
  u.cargoId=CARGOS[idx].id;
  saveUsers();logAcao("Alterou cargo do usuário",u.nome,CARGOS[idx].nome);buildNav();draw();
}

/* ============ LOGIN / SESSÃO ============ */
var session=ld("rad_session",null);
var login={step:"user",busca:"",userId:null,pin:"",msg:"",msgTipo:""};
var perfil="Colaborador",user="",cid=ld("rad_cid","");
var SIM="",perfilReal="";
function simularPerfil(cargoId){
  if(!cargoId){
    if(perfilReal)perfil=perfilReal;
    perfilReal="";SIM="";
  }else{
    var cg=cargoById(cargoId);if(!cg)return;
    if(!perfilReal)perfilReal=perfil;
    SIM=cargoId;perfil=cg.nome;
  }
  var a=acessosDoUsuario();
  if(a.indexOf(tab)<0&&tab!="perfil"&&tab!="cont")tab=a[0]||"perfil";
  closeMenu();buildNav();draw();
  window.scrollTo(0,0);
}
function abrirSimulacao(){
  var sim=document.getElementById("sbSim");if(!sim)return;
  sim.innerHTML='<div class="s" style="margin:8px 0 4px;text-align:center">Ver como:</div>'
    +CARGOS.map(function(c){return '<button class="btn alt" style="padding:8px;font-size:13px;margin-top:4px" onclick="simularPerfil(\''+c.id+'\')">'+esc(c.nome)+'</button>'}).join("")
    +'<button class="btn" style="padding:8px;font-size:13px;margin-top:4px;background:var(--mut);color:#fff" onclick="fecharSimulacao()">'+T("btn_cancel")+'</button>';
}
function fecharSimulacao(){updateUserInfo()}
var PINLOCK=ld("rad_pinlock",{});
function pinLockInfo(id){var l=PINLOCK[id]||{tentativas:0,ate:0};if(l.ate&&Date.now()>=l.ate){l.tentativas=0;l.ate=0}return l}
function pinLockSalvar(id,l){PINLOCK[id]=l;sv("rad_pinlock",PINLOCK)}
function showLogin(){
  document.getElementById("login").classList.remove("hide");
  document.getElementById("shell").classList.add("hide");
  login.step="user";login.pin="";login.msg="";login.msgTipo="";login.busca="";
  renderLangLogin();renderLogin();
}
function loginBusca(el){
  login.busca=el.value;var p=el.selectionStart;renderLogin();
  var e=document.getElementById("searchU");if(e){e.focus();try{e.setSelectionRange(p,p)}catch(x){}}
}
function renderLogin(){
  var l=document.getElementById("loginBody"),h="";
  if(login.step==="user"){
    if(!USERS.length)h='<div class="empty-users">'+T("login_nousers")+'</div>';
    else{
      var term=login.busca.trim().toLowerCase();
      var oculta=ld("rad_sys",{loginOculto:true}).loginOculto!==false&&USERS.length>1&&normT(term).length<3;
      var lista=oculta?[]:USERS.filter(function(u){return !u.bloqueado}).filter(function(u){return !term||normT(u.nome||"").indexOf(normT(term))>=0||(u.mat||"").toLowerCase().indexOf(term)>=0});
      h='<div class="s" style="text-align:center;margin-bottom:8px">'+T("login_tap")+'</div>'
       +'<input id="searchU" placeholder="'+esc(T("login_search"))+'" value="'+esc(login.busca)+'" oninput="loginBusca(this)">'
       +'<div class="user-list">'+(lista.length?lista.map(function(u){
           var cg=cargoById(u.cargoId)||cargoById("colab");var lang=u.idiomaComun&&u.idiomaComun!=="pt-BR"?idiomaComunicNome(u.idiomaComun):"";
           return '<button onclick="loginPick('+u.id+')"><span class="uav">'+esc(iniciais(u.nome))+'</span><span class="un"><b>'+esc(u.nome)+'</b><span>'+esc(cg?cg.nome:"—")+(u.mat?' · mat. '+esc(u.mat):'')+(lang?' · '+esc(lang):'')+'</span></span></button>';
         }).join(""):'<div class="empty-users">'+(oculta?"Digite as 3 primeiras letras do seu nome":T("login_nofound"))+'</div>')+'</div>';
    }
  }else if(login.step==="pin"){
    var u=userById(login.userId);if(!u){showLogin();return}
    var cg=cargoById(u.cargoId)||cargoById("colab");
    h='<div style="text-align:center;margin-bottom:6px"><span class="uav" style="display:inline-flex;width:48px;height:48px;font-size:18px">'+esc(iniciais(u.nome))+'</span></div>'
     +'<div style="text-align:center;font-weight:700;font-size:18px">'+esc(u.nome)+'</div>'
     +'<div style="text-align:center;color:var(--mut);font-size:13px;margin-bottom:6px">'+esc(cg?cg.nome:"")+'</div>'
     +'<div id="pinDots">'+[0,1,2,3].map(function(i){return '<i class="'+(login.pin.length>i?"f":"")+'"></i>'}).join("")+'</div>'
     +'<div class="keypad">'
     +[1,2,3,4,5,6,7,8,9].map(function(n){return '<button onclick="pinKey(\''+n+'\')">'+n+'</button>'}).join("")
     +'<button onclick="pinKey(\'C\')" class="act">C</button><button onclick="pinKey(\'0\')">0</button><button onclick="pinKey(\'OK\')" class="act">OK</button></div>'
     +'<div class="pin-msg'+(login.msgTipo?" "+login.msgTipo:"")+'">'+esc(login.msg)+'</div>'
     +'<div class="login-back"><a onclick="loginBack()">'+T("login_back")+'</a></div>';
  }
  l.innerHTML=h;
}
function loginPick(id){login.userId=id;login.step="pin";login.pin="";login.msg="";login.msgTipo="";renderLogin();document.removeEventListener("keydown",pinKeyHandler);document.addEventListener("keydown",pinKeyHandler)}
function loginBack(){document.removeEventListener("keydown",pinKeyHandler);login.step="user";login.pin="";login.msg="";login.msgTipo="";login.busca="";renderLogin()}
function pinKeyHandler(e){
  if(login.step!=="pin")return;
  var el=document.getElementById("login");if(!el||el.classList.contains("hide"))return;
  var tag=((e.target&&e.target.tagName)||"").toUpperCase();
  if(tag==="INPUT"||tag==="TEXTAREA"||tag==="SELECT")return;
  if(/^[0-9]$/.test(e.key))pinKey(e.key);
  else if(e.key==="Backspace"){e.preventDefault();pinKey("C")}
  else if(e.key==="Enter")pinKey("OK");
}
var pinConferindo=false;
function pinKey(k){
  if(pinConferindo)return;
  if(k==="C"){login.pin="";login.msg="";login.msgTipo="";renderLogin();return}
  if(k==="OK"){if(login.pin.length===4)conferePin();return}
  if(login.pin.length>=4)return;
  login.pin+=k;renderLogin();
  if(login.pin.length===4){pinConferindo=true;setTimeout(function(){pinConferindo=false;conferePin()},120)}
}
function conferePin(){
  if(login.step!=="pin")return;
  var u=userById(login.userId);if(!u){showLogin();return}
  if(u.bloqueado){login.pin="";login.msg="Acesso desativado (desligado no efetivo).";login.msgTipo="err";renderLogin();return}
  var lk=pinLockInfo(u.id);
  if(lk.ate&&Date.now()<lk.ate){
    var rest=Math.ceil((lk.ate-Date.now())/60000);
    login.pin="";login.msg=T("login_blocked")+" "+rest+" "+T("login_min");login.msgTipo="err";
    renderLogin();return;
  }
  if(login.pin===u.pin){
    pinLockSalvar(u.id,{tentativas:0,ate:0});
    login.msg=T("login_welcome");login.msgTipo="ok";renderLogin();
    document.removeEventListener("keydown",pinKeyHandler);login.step="";
    var sess={userId:u.id,nome:u.nome,ts:Date.now()};
    sv("rad_session",sess);session=sess;
    u.ultimoAcesso=Date.now();saveUsers();
    setTimeout(function(){aplicaSession(sess);startApp();logAcao("Login")},220);
  }else{
    lk.tentativas=(lk.tentativas||0)+1;
    if(lk.tentativas>=5){
      lk.ate=Date.now()+5*60*1000;pinLockSalvar(u.id,lk);
      login.pin="";login.msg=T("login_blocked")+" 5 "+T("login_min");login.msgTipo="err";
    }else{
      pinLockSalvar(u.id,lk);
      login.pin="";login.msg=T("login_wrong")+" ("+lk.tentativas+"/5).";login.msgTipo="err";
    }
    renderLogin();
  }
}
function aplicaSession(s){
  var u=userById(s.userId);
  if(!u){logout(true);return}
  var cg=cargoById(u.cargoId)||cargoById("colab");
  perfil=cg?cg.nome:"Colaborador";
  user=u.nome||"";
  if(cg&&cg.id=="colab"){
    if(u.colId){cid=u.colId;sv("rad_cid",cid)}
  }
  ultimaAtiv=Date.now();
}
function logout(silent){
  if(!silent&&!confirm("Sair do sistema?"))return;
  logAcao("Logout");
  stopCam();stopNFC();
  session=null;try{localStorage.removeItem("rad_session")}catch(e){}
  try{document.body.style.background=""}catch(e){}
  document.removeEventListener("keydown",pinKeyHandler);
  if(perfilReal){perfil=perfilReal}
  perfilReal="";SIM="";user="";
  showLogin();
}
function telaAlterarPin(){
  closeMenu();
  var u=userById(session.userId);if(!u)return;
  var atual=prompt("PIN atual de "+u.nome+":");
  if(atual===null)return;
  if(atual!==u.pin){alert("PIN atual incorreto.");return}
  var n1=prompt("Novo PIN (4 dígitos):");
  if(!/^\d{4}$/.test(n1||"")){alert("O PIN deve ter 4 dígitos.");return}
  var n2=prompt("Confirme o novo PIN:");
  if(n1!==n2){alert("Os PINs não conferem.");return}
  u.pin=n1;u.pinTrocar=false;saveUsers();PINLOCK[u.id]={tentativas:0,ate:0};sv("rad_pinlock",PINLOCK);
  logAcao("Alterou próprio PIN");alert("PIN atualizado com sucesso.");
}
var SESSAO_MIN=ld("rad_sessmin",30),ultimaAtiv=Date.now();
function setSessaoMin(v){if(!canEdit()){alert("Sem permissão.");return}SESSAO_MIN=Math.max(0,+v||0);sv("rad_sessmin",SESSAO_MIN);draw();alert(SESSAO_MIN===0?"Expiração desativada.":"Sessão expira após "+SESSAO_MIN+" min.")}
["click","keydown","touchstart","scroll"].forEach(function(ev){document.addEventListener(ev,function(){ultimaAtiv=Date.now()},{passive:true,capture:true})});
setInterval(function(){if(!session)return;if(SESSAO_MIN<=0)return;var min=(Date.now()-ultimaAtiv)/60000;if(min>=SESSAO_MIN){logout(true);alert("Sessão expirada.")}},30000);

/* ============ MENU ============ */
function openMenu(){document.getElementById("sidebar").classList.add("open");document.getElementById("overlay").classList.add("on")}
function closeMenu(){document.getElementById("sidebar").classList.remove("open");document.getElementById("overlay").classList.remove("on")}
function updateUserInfo(){
  document.getElementById("sbUser").textContent=user||"—";
  var u=session?userById(session.userId):null;
  var cg=cargoDoUsuario(u);
  document.getElementById("sbPerfil").textContent=(cg?cg.nome:perfil)+(SIM?" (simulado)":"");
  document.getElementById("topWho").textContent=(user||"").split(" ")[0]+" · "+(cg?cg.nome.charAt(0):"C");
  renderLangBar();
  var btns=document.getElementById("sbBtns");
  if(btns&&!document.getElementById("btnMeuPerfil")){
    var b=document.createElement("button");
    b.id="btnMeuPerfil";b.className="btn alt";
    b.style.cssText="padding:10px;font-size:14px;margin-top:6px";
    b.onclick=function(){go("perfil")};
    btns.insertBefore(b,btns.firstChild);
  }
  var bmp=document.getElementById("btnMeuPerfil");if(bmp)bmp.textContent="👤 "+T("menu_perfil");
  var sim=document.getElementById("sbSim");
  if(sim){
    if(perfilReal){sim.innerHTML='<button class="btn" style="background:var(--run);color:#222;border:0;padding:10px;font-size:13px;margin-top:8px;font-weight:700" onclick="simularPerfil(\'\')">'+T("sb_back")+' '+esc(perfilReal)+'</button>';if(btns)btns.style.display="none"}
    else if(u&&u.cargoId=="admin"){sim.innerHTML='<button class="btn alt" style="padding:10px;font-size:13px;margin-top:8px" onclick="abrirSimulacao()">'+T("sb_simulate")+'</button>';if(btns)btns.style.display=""}
    else{sim.innerHTML="";if(btns)btns.style.display=""}
  }
}


/* ============ DADOS OPERACIONAIS ============ */
var LOC_PADRAO=[
["1º Andar - CEN","Sala Entrega Material|Corredor Lavabos|DML|REPAE|Sala 1|Sala 2|Sala 3|Sala 4|Sala 5|Sala ADM|Arsenal - Esterilização|Vestiário F - Funcionário|Vestiário M - Funcionário|Expurgo|Sala de Lavação|Copa|Corredor Admissão|Sala CME|Sala VIP|Corredor Laser|Banheiro - Fora|Vestiário F|Vestiário M|Rouparia|Sala Enfermagem"],
["Recepção Principal","Recepção Principal|Sala Supervisão|Guichê Estacionamento|Sala T.I|SAME|Banheiro M|Banheiro F|Elevador|Consultório 1|Consultório 2|Sala Espera PA|Toalete PA|Corredor Principal|DML|Banheiro Restaurante M|Banheiro Restaurante F|Entrada Estacionamento|Entrada Recp Principal"],
["1º Andar - CD1","Recepção Exames e Diagnóstico|Sala de Argônio|Sala Supervisão|Corredor|Banheiro Funcionários|Banheiro|Frente Elevador|Angiografia|Campimetria|Optomap|Serviços|Olhos Secos|Sala de Impressão|Sala de Curativo 1|Sala de Curativo 2|Ultrason|Corredor Admissão|Farmácia|Almoxarifado|Suprimentos|Recebimento|Bolha Farmácia|Retinografia|OCT|Recepção Agendamento CC|Guarita Colaboradores"],
["1º Andar - LASI","Banheiro F|Banheiro M|Escada|Recepção Exames|Gerência Assistencial|Tesouraria|Sala Orientação 1|Sala Orientação 2|Sala Orientação 3|Sala Orientação 4|Sala de Espera CC|Banheiro Médico|Sala Anestesia|Sala da Fabi|Recepção Laser|Sala Orient. Cirurgia|Administrativo|Supervisão|Sala Orient. Cirurgia 2|Sala Maria|Autorização|Central de Comunicação|Corredor Médico|Copinha|Sala de Preparo|Laser - Pós|Laser - Escovação|Laser - Estoque Material|DML Laser|DML"],
["2º Andar - COF (café)","Copa Segundo Andar|Máquina de Café 2º Andar|Máquina de Café Sala VIP|Máquina de Café Recepção Exames|Máquina de Café Lente|Máquina de Café Recp Dr Emir|Máquina de Café Córnea|Máquina de Café Retina|Dispensa|DML + Carrinho|Bebedouros|Geladeiras"],
["2º Andar - COF (consultórios)","Consultório 1|Consultório 2|Consultório 3|Consultório 4|Consultório 5|Consultório 6|Consultório 7|Consultório 8|Consultório 9|Consultório 10|DML|Conforto Médico|Banheiro Conforto|Corredor|Escada|Sala Gerência|Sala Coordenação|Vestiário M|Vestiário F|Banheiro PCD|Frente Elevador|Recepção Córnea|Biometria|YAG Laser|Topografia|Sala Inclusão|Recepção 2"],
["2º Andar - LEN","Recepção Dr Emir|Banheiro Social F|Banheiro Social M|Consultório 11|Consultório 12|Consultório 13|Recepção CD2|Escada|DML|Recepção Lente|Consultório 45|Consultório 46|Consultório 47|Consultório 48|Financeiro Lente|Apoio Logístico|Administrativo|Teste de Lente|Copa|Banheiro Funcionários|Banheiro Médica|Consultório 44|Topografia|Biometria|Yag Laser"],
["3º Andar - EST","Recepção Especialidades|Escada|Banheiro Feminino|Banheiro Masculino|Consultório 23|Consultório 24|Recepção Plástica|Consultório 18|Consultório 19|Consultório 20|Consultório 21|Consultório 22|Recepção Estética|ADM Estética|Corredor|Consultório 1|Consultório 2|Procedimento 1|Procedimento 2|Psicologia|Hall Recepção|DML|Consultório 26|Consultório 27|Consultório 28|Consultório 36|Consultório 37|Consultório 38|Consultório 39|Consultório 40"],
["3º Andar - RET","Recepção Retina|Banheiro Médico|Banheiro Funcionários|Banheiro Conforto|Banheiro Corredor|Conforto Médico|Sala de Reunião|Consultório 41|Consultório 42|Consultório 43|Corredor|DML|CARE|Consultório 25|Consultório 29|Consultório 30|Consultório 31|Consultório 32|Consultório 33|Consultório 34|Consultório 35|Recepção Pediatria|Brinquedoteca|Escada|Sala Adamo"],
["4º Andar - ADM","Recepção - Corredor|Copa 1-2|Escada - Elevador|DML|Sala Reunião|Financeiro|Comercial 1|Telefonia|Copinha|Banheiro M|Banheiro F|CSI|Manutenção|RH|Qualidade|Diretora|Corredores|Recepção Neurovie|Copa Corredor Neurovie|WC Neurovie|Cons 1 - Neurovie|Cons 2 - Neurovie|Cons 3 - Neurovie|Cons 4 - Neurovie|WC F - PA|WC M - PA|WC Família - PA|Recepção PA"],
["Área Externa","Guarita Entrada Funcionários|Escada Área de Terceiros|Sala de Descanso Colaboradores|Sala Gávea|Banheiro Gávea|Oficina|Corredor Área Terceiros|Casa Lixo Comum|Casa Lixo Reciclável|Casa Lixo Infectante|Casa Perfuro|Lavanderia"]
].map(function(l){return{n:l[0],e:l[1].split("|")}});
/* CORREÇÃO: as áreas editadas eram salvas em rad_loc2 mas nunca recarregadas */
var LOC=ld("rad_loc2",null);
if(!Array.isArray(LOC)||!LOC.length)LOC=LOC_PADRAO;

var TRN=ld("rad_trn",[]);
var TRN_CAT=["Integração","Reciclagem","EPI","Produtos químicos","Biossegurança","Segurança do trabalho","Outros"];
var EXTRAS=["Limpeza de vidros","Limpeza de corrimãos","Lavar carrinho","Revisão de todos os andares (após 15h)","Limpeza de janelas","Mecânica no corredor","Outra atividade"];
var MOT=["Sem autorização","Área interditada (obra/manutenção)","Área em uso (reunião/atendimento)","Ambiente trancado","Outro"];
var MOT_GIRO=["Sala em uso (cirurgia em andamento)","Sem autorização médica","Sala interditada pela CCIH","Equipe médica não liberou","Paciente em recuperação","Falta material para o giro","Equipamento com defeito impede giro","Outro (CC)"];
var TIPOS=[["CO","Concorrente"],["TE","Terminal"],["PR","Manutenção/Programada"],["LA","Lavação"]];
var STAT=["Pendente","Reagendado","Realizado com sucesso","Realizado com observações","Não realizado"];
var rec=false;

/* ============ ATIVIDADES POR ÁREA (PG) — usadas no PTI e no Plano Geral ============ */
/* CORREÇÃO: PG, pgCard e colEditor eram usados mas não existiam (erro ao abrir PTI/Configurações) */
var ATIV_PADRAO=["Limpeza concorrente","Limpeza terminal","Recolhimento de resíduos","Reposição de insumos","Limpeza de pisos","Limpeza de sanitários","Limpeza de superfícies de contato"];
var PG=ld("rad_pg",null);
if(!Array.isArray(PG)){
  PG=[];var _pgId=1;
  LOC.forEach(function(l,i){ATIV_PADRAO.forEach(function(n){PG.push({id:_pgId++,l:i,nome:n})})});
  sv("rad_pg",PG);
}
/* ---------- "Como fazer" padrão das atividades ----------
   Passo a passo de referência (boas práticas de limpeza e Manual ANVISA de limpeza e desinfecção
   de superfícies). Preenchido automaticamente pelo nome da atividade; o supervisor pode editar. */
var COMO_FAZER=[
 [/terminal/,"Ambiente desocupado (programada ou após alta/transferência). 1) Use todos os EPIs da área. 2) Retire resíduos e roupas. 3) Limpe de cima para baixo: teto, luminárias, saídas de ar, paredes, janelas e portas. 4) Limpe e desinfete mobiliário e equipamentos, inclusive partes de baixo. 5) Lave o piso com dois baldes (solução e enxágue) ou máquina, do fundo para a saída. 6) Aplique o desinfetante padronizado nas superfícies de contato. 7) Reponha insumos e registre."],
 [/concorrente/,"Com o ambiente em uso, todos os dias. 1) Use os EPIs da área. 2) Recolha o lixo e troque os sacos. 3) Limpe as superfícies de contato (maçanetas, interruptores, bancadas, mesas, equipamentos) com pano úmido e detergente; depois aplique o desinfetante padronizado. Trabalhe do mais limpo para o mais sujo e de cima para baixo. 4) Limpe o piso com MOP/pano úmido, do fundo para a saída. 5) Reponha insumos. 6) Sinalize piso molhado."],
 [/desinfec/,"Sempre depois da limpeza com detergente. Aplique o desinfetante padronizado (ex.: álcool 70% por fricção ou quaternário de amônio, conforme a ficha técnica) em maçanetas, corrimãos, interruptores, grades de cama, bancadas e botões. Respeite o tempo de contato do produto. Use pano exclusivo e troque-o ao mudar de ambiente."],
 [/infectant|perfuro/,"Luva de borracha cano longo, avental impermeável e bota. Feche o saco branco leitoso com até 2/3 da capacidade, sem apertar e sem retirar o ar com as mãos; nunca passe resíduo de um saco para outro. Caixa de perfurocortante só até a linha tracejada. Leve ao abrigo no carro de coleta fechado, pelo roteiro definido. Higienize o carro e as luvas ao final."],
 [/coleta seletiva|reciclav/,"Separe papel, plástico, metal e vidro nos coletores identificados; não misture orgânico. Feche os sacos com até 2/3 da capacidade e leve ao ponto de coleta seletiva."],
 [/residuo|lixo/,"Feche o saco com até 2/3 da capacidade, sem apertar. Troque o saco, limpe a lixeira por fora (e por dentro na limpeza terminal) e tampe. Separe recicláveis conforme a coleta seletiva. Leve ao abrigo/casa de lixo pelo trajeto definido."],
 [/sanitari|banheiro|vestiari/,"1) Sinalize a área. 2) Recolha o lixo. 3) Aplique o produto no vaso e no mictório e deixe agir. 4) Limpe pia, torneira, espelho e dispensadores. 5) Escove o vaso por dentro e limpe por fora, assento e descarga. 6) Limpe paredes, divisórias e portas. 7) Lave ou limpe o piso por último, do fundo para a porta. 8) Reponha papel, sabonete e álcool gel. Use panos de cor exclusiva para o banheiro."],
 [/reposic|insumo/,"Confira e reponha papel higiênico, papel toalha, sabonete líquido e álcool gel. Troque o refil de sabonete, nunca complete. Falta de material: abra chamado operacional."],
 [/vidro|janela|fachada/,"Use limpa-vidros e rodo, de cima para baixo, retirando o excesso com pano de microfibra. Acima de 2 metros de altura, somente com treinamento NR-35 e equipamento adequado."],
 [/elevador/,"Botoeiras com pano levemente umedecido em álcool 70% (nunca borrife direto). Espelho e paredes com limpa-vidros ou pano úmido. Piso com MOP úmido e limpeza do trilho da porta."],
 [/garagem/,"Varra ou use a lavadora. Manchas de óleo: absorvente, depois desengraxante e enxágue. Esvazie lixeiras e mantenha a sinalização de piso molhado."],
 [/oleo|pisos? industria/,"Contenha o derramamento com material absorvente, recolha e descarte como resíduo contaminado. Aplique desengraxante, esfregue, enxágue e seque. Sinalize a área até secar."],
 [/piso/,"Em áreas de saúde não varra a seco: use MOP úmido. Nas demais, varra ou aspire antes. Lave com detergente e enxágue, ou passe MOP úmido, do fundo para a saída. Sinalize piso molhado e deixe secar."],
 [/mesa|esta[cç][aã]o|carteira/,"Não mova documentos sem autorização. Tampo com pano úmido e detergente neutro. Teclado, mouse e telefone com pano levemente umedecido em álcool 70% (nunca borrife). Limpe as cadeiras e organize o ambiente."],
 [/quadro|lousa/,"Apague o quadro, limpe com pano úmido ou produto próprio para o tipo de quadro e limpe a canaleta do apagador."],
 [/sala de aula|salas de aula/,"Recolha o lixo, limpe carteiras e mesas, apague e limpe os quadros, limpe maçanetas e interruptores, e por último o piso, do fundo para a porta."],
 [/laborat/,"Bancadas e pias com detergente e depois desinfetante. Não toque em reagentes, peças ou equipamentos ligados. Resíduos químicos e biológicos somente conforme orientação do responsável do laboratório."],
 [/copa|refeit|cantina/,"Limpe bancadas, pia, mesas, micro-ondas e cafeteira. Geladeira por fora diariamente e por dentro conforme cronograma. Use pano exclusivo da copa. Descarte resíduos orgânicos e lave o piso por último."],
 [/vitrine|expositor/,"Pano de microfibra e limpa-vidros sem amônia, sem borrifar sobre as peças. Nunca abra vitrines ou expositores sem o acompanhamento de um responsável da loja."],
 [/cofre|tesouraria|restrit|acompanhad/,"Somente acompanhado por um responsável do cliente. Não mova valores, documentos ou equipamentos. Registre o horário e o nome de quem acompanhou."],
 [/autoatendimento|terminai|\batm\b/,"Pano levemente umedecido em álcool 70% na tela, no teclado e no leitor de cartão; nunca borrife líquido. Não toque nos compartimentos de cédulas e envelopes."],
 [/porta giratoria/,"Vidros com limpa-vidros, de cima para baixo; estrutura e pegadores com pano úmido e álcool 70%. Não force o mecanismo."],
 [/guich|balc/,"Pano úmido com detergente neutro no balcão e superfícies de atendimento; álcool 70% em canetas, interfones e pontos de contato. Não mova documentos."],
 [/carrinho|cesta/,"Alças e pegadores com álcool 70% ou desinfetante; estrutura com pano úmido e detergente."],
 [/provador/,"Recolha cabides e peças esquecidas, limpe espelho, banco e maçaneta, e passe MOP no piso."],
 [/caixa/,"Esteira, balcão e leitor/máquina de cartão com pano levemente umedecido em álcool 70%, sem borrifar sobre equipamentos."],
 [/jardin|poda|roçad|ro[cç]ad/,"Rega, poda, retirada de folhas e adubação conforme o cronograma. Recolha os resíduos verdes. Use luva, óculos e perneira (protetor auricular na roçadeira) e isole a área durante o uso de roçadeira."],
 [/manuten|predial/,"Inspecione e registre vazamentos, iluminação, portas e fechaduras. Resolva o que estiver no escopo; o restante, abra chamado operacional com foto."],
 [/ronda/,"Percorra o roteiro definido nos horários previstos, verifique portas, janelas, iluminação e pessoas estranhas. Registre horário e pontos verificados. Comunique ocorrências imediatamente ao supervisor."],
 [/acesso|visitante/,"Identifique e registre todo visitante (nome, documento, destino e horário), confirme a autorização com o setor, entregue o crachá e registre a saída."],
 [/ocorr[eê]ncia/,"Registre no livro ou no sistema: data, hora, local, envolvidos, descrição objetiva do fato e providências tomadas. Comunique o supervisor."],
 [/abertura|fechamento/,"Siga o checklist do posto: portas, janelas, alarme, luzes e chaves. Registre o horário e qualquer anormalidade."],
 [/cftv|monitoramento/,"Acompanhe as câmeras, confirme que estão gravando, registre falhas e acione a ronda em situação suspeita."],
 [/superficie/,"Pano umedecido em detergente neutro, de cima para baixo e do mais limpo para o mais sujo; seque. Nas superfícies de contato, aplique depois o desinfetante padronizado."],
 [/pat[ií]o|area externa|externa/,"Varra, recolha folhas e lixo, esvazie lixeiras externas e lave quando necessário. Sinalize o piso molhado."],
 [/limpeza|higieniza|lavagem|lavar/,"1) Use os EPIs da área. 2) Recolha o lixo. 3) Limpe de cima para baixo e do mais limpo para o mais sujo: superfícies, mobiliário e pontos de contato com pano úmido e detergente; depois desinfete os pontos de contato. 4) Limpe o piso por último, do fundo para a saída. 5) Sinalize piso molhado e reponha insumos."]
];
function comoFazerPadrao(nome){var n=normT(nome);for(var i=0;i<COMO_FAZER.length;i++)if(COMO_FAZER[i][0].test(n))return COMO_FAZER[i][1];return ""}
function aplicarComoFazer(){
  var mud=false;
  PG.forEach(function(g){
    if(g.obs&&!g.obsAuto)return;
    var t=comoFazerPadrao(g.nome);
    if(t&&(g.obs!=t||!g.obsAuto)){g.obs=t;g.obsAuto=true;mud=true}
    else if(!t&&g.obsAuto){g.obs="";mud=true}
  });
  return mud;
}
function svPg(){aplicarComoFazer();sv("rad_pg",PG)}
var pgLoc=0;
function pgCard(){
  if(!LOC[pgLoc])pgLoc=0;
  var lista=PG.filter(function(g){return g.l==pgLoc});
  return '<div class="card"><b>🧩 Atividades por área</b><div class="s">Usadas na montagem do PTI e no Plano Geral.</div>'
    +'<select onchange="pgLoc=+this.value;draw()">'+LOC.map(function(l,i){return '<option value="'+i+'"'+(i==pgLoc?' selected':'')+'>'+esc(l.n)+'</option>'}).join("")+'</select>'
    +(lista.length?lista.map(function(g){return '<div class="esp-linha"><input value="'+esc(g.nome)+'" onchange="renPg('+g.id+',this.value)"><button class="pill no mini-btn" onclick="delPg('+g.id+')">✕</button></div>'}).join(""):'<div class="s">Nenhuma atividade nesta área.</div>')
    +'<div style="display:flex;gap:6px;margin-top:10px"><input id="npg" placeholder="Nova atividade" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px;cursor:pointer" onclick="addPg()">Add</button></div>'
    +'<button class="btn alt" onclick="copiarPgTodas()" style="margin-top:8px">Copiar atividades desta área para todas</button></div>';
}
function addPg(){
  if(!canEdit()){alert("Sem permissão.");return}
  var v=($("#npg").value||"").trim();if(!v)return;
  var id=PG.reduce(function(m,g){return Math.max(m,g.id)},0)+1;
  PG.push({id:id,l:pgLoc,nome:v});svPg();draw();
}
function renPg(id,v){
  var g=PG.filter(function(x){return x.id==id})[0];if(!g)return;
  if(!canEdit()&&!(LOC[g.l]&&podeAreasCt(LOC[g.l].ct))){alert("Sem permissão.");return}
  g.nome=(v||"").trim()||g.nome;svPg();
}
function delPg(id){
  var g0=PG.filter(function(x){return x.id==id})[0];
  if(!canEdit()&&!(g0&&LOC[g0.l]&&podeAreasCt(LOC[g0.l].ct))){alert("Sem permissão.");return}
  if(!confirm("Excluir esta atividade? Ela será removida dos PTIs."))return;
  PG=PG.filter(function(x){return x.id!=id});svPg();
  PTIS.forEach(function(p){(p.areas||[]).forEach(function(a){a.atividades=(a.atividades||[]).filter(function(x){return x!=id})})});svPti();
  draw();
}
function copiarPgTodas(){
  if(!canEdit()){alert("Sem permissão.");return}
  var base=PG.filter(function(g){return g.l==pgLoc}).map(function(g){return g.nome});
  if(!base.length){alert("Nenhuma atividade para copiar.");return}
  if(!confirm("Adicionar estas "+base.length+" atividades às demais áreas (sem duplicar)?"))return;
  var id=PG.reduce(function(m,g){return Math.max(m,g.id)},0);
  LOC.forEach(function(_,i){
    if(i==pgLoc)return;
    var ex=PG.filter(function(g){return g.l==i}).map(function(g){return g.nome.toLowerCase()});
    base.forEach(function(n){if(ex.indexOf(n.toLowerCase())<0)PG.push({id:++id,l:i,nome:n})});
  });
  svPg();draw();alert("Atividades copiadas.");
}

/* ============ ESTOQUE / REQUISIÇÕES / SETORES ============ */
var EST=ld("rad_est",[]);
var REQ=ld("rad_req",[]);
var SET=ld("rad_set",[
  {id:"limpeza",nome:"Limpeza",zap:"",mail:""},
  {id:"manut",nome:"Manutenção",zap:"",mail:""},
  {id:"ti",nome:"TI",zap:"",mail:""},
  {id:"qualidade",nome:"Qualidade",zap:"",mail:""},
  {id:"compras",nome:"Compras / Suprimentos",zap:"",mail:""},
  {id:"rh",nome:"RH / Pessoal",zap:"",mail:""},
  {id:"admin",nome:"Administração / Financeiro",zap:"",mail:""}
]);
function svEst(){sv("rad_est",EST)}
function svReq(){sv("rad_req",REQ)}
function svSet(){sv("rad_set",SET)}
function estAdd(nome,unidade,qtd,minimo,local,ct){
  nome=(nome||"").trim();
  if(!nome){alert("Informe o nome do item.");return false}
  EST.push({id:Date.now(),ct:ct===undefined?(typeof esCt!="undefined"?esCt:null):ct,nome:nome,unidade:(unidade||"un").trim(),qtd:Math.max(0,+qtd||0),min:Math.max(0,+minimo||0),local:(local||"").trim(),mov:[{t:Date.now(),d:Math.max(0,+qtd||0),m:"Cadastro inicial",u:user||"Admin"}]});
  svEst();return true;
}
function estMov(id,delta,motivo){
  var it=EST.filter(function(x){return x.id==id})[0];if(!it)return;
  var antes=it.qtd;it.qtd=Math.max(0,it.qtd+delta);it.mov=it.mov||[];
  it.mov.push({t:Date.now(),d:it.qtd-antes,m:motivo||"",u:user||"Admin"});svEst();
}
function estBaixa(id,qtd,motivo){
  var it=EST.filter(function(x){return x.id==id})[0];if(!it)return false;
  if(qtd>it.qtd){alert("Estoque insuficiente de "+it.nome+". Disponível: "+it.qtd+" "+it.unidade);return false}
  estMov(id,-qtd,motivo);return true;
}
function estCritico(it){return it.qtd<=it.min}
function reqNova(itemId,qtd,obs,setorDest){
  var it=EST.filter(function(x){return x.id==itemId})[0];
  if(!it){alert("Item não encontrado.");return false}
  qtd=Math.max(1,+qtd||1);
  REQ.push({id:Date.now(),itemId:itemId,itemNome:it.nome,unidade:it.unidade,qtd:qtd,obs:(obs||"").trim(),setor:setorDest||"compras",ini:Date.now(),colab:user||"",st:"Pendente",hist:[{t:Date.now(),x:"Requisição aberta por "+(user||"Colaborador")}]});
  svReq();notificarSetor(setorDest||"compras","🛒 Nova requisição de material",(user||"Colaborador")+" pediu "+qtd+" "+it.unidade+" de "+it.nome+(obs?"\n\nObs: "+obs:""),null);
  return true;
}
function reqSt(id,v){
  var r=REQ.filter(function(x){return x.id==id})[0];if(!r)return;
  if(v=="Aprovada"&&r.st!="Aprovada"){if(!estBaixa(r.itemId,r.qtd,"Requisição aprovada para "+r.colab))return}
  r.st=v;r.hist=r.hist||[];r.hist.push({t:Date.now(),x:v+" por "+(user||"")});svReq();
}
function reqPend(){return REQ.filter(function(r){return r.st=="Pendente"})}
function setor(id){return SET.filter(function(s){return s.id==id})[0]||{id:id,nome:id,zap:"",mail:""}}
function setSetor(id,k,v){var s=SET.filter(function(x){return x.id==id})[0];if(!s)return;s[k]=v.trim();svSet()}

/* ============ CANAIS DE CHAMADO (3 canais) ============ */
var CANAIS={
  OP:{id:"OP",nome:"Operacional",desc:"Problemas ou necessidades da operação",ico:"🛠️"},
  AD:{id:"AD",nome:"Administrativo",desc:"Documentos, pagamentos, benefícios, RH",ico:"📄"},
  CLI:{id:"CLI",nome:"Cliente",desc:"Solicitações, reclamações e elogios do cliente",ico:"💬"}
};
var CHT_CANAL=ld("rad_chtcanal",{
  OP:[
    {t:"Falta de material",d:"compras"},
    {t:"Equipamento com defeito",d:"manut"},
    {t:"Vazamento ou manutenção",d:"manut"},
    {t:"Sujeira excessiva",d:"limpeza"},
    {t:"Risco ou acidente",d:"qualidade"},
    {t:"Ar-condicionado com problema",d:"manut"},
    {t:"Outro (operacional)",d:"admin"}
  ],
  AD:[
    {t:"Documentos / declarações",d:"rh"},
    {t:"Pagamentos / adiantamento",d:"admin"},
    {t:"Benefícios (VT, VA, VR)",d:"rh"},
    {t:"Férias / folga",d:"rh"},
    {t:"Atestado / saúde",d:"rh"},
    {t:"Outro (administrativo)",d:"admin"}
  ],
  CLI:[
    {t:"🎯 Solicitação de serviço",d:"limpeza"},
    {t:"⚠️ Reclamação",d:"admin"},
    {t:"🌟 Elogio",d:"admin"},
    {t:"❓ Dúvida / Informação",d:"admin"},
    {t:"📋 Ajuste em procedimento",d:"qualidade"},
    {t:"🚨 Urgência",d:"admin"}
  ]
});
function svChCanal(){sv("rad_chtcanal",CHT_CANAL)}
function canalTipos(c){return CHT_CANAL[c]||[]}
function chDestinoCan(c,t){var a=canalTipos(c);for(var i=0;i<a.length;i++)if(a[i].t===t)return a[i].d;return "admin"}
function setDestCanal(canal,i,setorId){if(!CHT_CANAL[canal])return;CHT_CANAL[canal][i].d=setorId;svChCanal();draw()}
function notificarSetor(setorId,titulo,texto,canalId){
  var s=setor(setorId),canal=canalId?CANAIS[canalId]:null,prefixo=canal?(canal.ico+" ["+canal.nome+"] "):"";
  var msg="*"+prefixo+titulo+"*\n\n"+texto+"\n\n_"+new Date().toLocaleString("pt-BR")+"_";
  try{var hist=ld("rad_notif",[]);hist.push({id:Date.now(),setor:setorId,canal:canalId||"",titulo:titulo,texto:texto,ts:Date.now()});if(hist.length>200)hist=hist.slice(-200);sv("rad_notif",hist)}catch(e){}
  var auto=ld("rad_notifauto",{zap:true,mail:false});
  if(auto.zap&&s.zap){var zap=(s.zap||"").replace(/\D/g,"");if(zap){if(zap.length<=11)zap="55"+zap;window.open("https://wa.me/"+zap+"?text="+encodeURIComponent(msg),"_blank")}}
  if(auto.mail&&s.mail){var subj=prefixo+titulo+" - "+CONTRATO;window.location.href="mailto:"+s.mail+"?subject="+encodeURIComponent(subj)+"&body="+encodeURIComponent(texto)}
}
function setNotifAuto(k,v){var auto=ld("rad_notifauto",{zap:true,mail:false});auto[k]=!!v;sv("rad_notifauto",auto);draw()}

/* ============ CLASSIFICAÇÃO ACL ============ */
var ACL=ld("rad_acl",{}),AL={C:"Crítica",S:"Semicrítica",N:"Não crítica"};
function aclOf(l,e){return ACL[l+"_"+e]||"X"}
function acNome(k){return {C:"Crítica",S:"Semicrítica",N:"Não crítica",X:"Sem classificação"}[k]||"Sem classificação"}
function acCor(k){return {C:"#c62828",S:"#e0a100",N:"#2e8b57",X:"#5d6e69"}[k]||"#5d6e69"}
/* ---------- Classificação de risco pré-estabelecida ----------
   Hospitalar: critérios do Manual ANVISA "Segurança do Paciente em Serviços de Saúde:
   Limpeza e Desinfecção de Superfícies" (crítica / semicrítica / não crítica).
   Demais serviços: risco de contaminação cruzada (alimentos, sanitários, resíduos). */
var ACLA=ld("rad_acla",{});
var REGRA_RISCO={
  hosp:{
    C:/centro cir|sala cir|cirurg|\bcc\b|\bcme\b|esteriliz|expurgo|\buti\b|\bcti\b|terapia intensiva|isolamento|hemodi|di[aá]lise|laborat|banco de sangue|lact[aá]rio|nutri[cç][aã]o|diet[eé]tica|cozinha|lavanderia|necrot|morgue|repae|arsenal|recupera[cç][aã]o|\brpa\b|procedimento|curativo|parto|obst[eé]tr|queimad|transplant|quimio|endoscop|hemodin|pronto.?socorro|emerg[eê]ncia|sala vermelha|berç[aá]rio|farm[aá]cia|infectante|perfuro|abrigo de res|res[ií]duo|lixo infect|casa perfuro|vacina/,
    S:/enfermaria|apartamento|quarto|leito|consult[oó]rio|\bcons\b|ambulat|banheiro|\bwc\b|sanit[aá]rio|lavabo|toalete|posto de enfermagem|enfermagem|exame|diagn[oó]stico|raio.?x|tomograf|resson|ultras|\boct\b|retino|angio|biometria|topograf|laser|\byag\b|campimetria|optomap|olhos secos|sala de espera|espera|recep[cç][aã]o|admiss[aã]o|elevador|corredor|brinquedoteca|pediatria|fisioter|psicolog|sala de preparo|orienta[cç][aã]o|inclus[aã]o|conforto|care|\bpa\b|lixo|dml|rouparia|bebedouro|lavabo/,
    /* áreas sem pacientes (ANVISA: não críticas) — verificadas primeiro */
    N:/vesti[aá]rio|administrat|\badm\b|diretor|financeiro|comercial|\brh\b|qualidade|supervis[aã]o|ger[eê]ncia|coordena|tesouraria|telefonia|\bcsi\b|manuten[cç][aã]o|\bt\.?i\b|\bsame\b|almoxarif|suprimento|recebimento|oficina|guarita|estacionamento|\bcopa\b|copinha|m[aá]quina de caf|reuni[aã]o|impress[aã]o|autoriza[cç][aã]o|comunica[cç][aã]o|guich[eê]|escada|dispensa|geladeira|sala de descanso|entrega de material|sala da |sala gavea|apoio log/,
    /* atendimento a pacientes fora da área crítica: semicrítica mesmo que o nome cite o CC */
    P:/recep[cç][aã]o|espera|orienta[cç][aã]o|orient\.|agendamento|admiss[aã]o/
  },
  geral:{
    C:/anatomia|infectante|perfuro|res[ií]duo infect|cozinha|lact[aá]rio|c[aâ]mara fria|a[cç]ougue|manipula[cç][aã]o de alimento|ambulat[oó]rio|enfermaria/,
    S:/banheiro|\bwc\b|sanit[aá]rio|vesti[aá]rio|lavabo|toalete|chuveiro|copa|refeit|cantina|bebedouro|pra[cç]a de alimenta|laborat|creche|berç|frald|trocador|lixo|lixeira|res[ií]duo|provador|autoatendimento|\batm\b|terminai|elevador|academia|piscina|dml/
  }
};
function riscoSugerido(l,e){
  var L=LOC[l];if(!L)return "N";
  var nomeEsp=normT(L.e[e]),nomeArea=normT(L.n);
  var t=tipoDaArea(l)||((L.ct==null)?"hosp":"");
  if(t=="hosp"){
    var r=REGRA_RISCO.hosp;
    if(r.P.test(nomeEsp))return "S";
    if(/lava[cç][aã]o|lavagem de material/.test(nomeEsp))return "C";
    if(r.C.test(nomeEsp))return "C";
    if(r.N.test(nomeEsp))return "N";
    if(r.S.test(nomeEsp))return "S";
    /* área de centro cirúrgico: salas sem nome específico são críticas */
    if(/\bcen\b|centro cir|\bcc\b|bloco cir|cirurg/.test(nomeArea))return "C";
    if(r.C.test(nomeArea))return "C";
    /* hospital: na dúvida, semicrítica (mais seguro que não crítica) */
    return "S";
  }
  var g=REGRA_RISCO.geral;
  if(g.C.test(nomeEsp))return "C";
  if(g.S.test(nomeEsp))return "S";
  return "N";
}
/* preenche e mantém atualizada a classificação automática; respeita o que foi definido manualmente */
function autoClassificar(){
  var mud=false;
  LOC.forEach(function(L,l){
    if(L.del)return;
    L.e.forEach(function(_,e){
      if(espDel(l,e))return;
      var k=l+"_"+e;
      if(!ACL[k]||ACLA[k]){var v=riscoSugerido(l,e);if(ACL[k]!=v||!ACLA[k]){ACL[k]=v;ACLA[k]=true;mud=true}}
    });
  });
  if(mud){sv("rad_acl",ACL);sv("rad_acla",ACLA)}
}
function setAcl(l,e,v){var k=l+"_"+e;if(v=="auto"){delete ACL[k];ACLA[k]=true}else{ACL[k]=v;delete ACLA[k]}autoClassificar();sv("rad_acl",ACL);sv("rad_acla",ACLA);var y=window.scrollY;draw();window.scrollTo(0,y)}
function reclassificarArea(l){if(!arChk(l))return;if(!confirm("Voltar todos os espaços desta área para a classificação automática?"))return;LOC[l].e.forEach(function(_,e){var k=l+"_"+e;delete ACL[k];ACLA[k]=true});autoClassificar();draw()}
function setAclAll(v){if(!v)return;LOC[loc].e.forEach(function(x,i){ACL[loc+"_"+i]=v});sv("rad_acl",ACL);draw()}
function aclSel(l,e){var k=l+"_"+e,c=ACL[k]||"",au=!!ACLA[k];return '<select title="'+(au?'Classificação automática':'Definida manualmente')+'" style="width:auto;margin:0;'+(c?'border-left:5px solid '+acCor(c):'')+'" onchange="setAcl('+l+','+e+',this.value)">'+["C","S","N"].map(function(x){return '<option value="'+x+'"'+(c==x?' selected':'')+'>'+AL[x]+(c==x&&au?' (auto)':'')+'</option>'}).join("")+(au?'':'<option value="auto">↺ Automática</option>')+'</select>'}

/* ============ NFC ============ */
var NFCCFG=ld("rad_nfccfg",{on:true});
function setNFC(v){if(!canEdit()){alert("Sem permissão.");return}NFCCFG.on=!!v;sv("rad_nfccfg",NFCCFG);draw()}

/* ============ CHECKLISTS CUSTOMIZADOS ============ */
var CHECKLISTS=ld("rad_checklists",null);
if(!CHECKLISTS){
  CHECKLISTS=[
    {id:"lz",nome:"Luz Negra (padrão)",cat:"Luz Negra",ico:"🔦",areas:"*",ativo:true,fixed:true},
    {id:"rc016",nome:"Check list REG-016",cat:"REG-016",ico:"✅",areas:"*",ativo:true,fixed:true}
  ];
  sv("rad_checklists",CHECKLISTS);
}
function svChkl(){sv("rad_checklists",CHECKLISTS)}
function chklById(id){return CHECKLISTS.filter(function(c){return c.id==id})[0]||null}
function chklItens(id){
  if(id=="lz")return LZI;
  if(id=="rc016")return RCI;
  var c=chklById(id);
  return c&&c.itens?c.itens:[];
}
function addChkl(){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var nome=$("#ck_nome").value.trim();
  var cat=$("#ck_cat").value;
  var ico=$("#ck_ico").value.trim()||"📋";
  if(!nome){alert("Informe o nome.");return}
  var id="ck_"+Date.now();
  var tipos=$all(".ck_tp:checked").map(function(x){return x.value});if(!tipos.length)tipos=["*"];
  CHECKLISTS.push({id:id,nome:nome,cat:cat,ico:ico,tipos:tipos,areas:"*",ativo:true,fixed:false,itens:[]});
  svChkl();
  logAcao("Criou checklist",nome);
  draw();
  alert("Checklist criado. Agora adicione os itens.");
}
function togChklAtivo(id){if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}var c=chklById(id);if(!c)return;c.ativo=!c.ativo;svChkl();draw()}
function delChkl(id){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var c=chklById(id);if(!c)return;
  if(c.fixed){alert("Checklist padrão não pode ser excluído.");return}
  if(!confirm("Excluir checklist "+c.nome+"?"))return;
  CHECKLISTS=CHECKLISTS.filter(function(x){return x.id!=id});
  if(cklSel==id)cklSel="lz";
  svChkl();logAcao("Excluiu checklist",c.nome);draw();
}
function addChklItem(id){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var c=chklById(id);if(!c||c.fixed)return;
  var n=$("#ci_nome_"+id).value.trim();
  var crit=$("#ci_crit_"+id).value;
  var pts=$("#ci_pts_"+id).value.trim();
  if(!n){alert("Informe o nome do item.");return}
  c.itens=c.itens||[];
  var p=pts?pts.split(";").map(function(x){return x.trim()}).filter(Boolean):[];
  var item={n:n,c:crit};if(p.length)item.p=p;
  c.itens.push(item);
  svChkl();
  logAcao("Adicionou item em checklist",c.nome,n);
  var y=window.scrollY;draw();window.scrollTo(0,y);
}
function delChklItem(id,i){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var c=chklById(id);if(!c||c.fixed)return;
  if(!confirm("Excluir este item?"))return;
  c.itens.splice(i,1);svChkl();
  var y=window.scrollY;draw();window.scrollTo(0,y);
}

/* ============ CONTATOS E LINKS ============ */
var CONTATOS_PADRAO=[
  {id:1,tipo:"tel",nome:"Manutenção - Ramal",valor:"2201",ico:"🔧",cat:"Telefones"},
  {id:2,tipo:"tel",nome:"Limpeza - Ramal",valor:"2202",ico:"🧹",cat:"Telefones"},
  {id:3,tipo:"tel",nome:"TI - Ramal",valor:"2203",ico:"💻",cat:"Telefones"},
  {id:4,tipo:"tel",nome:"Supervisão",valor:"",ico:"👔",cat:"Telefones"},
  {id:5,tipo:"tel",nome:"Emergência (SAMU)",valor:"192",ico:"🚨",cat:"Emergência"},
  {id:6,tipo:"tel",nome:"Bombeiros",valor:"193",ico:"🚒",cat:"Emergência"},
  {id:7,tipo:"link",nome:"Portal RH",valor:"https://",ico:"🔗",cat:"Links úteis"},
  {id:8,tipo:"link",nome:"Manual de Limpeza",valor:"https://",ico:"📘",cat:"Documentos"}
];
var CONTATOS=ld("rad_contatos",CONTATOS_PADRAO);
/* migração: emergência visível a todos; demais só para a equipe interna */
CONTATOS.forEach(function(c){if(c.vis==null)c.vis=c.cat=="Emergência"?"todos":"interno"});
function svContatos(){sv("rad_contatos",CONTATOS)}
function addContato(){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var nome=$("#ctt_nome").value.trim();
  var tipo=$("#ctt_tipo").value;
  var valor=$("#ctt_valor").value.trim();
  var cat=$("#ctt_cat").value;
  var ico=$("#ctt_ico").value.trim()||"📌";
  if(!nome||!valor){alert("Informe nome e valor.");return}
  CONTATOS.push({id:Date.now(),tipo:tipo,nome:nome,valor:valor,ico:ico,cat:cat,vis:cat=="Emergência"?"todos":"interno"});
  svContatos();
  logAcao("Adicionou contato",nome);
  draw();
}
function delContato(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c)return;
  if(!confirm("Excluir "+c.nome+"?"))return;
  CONTATOS=CONTATOS.filter(function(x){return x.id!=id});
  svContatos();logAcao("Excluiu contato",c.nome);draw();
}
function editContato(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c)return;
  var novo=prompt("Novo valor para "+c.nome+":",c.valor);
  if(novo===null)return;
  c.valor=novo.trim();svContatos();logAcao("Editou contato",c.nome);draw();
}
function copiarContato(id){
  var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c)return;
  if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(c.valor).then(function(){alert("Copiado!")},function(){prompt("Copie:",c.valor)})}
  else prompt("Copie:",c.valor);
}
function categoriaContatos(){
  var cats={};
  CONTATOS.filter(contatoVisivel).forEach(function(c){(cats[c.cat]=cats[c.cat]||[]).push(c)});
  return cats;
}
function urlSegura(u){u=String(u||"").trim();return /^(https?:|mailto:|tel:)/i.test(u)?u:"#"}
function acaoContato(c){
  if(c.tipo=="tel"){
    var num=(c.valor||"").replace(/\D/g,"");
    if(!num)return '<span class="s">sem número</span>';
    return '<a class="pill ok" style="border:0;text-decoration:none;padding:6px 12px" href="tel:'+esc(num)+'">📞 Ligar</a>'
      +(num.length>=10?'<a class="pill run" style="border:0;text-decoration:none;padding:6px 12px;margin-left:4px" target="_blank" rel="noopener" href="https://wa.me/'+(num.length<12?"55"+num:num)+'">💬 WhatsApp</a>':'')
      +'<button class="pill pen" style="border:0;margin-left:4px" onclick="copiarContato('+c.id+')">📋</button>';
  }
  if(c.tipo=="link"||c.tipo=="mail"){
    var url=c.tipo=="mail"?"mailto:"+c.valor:urlSegura(c.valor);
    return '<a class="pill ok" style="border:0;text-decoration:none;padding:6px 12px" target="_blank" rel="noopener" href="'+esc(url)+'">🔗 Abrir</a>'
      +'<button class="pill pen" style="border:0;margin-left:4px" onclick="copiarContato('+c.id+')">📋</button>';
  }
  return '';
}

/* ============ ÁREAS / ESPAÇOS ============ */
function locOff(i){return !!(LOC[i]&&(LOC[i].off||LOC[i].del))}
function locDel(i){return !LOC[i]||!!LOC[i].del}
function espDel(l,e){return !!(LOC[l]&&LOC[l].delE&&LOC[l].delE[e])}
function espOff(l,e){return !!(LOC[l]&&((LOC[l].offE&&LOC[l].offE[e])||(LOC[l].delE&&LOC[l].delE[e])))}
function nomeEsp(l,e){var L=LOC[l];return L&&L.e[e]!=null?L.e[e]:"(espaço removido)"}
function nomeLoc(l){var L=LOC[l];return L?L.n:"(área removida)"}
function optsAll(sel){return LOC.map(function(l,i){return '<option value="'+i+'"'+(i==sel?' selected':'')+'>'+esc(l.n)+(l.off?' — desabilitada':'')+'</option>'}).join("")}
function togL(){var L=LOC[loc];if(!L)return;var ativas=LOC.filter(function(x){return !x.off}).length;if(!L.off&&ativas<=1){alert("Última área.");return}L.off=!L.off;svLoc();draw()}
function togE(e){var L=LOC[loc];if(!L)return;if(!L.offE)L.offE={};if(!L.offE[e]){var ativos=L.e.filter(function(_,i){return !L.offE[i]}).length;if(ativos<=1){alert("Último espaço.");return}}L.offE[e]=!L.offE[e];svLoc();draw()}
function delE(e){arDelE(loc,e)}
function delL(){arDelL(loc)}
function arrsComLocal(){return [regs,pa,chs,AVALS]}
function salvarArrsLocal(){sv("rad_regs",regs);sv("rad_pa",pa);sv("rad_chs",chs);svAvals()}
function reindexEspaco(l,eDel){
  arrsComLocal().forEach(function(arr){arr.forEach(function(r){if(r&&r.l==l&&r.e!=null&&r.e>=0){if(r.e==eDel)r.eRem=true;else if(r.e>eDel)r.e--}})});
  salvarArrsLocal();
  var no={};Object.keys(ACL).forEach(function(k){var p=k.split("_"),kl=+p[0],ke=+p[1];if(kl===l){if(ke===eDel)return;if(ke>eDel)ke--}no[kl+"_"+ke]=ACL[k]});ACL=no;sv("rad_acl",ACL);
}
function reindexArea(lDel){
  arrsComLocal().forEach(function(arr){arr.forEach(function(r){if(!r||r.l==null||r.l<0)return;if(r.l==lDel)r.l=-2;else if(r.l>lDel)r.l--})});
  salvarArrsLocal();
  var no={};Object.keys(ACL).forEach(function(k){var p=k.split("_"),kl=+p[0],ke=+p[1];if(kl===lDel)return;if(kl>lDel)kl--;no[kl+"_"+ke]=ACL[k]});ACL=no;sv("rad_acl",ACL);
  PG=PG.filter(function(g){return g.l!=lDel});PG.forEach(function(g){if(g.l>lDel)g.l--});svPg();
  CONTRATOS.forEach(function(c){c.locais=locaisCt(c).filter(function(x){return x!=lDel}).map(function(x){return x>lDel?x-1:x})});svCt();
  PTIS.forEach(function(p){p.areas=(p.areas||[]).filter(function(a){return a.locIdx!=lDel});p.areas.forEach(function(a){if(a.locIdx>lDel)a.locIdx--})});svPti();
}

/* ============ REGISTROS / RECUSAS (unificado) ============ */
var regs=ld("rad_regs",[]),loc=ld("rad_loc",0),open=ld("rad_open",null),tipo="CO",tab="c",stream=null;
if(!LOC[loc])loc=0;
function hoje(){return regs.filter(function(r){return dia(r.ini)==dia(Date.now())})}
function reg(l,e){return hoje().filter(function(r){return r.l==l&&r.e==e&&!r.x&&r.origem!="giro"}).pop()}
function nomeOpen(){return open.origem=="giro"?(GS[open.e]||"Sala"):nomeEsp(open.l,open.e)}
function formRec(){var n=nomeOpen();var motivos=open.origem=="giro"?MOT_GIRO:MOT;return '<div class="card run-card"><b>Recusa: '+esc(n)+'</b><div class="s">Origem: '+esc(open.origem=="giro"?"Giro de sala":(open.origem||"rotina"))+'</div><div class="s">Justificativa</div><select id="rm">'+motivos.map(function(x){return '<option>'+x+'</option>'}).join("")+'</select><input id="ro" placeholder="Se Outro, descreva"><input id="rq" placeholder="Nome de quem recusou"><div class="s">Agendamento</div><input id="rd" type="date"><input id="rh" type="time"><label class="s"><input id="rs" type="checkbox" style="width:auto;margin-right:8px">Sem agendamento</label><input id="rb" placeholder="Obs"><button class="btn" onclick="salvarRec()">'+T("btn_save")+'</button><button class="btn alt" onclick="cancelRec()">'+T("btn_back")+'</button></div>'}
function cancelRec(){
  /* CORREÇÃO: ao voltar de uma recusa de giro, o giro em andamento é restaurado */
  if(open&&open.origem=="giro"){gopen={s:open.e,ini:open.ini};sv("rad_gopen",gopen);open=null;sv("rad_open",null)}
  rec=false;draw();
}
function salvarRec(){var m=$("#rm").value,o=$("#ro").value.trim(),q=$("#rq").value.trim(),d=$("#rd").value,s=$("#rs").checked;
 if(m.indexOf("Outro")>=0&&!o){alert("Descreva.");return}
 if(!q){alert("Informe quem recusou.");return}
 if(!s&&!d){alert("Informe a data ou marque 'Sem agendamento'.");return}
 regs.push({l:open.l,e:open.e,ini:open.ini,fim:Date.now(),tipo:"",colab:user,st:"nao",motivo:m,outro:o,quem:q,agD:d,agH:$("#rh").value,semAg:s,obs:$("#rb").value.trim(),stAt:"Pendente",origem:open.origem||"rotina"});sv("rad_regs",regs);rec=false;open=null;sv("rad_open",null);logAcao("Registrou recusa",m);draw()}
function setStat(i,v){regs.forEach(function(r){if(r.ini==i&&r.st=="nao")r.stAt=v});sv("rad_regs",regs);draw()}
function recCard(){var l=regs.filter(function(r){var s=r.stAt||"Pendente";return r.st=="nao"&&r.origem!="giro"&&r.l==loc&&s!="Realizado com sucesso"&&s!="Realizado com observações"});
 return '<div class="card"><b>Recusas e pendências (REG-278)</b>'+(l.length?l.map(function(r){var s=r.stAt||"Pendente";return '<div class="row" style="flex-wrap:wrap"><span style="flex:1"><b>'+esc(nomeEsp(r.l,r.e))+'</b> <span class="s">📋 Rotina</span><div class="s">'+new Date(r.ini).toLocaleDateString("pt-BR")+' '+hm(r.ini)+' · '+esc(r.motivo+(r.outro?": "+r.outro:""))+(r.quem?' · recusou: '+esc(r.quem):'')+(r.agD?' · reagendado '+esc(r.agD)+(r.agH?' '+esc(r.agH):''):'')+'</div></span><select style="width:auto;margin:0" onchange="setStat('+r.ini+',this.value)">'+STAT.map(function(x){return '<option'+(x==s?' selected':'')+'>'+x+'</option>'}).join("")+'</select></div>'}).join(""):'<div class="s">Nenhuma.</div>')+'</div>'}

/* ============ CHAMADOS (3 canais + visibilidade) ============ */
var chs=ld("rad_chs",[]),CHS=["Aberto","Em atendimento","Resolvido"];
function chCtx(){
  if(open){if(open.origem=="giro")return "CC · "+(GS[open.e]||"");return nomeLoc(open.l)+" · "+nomeEsp(open.l,open.e)}
  if(gopen)return "CC · "+GS[gopen.s];
  if(loc<0||!LOC[loc]){var c=ctById(ctAtualId());return c?c.codigo+" · "+c.nome:"—"}
  return nomeLoc(loc);
}
function visiveisChs(canalFixo){
  var u=session?userById(session.userId):null;
  var cg=cargoAtual();
  var todos=chs.filter(function(c){return c.canal==canalFixo});
  if(!cg||cg.fixo||cg.id=="admin")return todos;
  if(cg.id=="colab"||cg.id=="cli")return todos.filter(function(c){return c.colab===(u?u.nome:"")});
  if(verTudo())return todos;
  var ids=meusCtIds();
  return todos.filter(function(c){return (c.ct!=null?ids.indexOf(c.ct)>=0:(c.l==null||c.l<0||ids.indexOf(ctDaArea(c.l))>=0))||c.colab===(u?u.nome:"")});
}
function podeResponderCh(){
  var cg=cargoAtual();
  if(!cg)return false;
  if(cg.fixo||cg.id=="admin")return true;
  return ["enc","sup","coord","ger"].indexOf(cg.id)>=0;
}
function novoCh(canalId,origem){
  var canal=canalId||"OP",tp=$("#ct").value,x=($("#cx").value||"").trim(),dest=chDestinoCan(canal,tp);
  if(!x){alert("Descreva o chamado.");return}
  var giroAberto=open&&open.origem=="giro";
  var c={id:Date.now(),canal:canal,tipo:tp,onde:chCtx(),l:(open&&!giroAberto?open.l:(gopen||giroAberto?-1:loc)),e:(open&&!giroAberto?open.e:-1),ini:Date.now(),colab:user,st:"Aberto",setor:dest,origem:origem||"",msgs:[{de:"c",t:Date.now(),x:x}]};
  chs.push(c);sv("rad_chs",chs);
  notificarSetor(dest,"Chamado: "+tp,"Onde: "+c.onde+"\nPor: "+(user||"")+"\n\n"+x,canal);
  logAcao("Abriu chamado",tp,canal);
  draw();
}
function respCh(id){
  var el=document.getElementById("r"+id);if(!el)return;
  var x=el.value.trim();if(!x)return;
  var c=chs.filter(function(c){return c.id==id})[0];if(!c)return;
  var cg=cargoAtual();
  var de=(!cg||cg.id=="colab"||cg.id=="cli")?"c":"s";
  c.msgs.push({de:de,t:Date.now(),x:x,por:user||""});
  if(de=="s"&&c.st=="Aberto")c.st="Em atendimento";
  sv("rad_chs",chs);draw();
}
function stCh(id,v){chs.forEach(function(c){if(c.id==id)c.st=v});sv("rad_chs",chs);logAcao("Alterou status de chamado",v);draw()}
function chItem(c,sup){
  var cl={"Aberto":"no","Em atendimento":"run","Resolvido":"ok"}[c.st];
  var sem=sup&&c.st=="Aberto"&&!c.msgs.some(function(m){return m.de=="s"})&&Date.now()-c.ini>1800000;
  return '<div class="card">'
    +'<div class="row" style="border:0;padding:0;flex-wrap:wrap"><b style="flex:1">'+esc(c.tipo)+'</b><span class="pill '+cl+'" style="cursor:default">'+c.st+'</span></div>'
    +'<div class="s">'+esc(c.onde)+' · '+new Date(c.ini).toLocaleDateString("pt-BR")+' '+hm(c.ini)+' · por '+esc(c.colab||"?")+(c.setor?' · <b>'+esc(setor(c.setor).nome)+'</b>':'')+(c.origem?' · '+esc(c.origem):'')+(c.prot?' · protocolo '+esc(c.prot):'')+(sem?' · <b style="color:var(--warn)">sem resposta</b>':'')+'</div>'
    +(c.foto?'<img src="'+c.foto+'" alt="" style="max-width:180px;border-radius:8px;margin-top:6px;display:block">':'')
    +c.msgs.map(function(m){return '<div style="margin-top:6px"><span class="s"><b>'+(m.de=="s"?"Equipe":"Solicitante")+'</b>'+(m.por?' ('+esc(m.por)+')':'')+' '+hm(m.t)+'</span><br>'+esc(m.x)+'</div>'}).join("")
    +(c.st!="Resolvido"?'<input id="r'+c.id+'" placeholder="Mensagem" style="margin-top:8px"><button class="btn alt" onclick="respCh('+c.id+')">'+T("btn_send")+'</button>':'')
    +(sup?'<select onchange="stCh('+c.id+',this.value)">'+CHS.map(function(s){return '<option'+(s==c.st?' selected':'')+'>'+s+'</option>'}).join("")+'</select>':'')
    +'</div>';
}
function vChamados(canalFixo){
  var sup=podeResponderCh(),canal=CANAIS[canalFixo],h="";
  var cg=cargoAtual();
  /* CORREÇÃO: o Cliente não conseguia abrir chamado nem no canal Cliente */
  if(!cg||cg.id!="cli"||canalFixo=="CLI"){
    h+='<div class="card noprint"><b>'+canal.ico+' '+T("btn_open")+' '+canal.nome+'</b><div class="s">'+esc(canal.desc)+'</div><div class="s" style="margin-top:12px">Onde: '+esc(chCtx())+'</div><select id="ct">'+canalTipos(canalFixo).map(function(x){return '<option value="'+esc(x.t)+'">'+esc(x.t)+' → '+esc(setor(x.d).nome)+'</option>'}).join("")+'</select><input id="cx" placeholder="Descreva"><button class="btn" onclick="novoCh(\''+canalFixo+'\')">'+T("btn_send")+'</button></div>';
  }
  var l=visiveisChs(canalFixo).slice().reverse();
  var abertos=l.filter(function(c){return c.st!="Resolvido"}).length;
  h+='<div class="card"><b>'+canal.ico+' '+canal.nome+'</b><div class="s">'+abertos+' em aberto · '+l.length+' total'+(cg&&(cg.id=="colab"||cg.id=="cli")?' (seus)':'')+'</div></div>';
  if(!l.length)h+='<div class="card s">Nenhum chamado.</div>';
  else h+=l.map(function(c){return chItem(c,sup)}).join("");
  return h;
}
function vOP(){return vChamados("OP")}
function vAD(){return vChamados("AD")}
function vCLI(){return vChamados("CLI")}


/* ============ COLABORADOR — MODELO eSOCIAL ============ */
var NR_LIST=["NR-06","NR-32","NR-33","NR-35","NR-10","NR-11","NR-12","NR-18","NR-20","NR-23","Brigada","1º Socorros"];
var ESC=["6X1","5X2","12X36","5X1"];
var EPL=ld("rad_epl",["Luva de látex","Luva de borracha","Bota","Máscara","Óculos","Avental","Touca"]);
["Luva de borracha cano longo","Avental impermeável","Calçado fechado antiderrapante","Máscara PFF2"].forEach(function(n){if(EPL.indexOf(n)<0)EPL.push(n)});
/* EPI obrigatório por classificação de risco (editável em Configurações → EPI por risco) */
var EPI_RISCO=ld("rad_epirisco",{
  C:["Luva de borracha cano longo","Avental impermeável","Máscara","Óculos","Bota","Touca"],
  S:["Luva de borracha","Avental","Óculos","Bota"],
  N:["Luva de borracha","Calçado fechado antiderrapante"]
});
function svEpiRisco(){sv("rad_epirisco",EPI_RISCO)}
function episEsp(l,e){var t=tipoDaArea(l);if(t=="vig")return [];return (EPI_RISCO[ACL[l+"_"+e]||"N"]||[]).slice()}
function episTxt(l,e){var a=episEsp(l,e);return a.length?a.join(", "):"—"}
function togEpiRisco(k,ix){if(!canEdit()){alert("Sem permissão.");return}var n=EPL[ix];if(!n)return;var a=EPI_RISCO[k]=EPI_RISCO[k]||[];var p=a.indexOf(n);if(p>=0)a.splice(p,1);else a.push(n);svEpiRisco();var y=window.scrollY;draw();window.scrollTo(0,y)}
function addEpiRisco(){if(!canEdit()){alert("Sem permissão.");return}var el=document.getElementById("nepi_r"),v=(el&&el.value||"").trim();if(!v)return;if(EPL.indexOf(v)<0){EPL.push(v);sv("rad_epl",EPL)}draw()}
function epiRiscoCard(){
  return '<div class="card"><b>🧤 EPI obrigatório por classificação de risco</b><div class="s">Aparece no mapeamento das áreas, na Rotina (o colaborador confirma o uso) e no PTI/PTG impressos. Áreas de Vigilância/Portaria não exigem EPI de limpeza.</div>'
    +["C","S","N"].map(function(k){return '<div style="border-top:1px solid var(--line);padding:8px 0;margin-top:6px"><b style="color:'+acCor(k)+'">'+AL[k]+'</b><div>'+EPL.map(function(n,ix){return chip((EPI_RISCO[k]||[]).indexOf(n)>=0,"togEpiRisco(\'"+k+"\',"+ix+")",n)}).join("")+'</div></div>'}).join("")
    +'<div style="display:flex;gap:6px;margin-top:8px"><input id="nepi_r" placeholder="Novo EPI" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px" onclick="addEpiRisco()">Add</button></div></div>';
}
var CONTRATO=ld("rad_contrato","GAVEA FACILITIES SERVIÇOS GERAIS LTDA");
function setCt(v){CONTRATO=v.trim()||CONTRATO;sv("rad_contrato",CONTRATO)}
function novoCol(o,i){
  return{
    id:Date.now()+i,
    empresa:o.empresa||"",matricula:o.matricula||"",empresaNome:o.empresaNome||"",situacao:o.situacao||"",
    funcionarioCod:o.funcionarioCod||"",nome:o.nome||"",funcao:o.funcao||"ASG",
    deptoCod:o.deptoCod||"",departamento:o.departamento||"",salario:o.salario||"",
    admissao:o.admissao||"",prorrogacao:o.prorrogacao||"",cpf:o.cpf||"",pis:o.pis||"",
    nascimento:o.nascimento||"",sexo:o.sexo||"",endereco:o.endereco||"",numero:o.numero||"",
    bairro:o.bairro||"",cidade:o.cidade||"",uf:o.uf||"",cep:o.cep||"",mae:o.mae||"",
    foto:o.foto||"",nrs:o.nrs||[],observacoes:o.observacoes||"",
    idiomaComun:o.idiomaComun||"pt-BR",idiomaIndigena:o.idiomaIndigena||"",
    mat:o.matricula||"",escala:o.escala||"6X1",horario:o.horario||"",
    atu:["G"],ep:[],lin:[],vinculos:[],
    dias:[0,0,0,0,0,0,0].map(function(){return false}),
    h1:o.h1||"",h2:o.h2||"",acts:[]
  };
}
var COL=ld("rad_col",[]),tsel=ld("rad_tsel","");
function svCol(){sv("rad_col",COL)}
function cget(id){return COL.filter(function(c){return c.id==id})[0]}
function selCol(id){tsel=(tsel==id?"":id);sv("rad_tsel",tsel);draw()}
function delCol(id){
  if(!temAcao("excluir_colaborador")&&!canEdit()){alert("Sem permissão.");return}
  var c=cget(id);if(!c)return;
  if(!confirm("Excluir "+c.nome+"?"))return;
  COL=COL.filter(function(c){return c.id!=id});if(tsel==id){tsel="";sv("rad_tsel",tsel)}svCol();logAcao("Excluiu colaborador",c.nome);draw();
}
function setF(k,v){if(!canEdit()){alert("Sem permissão.");draw();return}var c=cget(tsel);if(c){c[k]=v;svCol()}}
function toggleNR(nr){if(!canEdit()){alert("Sem permissão.");return}var c=cget(tsel);if(!c)return;c.nrs=c.nrs||[];var p=c.nrs.indexOf(nr);if(p>=0)c.nrs.splice(p,1);else c.nrs.push(nr);svCol();var y=window.scrollY;draw();window.scrollTo(0,y)}
function tep(i){var c=cget(tsel),n=EPL[i];if(!c)return;c.ep=c.ep||[];var p=c.ep.indexOf(n);if(p>=0)c.ep.splice(p,1);else c.ep.push(n);svCol();var y=window.scrollY;draw();window.scrollTo(0,y)}
function addEpi(){var v=$("#nep").value.trim();if(!v)return;if(EPL.indexOf(v)<0){EPL.push(v);sv("rad_epl",EPL)}var c=cget(tsel);if(c){c.ep=c.ep||[];if(c.ep.indexOf(v)<0)c.ep.push(v);svCol()}draw()}
function setEsc(i){var c=cget(tsel);if(!c)return;c.escala=ESC[i];svCol();var y=window.scrollY;draw();window.scrollTo(0,y)}
function chip(on,fn,txt){return '<button style="padding:6px 10px;margin:3px 3px 0 0;border-radius:99px;font-size:12px;border:1px solid '+(on?'var(--brand)':'var(--line)')+';background:'+(on?'var(--brand)':'var(--bg)')+';color:'+(on?'var(--on)':'var(--ink)')+';cursor:pointer" onclick="'+fn+'">'+esc(txt)+'</button>'}
function shrink(f,cb){var im=new Image();var url=URL.createObjectURL(f);im.onload=function(){var c=document.createElement("canvas"),s=Math.min(1,800/Math.max(im.width,im.height));c.width=Math.round(im.width*s);c.height=Math.round(im.height*s);c.getContext("2d").drawImage(im,0,0,c.width,c.height);URL.revokeObjectURL(url);cb(c.toDataURL("image/jpeg",.7))};im.onerror=function(){URL.revokeObjectURL(url);alert("Imagem inválida.")};im.src=url}
function fotoCol(inp){var f=inp.files[0];if(!f)return;shrink(f,function(d){var c=cget(tsel);if(c){c.foto=d;svCol();draw()}})}

/* ============ IMPORTAÇÃO eSOCIAL COMPLETA ============ */
function impRows(rows){
  if(!rows||!rows.length){alert("Nenhuma linha.");return}
  rows=rows.map(function(r){return (r||[]).map(function(x){return x==null?"":String(x).replace(/^\ufeff/,"")})});
  var h=-1;
  for(var i=0;i<Math.min(rows.length,12);i++){
    if(rows[i].some(function(x){return/funcion[aá]rio|nome|matr[ií]cula/i.test(String(x))})){h=i;break}
  }
  var M={};
  if(h>=0){
    var H=rows[h].map(function(x){return String(x||"").toLowerCase().trim()});
    H.forEach(function(x,i){
      if(M.empresa==null&&/^empresa$/.test(x))M.empresa=i;
      else if(M.matricula==null&&/matr[ií]cula/.test(x))M.matricula=i;
      else if(M.empresaNome==null&&/nome.*empresa|raz[aã]o/.test(x))M.empresaNome=i;
      else if(M.situacao==null&&/situa[cç][aã]o/.test(x))M.situacao=i;
      else if(M.funcionarioCod==null&&/^funcion[aá]rio$/.test(x))M.funcionarioCod=i;
      else if(M.nome==null&&/nome.*funcion/.test(x))M.nome=i;
      else if(M.funcao==null&&/fun[cç][aã]o|cargo/.test(x))M.funcao=i;
      else if(M.deptoCod==null&&/c[oó]d.*depart/.test(x))M.deptoCod=i;
      else if(M.departamento==null&&/departamento/.test(x))M.departamento=i;
      else if(M.salario==null&&/sal[aá]rio/.test(x))M.salario=i;
      else if(M.admissao==null&&/admiss[aã]o/.test(x))M.admissao=i;
      else if(M.prorrogacao==null&&/prorroga/.test(x))M.prorrogacao=i;
      else if(M.cpf==null&&/^cpf$/.test(x))M.cpf=i;
      else if(M.pis==null&&/pis|pasep/.test(x))M.pis=i;
      else if(M.nascimento==null&&/nascimento/.test(x))M.nascimento=i;
      else if(M.sexo==null&&/^sexo$/.test(x))M.sexo=i;
      else if(M.mae==null&&/m[aã]e/.test(x))M.mae=i;
      else if(M.endereco==null&&/^endere[cç]o$|rua|logradouro/.test(x))M.endereco=i;
      else if(M.numero==null&&/n[uú]mero|^nº$/.test(x))M.numero=i;
      else if(M.bairro==null&&/bairro/.test(x))M.bairro=i;
      else if(M.cidade==null&&/cidade|munic[ií]pio/.test(x))M.cidade=i;
      else if(M.uf==null&&/^uf$/.test(x))M.uf=i;
      else if(M.cep==null&&/^cep$/.test(x))M.cep=i;
      else if(M.nome==null&&/^nome$/.test(x))M.nome=i;
    });
  }
  if(h<0||M.nome==null){M={empresa:0,matricula:1,empresaNome:2,situacao:3,funcionarioCod:4,nome:5,funcao:6,deptoCod:7,departamento:8,salario:9,admissao:10,prorrogacao:11,cpf:12,pis:13,nascimento:14,sexo:15,endereco:16,numero:17,bairro:18,cidade:19,uf:20,cep:21,mae:22};h=-1}
  var add=0,dup=0;
  var CAMPOS_ESOCIAL=["empresa","empresaNome","situacao","funcionarioCod","funcao","deptoCod","departamento","salario","admissao","prorrogacao","cpf","pis","nascimento","sexo","endereco","numero","bairro","cidade","uf","cep","mae"];
  rows.slice(h+1).forEach(function(r,idx){
    var g=function(k){return M[k]!=null&&r[M[k]]!=null?String(r[M[k]]).trim():""};
    var nome=g("nome");if(!nome)return;
    var mat=g("matricula");
    var ex=COL.filter(function(c){return(mat&&c.matricula===mat)||(c.nome||"").toLowerCase()===nome.toLowerCase()})[0];
    if(ex){
      /* Reimportação: atualiza os dados do eSocial (inclui mudança de departamento) sem apagar o que foi preenchido no app */
      CAMPOS_ESOCIAL.forEach(function(k){var v=g(k);if(v)ex[k]=v});
      if(mat&&!ex.matricula){ex.matricula=mat;ex.mat=mat}
      dup++;return;
    }
    COL.push(novoCol({empresa:g("empresa"),matricula:mat,empresaNome:g("empresaNome"),situacao:g("situacao"),funcionarioCod:g("funcionarioCod"),nome:nome,funcao:g("funcao")||"ASG",deptoCod:g("deptoCod"),departamento:g("departamento"),salario:g("salario"),admissao:g("admissao"),prorrogacao:g("prorrogacao"),cpf:g("cpf"),pis:g("pis"),nascimento:g("nascimento"),sexo:g("sexo"),endereco:g("endereco"),numero:g("numero"),bairro:g("bairro"),cidade:g("cidade"),uf:g("uf"),cep:g("cep"),mae:g("mae")},idx));
    add++;
  });
  svCol();
  var sc=syncContratosDepto();
  var ss=syncSupervisores();
  prepararAreas();
  logAcao("Importou efetivo",add+" novo(s), "+dup+" atualizado(s)",sc.novos+" contrato(s) criado(s)");
  alert(add+" adicionado(s)"+(dup?"; "+dup+" atualizado(s)":"")+".\n"+sc.novos+" contrato(s) criado(s) a partir dos departamentos · "+sc.vinc+" vínculo(s) ajustado(s).\n"+ss.novos+" perfil(is) de supervisor criado(s) automaticamente.");draw();
}
/* CORREÇÃO: CSV com aspas (inclusive o próprio modelo/exportação) era importado com as aspas no nome */
function parseCSV(t){
  t=String(t||"").replace(/^\ufeff/,"");
  var first=t.split(/\r?\n/)[0]||"";
  var d=first.indexOf("\t")>=0?"\t":first.indexOf(";")>=0?";":",";
  var rows=[],row=[],f="",q=false;
  for(var i=0;i<t.length;i++){
    var ch=t[i];
    if(q){
      if(ch=='"'){if(t[i+1]=='"'){f+='"';i++}else q=false}
      else f+=ch;
    }else{
      if(ch=='"'&&f==="")q=true;
      else if(ch==d){row.push(f);f=""}
      else if(ch=="\n"||ch=="\r"){if(ch=="\r"&&t[i+1]=="\n")i++;row.push(f);f="";if(row.some(function(x){return x.trim()}))rows.push(row);row=[]}
      else f+=ch;
    }
  }
  row.push(f);if(row.some(function(x){return x.trim()}))rows.push(row);
  return rows;
}
function pasteImp(){var t=$("#pi").value;if(!t.trim()){alert("Cole as linhas.");return}impRows(parseCSV(t))}
function fileImp(inp){
  var f=inp.files[0];if(!f)return;var r=new FileReader();
  if(/\.xlsx?$/i.test(f.name)){
    if(typeof XLSX==="undefined"){alert("Biblioteca de Excel não carregou (verifique a internet).");inp.value="";return}
    r.onload=function(){try{var wb=XLSX.read(r.result,{type:"array"});impRows(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:"",raw:false}))}catch(e){alert("Erro: "+e.message)}};
    r.readAsArrayBuffer(f);
  }else{r.onload=function(){impRows(parseCSV(String(r.result)))};r.readAsText(f)}
  inp.value="";
}
function baixar(nome,conteudo,tipo){try{var b=new Blob(["\ufeff"+conteudo],{type:tipo+";charset=utf-8"});var u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download=nome;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);if(a.parentNode)a.parentNode.removeChild(a)},1200)}catch(e){alert("Erro: "+e.message)}}
function baixarModeloEfetivo(){
  var sep=";",cab=["Empresa","Matrícula eSocial","Nome da Empresa","Situação Funcionário","Funcionário","Nome do Funcionário","Função","Código Departamento","Departamento","Salário Base","Admissão","Prorrogação Experiência","CPF","PIS/PASEP","Nascimento","Sexo","Endereço","Número Endereço","Bairro","Cidade","UF","CEP","Mãe"];
  var ex=["00001","00001000001","EMPRESA EXEMPLO LTDA","Ativo","000001","FULANO DE TAL","ASG","0001","CONTRATO EXEMPLO","R$1.500,00","01/01/2025","","00000000000","00000000000","01/01/1990","Masculino","RUA EXEMPLO","100","CENTRO","Rio de Janeiro","RJ","20000-000","NOME DA MÃE"];
  var linhas=[cab.map(function(x){return '"'+x+'"'}).join(sep),ex.map(function(x){return '"'+x+'"'}).join(sep)].join("\n");
  baixar("Modelo_eSocial.csv",linhas,"text/csv");
}
function exportarEfetivo(){
  if(!temAcao("importar_efetivo")&&!canEdit()){alert("Sem permissão.");return}
  if(!COL.length){alert("Nada.");return}
  var sep=";",cab=["Empresa","Matrícula eSocial","Nome da Empresa","Situação","Funcionário","Nome do Funcionário","Função","Código Departamento","Departamento","Salário Base","Admissão","Prorrogação","CPF","PIS/PASEP","Nascimento","Sexo","Endereço","Número","Bairro","Cidade","UF","CEP","Mãe"];
  var linhas=[cab.map(function(x){return '"'+x+'"'}).join(sep)];
  COL.forEach(function(c){linhas.push([c.empresa||"",c.matricula||"",c.empresaNome||"",c.situacao||"",c.funcionarioCod||"",c.nome||"",c.funcao||"",c.deptoCod||"",c.departamento||"",c.salario||"",c.admissao||"",c.prorrogacao||"",c.cpf||"",c.pis||"",c.nascimento||"",c.sexo||"",c.endereco||"",c.numero||"",c.bairro||"",c.cidade||"",c.uf||"",c.cep||"",c.mae||""].map(function(x){return '"'+String(x==null?"":x).replace(/"/g,'""')+'"'}).join(sep))});
  baixar("Efetivo_"+dia(Date.now())+".csv",linhas.join("\n"),"text/csv");
  logAcao("Exportou efetivo");
}

/* ============ FICHA DO COLABORADOR ============ */
function linhaF(rot,val,small){return val?'<div class="row" style="border:0;padding:3px 0"><span class="s">'+rot+'</span><b style="font-size:'+(small?12:13)+'px;text-align:right;max-width:65%">'+esc(val)+'</b></div>':''}
function fichaCol(c){
  var h='<div class="card"><div style="display:flex;gap:12px;align-items:flex-start">'
    +'<div class="foto" onclick="'+(c.id==tsel?'document.getElementById(\'foto_inp_'+c.id+'\').click()':'selCol('+c.id+')')+'">'
    +(c.foto?'<img src="'+c.foto+'" alt="">':'<span style="font-size:22px">📷</span>')
    +'</div>'
    +'<div style="flex:1"><h4 style="margin:0 0 4px;font-size:16px">'+esc(c.nome)+'</h4>'
    +'<div class="s">'+esc(c.funcao||"")+(sitExib(c.situacao)?' · <b style="color:'+sitCor(c.situacao)+'">'+esc(sitExib(c.situacao))+'</b>':'')+'</div>'
    +(c.idiomaComun&&c.idiomaComun!=="pt-BR"?'<div class="s">'+esc(idiomaComunicNome(c.idiomaComun))+(c.idiomaIndigena?" · "+esc(c.idiomaIndigena):"")+'</div>':'')
    +(c.empresaNome?'<div class="s">'+esc(c.empresaNome)+'</div>':'')
    +'</div>'
    +'<button class="pill '+(c.id==tsel?'ok':'pen')+'" style="border:0;cursor:pointer" onclick="selCol('+c.id+')">'+(c.id==tsel?'▲':T("btn_open"))+'</button></div>'
    +'<input type="file" id="foto_inp_'+c.id+'" accept="image/*" style="display:none" onchange="fotoCol(this)">';
  if(c.id==tsel){
    if(c.empresa||c.matricula||c.departamento){
      h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">📋 eSocial</div>'
        +linhaF("Empresa",c.empresa)+linhaF("Matrícula",c.matricula)+linhaF("Cód. func.",c.funcionarioCod)
        +linhaF("Depto",(c.deptoCod||"")+(c.departamento?(c.deptoCod?" · ":"")+c.departamento:""),true)
        +linhaF("Salário",c.salario)+linhaF("Admissão",c.admissao)+linhaF("Prorrogação",c.prorrogacao)+'</div>';
    }
    if(c.cpf||c.pis||c.nascimento||c.sexo||c.mae){
      h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">👤 Pessoais</div>'
        +linhaF("CPF",c.cpf)+linhaF("PIS",c.pis)+linhaF("Nascimento",c.nascimento)+linhaF("Sexo",c.sexo)+linhaF("Mãe",c.mae,true)+'</div>';
    }
    if(c.endereco||c.cidade){
      h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">🏠 Endereço</div>'
        +linhaF("Logradouro",(c.endereco||"")+(c.numero?", "+c.numero:""))+linhaF("Bairro",c.bairro)
        +linhaF("Cidade/UF",(c.cidade||"")+(c.uf?"/"+c.uf:""))+linhaF("CEP",c.cep)+'</div>';
    }
    h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">🗣️ '+T("lang_comunic")+'</div>'
      +'<div class="row" style="border:0;padding:3px 0"><span class="s">Idioma</span><b style="font-size:13px">'+esc(idiomaComunicNome(c.idiomaComun||"nao"))+(c.idiomaIndigena?" · "+esc(c.idiomaIndigena):"")+'</b></div></div>';
    h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">🎓 Certificações NR</div><div class="nr-grid">'
      +NR_LIST.map(function(nr){var on=(c.nrs||[]).indexOf(nr)>=0;return '<button class="'+(on?'on':'')+'" onclick="toggleNR(\''+nr+'\')">'+(on?'✓ ':'')+nr+'</button>'}).join("")+'</div></div>';
    if(canEdit()){
      h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">⚙️ '+T("btn_edit")+'</div>'
        +'<input value="'+esc(c.nome)+'" onchange="setF(\'nome\',this.value)" placeholder="Nome">'
        +'<input value="'+esc(c.funcao||"")+'" onchange="setF(\'funcao\',this.value)" placeholder="Função">'
        +'<input value="'+esc(c.situacao||"")+'" onchange="setF(\'situacao\',this.value)" placeholder="Situação">'
        +'<div class="s">'+T("lang_comunic")+'</div>'
        +'<select onchange="colSetIdioma(this.value)">'
        +IDIOMAS_COMUNICACAO.map(function(x){return '<option value="'+x.id+'"'+(c.idiomaComun==x.id?' selected':'')+'>'+x.ico+' '+x.nome+'</option>'}).join("")
        +'</select>'
        +(c.idiomaComun=="ind"?'<input value="'+esc(c.idiomaIndigena||"")+'" onchange="setF(\'idiomaIndigena\',this.value)" placeholder="'+esc(T("lang_indigenous"))+'">':'')
        +'<textarea rows="2" placeholder="Observações" onchange="setF(\'observacoes\',this.value)" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:6px 0 8px">'+esc(c.observacoes||"")+'</textarea>'
        +'<div class="s">Escala</div><div>'+ESC.map(function(e,i){return chip(c.escala==e,"setEsc("+i+")",e)}).join("")+'</div>'
        +'<div class="s" style="margin-top:8px">EPIs</div><div>'+EPL.map(function(n,i){return chip((c.ep||[]).indexOf(n)>=0,"tep("+i+")",n)}).join("")+'</div>'
        +'<div style="display:flex;gap:6px;margin-top:8px"><input id="nep" placeholder="Novo EPI" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px;cursor:pointer" onclick="addEpi()">'+T("btn_add")+'</button></div>'
        +'</div>';
    }
    h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">📱 QR Code pessoal</div>'
      +'<div class="qrp" id="qrp_'+c.id+'"></div>'
      +'<button class="btn alt" onclick="baixarQRColab('+c.id+')">⬇️ Baixar QR</button></div>';
    if(canEdit()||temAcao("excluir_colaborador"))h+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><button class="btn red" onclick="delCol('+c.id+')">'+T("btn_delete")+'</button></div>';
  }
  return h+'</div>';
}
function colSetIdioma(v){
  var c=cget(tsel);if(!c)return;
  c.idiomaComun=v;
  if(v=="ind"&&!c.idiomaIndigena){var q=prompt(T("lang_indigenous"),"");c.idiomaIndigena=(q||"").trim()}
  if(v!="ind")c.idiomaIndigena="";
  svCol();draw();
}
function geraQRColab(){
  var c=cget(tsel);if(!c)return;
  var el=document.getElementById("qrp_"+c.id);if(!el)return;
  el.innerHTML="";
  if(typeof QRCode==="undefined"){el.innerHTML='<span class="s">Biblioteca de QR não carregou.</span>';return}
  /* Por privacidade (LGPD), o QR não leva mais CPF — apenas identificação funcional */
  try{new QRCode(el,{text:JSON.stringify({id:c.id,nome:c.nome,mat:c.matricula,funcao:c.funcao}),width:160,height:160})}catch(e){}
}
function baixarQRColab(id){
  var c=cget(id);if(!c)return;
  var el=document.getElementById("qrp_"+id);if(!el)return;
  var cv=el.querySelector("canvas");
  var img=el.querySelector("img");
  var nome="QR_"+c.nome.replace(/[^\w]/g,"_")+".png";
  if(cv){cv.toBlob(function(b){var u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download=nome;document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);if(a.parentNode)a.parentNode.removeChild(a)},1200)},"image/png");return}
  if(img&&img.src){var a=document.createElement("a");a.href=img.src;a.download=nome;document.body.appendChild(a);a.click();a.remove();return}
  alert("QR ainda não gerado.");
}

/* ============ TELA EQUIPE ============ */
function buscaColab(el){
  sv("rad_buscaCol",el.value);var p=el.selectionStart;var y=window.scrollY;draw();window.scrollTo(0,y);
  var e=document.getElementById("buscaColInp");if(e){e.focus();try{e.setSelectionRange(p,p)}catch(x){}}
}
function deptosLista(){
  var m={};
  COL.filter(colDoMeuEscopo).forEach(function(c){var k=chaveDepto(c)||"—";if(!m[k])m[k]={k:k,nome:k=="—"?"Sem departamento":nomeDepto(c),cod:c.deptoCod||"",n:0};m[k].n++});
  return Object.keys(m).map(function(k){return m[k]}).sort(function(a,b){if(a.k=="—")return 1;if(b.k=="—")return -1;return a.nome.localeCompare(b.nome)});
}
var deptoAberto=ld("rad_deptoAberto",{});
var DEPS_CACHE=[];
function togDeptoI(i){var d=DEPS_CACHE[i];if(d)togDepto(d.k)}
function togDepto(k){deptoAberto[k]=!deptoAberto[k];sv("rad_deptoAberto",deptoAberto);var y=window.scrollY;draw();window.scrollTo(0,y)}
function vEquipe(){
  var busca=ld("rad_buscaCol",""),fDep=ld("rad_filtroDepto",""),lista=COL.filter(colDoMeuEscopo);
  if(busca){var t=busca.toLowerCase();lista=lista.filter(function(c){return(c.nome||"").toLowerCase().indexOf(t)>=0||(c.cpf||"").indexOf(t)>=0||(c.matricula||"").indexOf(t)>=0||(c.funcao||"").toLowerCase().indexOf(t)>=0||(c.departamento||"").toLowerCase().indexOf(t)>=0})}
  if(fDep)lista=lista.filter(function(c){return (chaveDepto(c)||"—")==fDep});
  var deps=deptosLista();
  var h='<div class="card"><b>'+T("menu_eq")+'</b>'
    +'<div class="s">Departamento / contrato</div><select onchange="sv(\'rad_filtroDepto\',this.value);draw()"><option value="">Todos ('+deps.length+')</option>'
    +deps.map(function(d){return '<option value="'+esc(d.k)+'"'+(d.k==fDep?' selected':'')+'>'+esc((d.cod?d.cod+" · ":"")+d.nome)+' ('+d.n+')</option>'}).join("")+'</select>'
    +'<div class="s">'+T("login_search")+'</div><input id="buscaColInp" value="'+esc(busca)+'" oninput="buscaColab(this)" placeholder="🔍"><div class="s" style="margin-top:8px"><b>'+lista.length+'</b> de <b>'+COL.filter(colDoMeuEscopo).length+'</b> colaborador(es)</div></div>';
  if(!COL.length)return h+'<div class="card"><div class="s">Nenhum colaborador. Use <b>'+T("menu_imp")+'</b>.</div></div>';
  var grupos={};lista.forEach(function(c){var k=chaveDepto(c)||"—";(grupos[k]=grupos[k]||[]).push(c)});
  var filtrando=!!(busca||fDep);
  DEPS_CACHE=deps;
  deps.forEach(function(d,di){
    var g=grupos[d.k];if(!g)return;
    g.sort(function(a,b){return (a.nome||"").localeCompare(b.nome||"")});
    var aberto=filtrando||deptoAberto[d.k]||g.some(function(c){return c.id==tsel});
    h+='<div class="card noprint" style="border-left:4px solid var(--brand);cursor:pointer;margin-bottom:8px" onclick="togDeptoI('+di+')">'
      +'<div class="row" style="border:0;padding:0"><span style="flex:1"><b>🏢 '+esc(d.nome)+'</b>'+(d.cod?' <span class="s">· cód. '+esc(d.cod)+'</span>':'')+'</span><span class="pill ok" style="cursor:pointer">'+g.length+' '+(aberto?'▲':'▼')+'</span></div></div>';
    if(aberto)h+=g.map(fichaCol).join("");
  });
  if(canEdit())h+='<div class="card"><b>➕ '+T("btn_add")+'</b><div style="display:flex;gap:6px;margin-top:8px"><input id="cn" placeholder="Nome completo" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px;cursor:pointer" onclick="addColM()">'+T("btn_add")+'</button></div></div>';
  return h;
}
function addColM(){if(!canEdit()){alert("Sem permissão.");return}var n=$("#cn").value.trim();if(!n){alert("Informe o nome.");return}var c=novoCol({nome:n},0);COL.push(c);tsel=c.id;svCol();sv("rad_tsel",tsel);logAcao("Cadastrou colaborador",n);draw()}

/* ============ IMPORTAÇÃO — TELA ============ */
function vImp(){
  if(!temAcao("importar_efetivo")&&!canEdit())return '<div class="card s">Sem permissão.</div>';
  var ativos=COL.filter(function(c){return ehAtivo(c.situacao)}).length;
  var afast=COL.filter(function(c){return ehAfast(c.situacao)}).length;
  var deps=deptosLista();
  return '<div class="card"><b>'+T("menu_imp")+' (eSocial)</b><div class="s">Detecta as colunas automaticamente. Cada <b>Departamento</b> vira um contrato e os colaboradores são vinculados a ele. Reimportar atualiza os dados (inclusive troca de departamento).</div></div>'
    +'<div class="card"><b>📊 Base atual</b><div class="row"><span>Total</span><b>'+COL.length+'</b></div><div class="row"><span>🟢 Ativos</span><b style="color:var(--ok)">'+ativos+'</b></div><div class="row"><span>🔴 Afastados</span><b style="color:var(--warn)">'+afast+'</b></div><div class="row"><span>🏢 Departamentos / contratos</span><b>'+deps.filter(function(d){return d.k!="—"}).length+'</b></div></div>'
    +'<div class="card"><b>📋 Colar do Excel</b><textarea id="pi" rows="6" placeholder="Cole aqui as linhas copiadas do Excel (com o cabeçalho)" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:8px 0"></textarea><button class="btn" onclick="pasteImp()">Importar</button></div>'
    +'<div class="card"><b>📁 Enviar arquivo</b><input type="file" accept=".xlsx,.xls,.csv,.txt" onchange="fileImp(this)" style="padding:10px"></div>'
    +'<div class="card"><b>⬇️ Exportar</b><button class="btn alt" onclick="exportarEfetivo()">Baixar CSV</button></div>'
    +'<div class="card"><b>📄 Modelo</b><button class="btn alt" onclick="baixarModeloEfetivo()">Baixar modelo</button></div>';
}

/* ============ TREINAMENTOS ============ */
function addTrn(){
  if(!canEdit()){alert("Sem permissão.");return}
  var t=$("#trn_t").value.trim();
  if(!t){alert("Informe o título.");return}
  TRN.push({id:Date.now(),titulo:t,descricao:($("#trn_d").value||"").trim(),categoria:$("#trn_c").value,data:$("#trn_data").value||dia(Date.now()),carga:($("#trn_ch").value||"").trim(),resp:($("#trn_r").value||"").trim(),local:($("#trn_l").value||"").trim(),part:($("#trn_p").value||"").trim(),criado:Date.now()});
  sv("rad_trn",TRN);logAcao("Criou treinamento",t);draw();
}
function delTrn(id){
  if(!canEdit()){alert("Sem permissão.");return}
  if(!confirm("Excluir?"))return;
  TRN=TRN.filter(function(x){return x.id!=id});sv("rad_trn",TRN);logAcao("Excluiu treinamento");draw();
}
function fmtDataBR(s){var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(s||"");return m?m[3]+"/"+m[2]+"/"+m[1]:(s||"")}
function vT(){
  var pode=canEdit();
  var l=TRN.slice().sort(function(a,b){return(b.data||"").localeCompare(a.data||"")});
  var h='<div class="card"><b>📚 '+T("menu_trein")+' (REG-011)</b><div class="s">Total: '+l.length+'</div><button class="btn alt" onclick="go(\'dc\')">🎬 POPs, vídeos e documentos traduzidos</button></div>';
  if(l.length){h+='<div class="card">'+l.map(function(t){return '<div class="trn-card"><h4>'+esc(t.titulo)+'</h4><div class="trn-meta">'+esc(fmtDataBR(t.data))+' · '+esc(t.categoria)+(t.carga?' · '+esc(t.carga)+'h':'')+(t.resp?' · '+esc(t.resp):'')+(t.local?' · '+esc(t.local):'')+'</div>'+(t.descricao?'<div class="trn-desc">'+esc(t.descricao)+'</div>':'')+(t.part?'<div class="trn-part"><b>Participantes:</b> '+esc(t.part)+'</div>':'')+(pode?'<button class="pill no" style="border:0;margin-top:8px;cursor:pointer" onclick="delTrn('+t.id+')">'+T("btn_delete")+'</button>':'')+'</div>'}).join("")+'</div>'}
  else h+='<div class="card s">Nenhum.</div>';
  if(pode){h+='<div class="card"><b>➕ '+T("btn_add")+'</b><input id="trn_t" placeholder="Título"><textarea id="trn_d" rows="3" placeholder="Descrição" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:6px 0 8px"></textarea><select id="trn_c">'+TRN_CAT.map(function(c){return '<option>'+c+'</option>'}).join("")+'</select><input id="trn_data" type="date" value="'+dia(Date.now())+'"><input id="trn_ch" placeholder="Carga (h)"><input id="trn_r" placeholder="Instrutor"><input id="trn_l" placeholder="Local"><input id="trn_p" placeholder="Participantes"><button class="btn" onclick="addTrn()">'+T("btn_add")+'</button></div>'}
  return h;
}

/* ============ AVALIAÇÕES (Luz Negra / REG-016 / Customizáveis) ============ */
var LZI="Teto|N|Centro do teto;Cantos;Junção teto/parede~Saídas do ar-condicionado|S|Grade frontal;Aletas;Bordas;Entorno imediato~Janelas / Vidros / Espelhos|N|Centro;Bordas;Cantos;Puxador/moldura~Paredes|N|Área central;Cantos;Perto de interruptores;Perto de portas~Equipamentos assistenciais|C|Painel;Comandos;Alças;Superfícies de contato;Cabos/áreas externas acessíveis~Dispensadores|S|Botão/alavanca;Bico de saída;Frente;Laterais;Parte inferior~Mobiliários|S|Tampo;Bordas;Puxadores;Laterais;Áreas de contato~Porta / Batentes / Maçanetas|C|Maçaneta interna;Maçaneta externa;Batente direito;Batente esquerdo;Área de contato da porta~Pias de inox|C|Cuba;Torneira;Acionamento;Borda frontal;Bordas laterais;Válvula/ralo~Leito / Grade do leito|C|Grade interna;Grade externa;Cabeceira;Peseira;Estrutura acessível;Controles~Pisos|N|Área de circulação;Cantos;Entorno do mobiliário;Área próxima ao leito~Armários|S|Puxadores;Portas;Bordas;Superfície superior acessível~Mesa dos computadores|S|Tampo;Bordas;Teclado;Mouse;Apoio das mãos~Lixeiras|C|Tampa;Acionamento;Alça/pedal;Borda;Área externa frontal~Vaso sanitário|C|Assento superior;Assento inferior;Tampa;Descarga;Borda externa;Base acessível~Mictório|C|Cuba;Bordas;Descarga;Acionamento;Área frontal;Base acessível~Pia|C|Cuba;Torneira;Acionamento;Bordas;Válvula/ralo~Trocador de fraldas|C|Superfície de troca;Laterais;Cabeceira;Bordas;Acionamento/estrutura~Cadeira|S|Assento;Encosto;Braços;Bordas;Estrutura de contato~Poltrona|S|Assento;Encosto;Braços;Laterais;Bordas~Telefone|S|Fone;Microfone;Teclado;Botões;Base~Bancada|S|Tampo;Bordas;Cantos;Frente;Áreas de contato~Monitor|C|Tela;Bordas;Botões;Suporte;Cabos/áreas externas acessíveis~Interruptor|S|Tecla/botão;Bordas;Moldura;Área imediatamente adjacente~Corrimão|S|Parte superior;Parte inferior;Extremidade direita;Extremidade esquerda;Suportes~Torneiras / Acionamentos|C|Corpo da torneira;Bico/saída;Alavanca/acionamento;Base;Sensor ou botão de acionamento~Controles remotos / Painéis de comando|S|Botões;Superfície frontal;Laterais;Parte traseira;Cabo/suporte".split("~").map(function(s){var p=s.split("|");return{n:p[0],c:p[1],p:p[2].split(";")}});
var LZN={C:"Crítico",S:"Semicrítico",N:"Não crítico"},LZK={C:"#c62828",S:"#e0a100",N:"#2e8b57"},lzA={},lzO=-1,lzE=0,lzT="Concorrente",lzU="Manhã",lzObs="",cklSel="lz";
var RCI="Teto|N~Saídas do ar condicionado|S~Luminárias|N~Foco de luz|S~Janela / Vidros / Espelhos|N~Paredes|N~Cortinas|S~Equipamentos|C~Dispensadores|S~Mobiliários|S~Porta / Batentes / Maçanetas|C~Pias de inox|C~Leito / Grade do leito|C~Pisos|N~Armários|S~Mesa dos computadores|S~Lixeiras|C".split("~").map(function(s){var p=s.split("|");return{n:p[0],c:p[1]}});
var lzP={},pa=ld("rad_pa",[]),PAS=["Aberto","Em andamento","Concluído"];
var PRZ=ld("rad_prz3",{C:30,S:120,N:1440}),PRM=ld("rad_prm","Automático");
var AVALS=ld("rad_avals",[]);
function svAvals(){sv("rad_avals",AVALS)}
function avalsDoChkl(id){return AVALS.filter(function(a){return a.chklId==id})}
function lzCalc(A,itens){
  itens=itens||LZI;
  var g={C:[0,0],S:[0,0],N:[0,0]},cnc=0;
  itens.forEach(function(it,i){
    var k=g[it.c]?it.c:"N";
    if(it.p&&it.p.length){it.p.forEach(function(x,j){var v=A[i+"_"+j];if(!v)return;g[k][0]++;if(v=="C")g[k][1]++;else if(k=="C")cnc++})}
    else{var v=A[i];if(!v)return;g[k][0]++;if(v=="C")g[k][1]++;else if(k=="C")cnc++}
  });
  var a=g.C[0]+g.S[0]+g.N[0],c=g.C[1]+g.S[1]+g.N[1],p=a?Math.round(c/a*1000)/10:0,r=!a?"Sem pontos":cnc?"NC crítico":p==100?"Conforme":p>=90?"Conforme c/ desvio":p>=80?"Não conforme":"NC crítico";
  return{g:g,a:a,c:c,p:p,r:r,cnc:cnc};
}
function lzCor(r){return r=="Conforme"?"#2e8b57":r=="Conforme c/ desvio"?"#e0a100":r=="Sem pontos"?"#5d6e69":"#c62828"}
function lzPill(r,t){return '<span class="pill" style="background:'+lzCor(r)+';cursor:default">'+esc(t)+'</span>'}
function lzSet(k,v){var y=window.scrollY;lzA[k]=lzA[k]==v?0:v;draw();window.scrollTo(0,y)}
function lzOp(i){var y=window.scrollY;lzO=lzO==i?-1:i;draw();window.scrollTo(0,y)}
function prazoMs(c){return Date.now()+(PRZ[c]||0)*60000}
function toTs(s){return s?new Date(s).getTime():0}
function tsLoc(t){return t?new Date(t-new Date(t).getTimezoneOffset()*60000).toISOString().slice(0,16):""}
function fmtPrazo(m){return m%1440==0?(m/1440)+" dia(s)":m>=60&&m%60==0?(m/60)+" h":m+" min"}
function przTs(p){return typeof p.prazo=="number"?p.prazo:(p.prazo?new Date(p.prazo+"T23:59").getTime():0)}
function setPrz(k,v){PRZ[k]=Math.max(0,+v||0);sv("rad_prz3",PRZ);draw()}
function setPrm(v){PRM=v;sv("rad_prm",v);draw()}
function prCfg(){if(!canEdit())return "";return '<div class="card"><b>Prazo</b><select onchange="setPrm(this.value)">'+["Automático","Programado"].map(function(x){return '<option'+(x==PRM?' selected':'')+'>'+x+'</option>'}).join("")+'</select>'+(PRM=="Automático"?["C","S","N"].map(function(k){return '<div class="row"><span>'+LZN[k]+'</span><span><input type="number" style="width:80px;display:inline;margin:0" value="'+PRZ[k]+'" onchange="setPrz(\''+k+'\',this.value)"> min <span class="s">('+fmtPrazo(PRZ[k])+')</span></span></div>'}).join(""):'')+'</div>'}
function lzSel(a,c,v){return '<select onchange="'+v+'=this.value">'+a.map(function(x){return '<option'+(x==c?' selected':'')+'>'+x+'</option>'}).join("")+'</select>'}
function lzFoto(i,inp){var f=inp.files[0];if(!f)return;shrink(f,function(d){lzP[i]=lzP[i]||{};lzP[i].foto=d;var y=window.scrollY;draw();window.scrollTo(0,y)})}
function paFoto(id,k,inp){var f=inp.files[0];if(!f)return;shrink(f,function(d){paSet(id,k,d,1)})}
function paFotos(p){function im(s){return s?'<img src="'+s+'" alt="" style="max-width:100%;border-radius:8px;margin:4px 0">':''}return '<div class="s">Foto NC</div>'+im(p.fotoNC)+(p.fotoNC?'':'<input type="file" accept="image/*" capture="environment" onchange="paFoto('+p.id+',\'fotoNC\',this)">')+'<div class="s">Foto OK</div>'+im(p.fotoOK)+'<input type="file" accept="image/*" capture="environment" onchange="paFoto('+p.id+',\'fotoOK\',this)">'}
function lzPS(i,k,v){lzP[i]=lzP[i]||{};lzP[i][k]=v}
function lzForm(i,bad,c,itens){itens=itens||chklItens(cklSel);c=c||(itens[i]&&itens[i].c);var d=lzP[i]||{};return '<div style="border:2px solid var(--warn);border-radius:10px;padding:10px;margin:8px 0"><b style="color:var(--warn)">Plano de ação</b><div class="s">NC: '+esc(bad.join(", "))+'</div><div class="s">Causador da NC: <b>'+esc(lzCol?lzCol.nome:"(informe quem limpou o local)")+'</b></div><input list="dl_resolve" placeholder="Quem vai resolver" value="'+esc(d.resp||"")+'" onchange="lzPS('+i+',\'resp\',this.value)"><datalist id="dl_resolve">'+equipeDaArea(loc).map(function(c){return '<option value="'+esc(c.nome)+'">'}).join("")+'<option value="Equipe de Higienização"></datalist><input type="datetime-local" value="'+(d.prazo||"")+'" onchange="lzPS('+i+',\'prazo\',this.value)"><input placeholder="Ação" value="'+esc(d.acao||"")+'" onchange="lzPS('+i+',\'acao\',this.value)">'+(d.foto?'<img src="'+d.foto+'" alt="" style="max-width:100%;border-radius:8px;margin:4px 0">':'')+'<input type="file" accept="image/*" capture="environment" onchange="lzFoto('+i+',this)"></div>'}
/* CORREÇÃO: "Todos C/N" usava sempre a lista Luz Negra, mesmo em checklists personalizados */
function lzAll(i,v){var itens=chklItens(cklSel);var y=window.scrollY;var pontos=(itens[i]&&itens[i].p)||[];pontos.forEach(function(x,j){lzA[i+"_"+j]=v});draw();window.scrollTo(0,y)}
function paSet(id,k,v,r){pa.forEach(function(p){if(p.id==id){p[k]=v;if(k=="st"&&v=="Concluído"){p.fim=Date.now();p.conclusao=Date.now()}if(k=="st"&&v!="Concluído")p.conclusao=null}});sv("rad_pa",pa);if(r)draw()}
function paCard(){var ab=pa.filter(function(p){return p.st!="Concluído"});return '<div class="card"><b>Planos de ação</b><div class="s">'+ab.length+' aberto · '+(pa.length-ab.length)+' concluído</div>'+ab.slice().reverse().slice(0,15).map(function(p){var t=przTs(p),at=t&&Date.now()>t;return '<div style="border-top:1px solid var(--line);padding:10px 0"><div class="row" style="border:0;padding:0"><b>'+esc(p.item)+'</b><span class="pill" style="background:'+(LZK[p.c]||"#5d6e69")+';cursor:default">'+(LZN[p.c]||"—")+'</span></div><div class="s">'+esc(nomeLoc(p.l))+' · '+esc(nomeEsp(p.l,p.e))+(t?' · prazo '+new Date(t).toLocaleString("pt-BR"):'')+(at?' · <b style="color:var(--warn)">atrasado</b>':'')+'</div>'+(p.pts&&p.pts.length?'<div class="s">NC: '+esc(p.pts.join(", "))+'</div>':'')+(p.causador?'<div class="s">Causador: <b>'+esc(p.causador)+'</b></div>':'')+paFotos(p)+'<input placeholder="Quem vai resolver" value="'+esc(p.resp)+'" onchange="paSet('+p.id+',\'resp\',this.value)"><input placeholder="Ação" value="'+esc(p.acao)+'" onchange="paSet('+p.id+',\'acao\',this.value)"><select onchange="paSet('+p.id+',\'st\',this.value,1)">'+PAS.map(function(s){return '<option'+(s==p.st?' selected':'')+'>'+s+'</option>'}).join("")+'</select></div>'}).join("")+'</div>'}
function planosVencidos(){return pa.filter(function(p){if(p.st=="Concluído")return false;var t=przTs(p);return t&&Date.now()>t})}
function alertaVencidos(){if(!session)return "";if(acessosDoUsuario().indexOf("l")<0)return "";var v=planosVencidos();if(!v.length)return "";return '<div class="card noprint" style="border-left:4px solid var(--warn);background:rgba(198,40,40,.06);cursor:pointer" onclick="go(\'l\')"><div style="display:flex;align-items:center;gap:10px"><span style="font-size:24px">⚠️</span><span style="flex:1"><b style="color:var(--warn)">'+v.length+' plano(s) de ação vencido(s)</b></span></div></div>'}
function alertaPtisPendentes(){
  if(!session)return "";
  if(acessosDoUsuario().indexOf("pti")<0)return "";
  var pend=ptisPendentesParaMim();
  if(!pend.length)return "";
  return '<div class="card noprint" style="border-left:4px solid var(--run);background:rgba(224,161,0,.08);cursor:pointer" onclick="go(\'pti\')"><div style="display:flex;align-items:center;gap:10px"><span style="font-size:24px">📋</span><span style="flex:1"><b style="color:#8a6400">'+pend.length+' PTI(s) aguardando</b><div class="s">Toque para analisar.</div></span></div></div>';
}
function modoSel(){
  var ativos=chklDoLoc();
  var t=tipoDoLoc(loc);
  return '<div class="card"><div class="s">Tipo de avaliação'+(t?' · '+esc(tipoCtNome(t)):'')+'</div><select onchange="cklSel=this.value;lzA={};lzP={};lzO=-1;draw()">'
    +ativos.map(function(c){return '<option value="'+c.id+'"'+(cklSel==c.id?' selected':'')+'>'+c.ico+' '+esc(c.nome)+'</option>'}).join("")
    +'</select></div>';
}
function lzSave(){
  var c=chklById(cklSel);
  var itens=chklItens(cklSel);
  var s=lzCalc(lzA,itens);
  if(!s.a){alert("Avalie pelo menos um ponto.");return}
  if(espOff(loc,lzE)||!LOC[loc]||LOC[loc].e[lzE]==null){alert("Selecione um espaço válido.");return}
  if(!lzCol){alert("Informe quem limpou o local (ou escolha \"Não identificado\").");return}
  AVALS.push({
    id:Date.now(),
    chklId:cklSel,
    chklNome:c?c.nome:"",
    l:loc,e:lzE,
    ini:Date.now(),
    av:user,
    t:lzT,u:lzU,o:lzObs,
    A:lzA,
    limpou:{colId:lzCol.colId,nome:lzCol.nome},
    s:{p:s.p,r:s.r,a:s.a,c:s.c,cnc:s.cnc,g:s.g}
  });
  svAvals();
  var np=0,base=Date.now();
  itens.forEach(function(it,i){
    var bad=[];
    if(it.p&&it.p.length){
      it.p.forEach(function(x,j){if(lzA[i+"_"+j]=="N")bad.push(x)});
    }else if(lzA[i]=="N")bad.push(it.n);
    if(bad.length){
      pa.push({id:base+i+1,ct:ctDaArea(loc),origem:"inspeção",causadorId:lzCol.colId,causador:lzCol.nome,quem:(lzP[i]||{}).resp||"",l:loc,e:lzE,item:(it.n||"Item "+(i+1))+" ["+(c?c.nome:"")+"]",c:it.c,pts:bad,ini:Date.now(),resp:(lzP[i]||{}).resp||"",prazo:toTs((lzP[i]||{}).prazo)||(PRM=="Automático"?prazoMs(it.c):0),acao:(lzP[i]||{}).acao||"",fotoNC:(lzP[i]||{}).foto||"",fotoOK:"",st:"Aberto"});
      np++;
    }
  });
  sv("rad_pa",pa);
  lzA={};lzP={};lzObs="";lzO=-1;lzCol=null;
  logAcao("Salvou avaliação",(c?c.nome:"")+" · "+s.r);
  alert("Salvo: "+s.r+" ("+s.p+"%)"+(np?". "+np+" plano(s) aberto(s).":""));
  draw();
}
function tipoDoLoc(l){return tipoDaArea(l)}
/* equipe do contrato da área (para "quem limpou" e "causador da NC") */
function equipeDaArea(l){
  var ct=ctDaArea(l);
  return COL.filter(function(c){return !ehDeslig(c.situacao)&&(ct==null||(c.vinculos||[]).some(function(v){return v.contratoId==ct}))}).sort(function(a,b){return (a.nome||"").localeCompare(b.nome||"")});
}
/* sugestão: quem registrou a última limpeza desse espaço nas últimas 24 h */
function ultimoQueLimpou(l,e){
  var r=regs.filter(function(x){return x.l==l&&x.e==e&&x.st=="ok"&&Date.now()-x.ini<86400000}).sort(function(a,b){return b.ini-a.ini})[0];
  if(!r)return null;
  var c=COL.filter(function(x){return normT(x.nome)==normT(r.colab)})[0];
  return {colId:c?c.id:null,nome:r.colab,quando:r.fim||r.ini};
}
var lzCol=null;
function lzColSel(v){
  if(v=="__ni")lzCol={colId:null,nome:"Não identificado"};
  else if(!v)lzCol=null;
  else{var c=cget(+v);lzCol=c?{colId:c.id,nome:c.nome}:null}
}
function lzColCard(){
  var eq=equipeDaArea(loc),sug=ultimoQueLimpou(loc,lzE);
  if(!lzCol&&sug&&sug.colId)lzCol={colId:sug.colId,nome:sug.nome,auto:true};
  var at=lzCol?(lzCol.colId||"__ni"):"";
  return '<div class="s" style="margin-top:6px"><b>Quem limpou este local</b> (obrigatório)</div>'
    +'<select onchange="lzColSel(this.value)"><option value="">— selecione —</option>'
    +eq.map(function(c){return '<option value="'+c.id+'"'+(at==c.id?' selected':'')+'>'+esc(c.nome)+(c.funcao?' · '+esc(c.funcao):'')+'</option>'}).join("")
    +'<option value="__ni"'+(at=="__ni"?' selected':'')+'>Não identificado</option></select>'
    +(sug?'<div class="s">💡 Último registro no espaço: <b>'+esc(sug.nome)+'</b> às '+hm(sug.quando)+'</div>':'<div class="s">Sem limpeza registrada neste espaço nas últimas 24 h.</div>');
}
function chklDoLoc(){var t=tipoDoLoc(loc);return CHECKLISTS.filter(function(c){return c.ativo&&chklServeTipo(c,t)})}
function vL(){
  if(loc<0||!LOC[loc])return ctSelCard()+semAreasCard();
  var L=LOC[loc];
  var disp=chklDoLoc();
  if(!disp.some(function(c){return c.id==cklSel})){cklSel=disp.length?disp[0].id:"";lzA={};lzP={};lzO=-1}
  if(!disp.length)return ctSelCard()+'<div class="card s">Nenhum checklist ativo para o tipo de serviço desta área ('+esc(tipoCtNome(tipoDoLoc(loc)))+'). Crie um em Configurações → Checklists.</div>';
  var c=chklById(cklSel);
  var itens=chklItens(cklSel);
  if(!itens.length)return modoSel()+'<div class="card s">Checklist sem itens. Edite em Configurações → Checklists.</div>';
  if(L.e[lzE]==null||espOff(loc,lzE)){lzE=L.e.findIndex(function(_,i){return !espOff(loc,i)});if(lzE<0)lzE=0}
  var s=lzCalc(lzA,itens);
  var h=ctSelCard()+modoSel()+'<div class="card"><b>'+(c?c.ico+" "+esc(c.nome):"Avaliação")+'</b><div class="s">Data: '+new Date().toLocaleDateString("pt-BR")+'</div><select onchange="lzE=0;lzA={};lzP={};lzCol=null;setLoc(this.value)">'+opts(loc)+'</select><select onchange="lzE=+this.value;lzCol=null;draw()">'+L.e.map(function(n,i){if(espOff(loc,i))return "";return '<option value="'+i+'"'+(i==lzE?' selected':'')+'>'+esc(n)+'</option>'}).join("")+'</select><div class="s">Tipo</div>'+lzSel(TIPOS.map(function(t){return t[1]}),lzT,"lzT")+'<div class="s">Turno</div>'+lzSel(["Manhã","Tarde","Noite"],lzU,"lzU")+lzColCard()+'</div>';
  h+='<div class="card">'+itens.map(function(it,i){
    var n=0,bad=[],temP=it.p&&it.p.length;
    if(temP){
      it.p.forEach(function(x,j){var v=lzA[i+"_"+j];if(v)n++;if(v=="N")bad.push(x)});
    }else{
      if(lzA[i])n++;
      if(lzA[i]=="N")bad.push(it.n);
    }
    var total=temP?it.p.length:1;
    var ico=temP?"▼":"•";
    return '<div class="row" onclick="lzOp('+i+')" style="cursor:pointer"><span>'+(i+1)+'. <b>'+esc(it.n)+'</b> <span class="pill" style="background:'+(LZK[it.c]||"#5d6e69")+';cursor:default">'+(LZN[it.c]||"—")+'</span></span><span>'+(bad.length?'<span class="pill" style="background:var(--warn);cursor:default">plano</span> ':'')+'<span class="s">'+n+'/'+total+' '+ico+'</span></span></div>'
      +(lzO==i?(temP?
        '<div class="chips" style="margin:6px 0"><button onclick="lzAll('+i+',\'C\')">Todos C</button><button onclick="lzAll('+i+',\'N\')">Todos N</button></div>'+it.p.map(function(x,j){var v=lzA[i+"_"+j];return '<div style="padding:6px 4px 10px"><div class="s">'+esc(x)+'</div><div class="chips" style="margin:4px 0 0"><button class="'+(v=="C"?'sel':'')+'" onclick="lzSet(\''+i+'_'+j+'\',\'C\')">C</button><button style="'+(v=="N"?'background:var(--warn);color:#fff;border-color:var(--warn)':'')+'" onclick="lzSet(\''+i+'_'+j+'\',\'N\')">N</button></div></div>'}).join("")+(bad.length?lzForm(i,bad,it.c,itens):'')
        :'<div style="padding:6px 4px 10px"><div class="chips" style="margin:4px 0 0"><button class="'+(lzA[i]=="C"?'sel':'')+'" onclick="lzSet(\''+i+'\',\'C\')">Conforme</button><button style="'+(lzA[i]=="N"?'background:var(--warn);color:#fff;border-color:var(--warn)':'')+'" onclick="lzSet(\''+i+'\',\'N\')">Não conforme</button></div>'+(lzA[i]=="N"?lzForm(i,[it.n],it.c,itens):'')+'</div>')
      :'');
  }).join("")+'</div>';
  h+='<div class="card tbl"><b>Resumo</b>'+clsTbl(s.g)+'<div style="margin-top:10px">'+lzPill(s.r,s.r)+'</div>'+(s.cnc?'<div class="s" style="color:var(--warn)">'+s.cnc+' ponto(s) crítico(s) NC.</div>':'')+'<input placeholder="Obs" value="'+esc(lzObs)+'" onchange="lzObs=this.value" style="margin-top:10px"><button class="btn" onclick="lzSave()">Finalizar</button></div>';
  var avals=avalsDoChkl(cklSel).slice().reverse().slice(0,10);
  h+='<div class="card"><b>Avaliações salvas</b>'+(avals.length?avals.map(function(r){return '<div class="row"><span><b>'+esc(nomeEsp(r.l,r.e))+'</b><div class="s">'+esc(nomeLoc(r.l))+' · '+new Date(r.ini).toLocaleDateString("pt-BR")+' '+hm(r.ini)+(r.av?' · avaliou: '+esc(r.av):'')+(r.limpou?' · limpou: '+esc(r.limpou.nome):'')+'</div></span>'+lzPill(r.s.r,r.s.p+'% · '+r.s.r)+'</div>'}).join(""):'<div class="s">Nenhuma.</div>')+'</div>';
  return h+prCfg()+paCard();
}
function clsTbl(g){var Tt=[0,0],pc=function(x){return x[0]?(Math.round(x[1]/x[0]*1000)/10).toString().replace(".",",")+"%":"—"};return '<table style="width:100%;margin-top:10px"><tr><th>Grupo</th><th>Avaliados</th><th>Conf.</th><th>%</th></tr>'+["C","S","N"].map(function(k){var x=g[k];Tt[0]+=x[0];Tt[1]+=x[1];return '<tr><td>'+LZN[k]+'</td><td>'+x[0]+'</td><td>'+x[1]+'</td><td>'+pc(x)+'</td></tr>'}).join("")+'<tr><td><b>Total</b></td><td>'+Tt[0]+'</td><td>'+Tt[1]+'</td><td>'+pc(Tt)+'</td></tr></table>'}



/* ============ GIRO DE SALA (com Recusa + Chamado) ============ */
var giro=ld("rad_giro",[]),GT=ld("rad_gtempo",{CO:[5,15],TE:[30,60],PR:[30,60]}),gtipo="CO";
var gopen=ld("rad_gopen",null),GS=["Sala 01","Sala 02","Sala 03","Sala 04","Sala 05","Laser"];
function gmin(r){return Math.max(0,Math.round((r.fim-r.ini)/60000))}
function gst(r){var t=GT[r.tipo];if(!t)return"";var m=gmin(r);return m<t[0]?"Abaixo":m>t[1]?"Acima":"OK"}
function ghoje(){return giro.filter(function(r){return dia(r.ini)==dia(Date.now())})}
function gini(s){if(gopen){alert("Termine: "+GS[gopen.s]);return}if(open){alert("Termine primeiro a atividade da rotina em andamento.");return}gopen={s:s,ini:Date.now()};gtipo="CO";sv("rad_gopen",gopen);draw()}
function gfim(){var obs=(($("#go")||{}).value||"").trim();giro.push({s:gopen.s,ini:gopen.ini,fim:Date.now(),tipo:gtipo,colab:user||"",obs:obs});sv("rad_giro",giro);gopen=null;sv("rad_gopen",null);draw()}
/* CORREÇÃO: a recusa de giro usava um espaço da área 1 (índice errado) e o formulário nunca aparecia na tela de Giro */
function gRecusar(){open={l:-1,e:gopen.s,ini:gopen.ini,origem:"giro"};sv("rad_open",open);rec=true;gopen=null;sv("rad_gopen",null);draw()}
function gChamado(){var s=gopen.s;tab="hOP";buildNav();draw();setTimeout(function(){var x=document.getElementById("cx");if(x){x.value="Giro de Sala: "+GS[s];x.focus()}},300)}
function setG(k,i,v){GT[k][i]=Math.max(0,+v||0);sv("rad_gtempo",GT);draw()}
function gcor(r){return{"OK":"ok","Abaixo":"no","Acima":"run"}[gst(r)]||"ok"}
function vG(){
  if(open&&rec&&open.origem=="giro")return formRec();
  var h='<div class="card"><b>Giro de sala: CC</b> <span class="s">(REG-202 B)</span></div>';
  if(gopen){
    h+='<div class="card run-card"><b>Em andamento: '+GS[gopen.s]+'</b><div class="s">Início '+hm(gopen.ini)+'</div>'+semaf(gopen.ini,GT[gtipo]||[0,9999])
      +'<div class="chips">'+TIPOS.map(function(t){return '<button class="'+(gtipo==t[0]?'sel':'')+'" onclick="gtipo=\''+t[0]+'\';draw()">'+t[1]+'</button>'}).join("")+'</div>'
      +'<input id="go" placeholder="Observação">'
      +'<button class="btn" onclick="gfim()">Terminei</button>'
      +'<div class="grid-acoes" style="margin-top:8px">'
        +'<button class="btn alt" onclick="gRecusar()">🚫 Não realizei</button>'
        +'<button class="btn alt" onclick="gChamado()">💬 Abrir chamado</button>'
      +'</div></div>';
  }else{
    h+='<div class="card"><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">'+GS.map(function(n,i){return '<button class="btn" style="margin:0" onclick="gini('+i+')">'+n+'</button>'}).join("")+'</div></div>';
  }
  var g=ghoje();
  h+='<div class="card"><b>Hoje</b>'+(g.length?g.slice().reverse().map(function(r){return '<div class="row"><span><b>'+GS[r.s]+'</b> <span class="s">'+esc(tnome(r.tipo))+(r.colab?' · '+esc(r.colab):'')+'</span></span><span class="pill '+gcor(r)+'" style="cursor:default">'+hm(r.ini)+'–'+hm(r.fim)+' · '+gmin(r)+'m</span></div>'}).join(""):'<div class="s">Nenhum giro.</div>')+'</div>';
  var gRecs=regs.filter(function(r){var s=r.stAt||"Pendente";return r.st=="nao"&&r.origem=="giro"&&s!="Realizado com sucesso"&&s!="Realizado com observações"});
  if(gRecs.length){
    h+='<div class="card"><b>🚫 Recusas de giro pendentes</b>'+gRecs.map(function(r){var s=r.stAt||"Pendente";return '<div class="row" style="flex-wrap:wrap"><span style="flex:1"><b>'+(GS[r.e]||"Sala")+'</b><div class="s">'+new Date(r.ini).toLocaleDateString("pt-BR")+' '+hm(r.ini)+' · '+esc(r.motivo+(r.outro?": "+r.outro:""))+(r.quem?' · '+esc(r.quem):'')+'</div></span><select style="width:auto;margin:0" onchange="setStat('+r.ini+',this.value)">'+STAT.map(function(x){return '<option'+(x==s?' selected':'')+'>'+x+'</option>'}).join("")+'</select></div>'}).join("")+'</div>';
  }
  return h;
}

/* ============ ROTINA ============ */
function vC(){
  /* CORREÇÃO: o formulário de recusa da rotina nunca abria (a verificação vinha depois de um return) */
  if(open&&rec)return formRec();
  if(open&&open.origem=="giro"){
    return '<div class="card run-card"><b>Recusa de giro em aberto: '+esc(GS[open.e]||"")+'</b><div class="s">Conclua ou volte para retomar o giro.</div><button class="btn" onclick="rec=true;draw()">Continuar recusa</button><button class="btn alt" onclick="cancelRec()">'+T("btn_back")+'</button></div>';
  }
  if(open){
    var tt=TEMPO[tipo]||[0,9999];
    return '<div class="card run-card"><b>'+esc(nomeEsp(open.l,open.e))+'</b><div class="s">'+esc(nomeLoc(open.l))+' · Início '+hm(open.ini)+'</div>'+semaf(open.ini,tt)+'<div class="chips">'+TIPOS.map(function(t){return '<button class="'+(tipo==t[0]?'sel':'')+'" onclick="tipo=\''+t[0]+'\';draw()">'+t[1]+'</button>'}).join("")+'</div>'+(episEsp(open.l,open.e).length?'<div style="border:2px solid '+acCor(aclOf(open.l,open.e))+';border-radius:10px;padding:10px;margin:8px 0"><b>🧤 EPI obrigatório — área '+esc(acNome(aclOf(open.l,open.e)).toLowerCase())+'</b><div class="s">'+esc(episTxt(open.l,open.e))+'</div><label style="display:flex;align-items:center;gap:8px;margin-top:8px;font-weight:600"><input id="epiok" type="checkbox" style="width:22px;height:22px;margin:0"> Estou usando todos os EPIs</label></div>':'')+comoFazerCard(open.l,open.e)+'<input id="exo" placeholder="Observações"><button class="btn" onclick="fim()">Terminei</button><div class="grid-acoes" style="margin-top:8px"><button class="btn alt" onclick="rec=true;draw()">🚫 Não realizei</button><button class="btn alt" onclick="go(\'hOP\')">💬 Abrir chamado</button></div></div>';
  }
  if(loc<0||!LOC[loc])return ctSelCard()+semAreasCard();
  var L=LOC[loc];
  var tot=0,feitas=0;
  L.e.forEach(function(_,i){if(espOff(loc,i))return;tot++;if(reg(loc,i))feitas++});
  var pctF=tot?Math.round(feitas/tot*100):0;
  var h=ctSelCard()+'<div class="card"><b>Olá, '+esc(user)+'</b><select onchange="setLoc(this.value)">'+opts(loc)+'</select><div class="s">'+feitas+'/'+tot+' hoje</div><div class="bar"><i style="width:'+pctF+'%"></i></div>'
    +'<div style="display:grid;grid-template-columns:'+(NFCCFG.on?'1fr 1fr':'1fr')+';gap:8px;margin-top:6px"><button class="btn alt" style="margin:0" onclick="startCam()">📷 QR</button>'+(NFCCFG.on?'<button class="btn nfc" style="margin:0" onclick="startNFC()">NFC</button>':'')+'</div>'
    +'<div style="margin-top:8px">'+campoBusca("rot","Buscar espaço em todas as minhas áreas")+'</div><video id="v" class="hide" playsinline muted></video><div id="nfclive"></div><div class="s" id="cm" style="margin-top:6px"></div></div>';
  if(termo("rot")){
    var achados=[];
    areasPermitidas().forEach(function(l){LOC[l].e.forEach(function(n,i){if(!espOff(l,i)&&bate("rot",n,LOC[l].n))achados.push([l,i])})});
    h+='<div class="card"><b>Resultado da busca ('+achados.length+')</b>'+(achados.length?achados.map(function(a){var l=a[0],i=a[1],r=reg(l,i);return '<div class="row"><span>'+marca(LOC[l].e[i],"rot")+'<div class="s">'+marca(LOC[l].n,"rot")+'</div></span>'+(r?(r.st=="nao"?'<span class="pill no" style="cursor:default">Não feito</span>':'<span class="pill ok" style="cursor:default">'+hm(r.ini)+'–'+hm(r.fim)+'</span>'):'<button class="pill pen" style="border:0;cursor:pointer" onclick="loc='+l+';ini('+l+','+i+')">Iniciar</button>')+'</div>'}).join(""):'<div class="s">Nenhum espaço encontrado.</div>')+'</div>';
    return h;
  }
  h+='<div class="card">'+L.e.map(function(n,i){if(espOff(loc,i))return "";var r=reg(loc,i);return '<div class="row"><span>'+esc(n)+'</span>'+(r?(r.st=="nao"?'<span class="pill no" style="cursor:default">Não feito</span>':'<span class="pill ok" style="cursor:default">'+hm(r.ini)+'–'+hm(r.fim)+'</span>'):'<button class="pill pen" style="border:0;cursor:pointer" onclick="ini('+loc+','+i+')">Iniciar</button>')+'</div>'}).join("")+'</div>';
  h+='<div class="card"><b>Extra</b><select id="ex">'+EXTRAS.map(function(x){return '<option>'+x+'</option>'}).join("")+'</select><button class="btn alt" onclick="extra()">Registrar</button></div>';
  h+=recCard();
  return h;
}
function ini(l,e){if(areasPermitidas().indexOf(+l)<0){alert("Área não liberada para você.");return}if(open){alert("Termine a atividade em andamento.");return}if(gopen){alert("Termine o giro em andamento.");return}open={l:l,e:e,ini:Date.now(),origem:"rotina"};sv("rad_open",open);tipo="CO";draw()}
function fim(){if(!open)return;var obs=(($("#exo")||{}).value||"").trim();var exige=episEsp(open.l,open.e).length>0,epiEl=$("#epiok"),epiOk=!exige||!!(epiEl&&epiEl.checked);
  if(exige&&!epiOk&&!confirm("Você não confirmou o uso dos EPIs. Registrar mesmo assim? Ficará anotado como EPI não confirmado."))return;
  regs.push({l:open.l,e:open.e,ini:open.ini,fim:Date.now(),tipo:tipo,colab:user,st:"ok",obs:obs,origem:open.origem||"rotina",risco:aclOf(open.l,open.e),epi:exige?(epiOk?"sim":"não"):"n/a"});sv("rad_regs",regs);open=null;sv("rad_open",null);draw()}
function extra(){if(loc<0){alert("Selecione uma área.");return}var x=$("#ex").value;if(x.indexOf("Outra")>=0){var d=prompt("Descreva a atividade:","");if(!d)return;x=d.trim()}regs.push({l:loc,x:x,ini:Date.now(),fim:Date.now(),tipo:"",colab:user,st:"ok",origem:"rotina"});sv("rad_regs",regs);alert("Registrado.");draw()}
var TEMPO=ld("rad_tempo",{CO:[5,20],TE:[20,60],PR:[10,30],LA:[20,60]});
if(!TEMPO.LA)TEMPO.LA=[20,60];
function mins(r){return Math.max(0,Math.round((r.fim-r.ini)/60000))}
function stt(r){var t=TEMPO[r.tipo];if(!t)return"";var m=mins(r);return m<t[0]?"Abaixo":m>t[1]?"Acima":"No tempo"}
function tnome(c){return(TIPOS.filter(function(t){return t[0]==c})[0]||["",c||""])[1]}
function setT(k,i,v){TEMPO[k]=TEMPO[k]||[0,0];TEMPO[k][i]=Math.max(0,+v||0);sv("rad_tempo",TEMPO);draw()}
var sfCtx=null;
function fmt(s){return p2(Math.floor(s/60))+":"+p2(s%60)}
function semaf(ini,t){sfCtx={ini:ini,t:t||[0,9999]};return '<div id="sf" style="text-align:center;border-radius:10px;padding:14px;margin:8px 0;color:#fff;background:var(--mut)"><div id="sft" style="font-size:44px;font-weight:700">00:00</div><div id="sfm"></div></div>'}
function tickSf(){var el=document.getElementById("sf");if(!el||!sfCtx)return;var s=Math.floor((Date.now()-sfCtx.ini)/1000),mn=sfCtx.t[0]*60,mx=sfCtx.t[1]*60,st=s<mn?0:s<=mx?1:2;el.style.background=["var(--warn)","var(--ok)","var(--run)"][st];document.getElementById("sft").textContent=fmt(s);document.getElementById("sfm").textContent=st==0?"Faltam "+fmt(mn-s):st==1?"No tempo":"Passou "+fmt(s-mx)}
setInterval(tickSf,1000);

/* ============ CÂMERA / NFC ============ */
/* CORREÇÃO: a câmera abria mas nunca lia o QR (jsQR não era usado); NFC só mostrava texto */
var scanRaf=null,scanCv=null;
function startCam(){
  var cm=document.getElementById("cm");function erro(m){if(cm)cm.innerHTML='<span class="s">'+m+'</span>'}
  if(!window.isSecureContext){erro("A câmera exige HTTPS. Hospede o arquivo (veja QR Codes).");return}
  if(!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia){erro("Navegador sem suporte à câmera.");return}
  stopCam();stopNFC();
  if(cm)cm.textContent="";
  navigator.mediaDevices.getUserMedia({video:{facingMode:{ideal:"environment"}},audio:false}).then(function(s){
    stream=s;var v=document.getElementById("v");if(!v){stopCam();return}
    v.classList.remove("hide");v.setAttribute("playsinline","");v.muted=true;v.srcObject=s;
    var p=v.play();if(p&&p.catch)p.catch(function(){});
    erro("Aponte para o QR do espaço");
    scanRaf=requestAnimationFrame(scanLoop);
  }).catch(function(err){erro("Erro na câmera: "+err.name)});
}
function scanLoop(){
  var v=document.getElementById("v");
  if(!stream)return;
  if(!v){stopCam();return}
  if(v.readyState>=2&&typeof jsQR==="function"&&v.videoWidth){
    var w=v.videoWidth,h=v.videoHeight,esc_=Math.min(1,640/Math.max(w,h));
    w=Math.round(w*esc_);h=Math.round(h*esc_);
    scanCv=scanCv||document.createElement("canvas");scanCv.width=w;scanCv.height=h;
    var ctx=scanCv.getContext("2d",{willReadFrequently:true});ctx.drawImage(v,0,0,w,h);
    var code=jsQR(ctx.getImageData(0,0,w,h).data,w,h,{inversionAttempts:"dontInvert"});
    if(code&&code.data){lido(code.data);return}
  }
  scanRaf=requestAnimationFrame(scanLoop);
}
function stopCam(){
  if(scanRaf){cancelAnimationFrame(scanRaf);scanRaf=null}
  if(stream){stream.getTracks().forEach(function(t){t.stop()});stream=null}
  var v=document.getElementById("v");if(v){v.srcObject=null;v.classList.add("hide")}
}
var nfcAbort=null;
function nfcUI(t,e){var el=document.getElementById("nfclive");if(!el)return;if(!t){el.className="";el.textContent="";return}el.textContent=t;el.className="on"+(e?" err":"")}
function startNFC(){
  if(!NFCCFG.on){nfcUI("NFC desativado nas configurações.",true);return}
  if(!("NDEFReader" in window)){nfcUI("Este aparelho/navegador não suporta NFC pela web (use o Chrome no Android, com HTTPS).",true);return}
  stopCam();stopNFC();
  try{
    var r=new NDEFReader();nfcAbort=new AbortController();
    r.scan({signal:nfcAbort.signal}).then(function(){nfcUI("Aproxime a tag NFC...",false)}).catch(function(e){nfcUI("Erro NFC: "+(e.message||e.name),true)});
    r.onreading=function(ev){
      var txt="";
      try{ev.message.records.forEach(function(rec){if(txt||!rec.data)return;if(["url","absolute-url","text"].indexOf(rec.recordType)>=0)txt=new TextDecoder(rec.encoding||"utf-8").decode(rec.data)})}catch(x){}
      lido(txt||"");
    };
    r.onreadingerror=function(){nfcUI("Não foi possível ler a tag.",true)};
  }catch(e){nfcUI("Erro: "+e.message,true)}
}
function stopNFC(){if(nfcAbort){try{nfcAbort.abort()}catch(e){}nfcAbort=null}nfcUI("")}
function parseQR(c){var m=/[?&]l=(\d+)&e=(\d+)/.exec(String(c||""))||/SEM_URL_(\d+)_(\d+)/.exec(String(c||""));return m?{l:+m[1],e:+m[2]}:null}
function lido(c){
  stopCam();stopNFC();
  var q=parseQR(c);
  if(q&&LOC[q.l]&&!LOC[q.l].del&&LOC[q.l].e[q.e]!=null&&!espOff(q.l,q.e)){
    var ctq=LOC[q.l].ct;
    if(ctq&&ctq!=ctAtualId()&&meusContratos().some(function(c){return c.id==ctq})){CT_ATUAL=ctq;sv("rad_ctAtual",CT_ATUAL)}
    if(areasPermitidas().indexOf(q.l)<0){alert("Esta área não está liberada para você.");draw();return}
    loc=q.l;sv("rad_loc",q.l);
    if(open&&open.l==q.l&&open.e==q.e&&open.origem!="giro")fim();else ini(q.l,q.e);
    return;
  }
  alert("Código não reconhecido para esta instalação.");draw();
}

/* ============ PAINEL ============ */
function vP(){
  var resEst=((podeEstMov()&&!ehColabCargo())?resumoEstoqueCard():"")+destaquesPainel();
  var rh=hoje().filter(function(r){return !r.x&&r.st=="ok"&&regDoMeuEscopo(r)});
  return bannerInterno()+'<div class="card"><b>📊 '+T("menu_painel")+'</b><div class="s">'+new Date().toLocaleDateString("pt-BR")+'</div></div>'
    +'<div class="card"><div class="row"><span>Rotinas realizadas hoje</span><b>'+rh.length+'</b></div><div class="row"><span>Giros hoje</span><b>'+ghoje().length+'</b></div><div class="row"><span>Planos de ação abertos</span><b>'+pa.filter(function(p){return p.st!="Concluído"}).length+'</b></div><div class="row"><span>Planos vencidos</span><b style="color:var(--warn)">'+planosVencidos().length+'</b></div><div class="row"><span>🛠️ Chamados OP abertos</span><b>'+chs.filter(function(c){return c.canal=="OP"&&c.st!="Resolvido"}).length+'</b></div><div class="row"><span>📄 Chamados AD abertos</span><b>'+chs.filter(function(c){return c.canal=="AD"&&c.st!="Resolvido"}).length+'</b></div><div class="row"><span>💬 Chamados Cliente abertos</span><b>'+chs.filter(function(c){return c.canal=="CLI"&&c.st!="Resolvido"}).length+'</b></div></div>'+resEst;
}

/* ============ INDICADORES ============ */

/* ============ SETORES ============ */
function vSetores(){
  if(!canEdit()&&!temAcao("gerenciar_usuarios"))return '<div class="card s">Sem permissão.</div>';
  var auto=ld("rad_notifauto",{zap:true,mail:false});
  var h='<div class="card"><b>📣 '+T("menu_set")+'</b><div class="row"><span>WhatsApp</span><button class="pill '+(auto.zap?'ok':'no')+'" style="border:0;cursor:pointer" onclick="setNotifAuto(\'zap\','+(!auto.zap)+')">'+(auto.zap?'✅':'🚫')+'</button></div><div class="row"><span>E-mail</span><button class="pill '+(auto.mail?'ok':'no')+'" style="border:0;cursor:pointer" onclick="setNotifAuto(\'mail\','+(!auto.mail)+')">'+(auto.mail?'✅':'🚫')+'</button></div></div>';
  h+='<div class="card">'+SET.map(function(s){return '<div style="border-top:1px solid var(--line);padding:10px 0"><b>'+esc(s.nome)+'</b><input value="'+esc(s.zap||"")+'" placeholder="WhatsApp (com DDD)" inputmode="tel" onchange="setSetor(\''+s.id+'\',\'zap\',this.value)" style="margin-top:6px"><input value="'+esc(s.mail||"")+'" placeholder="E-mail" type="email" onchange="setSetor(\''+s.id+'\',\'mail\',this.value)"></div>'}).join("")+'</div>';
  ["OP","AD","CLI"].forEach(function(k){
    var c=CANAIS[k];
    if(!c)return;
    h+='<div class="card"><b>'+c.ico+' '+c.nome+'</b>'+canalTipos(k).map(function(x,i){return '<div class="row"><span>'+esc(x.t)+'</span><select style="width:auto;margin:0" onchange="setDestCanal(\''+k+'\','+i+',this.value)">'+SET.map(function(s){return '<option value="'+s.id+'"'+(s.id==x.d?' selected':'')+'>'+esc(s.nome)+'</option>'}).join("")+'</select></div>'}).join("")+'</div>';
  });
  return h;
}

/* ============ CONTATOS (visualização geral) ============ */
function vContatos(){
  var cats=categoriaContatos();
  var h='<div class="card"><b>📞 '+T("menu_cont")+'</b><div class="s">Telefones, ramais e links essenciais.</div></div>';
  if(!CONTATOS.filter(contatoVisivel).length)return h+'<div class="card s">Nenhum contato disponível.</div>';
  Object.keys(cats).forEach(function(cat){
    h+='<div class="card"><b>'+esc(cat)+'</b>';
    cats[cat].forEach(function(c){
      h+='<div class="row" style="flex-wrap:wrap;align-items:flex-start"><span style="flex:1;min-width:180px">'
        +'<b>'+esc(c.ico||"📌")+' '+esc(c.nome)+'</b>'
        +'<div class="s" style="word-break:break-all">'+esc(c.valor)+'</div></span>'
        +'<span style="display:flex;gap:4px;flex-wrap:wrap">'+acaoContato(c)+'</span></div>';
    });
    h+='</div>';
  });
  if(temAcao("gerenciar_usuarios")&&acessosDoUsuario().indexOf("k")>=0){
    h+='<div class="card"><b>⚙️ Editar contatos</b><button class="btn alt" onclick="cfg=\'Contatos\';go(\'k\')">Ir para Configurações</button></div>';
  }
  return h;
}

/* ============ CONFIG ============ */
var cfg="Áreas";
var PERM=ld("rad_perm",{sup:false});
function canEdit(){
  var cg=cargoAtual();
  if(!cg)return false;
  if(cg.fixo||cg.id=="admin")return true;
  if(cg.id=="sup"&&PERM.sup)return true;
  return (cg.acoes||[]).indexOf("editar_colaboradores")>=0||(cg.acoes||[]).indexOf("editar_areas")>=0;
}
function togPerm(){if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}PERM.sup=!PERM.sup;sv("rad_perm",PERM);draw()}
function stepNav(S){var i=S.indexOf(cfg),ir=function(s){return"cfg=\'"+s+"\';draw();window.scrollTo(0,0)"};return '<div class="card"><div>'+S.map(function(s){return '<button style="padding:6px 10px;margin:3px 3px 0 0;border-radius:99px;font-size:12px;border:1px solid '+(s==cfg?'var(--brand)':'var(--line)')+';background:'+(s==cfg?'var(--brand)':'var(--bg)')+';color:'+(s==cfg?'var(--on)':'var(--ink)')+';cursor:pointer" onclick="'+ir(s)+'">'+s+'</button>'}).join("")+'</div><div style="display:flex;gap:8px;margin-top:10px">'+(i>0?'<button class="btn alt" style="margin:0" onclick="'+ir(S[i-1])+'">← '+T("btn_back")+'</button>':'')+(i<S.length-1?'<button class="btn" style="margin:0" onclick="'+ir(S[i+1])+'">'+T("btn_next")+' →</button>':'')+'</div></div>'}
function vK(){
  var S=[];
  if(canEdit())S.push("Equipe","Escala","Tempos","EPI por risco");
  if(temAcao("gerenciar_cargos"))S.push("Checklists");
  if(temAcao("gerenciar_usuarios"))S.push("Contatos");
  if(temAcao("gerenciar_cargos"))S.push("Permissões","Cargos");
  if(temAcao("gerenciar_usuarios"))S.push("Acessos");
  if(temAcao("ver_auditoria"))S.push("Auditoria");
  if(canEdit())S.push("Backup");
  if(canEdit()||temAcao("gerenciar_usuarios"))S.unshift("Sistema","Aparência");
  if(acessosDoUsuario().indexOf("fc")>=0)S.push("Fábrica");
  if(!S.length)return '<div class="card s">Sem permissão para configurações.</div>';
  var atalhoAr=acessosDoUsuario().indexOf("ar")>=0?'<div class="card noprint"><b>🗺️ Áreas e atividades</b><div class="s">Áreas, espaços e atividades são editados por contrato.</div><button class="btn alt" onclick="go(\'ar\')">Editar áreas e atividades</button></div>':'';
  if(S.indexOf(cfg)<0)cfg=S[0];
  var body="";
  if(cfg=="Áreas")body=spCard();
  else if(cfg=="Atividades")body=pgCard();
  else if(cfg=="Equipe")body=assignEd();
  else if(cfg=="Escala")body=vEscala();
  else if(cfg=="Tempos")body=temposCard()+prCfg();
  else if(cfg=="EPI por risco")body=epiRiscoCard();
  else if(cfg=="Checklists")body=vChecklistsConfig();
  else if(cfg=="Contatos")body=vContatosConfig();
  else if(cfg=="Permissões")body=permCard();
  else if(cfg=="Cargos")body=vCargos();
  else if(cfg=="Sistema")body=sistemaCard();
  else if(cfg=="Aparência")body=aparenciaCard();
  else if(cfg=="Fábrica")body='<div class="card"><button class="btn" onclick="go(\'fc\')">🏭 Abrir Fábrica de Contratos</button></div>';
  else if(cfg=="Acessos")body='<div class="card"><b>🔑 Acessos</b><div class="s">Os acessos agora ficam no menu <b>Administração → Acessos e clientes</b>.</div><button class="btn" onclick="go(\'ac\')">Abrir Acessos e clientes</button></div>';
  else if(cfg=="Auditoria")body=vAuditoria();
  else body=backupCard();
  var hub='<div class="card"><b>⚙️ Configuração do Sistema</b>'+CFG_GRUPOS.map(function(g){var it=g[1].filter(function(x){return S.indexOf(x[0])>=0});if(!it.length)return "";
    return '<div class="s" style="font-weight:700;margin:10px 0 4px">'+g[0]+'</div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:6px">'+it.map(function(x){var on=cfg==x[0];return '<button style="text-align:left;padding:10px;border-radius:10px;border:1px solid '+(on?'var(--brand)':'var(--line)')+';background:'+(on?'var(--brand)':'var(--bg)')+';color:'+(on?'var(--on)':'var(--ink)')+';font:inherit;font-size:13px;font-weight:600;cursor:pointer" onclick="'+(x[0]=="Fábrica"?"go(\'fc\')":"cfg=\'"+x[0]+"\';draw()")+'">'+x[1]+' '+esc(x[2])+'</button>'}).join("")+'</div>'}).join("")+'</div>';
  return hub+atalhoAr+body;
}
function permCard(){return '<div class="card"><b>Permissões</b><div class="row"><span>Supervisor edita (atalho global)</span><button class="pill '+(PERM.sup?'ok':'pen')+'" style="border:0;cursor:pointer" onclick="togPerm()">'+(PERM.sup?'Liberado':'Bloqueado')+'</button></div><div class="s">Para controle fino, use a aba Cargos.</div></div>'}
function temposCard(){return '<div class="card"><b>⏱️ Tempos (min)</b>'+TIPOS.map(function(t){var a=TEMPO[t[0]]||[0,0];function inp(v,f){return '<input type="number" style="width:70px;display:inline;margin:0" value="'+v+'" onchange="'+f+'">'}return '<div class="row" style="flex-wrap:wrap"><b style="width:100%">'+t[1]+'</b><span>'+inp(a[0],"setT(\'"+t[0]+"\',0,this.value)")+' a '+inp(a[1],"setT(\'"+t[0]+"\',1,this.value)")+'</span></div>'}).join("")+'<div class="row"><span>Giro — Concorrente</span><span><input type="number" style="width:70px;display:inline;margin:0" value="'+GT.CO[0]+'" onchange="setG(\'CO\',0,this.value)"> a <input type="number" style="width:70px;display:inline;margin:0" value="'+GT.CO[1]+'" onchange="setG(\'CO\',1,this.value)"></span></div><div class="row"><span>NFC</span><button class="pill '+(NFCCFG.on?'ok':'no')+'" style="border:0;cursor:pointer" onclick="setNFC('+(!NFCCFG.on)+')">'+(NFCCFG.on?'✅':'🚫')+'</button></div><div class="row"><span>Sessão (min, 0 = nunca)</span><input type="number" style="width:80px;margin:0" value="'+SESSAO_MIN+'" onchange="setSessaoMin(this.value)"></div></div>'}
function spCard(){var L=LOC[loc];return '<div class="card"><b>🏢 Áreas</b><select onchange="setLoc(this.value)">'+optsAll(loc)+'</select><input value="'+esc(L.n)+'" onchange="renL(this.value)"><div class="row"><span>Ativa?</span><button class="pill '+(L.off?'no':'ok')+'" style="border:0;cursor:pointer" onclick="togL()">'+(L.off?'🚫':'✅')+'</button></div><button class="btn red" style="margin-top:8px" onclick="delL()">Excluir área</button><div class="s" style="margin-top:14px"><b>Espaços</b> (nome · classificação · ativo · excluir)</div>'+L.e.map(function(n,i){var off=espOff(loc,i);return '<div class="esp-linha'+(off?' off':'')+'"><input value="'+esc(n)+'" onchange="renE('+i+',this.value)">'+aclSel(loc,i)+'<button class="pill '+(off?'pen':'ok')+' mini-btn" onclick="togE('+i+')">'+(off?'▶':'⏸')+'</button><button class="pill no mini-btn" onclick="delE('+i+')">✕</button></div>'}).join("")
  +'<div class="s" style="margin-top:8px">Classificar todos os espaços como:</div><select onchange="setAclAll(this.value)"><option value="">—</option>'+["C","S","N"].map(function(k){return '<option value="'+k+'">'+AL[k]+'</option>'}).join("")+'</select>'
  +'<div style="display:flex;gap:6px;margin-top:10px"><input id="ne" placeholder="Novo espaço" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px;cursor:pointer" onclick="addE()">Add</button></div>'
  +'<div style="display:flex;gap:6px;margin-top:8px"><input id="nl" placeholder="Nova área" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px;cursor:pointer" onclick="addL()">Add</button></div></div>'}

/* ============ CHECKLISTS CONFIG ============ */
/* CORREÇÃO: o HTML desta tela era montado sobre a própria variável externa e saía duplicado */
function vChecklistsConfig(){
  if(!temAcao("gerenciar_cargos"))return '<div class="card s">Sem permissão.</div>';
  var h='<div class="card"><b>📋 Checklists configuráveis</b><div class="s">Crie checklists personalizados para suas áreas.</div></div>';
  h+=CHECKLISTS.map(function(c){
    var itens=chklItens(c.id);
    var b='<div class="card"><div class="row" style="border:0;padding:0"><span style="flex:1"><b>'+esc(c.ico||"📋")+' '+esc(c.nome)+'</b><div class="s">'+esc(c.cat)+' · '+itens.length+' item(ns)'+(c.fixed?' · <b>padrão</b>':'')+'</div><div class="s">'+((c.tipos||["*"]).indexOf("*")>=0?'Todos os tipos':(c.tipos||[]).map(tipoCtNome).join(" · "))+'</div></span><button class="pill '+(c.ativo?'ok':'pen')+'" style="border:0;cursor:pointer;margin-right:4px" onclick="togChklAtivo(\''+c.id+'\')">'+(c.ativo?'Ativo':'Inativo')+'</button>'+(c.fixed?'':'<button class="pill no" style="border:0;cursor:pointer" onclick="delChkl(\''+c.id+'\')">🗑️</button>')+'</div>';
    if(c.fixed){
      b+='<div class="s" style="margin-top:8px;font-style:italic">Checklist padrão do sistema (não editável).</div>';
    }else{
      b+='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px">';
      b+=itens.map(function(it,i){
        var pts=it.p&&it.p.length?it.p.join(", "):"";
        return '<div class="row" style="padding:6px 0;flex-wrap:wrap"><span style="flex:1"><b>'+esc(it.n)+'</b> <span class="pill" style="background:'+LZK[it.c]+';cursor:default">'+LZN[it.c]+'</span>'+(pts?'<div class="s">'+esc(pts)+'</div>':'')+'</span><button class="pill no" style="border:0;cursor:pointer" onclick="delChklItem(\''+c.id+'\','+i+')">✕</button></div>';
      }).join("");
      b+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:4px">➕ Novo item</div>'
        +'<input id="ci_nome_'+c.id+'" placeholder="Nome do item (ex: Piso)">'
        +'<select id="ci_crit_'+c.id+'"><option value="C">Crítico</option><option value="S">Semicrítico</option><option value="N">Não crítico</option></select>'
        +'<input id="ci_pts_'+c.id+'" placeholder="Pontos (separados por ; ) ex: Cantos; Bordas; Centro">'
        +'<button class="btn alt" onclick="addChklItem(\''+c.id+'\')">Adicionar item</button></div>';
    }
    return b+'</div>';
  }).join("");
  h+='<div class="card"><b>➕ Novo checklist</b>'
    +'<input id="ck_nome" placeholder="Nome (ex: Checklist de Banheiros)">'
    +'<select id="ck_cat"><option>Personalizado</option><option>Luz Negra</option><option>REG-016</option><option>Auditoria</option></select>'
    +'<input id="ck_ico" placeholder="Ícone (emoji)" value="📋">'
    +'<div class="s">Tipos de serviço em que aparece (nenhum marcado = todos):</div><div>'+TIPOS_CT.map(function(t){return '<label class="s" style="display:inline-block;margin:3px 8px 3px 0"><input type="checkbox" class="ck_tp" value="'+t.id+'" style="width:auto;margin-right:4px">'+t.ico+' '+esc(t.nome)+'</label>'}).join("")+'</div>'
    +'<button class="btn" onclick="addChkl()">Criar checklist</button></div>';
  return h;
}

/* ============ CONTATOS CONFIG ============ */
function vContatosConfig(){
  if(!temAcao("gerenciar_usuarios"))return '<div class="card s">Sem permissão.</div>';
  var cats=["Telefones","Emergência","Links úteis","Documentos","Outros"];
  var h='<div class="card"><b>📞 Gerenciar contatos e links</b><div class="s">Configure telefones, ramais e links que aparecem no menu Contatos.</div></div>';
  h+=CONTATOS.map(function(c){
    var v=c.vis||"interno",modo=Array.isArray(v)?"cargos":v;
    return '<div class="card"><div class="row" style="flex-wrap:wrap"><span style="flex:1;min-width:180px"><b>'+esc(c.ico||"📌")+' '+esc(c.nome)+'</b><div class="s">'+esc(c.tipo)+' · '+esc(c.cat)+'</div><div class="s" style="word-break:break-all">'+esc(c.valor)+'</div><div class="s">👁️ '+esc(visTxt(c))+(c.ct?' · só contrato '+esc((ctById(c.ct)||{}).codigo||"?"):'')+'</div></span><span style="display:flex;gap:4px;flex-wrap:wrap"><button class="pill pen" style="border:0" onclick="editContato('+c.id+')">✏️</button><button class="pill no" style="border:0" onclick="delContato('+c.id+')">✕</button></span></div>'
      +'<div class="s" style="margin-top:6px">Quem pode ver</div><select onchange="setVisContato('+c.id+',this.value)"><option value="interno"'+(modo=="interno"?' selected':'')+'>Equipe interna (sem cliente)</option><option value="todos"'+(modo=="todos"?' selected':'')+'>Todos, inclusive cliente</option><option value="cargos"'+(modo=="cargos"?' selected':'')+'>Só cargos escolhidos</option></select>'
      +(modo=="cargos"?'<div>'+CARGOS.filter(function(g){return g.id!="admin"}).map(function(g){return chip(v.indexOf(g.id)>=0,"togVisCargo("+c.id+",\'"+g.id+"\')",g.nome)}).join("")+'</div>':'')
      +'<div class="s" style="margin-top:6px">Contrato</div><select onchange="setCtContato('+c.id+',this.value)"><option value="">Todos os contratos</option>'+CONTRATOS.map(function(x){return '<option value="'+x.id+'"'+(c.ct==x.id?' selected':'')+'>'+esc(x.codigo+" · "+x.nome)+'</option>'}).join("")+'</select></div>';
  }).join("");
  h+='<div class="card"><b>➕ Novo contato</b>'
    +'<input id="ctt_nome" placeholder="Nome (ex: Ramal Manutenção)">'
    +'<select id="ctt_tipo"><option value="tel">📞 Telefone / Ramal</option><option value="link">🔗 Link / Site</option><option value="mail">✉️ E-mail</option></select>'
    +'<input id="ctt_valor" placeholder="Valor (ex: 2201 ou https://...)">'
    +'<select id="ctt_cat">'+cats.map(function(x){return '<option>'+x+'</option>'}).join("")+'</select>'
    +'<input id="ctt_ico" placeholder="Ícone (emoji)" value="📌">'
    +'<button class="btn" onclick="addContato()">Adicionar</button></div>';
  return h;
}

/* ============ AUDITORIA ============ */
function vAuditoria(){
  if(!temAcao("ver_auditoria"))return '<div class="card s">Sem permissão.</div>';
  var l=LOG.slice().reverse().slice(0,200);
  var h='<div class="card"><b>🔍 Log de auditoria</b><div class="s">Últimas 200 ações (máx. 500 guardadas).</div>'
    +'<button class="btn alt" onclick="exportarLog()">⬇️ Exportar CSV</button></div>';
  if(!l.length)return h+'<div class="card s">Nenhuma ação.</div>';
  h+='<div class="card">'+l.map(function(x){
    return '<div class="row" style="padding:6px 0"><span style="flex:1"><b>'+esc(x.user)+'</b> <span class="s">'+esc(x.acao)+(x.alvo?' → '+esc(x.alvo):'')+'</span>'+(x.extra?'<div class="s">'+esc(x.extra)+'</div>':'')+'</span><span class="s">'+new Date(x.t).toLocaleString("pt-BR")+'</span></div>';
  }).join("")+'</div>';
  return h;
}
function exportarLog(){
  var linhas=[["Data/hora","Usuário","Ação","Alvo","Detalhe"]].concat(LOG.map(function(x){return [new Date(x.t).toLocaleString("pt-BR"),x.user,x.acao,x.alvo,x.extra]}));
  baixar("Auditoria_"+dia(Date.now())+".csv",linhas.map(function(l){return l.map(function(v){return '"'+String(v==null?"":v).replace(/"/g,'""')+'"'}).join(";")}).join("\n"),"text/csv");
}

function assignEd(){return '<div class="card"><b>Colaboradores</b><div class="s">Toque em Ver para abrir a ficha na tela Equipe.</div>'+(COL.length?COL.map(function(x){return '<div class="row"><span style="flex:1"><b>'+esc(x.nome)+'</b><div class="s">'+esc(x.funcao||"")+' · '+esc(x.escala||"")+' · '+(x.vinculos||[]).length+' contrato(s)</div></span><button class="pill pen" style="border:0;cursor:pointer" onclick="tsel='+x.id+';sv(\'rad_tsel\',tsel);go(\'eq\')">Ver</button></div>'}).join(""):'<div class="s">Nenhum.</div>')+'</div>'}
function renL(v){LOC[loc].n=v.trim()||LOC[loc].n;svLoc()}
function renE(i,v){LOC[loc].e[i]=v.trim()||LOC[loc].e[i];svLoc()}
function addE(){var v=$("#ne").value.trim();if(!v)return;LOC[loc].e.push(v);svLoc();draw()}
function addL(){var v=$("#nl").value.trim();if(!v)return;LOC.push({n:v,e:[]});loc=LOC.length-1;sv("rad_loc",loc);svLoc();draw()}
function svLoc(){sv("rad_loc2",LOC);autoClassificar()}
function hideCfg(){if(canEdit())return;$all("#app .card").forEach(function(c){var b=c.querySelector("b");if(b&&/^(Tempo|Prazo|Classificação)/.test(b.textContent))c.style.display="none"})}
["addE","addL","togE","delE","delL","togL","renL","renE","setT","setG","setPrz","setPrm","setAcl","setAclAll","addEpi","setEsc","tep"].forEach(function(n){var f=window[n];if(!f)return;window[n]=function(){if(!canEdit()){alert("Sem permissão.");draw();return}return f.apply(this,arguments)}});

/* ============ BACKUP ============ */
var BACKUP_KEYS=["rad_users","rad_col","rad_tsel","rad_cid","rad_loc","rad_loc2","rad_chs","rad_regs","rad_open","rad_pa","rad_giro","rad_gopen","rad_gtempo","rad_tempo","rad_trn","rad_set","rad_perm","rad_contrato","rad_epl","rad_chtcanal","rad_notifauto","rad_notif","rad_pinlock","rad_buscaCol","rad_contratos","rad_ptis","rad_ctSel","rad_planoCt","rad_idioma_ui","rad_cargos","rad_log","rad_checklists","rad_contatos","rad_avals","rad_acl","rad_est","rad_req","rad_prz3","rad_prm","rad_sessmin","rad_nfccfg","rad_base","rad_pg","rad_mig_cli","rad_mig_ar","rad_ctAtual","rad_arCt","rad_acla","rad_epirisco","rad_agenda","rad_mig_ag","rad_indman","rad_mig_ind","rad_receb","rad_ocorr","rad_mig_es","rad_portal_meus","rad_sys","rad_fcCt","rad_tema","rad_elog","rad_docs"];
function exportarBackup(){if(!canEdit()){alert("Sem permissão.");return}var obj={_meta:{gerado:new Date().toISOString(),app:"RAD Digital"}};BACKUP_KEYS.forEach(function(k){var v=ld(k,null);if(v!==null)obj[k]=v});baixar("RAD_backup_"+dia(Date.now())+".json",JSON.stringify(obj,null,2),"application/json");logAcao("Exportou backup");alert("Backup exportado.")}
function importarBackup(inp){if(!temAcao("gerenciar_usuarios")){alert("Somente quem gerencia usuários pode restaurar backup.");inp.value="";return}var f=inp.files[0];if(!f)return;if(!confirm("Substitui todos os dados deste aparelho. Continuar?")){inp.value="";return}var r=new FileReader();r.onload=function(){try{var obj=JSON.parse(r.result);if(!obj||typeof obj!=="object"||!obj.rad_users){alert("Arquivo não parece um backup do RAD Digital.");return}var count=0;BACKUP_KEYS.forEach(function(k){if(obj[k]!==undefined){localStorage.setItem(k,JSON.stringify(obj[k]));count++}});localStorage.removeItem("rad_session");alert(count+" conjuntos restaurados. Recarregando.");location.reload()}catch(e){alert("Erro: "+e.message)}};r.readAsText(f);inp.value=""}
function backupCard(){var meta=[["Usuários",USERS.length],["Cargos",CARGOS.length],["Colaboradores",COL.length],["Registros",regs.length],["Chamados",chs.length],["Avaliações",AVALS.length],["Planos de ação",pa.length],["Checklists",CHECKLISTS.length],["Contatos",CONTATOS.length]];var uso=0;try{for(var k in localStorage)if(localStorage.hasOwnProperty(k)&&k.indexOf("rad_")==0)uso+=(localStorage.getItem(k)||"").length}catch(e){}return '<div class="card"><b>💾 Backup</b><div class="s">Os dados ficam só neste navegador. Exporte com frequência.</div>'+meta.map(function(m){return '<div class="row"><span>'+m[0]+'</span><b>'+m[1]+'</b></div>'}).join("")+'<div class="row"><span>Espaço usado</span><b>'+(Math.round(uso*2/1024/1024*10)/10).toString().replace(".",",")+' MB</b></div><button class="btn" onclick="exportarBackup()" style="margin-top:10px">⬇️ Exportar</button><div class="s" style="margin-top:10px">Restaurar backup:</div><input type="file" accept=".json" onchange="importarBackup(this)" style="padding:10px"></div>'}

/* ============ CADASTRO / QR ============ */
var base=ld("rad_base","");
function urlLocal(){return location.protocol==="file:"||location.hostname===""||location.hostname==="localhost"}
function setBase(v){base=(v||"").trim().split("?")[0];sv("rad_base",base);draw()}
function autoBase(){setBase(location.href.split("?")[0].split("#")[0])}
function qrs(){if(typeof QRCode==="undefined")return;$all(".q").forEach(function(e){e.innerHTML="";try{new QRCode(e,{text:e.dataset.c,width:120,height:120})}catch(x){}})}
function vQ(){
  if(loc<0||!LOC[loc])return ctSelCard()+semAreasCard();
  var L=LOC[loc];
  var local=urlLocal();
  var h='<div class="card noprint"><b>🏷️ QR Codes dos espaços</b><div class="s">Cada QR aponta para um espaço específico.</div><button class="btn alt" onclick="window.print()">🖨️ '+T("btn_print")+'</button></div>';
  if(local){
    h+='<div class="card noprint" style="border-left:4px solid var(--warn);background:rgba(198,40,40,.06)">'
      +'<b style="color:var(--warn)">⚠️ URL local detectada</b>'
      +'<div class="s">Você está usando <code>file://</code>. Câmera e QRs só funcionam com o arquivo hospedado em HTTPS.</div>'
      +'<details style="margin-top:10px"><summary style="cursor:pointer;font-weight:700;color:var(--brand)">▶ Como hospedar gratuitamente (Netlify Drop)</summary>'
      +'<ol style="line-height:1.7;padding-left:22px;margin:8px 0">'
      +'<li>Acesse <b>app.netlify.com/drop</b></li>'
      +'<li>Arraste a pasta contendo <code>rad.html</code></li>'
      +'<li>Copie o link gerado (ex.: <code>https://seunome.netlify.app/rad.html</code>)</li>'
      +'<li>Cole o link no campo abaixo e salve</li>'
      +'</ol>'
      +'<div class="s">Atenção: os dados ficam no navegador de cada aparelho — não são compartilhados entre celulares.</div>'
      +'</details></div>';
  }else{
    h+='<div class="card noprint" style="border-left:4px solid var(--ok);background:rgba(46,139,87,.05)"><b style="color:var(--ok)">✅ URL pública detectada — QRs prontos para uso</b></div>';
  }
  h+='<div class="card noprint"><b>URL base</b>'
    +'<input id="base_input" value="'+esc(base)+'" placeholder="https://seu-site.netlify.app/rad.html">'
    +'<div class="grid-acoes" style="margin-top:8px">'
      +'<button class="btn" onclick="setBase(document.getElementById(\'base_input\').value)">💾 Salvar URL</button>'
      +'<button class="btn alt" onclick="autoBase()">🔄 Detectar automaticamente</button>'
    +'</div>'
    +'<div class="s" style="margin-top:8px">URL atual usada nos QRs:</div>'
    +'<div class="s" style="word-break:break-all;background:var(--bg);padding:8px;border-radius:6px;margin-top:4px">'+esc(base||"(vazio — os QRs só funcionarão pelo leitor interno do app)")+'</div></div>';
  h+=ctSelCard()+'<div class="card noprint"><select onchange="setLoc(this.value)">'+opts(loc)+'</select></div>';
  h+='<div class="card"><b>'+esc(L.n)+'</b><div class="qrs" style="margin-top:8px">'+L.e.map(function(n,i){if(espOff(loc,i))return "";var url=base?base+'?l='+loc+'&e='+i:'SEM_URL_'+loc+'_'+i;return '<div><span class="q" data-c="'+esc(url)+'"></span><br>'+esc(n)+'</div>'}).join("")+'</div></div>';
  return h;
}


/* ===================================================================
   CONTRATOS
   =================================================================== */
var CONTRATOS=ld("rad_contratos",[]);
function svCt(){sv("rad_contratos",CONTRATOS)}
function ctById(id){return CONTRATOS.filter(function(c){return c.id==id})[0]||null}
function ctAtivo(){return CONTRATOS.filter(function(c){return c.status=="Ativo"})}
function novoContrato(o){
  return{
    id:Date.now(),
    codigo:o.codigo||"",
    nome:o.nome||"",
    cnpj:o.cnpj||"",
    endereco:o.endereco||"",
    responsavel:o.responsavel||"",
    telefone:o.telefone||"",
    email:o.email||"",
    inicio:o.inicio||"",
    fim:o.fim||"",
    status:o.status||"Ativo",
    tipo:o.tipo||"out",
    tipos:o.tipos&&o.tipos.length?o.tipos:[o.tipo||"out"],
    criado:Date.now()
  };
}
function addContrato(){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var nome=$("#ct_nome").value.trim();
  var codigo=$("#ct_codigo").value.trim();
  if(!nome||!codigo){alert("Informe código e nome.");return}
  if(CONTRATOS.filter(function(c){return c.codigo===codigo})[0]){alert("Código já existe.");return}
  var ini_=$("#ct_ini").value,fim_=$("#ct_fim").value;
  if(ini_&&fim_&&fim_<ini_){alert("A data de fim é anterior à de início.");return}
  var c=novoContrato({
    codigo:codigo,nome:nome,
    cnpj:$("#ct_cnpj").value.trim(),
    endereco:$("#ct_end").value.trim(),
    responsavel:$("#ct_resp").value.trim(),
    telefone:$("#ct_fone").value.trim(),
    email:$("#ct_email").value.trim(),
    inicio:ini_,
    fim:fim_,
    tipo:($all(".ct_tp:checked")[0]||{value:"out"}).value,
    tipos:$all(".ct_tp:checked").map(function(x){return x.value})
  });
  CONTRATOS.push(c);svCt();prepararAreas();
  logAcao("Criou contrato",codigo+" · "+nome);
  alert("Contrato cadastrado.");draw();
}
function delContrato(id){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var c=ctById(id);if(!c)return;
  var usos=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==id})}).length;
  if(usos){alert("Há "+usos+" colaborador(es) vinculados. Desvincule antes de excluir.");return}
  if(PTIS.some(function(p){return p.contratoId==id})){alert("Há PTIs neste contrato. Exclua-os antes.");return}
  if(!confirm("Excluir contrato "+c.nome+"?"))return;
  CONTRATOS=CONTRATOS.filter(function(x){return x.id!=id});
  LOC.forEach(function(L){if(L.ct==id)L.ct=null});svLoc();
  USERS.forEach(function(u){if(u.contratosIds)u.contratosIds=u.contratosIds.filter(function(x){return x!=id})});saveUsers();
  svCt();logAcao("Excluiu contrato",c.codigo);draw();
}
function setCtStatus(id,v){
  if(!temAcao("gerenciar_contratos")&&!canEdit())return;
  var c=ctById(id);if(!c)return;
  c.status=v;svCt();logAcao("Mudou status do contrato",c.codigo,v);draw();
}
function togCtLocal(id,locIdx,on){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var c=ctById(id);if(!c)return;
  c.locais=c.locais||[];
  var p=c.locais.indexOf(locIdx);
  if(on&&p<0)c.locais.push(locIdx);
  if(!on&&p>=0)c.locais.splice(p,1);
  svCt();var y=window.scrollY;draw();window.scrollTo(0,y);
}

/* ===================================================================
   CONTRATOS A PARTIR DOS DEPARTAMENTOS DO EFETIVO
   Cada Departamento do eSocial = um contrato. O colaborador é vinculado
   automaticamente ao contrato do seu departamento.
   =================================================================== */
function chaveDepto(c){return (c.deptoCod||"").trim()?"cod:"+c.deptoCod.trim():((c.departamento||"").trim()?"nome:"+c.departamento.trim().toUpperCase():"")}
function nomeDepto(c){return (c.departamento||"").trim()||(c.deptoCod?"Departamento "+c.deptoCod:"Sem departamento")}
function ctDoDepto(c){
  var k=chaveDepto(c);if(!k)return null;
  return CONTRATOS.filter(function(x){
    if(x.deptoChave)return x.deptoChave==k;
    return (c.deptoCod&&x.codigo==c.deptoCod.trim())||(x.nome||"").trim().toUpperCase()==(c.departamento||"").trim().toUpperCase();
  })[0]||null;
}
function syncContratosDepto(){
  var novos=0,vinc=0,base=Date.now();
  COL.forEach(function(c){
    var k=chaveDepto(c);
    c.vinculos=c.vinculos||[];
    if(!k){
      var antes=c.vinculos.length;
      c.vinculos=c.vinculos.filter(function(v){return !v.auto});
      vinc+=antes-c.vinculos.length;return;
    }
    var ct=ctDoDepto(c);
    if(!ct){
      ct=novoContrato({codigo:(c.deptoCod||"").trim()||("DEP-"+(++novos)),nome:nomeDepto(c)});
      ct.id=base+novos+Math.floor(Math.random()*1000);
      ct.origem="efetivo";ct.deptoChave=k;
      CONTRATOS.push(ct);novos++;
    }else{
      if(!ct.deptoChave)ct.deptoChave=k;
      if(ct.origem=="efetivo"&&c.departamento&&ct.nome!=c.departamento.trim())ct.nome=c.departamento.trim();
    }
    /* remove vínculo automático de departamento antigo (mudança de lotação) */
    var antes2=c.vinculos.length;
    c.vinculos=c.vinculos.filter(function(v){return !v.auto||v.contratoId==ct.id});
    vinc+=antes2-c.vinculos.length;
    var v=c.vinculos.filter(function(x){return x.contratoId==ct.id})[0];
    if(!v){c.vinculos.push({contratoId:ct.id,supervisorId:null,funcao:c.funcao,escala:c.escala,carga:"44h",areas:[],auto:true,criado:Date.now()});vinc++}
  });
  svCt();svCol();
  return {novos:novos,vinc:vinc};
}

/* ===================================================================
   VÍNCULOS COLABORADOR ↔ CONTRATO
   =================================================================== */
function vinculosDe(colabId){
  var c=cget(colabId);if(!c)return[];
  c.vinculos=c.vinculos||[];
  return c.vinculos;
}
function addVinculo(colabId,contratoId,supervisorId,funcao,escala,carga,silencioso){
  var c=cget(colabId);if(!c)return false;
  c.vinculos=c.vinculos||[];
  if(c.vinculos.some(function(v){return v.contratoId==contratoId})){
    if(!silencioso)alert("Colaborador já vinculado a este contrato.");return false;
  }
  c.vinculos.push({
    contratoId:+contratoId,
    supervisorId:supervisorId?+supervisorId:null,
    funcao:funcao||c.funcao,
    escala:escala||c.escala,
    carga:carga||"44h",
    areas:[],
    criado:Date.now()
  });
  svCol();return true;
}
function delVinculo(colabId,contratoId){
  var c=cget(colabId);if(!c)return;
  var vv=(c.vinculos||[]).filter(function(v){return v.contratoId==contratoId})[0];
  if(vv&&vv.auto){alert("Este vínculo vem do departamento do colaborador no efetivo. Para mudar, atualize o departamento no eSocial e reimporte.");return}
  if(!confirm("Desvincular este contrato?"))return;
  c.vinculos=(c.vinculos||[]).filter(function(v){return v.contratoId!=contratoId});
  svCol();logAcao("Desvinculou colaborador do contrato",c.nome);draw();
}
function nomesSupCt(id){return supervisoresDoContrato(id).filter(function(u){var cg=cargoById(u.cargoId);return cg&&cg.id=="sup"&&!u.bloqueado}).map(function(u){return u.nome}).join(", ")}
function supervisoresDoContrato(contratoId){
  return USERS.filter(function(u){
    var cg=cargoById(u.cargoId);
    if(!cg)return false;
    return ["sup","coord","ger","admin"].indexOf(cg.id)>=0
      && (u.contratosIds||[]).indexOf(contratoId)>=0;
  });
}
function contratosDoUsuario(u){
  if(!u)return CONTRATOS.slice();
  var cg=cargoDoUsuario(u);
  if(cg&&(cg.fixo||cg.id=="admin"))return CONTRATOS.slice();
  if(cg&&(cg.acoes||[]).indexOf("ver_todos_contratos")>=0)return CONTRATOS.slice();
  var ids=(u.contratosIds||[]).slice();
  /* supervisão/coordenação/gerência sem contrato definido continuam vendo todos (comportamento anterior) */
  if(!ids.length&&cg&&["sup","coord","ger"].indexOf(cg.id)>=0)return CONTRATOS.slice();
  /* demais (cliente, encarregado, cargos novos): só os contratos atribuídos ou do próprio vínculo no efetivo */
  var col=u.colId?cget(u.colId):null;
  if(col)(col.vinculos||[]).forEach(function(v){if(ids.indexOf(v.contratoId)<0)ids.push(v.contratoId)});
  return CONTRATOS.filter(function(c){return ids.indexOf(c.id)>=0});
}

/* ===================================================================
   TELA: CONTRATOS
   =================================================================== */
var ctSel=ld("rad_ctSel",null);

function vContratos(){
  var pode=temAcao("gerenciar_contratos")||canEdit();
  var u=session?userById(session.userId):null;
  var meus=contratosDoUsuario(u);
  var lista=meus.slice().sort(function(a,b){return (a.codigo||"").localeCompare(b.codigo||"")}).filter(function(c){return bate("ct",c.codigo,c.nome,c.responsavel,c.cnpj,nomesSupCt(c.id))});
  var h='<div class="card"><b>'+T("menu_ct")+'</b>'
    +'<div class="s">Cadastre contratos e vincule colaboradores.</div>'
    +'<div class="s" style="margin-top:8px"><b>Total:</b> '+meus.length+' · Ativos: '+meus.filter(function(c){return c.status=="Ativo"}).length+' · Do efetivo: '+meus.filter(function(c){return c.origem=="efetivo"}).length+'</div>'
    +campoBusca("ct","Buscar por contrato, código ou supervisor")
    +(pode?'<button class="btn alt" onclick="syncCtUI()">🔄 Sincronizar com departamentos do efetivo</button>':'')
    +'<button class="btn" onclick="go(\'fc\')">🏭 Fábrica de Contratos — criar e configurar</button>'
    +'</div>';

  if(!lista.length){
    h+='<div class="card s">Nenhum contrato disponível.</div>';
  } else {
    h+=lista.map(function(c){
      var colabs=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==c.id})}).length;
      var ptis=PTIS.filter(function(p){return p.contratoId==c.id}).length;
      var pend=PTIS.filter(function(p){return p.contratoId==c.id&&p.status!="Aprovado"&&p.status!="Rejeitado"}).length;
      var aberto=ctSel==c.id;
      /* CORREÇÃO: o botão abrir/fechar estava com aspas trocadas (quebrava o script inteiro) */
      return '<div class="card'+(aberto?' run-card':'')+'">'
        +'<div class="row" style="border:0;padding:0;flex-wrap:wrap">'
          +'<b style="flex:1">'+esc(c.codigo)+' · '+esc(c.nome)+(c.origem=="efetivo"?' <span class="s">· 🏢 departamento</span>':'')+'</b><span class="s" style="margin-right:6px">'+tiposCtIco(c)+'</span>'
          +'<span class="pill '+(c.status=="Ativo"?'ok':'pen')+'" style="cursor:default">'+esc(c.status)+'</span>'
        +'</div>'
        +'<div class="s">👔 Supervisor: '+(nomesSupCt(c.id)?marca(nomesSupCt(c.id),"ct"):'<span style="color:var(--warn)">não definido</span>')+'</div>'
        +'<div class="s">'+esc(c.cnpj||"—")+' · '+esc(c.responsavel||"—")+(c.fim?' · até '+esc(fmtDataBR(c.fim)):'')+'</div>'
        +'<div class="s">👥 '+colabs+' colab. · 📋 '+ptis+' PTI'+(pend?' · <b style="color:var(--run)">'+pend+' pendente(s)</b>':'')+'</div>'
        +'<div class="chips" style="margin-top:8px">'
          +'<button class="'+(aberto?"sel":"")+'" onclick="toggleCt('+c.id+')">'+(aberto?'▲ '+T("btn_close"):'▼ '+T("btn_open"))+'</button>'
          +(pode?'<button onclick="setCtStatus('+c.id+',\''+(c.status=="Ativo"?"Encerrado":"Ativo")+'\')">↔️ '+(c.status=="Ativo"?"Encerrar":"Reativar")+'</button>':'')
        +'</div>'
        +(aberto?detalheContrato(c,pode):'')
        +'</div>';
    }).join("");
  }

  if(pode){
    h+='<div class="card"><b>➕ Novo contrato</b>'
      +'<input id="ct_codigo" placeholder="Código (ex: CT-2026-001)">'
      +'<input id="ct_nome" placeholder="Nome do contrato">'
      +'<input id="ct_cnpj" placeholder="CNPJ">'
      +'<input id="ct_end" placeholder="Endereço">'
      +'<input id="ct_resp" placeholder="Responsável">'
      +'<input id="ct_fone" placeholder="Telefone">'
      +'<input id="ct_email" placeholder="E-mail">'
      +'<div class="s">Início</div><input id="ct_ini" type="date">'
      +'<div class="s">Fim</div><input id="ct_fim" type="date">'
      +'<div class="s" style="margin-top:8px"><b>Tipos de serviço</b> (marque todos; definem checklists e modelos de áreas)</div>'
      +'<div>'+TIPOS_CT.map(function(t){return '<label class="s" style="display:inline-block;margin:3px 8px 3px 0;cursor:pointer"><input type="checkbox" class="ct_tp" value="'+t.id+'" style="width:auto;margin-right:4px">'+t.ico+' '+esc(t.nome)+'</label>'}).join("")+'</div>'
      +'<div class="s">As áreas são mapeadas depois, pelo supervisor, em 🗺️ Áreas do Contrato.</div>'
      +'<button class="btn" onclick="addContrato()" style="margin-top:10px">Cadastrar contrato</button>'
      +'</div>';
  }
  return h;
}
function syncCtUI(){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var r=syncContratosDepto();syncSupervisores();prepararAreas();logAcao("Sincronizou contratos com departamentos",r.novos+" novo(s)");
  alert(r.novos+" contrato(s) criado(s) · "+r.vinc+" vínculo(s) ajustado(s).");draw();
}
function toggleCt(id){ctSel=(ctSel==id?null:id);sv("rad_ctSel",ctSel);draw()}

function detalheContrato(c,pode){
  var colabs=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==c.id})});
  var h='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px">';
  h+=atalhosCtBloco(c);
  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin:12px 0 6px">🧭 Tipos de serviço</div>';
  if(pode||podeAreasCt(c.id))h+=selTipoCt(c);
  else h+='<div class="s">'+esc(tiposCtNome(c))+'</div>';
  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin:12px 0 6px">🗺️ Áreas do contrato ('+locaisCt(c).length+')</div>'
    +'<div class="s">'+(locaisCt(c).map(function(i){return esc(nomeLoc(i))}).join(" · ")||"Ainda não mapeadas.")+'</div>'
    +(acessosDoUsuario().indexOf("ar")>=0?'<button class="btn alt" onclick="abrirAreasCt('+c.id+')">🗺️ Mapear áreas e liberar acesso</button>':'');
  h+=acessosCtBloco(c,pode);
  h+=portalCtBloco(c,pode);

  h+=supCtBloco(c,pode);
  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:6px">👥 Colaboradores ('+colabs.length+')</div>';
  if(colabs.length>10)h+=campoBusca("colct_"+c.id,"Buscar colaborador deste contrato");
  var colabsVis=colabs.filter(function(x){return bate("colct_"+c.id,x.nome,x.funcao,x.matricula)});
  if(colabs.length>10&&!termo("colct_"+c.id)){h+='<div class="s">Digite para localizar um colaborador. A equipe completa está em 👥 Equipe e no PTG.</div>';colabsVis=[]}
  if(colabs.length){
    h+=colabsVis.map(function(x){
      var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==c.id})[0]||{};
      var sup=v.supervisorId?userById(v.supervisorId):null;
      return '<div class="row" style="padding:6px 0;flex-wrap:wrap">'
        +'<span style="flex:1;min-width:180px"><b>'+marca(x.nome,"colct_"+c.id)+'</b> <span class="s">· '+esc(x.funcao)+((x.vinculos||[]).some(function(vv){return vv.contratoId==c.id&&vv.auto})?' · do departamento':' · cobertura')+'</span>'
        +(sup?'<div class="s">Supervisor: '+esc(sup.nome)+'</div>':'<div class="s" style="color:var(--warn)">Sem supervisor</div>')
        +'</span>'
        +(pode?'<button class="pill no" style="border:0;cursor:pointer" onclick="delVinculo('+x.id+','+c.id+')">✕</button>':'')
        +'</div>';
    }).join("");
  } else {
    h+='<div class="s">Nenhum colaborador vinculado.</div>';
  }

  if(pode){
    h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:6px">➕ Incluir colaborador de outro departamento (cobertura)</div>'
      +'<div class="s">Os colaboradores do departamento já entram sozinhos. Use apenas para quem atua aqui vindo de outro contrato.</div>'
      +campoBusca("vinc_"+c.id,"Digite o nome ou a matrícula");
    if(termo("vinc_"+c.id)){
      var disp=COL.filter(function(x){return !ehDeslig(x.situacao)&&!(x.vinculos||[]).some(function(v){return v.contratoId==c.id})&&bate("vinc_"+c.id,x.nome,x.matricula)});
      h+=disp.length?disp.slice(0,12).map(function(x){return '<div class="row" style="padding:6px 0"><span style="flex:1"><b>'+marca(x.nome,"vinc_"+c.id)+'</b><div class="s">'+esc(x.funcao||"")+(x.departamento?' · '+esc(x.departamento):'')+'</div></span><button class="pill ok" style="border:0;cursor:pointer" onclick="vincularUI('+c.id+','+x.id+')">Incluir</button></div>'}).join("")+(disp.length>12?'<div class="s">+'+(disp.length-12)+' resultado(s). Continue digitando.</div>':'')
        :'<div class="s">Ninguém encontrado.</div>';
    }
    h+='<div style="display:flex;gap:6px;margin-top:12px;flex-wrap:wrap">'
      +'<button class="pill no" style="border:0;cursor:pointer;padding:8px 12px" onclick="delContrato('+c.id+')">🗑️ '+T("btn_delete")+'</button>'
      +'</div>';
  }

  h+='</div>';
  return h;
}

function vincularUI(ctId,colId){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var sups=supervisoresDoContrato(ctId).filter(function(u){return u.cargoId=="sup"});
  if(addVinculo(colId,ctId,sups.length==1?sups[0].id:null)){BUSCA["vinc_"+ctId]="";logAcao("Incluiu colaborador em cobertura",(cget(colId)||{}).nome,(ctById(ctId)||{}).codigo);draw()}
}
/* Supervisores que podem ser adicionados ao contrato: do efetivo (função supervisor) e usuários com cargo de supervisão */
function candidatosSup(ctId){
  var r=[],vistos={};
  USERS.forEach(function(u){
    var cg=cargoById(u.cargoId);if(!cg||["sup","coord","ger"].indexOf(cg.id)<0||u.bloqueado)return;
    if((u.contratosIds||[]).indexOf(ctId)>=0)return;
    var col=u.colId?cget(u.colId):null;
    r.push({key:"u"+u.id,uid:u.id,nome:u.nome,funcao:col?col.funcao:cg.nome,dep:col?col.departamento:"",mat:u.mat||""});
    if(u.colId)vistos[u.colId]=1;
  });
  COL.forEach(function(c){
    if(vistos[c.id]||cargoPorFuncao(c.funcao)!="sup"||ehDeslig(c.situacao))return;
    if(userDoColab(c.id))return;
    r.push({key:"c"+c.id,uid:null,nome:c.nome,funcao:c.funcao,dep:c.departamento,mat:c.matricula||""});
  });
  return r.sort(function(a,b){return (a.nome||"").localeCompare(b.nome||"")});
}
function adicionarSupCt(key,ctId){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var u=null;
  if(key.charAt(0)=="u")u=userById(+key.slice(1));
  else{var c=cget(+key.slice(1));if(!c)return;u=userDoColab(c.id)||criarUserColab(c,"sup")}
  if(!u)return;
  u.contratosIds=u.contratosIds||[];
  if(u.contratosIds.indexOf(ctId)<0)u.contratosIds.push(ctId);
  if(u.ctAutoBloq==ctId)u.ctAutoBloq=null;
  saveUsers();BUSCA["supct_"+ctId]="";
  logAcao("Adicionou supervisor ao contrato",u.nome,(ctById(ctId)||{}).codigo);draw();
}
function removerSupCt(uid,ctId){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var u=userById(uid);if(!u)return;
  if(!confirm("Retirar "+u.nome+" da supervisão deste contrato?"))return;
  u.contratosIds=(u.contratosIds||[]).filter(function(x){return x!=ctId});
  if(u.ctAuto==ctId)u.ctAutoBloq=ctId;
  COL.forEach(function(c){(c.vinculos||[]).forEach(function(v){if(v.contratoId==ctId&&v.supervisorId==uid)v.supervisorId=null})});
  saveUsers();svCol();logAcao("Retirou supervisor do contrato",u.nome,(ctById(ctId)||{}).codigo);draw();
}
function vincularTodos(ctId){
  if(!confirm("Vincular a este contrato TODOS os colaboradores ainda não vinculados a ele?"))return;
  var add=0;
  COL.forEach(function(x){
    if(addVinculo(x.id,ctId,null,null,null,null,true))add++;
  });
  logAcao("Vinculou em massa",add+" colab.");draw();alert(add+" vinculado(s).");
}

/* ===================================================================
   PTI
   =================================================================== */
var PTIS=ld("rad_ptis",[]);
/* CORREÇÃO: ptiSel não era declarada — a tela PTI quebrava ao abrir */
var ptiSel=ld("rad_ptiSel",null);
function svPti(){sv("rad_ptis",PTIS)}
function ptiById(id){return PTIS.filter(function(p){return p.id==id})[0]||null}
function ptiDoColab(colabId,contratoId){
  return PTIS.filter(function(p){return p.colaboradorId==colabId&&p.contratoId==contratoId})[0]||null;
}
var PTI_STATUS={
  "Rascunho":"pen",
  "Aguardando Coordenação":"run",
  "Aguardando Gerência":"run",
  "Aguardando Administrador":"run",
  "Aprovado":"ok",
  "Rejeitado":"no"
};
var PTI_FLUXO=["Coordenação","Gerência","Administrador"];
var PTI_ACAO={"Coordenação":"aprovar_pti_coord","Gerência":"aprovar_pti_ger","Administrador":"aprovar_pti_admin"};
function novoPti(colabId,contratoId){
  var c=cget(colabId);if(!c)return null;
  var v=(c.vinculos||[]).filter(function(x){return x.contratoId==contratoId})[0];
  if(!v){alert("Colaborador não está vinculado a este contrato.");return null}
  var p={
    id:Date.now(),
    colaboradorId:colabId,
    contratoId:contratoId,
    supervisorId:v.supervisorId||(session?session.userId:null),
    areas:[],
    metas:"",
    obs:"",
    status:"Rascunho",
    aprovacoes:[],
    criado:Date.now(),
    atualizado:Date.now()
  };
  PTIS.push(p);svPti();
  return p;
}
function ptiEditavel(p){return p&&(p.status=="Rascunho"||p.status=="Rejeitado")&&(canEdit()||temAcao("aprovar_pti_sup"))}
function ptiSet(id,k,v){
  var p=ptiById(id);if(!p)return;
  if(!ptiEditavel(p)){alert("PTI só pode ser editado em Rascunho ou após rejeição.");draw();return}
  p[k]=v;p.atualizado=Date.now();svPti();
}
function ptiTogAtiv(ptiId,areaLocIdx,ativId){
  var p=ptiById(ptiId);if(!p)return;
  if(!ptiEditavel(p)){alert("PTI bloqueado para edição.");draw();return}
  var a=p.areas.filter(function(x){return x.locIdx==areaLocIdx})[0];
  if(!a){a={locIdx:areaLocIdx,atividades:[],periodicidade:"diaria",horario:""};p.areas.push(a)}
  var pos=a.atividades.indexOf(ativId);
  if(pos>=0)a.atividades.splice(pos,1);else a.atividades.push(ativId);
  p.atualizado=Date.now();svPti();
}
function ptiTogAtivArea(ptiId,locIdx,on){
  var p=ptiById(ptiId);if(!p)return;
  if(!ptiEditavel(p)){alert("PTI bloqueado para edição.");draw();return}
  var pos=-1;
  p.areas.forEach(function(a,i){if(a.locIdx==locIdx)pos=i});
  if(on&&pos<0){p.areas.push({locIdx:locIdx,atividades:[],periodicidade:"diaria",horario:""})}
  else if(!on&&pos>=0){p.areas.splice(pos,1)}
  p.atualizado=Date.now();svPti();
  var y=window.scrollY;draw();window.scrollTo(0,y);
}
function ptiMarcarTodas(ptiId,locIdx,on){
  var p=ptiById(ptiId);if(!p)return;
  if(!ptiEditavel(p)){alert("PTI bloqueado para edição.");return}
  var area=p.areas.filter(function(a){return a.locIdx==locIdx})[0];
  if(!area)return;
  var ativs=PG.filter(function(g){return g.l==locIdx}).map(function(a){return a.id});
  area.atividades=on?ativs.slice():[];
  p.atualizado=Date.now();svPti();
  var y=window.scrollY;draw();window.scrollTo(0,y);
}
function ptiSetArea(id,locIdx,k,v){
  var p=ptiById(id);if(!p)return;
  if(!ptiEditavel(p)){alert("PTI bloqueado para edição.");draw();return}
  var a=p.areas.filter(function(x){return x.locIdx==locIdx})[0];
  if(!a)return;
  a[k]=v;p.atualizado=Date.now();svPti();
}
/* CORREÇÃO: o fluxo de aprovação estava deslocado em uma etapa
   (quem estava em "Aguardando Coordenação" só podia ser aprovado pela Gerência, e
   a Coordenação conseguia aprovar um Rascunho ainda não enviado). */
function ptiNivelAtual(p){var m=/^Aguardando (.+)$/.exec((p&&p.status)||"");return m?m[1]:null}
function ptiAprovar(id,obs){
  var p=ptiById(id);if(!p)return;
  var nivel=ptiNivelAtual(p);
  if(!nivel){alert("Este PTI não está aguardando aprovação.");return}
  if(!ptiPodeAprovar(p)){alert("Sem permissão para esta etapa.");return}
  p.aprovacoes.push({nivel:nivel,userId:session.userId,nome:user,status:"Aprovado",data:Date.now(),obs:obs||""});
  var i=PTI_FLUXO.indexOf(nivel);
  p.status=(i>=0&&i<PTI_FLUXO.length-1)?"Aguardando "+PTI_FLUXO[i+1]:"Aprovado";
  if(p.status=="Aprovado"){var cc=cget(p.colaboradorId);var vv=cc&&(cc.vinculos||[]).filter(function(v){return v.contratoId==p.contratoId})[0];if(vv){vv.areas=vv.areas||[];p.areas.forEach(function(a){if(vv.areas.indexOf(a.locIdx)<0)vv.areas.push(a.locIdx)});svCol()}}
  p.atualizado=Date.now();svPti();
  logAcao("Aprovou PTI",(cget(p.colaboradorId)||{}).nome,nivel);
  alert("PTI: "+p.status);draw();
}
function ptiRejeitar(id,obs){
  var p=ptiById(id);if(!p)return;
  if(!ptiPodeAprovar(p)){alert("Sem permissão para esta etapa.");return}
  if(!obs||!obs.trim()){alert("Informe o motivo.");return}
  p.aprovacoes.push({nivel:ptiNivelAtual(p)||"Rejeição",userId:session.userId,nome:user,status:"Rejeitado",data:Date.now(),obs:obs.trim()});
  p.status="Rejeitado";p.atualizado=Date.now();svPti();
  logAcao("Rejeitou PTI",(cget(p.colaboradorId)||{}).nome);alert("Rejeitado. O supervisor pode corrigir e reenviar.");draw();
}
function ptiPodeAprovar(p){
  var n=ptiNivelAtual(p);if(!n)return false;
  var cg=cargoAtual();
  if(!cg)return false;
  if(cg.fixo||cg.id=="admin")return true;
  return (cg.acoes||[]).indexOf(PTI_ACAO[n])>=0;
}
function ptiEnviar(id){
  var p=ptiById(id);if(!p)return;
  if(!ptiEditavel(p)){alert("Sem permissão.");return}
  if(!p.areas.some(function(a){return a.atividades.length})){alert("Adicione atividades.");return}
  p.aprovacoes.push({nivel:"Supervisor",userId:session.userId,nome:user,status:"Enviado",data:Date.now(),obs:""});
  p.status="Aguardando "+PTI_FLUXO[0];p.atualizado=Date.now();svPti();
  logAcao("Enviou PTI",(cget(p.colaboradorId)||{}).nome);alert("PTI enviado para "+PTI_FLUXO[0]+".");draw();
}
/* ---- programação por atividade no PTI ---- */
var PTI_PER=[["diaria","Diária"],["2xdia","2× ao dia"],["3xdia","3× ao dia"],["semanal","Semanal"],["quinzenal","Quinzenal"],["mensal","Mensal"],["bimestral","Bimestral"],["trimestral","Trimestral"],["semestral","Semestral"],["anual","Anual"],["eventual","Eventual / sob demanda"]];
var DIAS_SEM=[[1,"Seg"],[2,"Ter"],[3,"Qua"],[4,"Qui"],[5,"Sex"],[6,"Sáb"],[0,"Dom"]];
var FREQ_PADRAO={"Diária":["diaria"],"Mais de uma vez ao dia":["2xdia"],"Semanal":["semanal"],"Quinzenal":["quinzenal"],"Mensal":["mensal"],"Eventual / sob demanda":["eventual"]};
function ptiCfg(p,locIdx,ativId){
  var a=p.areas.filter(function(x){return x.locIdx==locIdx})[0];if(!a)return null;
  a.cfg=a.cfg||{};
  if(!a.cfg[ativId]){var g=PG.filter(function(x){return x.id==ativId})[0]||{};a.cfg[ativId]={per:(FREQ_PADRAO[g.freq||"Diária"]||["diaria"]).slice(),dias:[1,2,3,4,5,6],hor:""}}
  return a.cfg[ativId];
}
function ptiAtivCfgHtml(p,locIdx,a,edit){
  var c=ptiCfg(p,locIdx,a.id),pid=p.id;
  if(!edit)return '<div class="s" style="margin-left:28px">'+esc(ptiCfgTxt(c))+'</div>';
  var temSem=c.per.some(function(x){return /dia|semanal|quinzenal/.test(x)});
  return '<div style="margin-left:28px">'
    +'<div class="s">Periodicidade (marque todas que se aplicam)</div><div>'+PTI_PER.map(function(x){return chip(c.per.indexOf(x[0])>=0,"ptiTogCfg("+pid+","+locIdx+","+a.id+",\'per\',\'"+x[0]+"\')",x[1])}).join("")+'</div>'
    +(temSem?'<div class="s" style="margin-top:4px">Dias da semana</div><div>'+DIAS_SEM.map(function(d){return chip(c.dias.indexOf(d[0])>=0,"ptiTogCfg("+pid+","+locIdx+","+a.id+",\'dias\',"+d[0]+")",d[1])}).join("")+'</div>':'')
    +'<input placeholder="Horário(s) — ex.: 08h, 12h e 16h · ou dia do mês (ex.: todo dia 10)" value="'+esc(c.hor||"")+'" onchange="ptiSetCfgHor('+pid+','+locIdx+','+a.id+',this.value)" style="margin-top:4px">'
    +'</div>';
}
function ptiCfgTxt(c){
  if(!c)return "";
  var per=c.per.map(function(k){return (PTI_PER.filter(function(x){return x[0]==k})[0]||[0,k])[1]}).join(" + ");
  var temSem=c.per.some(function(x){return /dia|semanal|quinzenal/.test(x)});
  var dias=temSem?(c.dias.length==7?"todos os dias":(c.dias.length==5&&c.dias.indexOf(0)<0&&c.dias.indexOf(6)<0?"seg a sex":(c.dias.length==6&&c.dias.indexOf(0)<0?"seg a sáb":DIAS_SEM.filter(function(d){return c.dias.indexOf(d[0])>=0}).map(function(d){return d[1]}).join(", ")))):"";
  return [per,dias,c.hor].filter(Boolean).join(" · ");
}
function ptiTogCfg(pid,locIdx,ativId,campo,v){
  var p=ptiById(pid);if(!p)return;if(!ptiEditavel(p)){alert("PTI bloqueado para edição.");return}
  var c=ptiCfg(p,locIdx,ativId),arr=c[campo],i=arr.indexOf(v);
  if(i>=0){if(arr.length==1){alert(campo=="per"?"Mantenha ao menos uma periodicidade.":"Mantenha ao menos um dia.");return}arr.splice(i,1)}else arr.push(v);
  p.atualizado=Date.now();svPti();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function ptiSetCfgHor(pid,locIdx,ativId,v){var p=ptiById(pid);if(!p||!ptiEditavel(p))return;ptiCfg(p,locIdx,ativId).hor=(v||"").trim();p.atualizado=Date.now();svPti()}
function delPti(id){
  if(!canEdit()){alert("Sem permissão.");return}
  if(!confirm("Excluir este PTI?"))return;
  PTIS=PTIS.filter(function(p){return p.id!=id});svPti();logAcao("Excluiu PTI");draw();
}
function ptisPendentesParaMim(){
  var cg=cargoAtual();
  if(!cg||!session)return [];
  var me=session.userId;
  return PTIS.filter(function(p){
    if(ptiPodeAprovar(p))return true;
    if((p.status=="Rascunho"||p.status=="Rejeitado")&&(cg.acoes||[]).indexOf("aprovar_pti_sup")>=0&&p.supervisorId==me)return true;
    return false;
  });
}

/* ===================================================================
   PLANO GERAL
   =================================================================== */
function planoGeral(contratoId){
  var c=ctById(contratoId);if(!c)return null;
  var ptisAprov=PTIS.filter(function(p){return p.contratoId==contratoId&&p.status=="Aprovado"});
  var ptisPend=PTIS.filter(function(p){return p.contratoId==contratoId&&p.status!="Aprovado"&&p.status!="Rejeitado"});
  var colabs=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==contratoId})});
  var cobertura={};
  ptisAprov.forEach(function(p){
    p.areas.forEach(function(a){
      cobertura[a.locIdx]=cobertura[a.locIdx]||{colabs:[],atividades:{}};
      if(cobertura[a.locIdx].colabs.indexOf(p.colaboradorId)<0)cobertura[a.locIdx].colabs.push(p.colaboradorId);
      a.atividades.forEach(function(ativId){
        var atv=PG.filter(function(g){return g.id==ativId})[0];
        if(atv){var k=atv.nome;cobertura[a.locIdx].atividades[k]=(cobertura[a.locIdx].atividades[k]||0)+1}
      });
    });
  });
  return {contrato:c,totalColaboradores:colabs.length,totalPtis:PTIS.filter(function(p){return p.contratoId==contratoId}).length,aprovados:ptisAprov.length,pendentes:ptisPend.length,cobertura:cobertura,colaboradores:colabs};
}

function vPTI(){
  var lista;
  var cg=cargoAtual();
  if(!cg)lista=[];
  else if(cg.fixo||cg.id=="admin"||cg.id=="ger"||cg.id=="coord"||canEdit())lista=PTIS.slice();
  else if(cg.id=="sup")lista=PTIS.filter(function(p){return p.supervisorId==session.userId});
  else lista=PTIS.filter(function(p){return p.colaboradorId==cid});

  lista.sort(function(a,b){return b.atualizado-a.atualizado});

  var pend=ptisPendentesParaMim();
  var h='<div class="card"><b>'+T("menu_pti")+'</b>'
    +'<div class="s">Estruture as atividades. Fluxo: Supervisor envia → Coordenação → Gerência → Administrador.</div>';
  if(pend.length)h+='<div class="s" style="margin-top:8px;color:var(--run);font-weight:700">⚡ '+pend.length+' PTI(s) aguardando você</div>';
  h+='</div>';

  if(!lista.length){
    h+='<div class="card s">Nenhum PTI registrado.</div>';
  } else {
    h+=lista.map(function(p){
      var c=cget(p.colaboradorId);
      var ct=ctById(p.contratoId);
      var aberto=ptiSel==p.id;
      var stCl=PTI_STATUS[p.status]||"pen";
      return '<div class="card'+(aberto?' run-card':'')+'">'
        +'<div class="row" style="border:0;padding:0;flex-wrap:wrap"><b style="flex:1">'+esc(c?c.nome:"?")+'</b><span class="pill '+stCl+'" style="cursor:default">'+esc(p.status)+'</span></div>'
        +'<div class="s">📁 '+esc(ct?ct.codigo+" · "+ct.nome:"?")+'</div>'
        +'<div class="s">'+p.areas.length+' área(s) · '+p.areas.reduce(function(s,a){return s+a.atividades.length},0)+' atividade(s)</div>'
        +'<div class="chips" style="margin-top:8px"><button class="'+(aberto?"sel":"")+'" onclick="togglePti('+p.id+')">'+(aberto?'▲ '+T("btn_close"):'▼ '+T("btn_open"))+'</button></div>'
        +(aberto?detalhePti(p):'')+'</div>';
    }).join("");
  }

  if(canEdit()||(cg&&(cg.acoes||[]).indexOf("aprovar_pti_sup")>=0))h+=novoPtiCard();
  return h;
}
function togglePti(id){ptiSel=(ptiSel==id?null:id);sv("rad_ptiSel",ptiSel);draw()}

var ptiNovoCol=null;
function novoPtiCard(){
  var meus=meusContratos().map(function(c){return c.id});
  var colabs=COL.filter(function(x){return !ehDeslig(x.situacao)&&(x.vinculos||[]).some(function(v){return meus.indexOf(v.contratoId)>=0})});
  if(!colabs.length)return '<div class="card"><b>➕ '+T("btn_add")+' PTI</b><div class="s">Nenhum colaborador nos seus contratos.</div></div>';
  var h='<div class="card"><b>➕ '+T("btn_add")+' PTI</b>';
  var sel=ptiNovoCol?cget(ptiNovoCol):null;
  if(sel){
    var conts=(sel.vinculos||[]).filter(function(v){return meus.indexOf(v.contratoId)>=0});
    h+='<div class="row" style="border:0;padding:6px 0"><span style="flex:1"><b>'+esc(sel.nome)+'</b><div class="s">'+esc(sel.funcao||"")+'</div></span><button class="pill pen" style="border:0;cursor:pointer" onclick="ptiNovoCol=null;draw()">Trocar</button></div>'
      +'<input type="hidden" id="novo_pti_colab" value="'+sel.id+'">'
      +'<div class="s">Contrato</div><select id="novo_pti_contrato">'+conts.map(function(v){var ct=ctById(v.contratoId);return '<option value="'+v.contratoId+'">'+esc(ct?ct.codigo+" · "+ct.nome:"?")+'</option>'}).join("")+'</select>'
      +'<button class="btn" onclick="criarPtiUI()">Criar PTI</button>';
  }else{
    h+=campoBusca("ptic","Digite o nome do colaborador ("+colabs.length+" nos seus contratos)");
    if(termo("ptic")){
      var r=colabs.filter(function(x){return bate("ptic",x.nome,x.matricula,x.funcao)});
      h+=r.length?r.slice(0,12).map(function(x){return '<div class="row" style="padding:6px 0;cursor:pointer" onclick="ptiNovoCol='+x.id+';BUSCA.ptic=\'\';draw()"><span style="flex:1"><b>'+marca(x.nome,"ptic")+'</b><div class="s">'+esc(x.funcao||"")+(x.departamento?' · '+esc(x.departamento):'')+(ptiDoColab(x.id,(x.vinculos[0]||{}).contratoId)?' · já tem PTI':'')+'</div></span><span class="pill ok">Escolher</span></div>'}).join("")+(r.length>12?'<div class="s">+'+(r.length-12)+' resultado(s). Continue digitando.</div>':'')
        :'<div class="s">Ninguém encontrado.</div>';
    }
  }
  return h+'</div>';
}
function novoPtiContratos(){
  var sel=$("#novo_pti_colab");if(!sel||sel.type=="hidden")return;
  var c=cget(+sel.value);if(!c)return;
  var conts=c.vinculos||[];
  var el=document.getElementById("novo_pti_contratos");
  if(!el)return;
  el.innerHTML='<select id="novo_pti_contrato">'
    +conts.map(function(v){
      var ct=ctById(v.contratoId);
      return '<option value="'+v.contratoId+'">'+esc(ct?ct.codigo+" · "+ct.nome:"?")+'</option>';
    }).join("")
    +'</select>';
}
function criarPtiUI(){
  var cid_=+$("#novo_pti_colab").value;
  var ctEl=$("#novo_pti_contrato");
  var ctid=ctEl?+ctEl.value:0;
  if(!cid_||!ctid){alert("Selecione colaborador e contrato.");return}
  if(ptiDoColab(cid_,ctid)){alert("Já existe PTI deste colaborador neste contrato.");return}
  var p=novoPti(cid_,ctid);
  if(p){ptiNovoCol=null;ptiSel=p.id;sv("rad_ptiSel",ptiSel);logAcao("Criou PTI",(cget(cid_)||{}).nome);draw();alert("PTI criado.")}
}

function detalhePti(p){
  var ct=ctById(p.contratoId);
  var pode=ptiPodeAprovar(p);
  var edit=ptiEditavel(p);
  var dis=edit?'':' disabled';
  var h='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px">';

  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">🎯 Áreas e atividades</div>';
  if(!edit)h+='<div class="s" style="margin-bottom:6px">🔒 Em análise/aprovado — edição bloqueada.</div>';

  var locaisDisponiveis=ct?locaisCt(ct):[];
  if(!locaisDisponiveis.length)h+='<div class="s">O contrato ainda não tem áreas mapeadas. O supervisor deve criá-las em 🗺️ Áreas do Contrato.</div>';

  locaisDisponiveis.forEach(function(locIdx){
    var L=LOC[locIdx];if(!L)return;
    var ativs=PG.filter(function(g){return g.l==locIdx});
    var area=p.areas.filter(function(a){return a.locIdx==locIdx})[0];
    var on=!!area;
    if(!edit&&!on)return;
    h+='<div style="margin-top:10px;border:1px solid var(--line);border-radius:8px;padding:10px">'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer">'
        +'<input type="checkbox" '+(on?'checked':'')+dis+' onchange="ptiTogAtivArea('+p.id+','+locIdx+',this.checked)" style="width:auto;margin:0">'
        +'<b>'+esc(L.n)+'</b>'
      +'</label>';
    if(on){
      if(edit)h+='<div class="chips" style="margin:8px 0 4px"><button onclick="ptiMarcarTodas('+p.id+','+locIdx+',1)">Marcar todas</button><button onclick="ptiMarcarTodas('+p.id+','+locIdx+',0)">Limpar</button></div>';
      h+=(ativs.length?'<div style="display:flex;flex-direction:column;gap:4px">'
        +ativs.map(function(a){var onAt=area.atividades.indexOf(a.id)>=0;if(!edit&&!onAt)return "";return '<div style="border-bottom:1px dashed var(--line);padding:4px 0"><label class="s" style="display:flex;align-items:center;gap:8px;padding:4px 0;cursor:pointer"><input type="checkbox" '+(onAt?'checked':'')+dis+' onchange="ptiTogAtiv('+p.id+','+locIdx+','+a.id+')" style="width:auto;margin:0"><b style="color:var(--ink)">'+esc(a.nome)+'</b></label>'+(onAt?ptiAtivCfgHtml(p,locIdx,a,edit):'')+'</div>'}).join("")
        +'</div>':'<div class="s">Sem atividades cadastradas para esta área (Configurações → Atividades).</div>');
      h+='<div class="s" style="margin-top:6px">Horário geral da área / turno (opcional)</div>'
        +'<input'+dis+' placeholder="Ex.: 07h às 15h20" value="'+esc(area.horario||"")+'" onchange="ptiSetArea('+p.id+','+locIdx+',\'horario\',this.value)">';
    }
    h+='</div>';
  });

  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:6px">📝 Metas</div>'
    +'<textarea rows="3"'+dis+' placeholder="Metas e observações" onchange="ptiSet('+p.id+',\'metas\',this.value)" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:6px 0">'+esc(p.metas||"")+'</textarea>';

  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:6px">👣 Histórico</div>';
  if(p.aprovacoes.length){
    h+=p.aprovacoes.map(function(a){return '<div class="row" style="padding:4px 0;flex-wrap:wrap"><span class="s" style="flex:1"><b>'+esc(a.nivel)+'</b> · '+esc(a.nome)+' · '+new Date(a.data).toLocaleString("pt-BR")+(a.obs?'<br>'+esc(a.obs):'')+'</span><span class="pill '+(a.status=="Rejeitado"?"no":a.status=="Enviado"?"pen":"ok")+'" style="cursor:default">'+esc(a.status)+'</span></div>'}).join("");
  } else {
    h+='<div class="s">Nenhuma movimentação.</div>';
  }

  h+='<div class="chips" style="margin-top:12px">';
  if(edit){
    h+='<button class="sel" onclick="ptiEnviar('+p.id+')">📤 '+(p.status=="Rejeitado"?"Reenviar":"Enviar")+'</button>';
  }
  if(pode){
    h+='<button class="sel" onclick="aprovarPtiUI('+p.id+')">✅ '+T("btn_approve")+' ('+esc(ptiNivelAtual(p))+')</button><button onclick="rejeitarPtiUI('+p.id+')">❌ '+T("btn_reject")+'</button>';
  }
  h+='<button onclick="imprimirPTI('+p.id+')">🖨️ '+T("btn_print")+' PTI</button>';
  if(canEdit())h+='<button onclick="delPti('+p.id+')">🗑️</button>';
  h+='</div></div>';
  return h;
}
function aprovarPtiUI(id){var obs=prompt("Observação (opcional):");if(obs===null)return;ptiAprovar(id,obs)}
function rejeitarPtiUI(id){var obs=prompt("Motivo da rejeição:");if(obs===null)return;ptiRejeitar(id,obs)}

function vPlanoGeral(){
  if(!CONTRATOS.length)return '<div class="card s">Nenhum contrato. Vá em 📁 Contratos.</div>';
  var mc=meusContratos();
  if(!mc.length)return '<div class="card s">Nenhum contrato disponível para você.</div>';
  var ctId=ld("rad_planoCt",mc[0].id);
  var c=mc.filter(function(x){return x.id==ctId})[0]||mc[0];
  var h='<div class="card noprint"><b>'+T("menu_plano")+'</b>'
    +'<div class="s">Consolida PTIs aprovados.</div>'
    +pickContrato("plct",mc,c.id,function(id){sv("rad_planoCt",id);draw();window.scrollTo(0,0)})
    +'<div class="grid-acoes" style="margin-top:8px"><button class="btn" onclick="imprimirPTG('+c.id+')">🖨️ Imprimir PTG</button><button class="btn alt" onclick="imprimirPTIsContrato('+c.id+',true)">🖨️ PTIs aprovados</button><button class="btn alt" onclick="imprimirPTIsContrato('+c.id+',false)">🖨️ Todos os PTIs</button><button class="btn alt" onclick="enviarWhatsAppPlano('+c.id+')">💬 WhatsApp</button></div></div>';
  var pg=planoGeral(c.id);
  if(!pg)return h+'<div class="card s">Não encontrado.</div>';
  h+='<div class="card"><b>📈 Resumo</b><div class="row"><span>Contrato</span><b>'+esc(c.codigo)+'</b></div><div class="row"><span>Status</span><b>'+esc(c.status)+'</b></div><div class="row"><span>Colaboradores</span><b>'+pg.totalColaboradores+'</b></div><div class="row"><span>PTIs totais</span><b>'+pg.totalPtis+'</b></div><div class="row"><span>✅ Aprovados</span><b style="color:var(--ok)">'+pg.aprovados+'</b></div><div class="row"><span>⏳ Pendentes</span><b style="color:var(--run)">'+pg.pendentes+'</b></div></div>';
  h+='<div class="card"><b>🎯 Cobertura por área</b>';
  var locais=locaisCt(c);
  if(!locais.length)h+='<div class="s">Sem áreas definidas no contrato.</div>';
  else{
    locais.forEach(function(locIdx){
      var L=LOC[locIdx];if(!L)return;
      var cob=pg.cobertura[locIdx]||{colabs:[],atividades:{}};
      var nCol=cob.colabs.length;
      var cls=nCol?'ok':'no';
      h+='<div style="border-top:1px solid var(--line);margin-top:10px;padding-top:8px"><div class="row" style="border:0;padding:0"><b>'+esc(L.n)+'</b><span class="pill '+cls+'" style="cursor:default">'+nCol+' colab.</span></div>';
      if(nCol){
        var ativs=Object.keys(cob.atividades);
        if(ativs.length)h+='<div class="s" style="margin-top:4px">'+ativs.map(function(a){return esc(a)+' ('+cob.atividades[a]+')'}).join(" · ")+'</div>';
      } else h+='<div class="s" style="color:var(--warn);margin-top:4px">⚠️ Sem cobertura</div>';
      h+='</div>';
    });
  }
  h+='</div>';
  h+='<div class="card"><b>👥 Colaboradores</b>';
  if(!pg.colaboradores.length)h+='<div class="s">Nenhum.</div>';
  else{
    pg.colaboradores.forEach(function(x){
      var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==c.id})[0]||{};
      var pti=ptiDoColab(x.id,c.id);
      var stCl=pti?PTI_STATUS[pti.status]:"pen";
      var stTxt=pti?pti.status:"Sem PTI";
      h+='<div class="row" style="padding:8px 0"><span style="flex:1"><b>'+esc(x.nome)+'</b> <span class="s">· '+esc(v.funcao||x.funcao)+'</span></span><span class="pill '+stCl+'" style="cursor:default">'+esc(stTxt)+'</span></div>';
    });
  }
  h+='</div>';
  var lacunas=locaisCt(c).filter(function(locIdx){var cob=pg.cobertura[locIdx];return LOC[locIdx]&&(!cob||!cob.colabs.length)});
  if(!locais.length){}
  else if(lacunas.length){
    h+='<div class="card" style="border-left:4px solid var(--warn);background:rgba(198,40,40,.04)"><b style="color:var(--warn)">⚠️ '+lacunas.length+' área(s) sem cobertura</b><ul style="margin:8px 0 0;padding-left:22px;line-height:1.8">'+lacunas.map(function(i){return '<li>'+esc(LOC[i].n)+'</li>'}).join("")+'</ul></div>';
  } else {
    h+='<div class="card" style="border-left:4px solid var(--ok);background:rgba(46,139,87,.04)"><b style="color:var(--ok)">✅ Todas as áreas cobertas</b></div>';
  }
  return h;
}
function enviarWhatsAppPlano(id){
  var c=ctById(id)||CONTRATOS[0];if(!c)return;
  var pg=planoGeral(c.id);
  var L=['📊 *PLANO GERAL — '+c.codigo+' '+c.nome+'*','','👥 Colaboradores: '+pg.totalColaboradores,'📋 PTIs: '+pg.totalPtis,'✅ Aprovados: '+pg.aprovados,'⏳ Pendentes: '+pg.pendentes,''];
  var lacunas=locaisCt(c).filter(function(locIdx){var cob=pg.cobertura[locIdx];return LOC[locIdx]&&(!cob||!cob.colabs.length)});
  if(lacunas.length){L.push('⚠️ Sem cobertura:');lacunas.forEach(function(i){L.push('• '+LOC[i].n)})}
  else L.push('✅ Todas as áreas cobertas');
  L.push('');L.push('_'+new Date().toLocaleString("pt-BR")+'_');
  window.open("https://wa.me/?text="+encodeURIComponent(L.join("\n")),"_blank");
}

/* ===================================================================
   TELA: CARGOS
   =================================================================== */
var cargoSel=ld("rad_cargoSel",null);
function toggleCargo(id){cargoSel=(cargoSel==id?null:id);draw()}
function vCargos(){
  if(!temAcao("gerenciar_cargos"))return '<div class="card s">Sem permissão.</div>';
  var lista=CARGOS.slice().sort(function(a,b){
    if(a.fixo&&!b.fixo)return -1;if(!a.fixo&&b.fixo)return 1;
    return (a.nome||"").localeCompare(b.nome||"");
  });
  var h='<div class="card"><b>⚙️ Cargos e Permissões</b><div class="s">Defina o que cada cargo acessa.</div><div class="s" style="margin-top:8px"><b>Total:</b> '+CARGOS.length+'</div></div>';
  h+=lista.map(function(cg){
    var abertos=cg.acessos.indexOf("*")>=0;
    var nAcessos=abertos?Object.keys(NAV).length-1:cg.acessos.filter(function(k){return k!="perfil"}).length;
    var nAcoes=(cg.acoes||[]).length;
    var nUsuarios=USERS.filter(function(u){return u.cargoId==cg.id}).length;
    var aberto=cargoSel==cg.id;
    return '<div class="card'+(aberto?' run-card':'')+'">'
      +'<div class="row" style="border:0;padding:0;flex-wrap:wrap">'
        +'<span style="display:flex;align-items:center;gap:8px;flex:1"><span style="width:14px;height:14px;border-radius:50%;background:'+esc(cg.cor||"#5d6e69")+'"></span><b>'+esc(cg.nome)+'</b>'+(cg.fixo?' <span class="s">(fixo)</span>':'')+'</span>'
      +'</div>'
      +'<div class="s">'+nAcessos+' menu · '+nAcoes+' ação(ões) · '+nUsuarios+' usuário(s)</div>'
      +'<div class="chips" style="margin-top:8px"><button class="'+(aberto?"sel":"")+'" onclick="toggleCargo(\''+cg.id+'\')">'+(aberto?'▲ Fechar':'▼ Editar')+'</button>'+(cg.fixo?'':'<button onclick="delCargo(\''+cg.id+'\')">🗑️ Excluir</button>')+'</div>'
      +(aberto?editorCargo(cg):'')
      +'</div>';
  }).join("");
  h+='<div class="card"><b>➕ Novo cargo</b>'
    +'<input id="cg_nome" placeholder="Nome do cargo">'
    +'<div class="s">Cor</div><input id="cg_cor" type="color" value="#0f5c4d" style="height:44px">'
    +'<button class="btn" onclick="addCargo()" style="margin-top:8px">Criar cargo</button>'
    +'</div>';
  return h;
}
function editorCargo(cg){
  if(cg.fixo&&cg.id=="admin")return '<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px"><b style="color:var(--ok)">✅ Acesso total</b></div>';
  var h='<div style="border-top:1px dashed var(--line);margin-top:10px;padding-top:8px">';
  h+='<div class="s" style="font-weight:700;font-size:11px;margin-bottom:6px">📱 MENU</div>';
  var itensTodos=Object.keys(NAV).filter(function(k){return k!=="perfil"});
  var abertos=cg.acessos.indexOf("*")>=0;
  h+='<div class="chips"><button onclick="togCargoAll(\''+cg.id+'\',true)">Todas</button><button onclick="togCargoAll(\''+cg.id+'\',false)">Limpar</button></div>';
  h+='<div style="margin-top:6px">'+itensTodos.map(function(k){
    var on=abertos||cg.acessos.indexOf(k)>=0;
    return '<label class="s" style="display:inline-block;margin:4px 10px 4px 0;cursor:pointer"><input type="checkbox" '+(on?'checked':'')+' onchange="togCargoAcesso(\''+cg.id+'\',\''+k+'\',this.checked)" style="width:auto;margin-right:4px">'+NN_ICONE(k)+' '+esc(nomeNav(k))+'</label>';
  }).join("")+'</div>';
  h+='<div class="s" style="font-weight:700;font-size:11px;margin-top:14px;margin-bottom:6px">⚡ AÇÕES</div>';
  h+='<div style="margin-top:6px">'+ACOES_LIST.map(function(a){
    var on=(cg.acoes||[]).indexOf(a.id)>=0;
    return '<label class="s" style="display:block;padding:6px 0;cursor:pointer"><input type="checkbox" '+(on?'checked':'')+' onchange="togCargoAcao(\''+cg.id+'\',\''+a.id+'\',this.checked)" style="width:auto;margin-right:8px">'+esc(a.nome)+'</label>';
  }).join("")+'</div>';
  if(!cg.fixo){
    h+='<div class="s" style="font-weight:700;font-size:11px;margin-top:14px;margin-bottom:6px">✏️ RENOMEAR</div>'
      +'<input value="'+esc(cg.nome)+'" onchange="renCargo(\''+cg.id+'\',\'nome\',this.value)">'
      +'<div class="s">Cor</div><input type="color" value="'+esc(cg.cor||"#0f5c4d")+'" onchange="renCargo(\''+cg.id+'\',\'cor\',this.value)" style="height:44px">';
  }
  h+='</div>';
  return h;
}
function NN_ICONE(k){return {k:"⚙️",r:"📊",l:"🔍",c:"📋",g:"🔄",hOP:"🛠️",hAD:"📄",CLI:"💬",imp:"📥",eq:"👥",ct:"📁",ar:"🗺️",ag:"📅",pti:"📋",plano:"📊",p:"📈",t:"📚",i:"🎯",q:"🏷️",es:"📦",rq:"🛒",st:"📣",perfil:"👤",cont:"📞",ac:"🔑",fc:"🏭",dc:"🎬"}[k]||"•"}
function addCargo(){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var nome=$("#cg_nome").value.trim();
  var cor=$("#cg_cor").value||"#0f5c4d";
  if(!nome){alert("Informe o nome.");return}
  if(CARGOS.some(function(c){return c.nome.toLowerCase()===nome.toLowerCase()})){alert("Já existe.");return}
  var id="cargo_"+Date.now();
  CARGOS.push({id:id,nome:nome,cor:cor,fixo:false,acessos:["perfil"],acoes:[]});
  svCargos();logAcao("Criou cargo",nome);cargoSel=id;draw();
}
function delCargo(id){
  if(!temAcao("gerenciar_cargos")){alert("Sem permissão.");return}
  var cg=cargoById(id);if(!cg)return;
  if(cg.fixo){alert("Fixo não pode excluir.");return}
  if(["colab"].indexOf(id)>=0){alert("O cargo Colaborador é usado como padrão e não pode ser excluído.");return}
  var usos=USERS.filter(function(u){return u.cargoId==id}).length;
  if(usos){alert(usos+" usuário(s) usam este cargo.");return}
  if(!confirm("Excluir cargo "+cg.nome+"?"))return;
  CARGOS=CARGOS.filter(function(c){return c.id!=id});svCargos();logAcao("Excluiu cargo",cg.nome);cargoSel=null;draw();
}
function togCargoAcesso(id,k,on){
  if(!temAcao("gerenciar_cargos"))return;
  var cg=cargoById(id);if(!cg||cg.id=="admin")return;
  if(cg.acessos.indexOf("*")>=0)cg.acessos=["perfil"].concat(Object.keys(NAV).filter(function(x){return x!="perfil"}));
  var p=cg.acessos.indexOf(k);
  if(on&&p<0)cg.acessos.push(k);
  if(!on&&p>=0)cg.acessos.splice(p,1);
  if(cg.acessos.indexOf("perfil")<0)cg.acessos.unshift("perfil");
  svCargos();buildNav();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function togCargoAcao(id,a,on){
  if(!temAcao("gerenciar_cargos"))return;
  var cg=cargoById(id);if(!cg||cg.id=="admin")return;
  cg.acoes=cg.acoes||[];
  var p=cg.acoes.indexOf(a);
  if(on&&p<0)cg.acoes.push(a);
  if(!on&&p>=0)cg.acoes.splice(p,1);
  svCargos();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function togCargoAll(id,on){
  if(!temAcao("gerenciar_cargos"))return;
  var cg=cargoById(id);if(!cg||cg.id=="admin")return;
  if(on)cg.acessos=["*"];else cg.acessos=["perfil"];
  svCargos();buildNav();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function renCargo(id,k,v){
  if(!temAcao("gerenciar_cargos"))return;
  var cg=cargoById(id);if(!cg||cg.fixo)return;
  cg[k]=(v||"").trim()||cg[k];svCargos();
}


/* ===================================================================
   MEU PERFIL
   =================================================================== */
function vMeuPerfil(){
  var u=userById(session.userId);
  if(!u)return '<div class="card s">Sessão inválida.</div>';
  var cg=cargoById(u.cargoId)||cargoById("colab");
  var incompleto=!u.perfilCompleto;
  var h='';
  if(incompleto){
    h+='<div class="card noprint" style="border-left:4px solid var(--run);background:rgba(224,161,0,.08)">'
      +'<b style="color:#8a6400">⚠️ Complete seu perfil</b>'
      +'<div class="s">Você precisa completar o cadastro antes de acessar o sistema.</div></div>';
  }
  if(u.pinTrocar){
    return h+'<div class="card" style="border-left:4px solid var(--warn)"><b>🔑 Crie seu PIN pessoal</b>'
      +'<div class="s">Olá, '+esc(u.nome)+'! Você entrou com o PIN inicial. Antes de continuar, escolha um PIN de 4 dígitos só seu.</div>'
      +'<input id="np1" type="password" inputmode="numeric" maxlength="4" placeholder="Novo PIN (4 dígitos)" style="margin-top:8px">'
      +'<input id="np2" type="password" inputmode="numeric" maxlength="4" placeholder="Repita o novo PIN">'
      +'<button class="btn" onclick="salvarPinInicial()">Salvar PIN</button></div>';
  }
  if(u.pin==="0000"){
    h+='<div class="card noprint" style="border-left:4px solid var(--warn)"><b style="color:var(--warn)">🔑 Você está usando o PIN padrão 0000</b><div class="s">Troque agora em Segurança, abaixo.</div></div>';
  }
  h+='<div class="card"><div style="display:flex;gap:12px;align-items:flex-start">'
    +'<div class="foto" onclick="document.getElementById(\'myfoto\').click()">'
    +(u.foto?'<img src="'+u.foto+'" alt="">':'<span style="font-size:22px">📷</span>')
    +'</div>'
    +'<div style="flex:1"><h4 style="margin:0 0 4px;font-size:16px">'+esc(u.nome)+'</h4>'
    +'<div class="s">'+esc(cg?cg.nome:"—")+(u.mat?' · mat. '+esc(u.mat):'')+'</div>'
    +'<div class="s">'+esc(u.email||u.telefone||"—")+'</div>'
    +'</div></div>'
    +'<input type="file" id="myfoto" accept="image/*" style="display:none" onchange="fotoMeuPerfil(this)">'
    +'</div>';

  if(incompleto && cg && cg.id=="colab" && !u.colId){
    h+='<div class="card"><b>🔗 Vincule seu cadastro ao efetivo</b>'
      +'<div class="s">Procure seu nome na lista e aproveite as informações já cadastradas.</div>'
      +'<input id="bcol" placeholder="🔍 Digite seu nome ou matrícula" oninput="renderBuscaColab(this.value)" autocomplete="off">'
      +'<div id="rescol" style="max-height:340px;overflow-y:auto;margin-top:8px"></div>'
      +'<button class="btn alt" onclick="pularVinculo()" style="margin-top:10px">Não encontrei meu nome — preencher manualmente</button>'
      +'</div>';
  }

  h+='<div class="card"><b>📝 Dados pessoais</b>'
    +'<input value="'+esc(u.nome||"")+'" onchange="setMeuF(\'nome\',this.value)" placeholder="Nome completo">'
    +'<input value="'+esc(u.mat||"")+'" onchange="setMeuF(\'mat\',this.value)" placeholder="Matrícula">'
    +'<input value="'+esc(u.cpf||"")+'" onchange="setMeuF(\'cpf\',this.value)" placeholder="CPF" inputmode="numeric">'
    +'<input value="'+esc(u.telefone||"")+'" onchange="setMeuF(\'telefone\',this.value)" placeholder="Telefone / WhatsApp" inputmode="tel">'
    +'<input value="'+esc(u.email||"")+'" onchange="setMeuF(\'email\',this.value)" placeholder="E-mail" type="email">'
    +'<input value="'+esc(u.nascimento||"")+'" onchange="setMeuF(\'nascimento\',this.value)" placeholder="Nascimento (dd/mm/aaaa)">'
    +'<textarea rows="2" placeholder="Endereço completo" onchange="setMeuF(\'endereco\',this.value)" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:6px 0 8px">'+esc(u.endereco||"")+'</textarea>'
    +'</div>';

  h+='<div class="card"><b>🗣️ '+T("lang_comunic")+'</b>'
    +'<select onchange="setMeuIdioma(this.value)">'
    +IDIOMAS_COMUNICACAO.map(function(x){return '<option value="'+x.id+'"'+(u.idiomaComun==x.id?' selected':'')+'>'+x.ico+' '+x.nome+'</option>'}).join("")
    +'</select>'
    +(u.idiomaComun=="ind"?'<input value="'+esc(u.idiomaIndigena||"")+'" onchange="setMeuF(\'idiomaIndigena\',this.value)" placeholder="'+esc(T("lang_indigenous"))+'">':'')
    +'</div>';

  h+='<div class="card"><b>🏢 Cargo e vínculos</b>'
    +'<div class="row"><span class="s">Cargo</span><b>'+esc(cg?cg.nome:"—")+'</b></div>';
  var ctsIds=u.contratosIds||[];
  if(ctsIds.length){
    h+=ctsIds.map(function(id){
      var ct=ctById(id);
      return '<div class="row"><span class="s">Contrato</span><b>'+esc(ct?ct.codigo+" · "+ct.nome:"(removido)")+'</b></div>';
    }).join("");
  } else {
    h+='<div class="row"><span class="s">Contratos</span><b>—</b></div>';
  }
  h+='<div class="s" style="margin-top:6px">Cargo e contratos são definidos pelo administrador.</div></div>';

  if(u.colId){
    var col=cget(u.colId);
    if(col){
      h+='<div class="card"><b>🔗 Cadastro no efetivo</b>'
        +'<div class="row"><span class="s">Colaborador</span><b>'+esc(col.nome)+'</b></div>'
        +'<div class="row"><span class="s">Função</span><b>'+esc(col.funcao||"—")+'</b></div>'
        +(col.departamento?'<div class="row"><span class="s">Departamento</span><b>'+esc(col.departamento)+'</b></div>':'')
        +(sitExib(col.situacao)?'<div class="row"><span class="s">Situação</span><b style="color:'+sitCor(col.situacao)+'">'+esc(sitExib(col.situacao))+'</b></div>':'')
        +'<button class="btn alt" onclick="desvincularMeuColab()" style="margin-top:8px">Desvincular</button>'
        +'</div>';
    }
  }

  if(!(u.pinTrocar))h+=wpCard();
  h+='<div class="card"><b>🔑 Segurança</b>'
    +'<button class="btn alt" onclick="telaAlterarPin()">Alterar meu PIN</button>'
    +'<div class="s" style="margin-top:8px">Último acesso: '+new Date(u.ultimoAcesso||u.criado||Date.now()).toLocaleString("pt-BR")+'</div>'
    +'</div>';

  if(incompleto){
    h+='<div class="card noprint"><button class="btn" onclick="concluirPerfil()">✅ Concluir e entrar no sistema</button></div>';
  }
  return h;
}

function setMeuF(k,v){
  var u=userById(session.userId);if(!u)return;
  var val=(v||"").trim();
  if(k=="nome"){
    if(!val){alert("O nome não pode ficar vazio.");draw();return}
    var dup=userByName(val);if(dup&&dup.id!=u.id){alert("Já existe um usuário com este nome.");draw();return}
    user=val;
  }
  u[k]=val;saveUsers();
  logAcao("Editou perfil",k);
  if(k=="nome")updateUserInfo();
}
function setMeuIdioma(v){
  var u=userById(session.userId);if(!u)return;
  u.idiomaComun=v;
  if(v=="ind"&&!u.idiomaIndigena){var q=prompt(T("lang_indigenous"),"");u.idiomaIndigena=(q||"").trim()}
  if(v!="ind")u.idiomaIndigena="";
  saveUsers();logAcao("Alterou idioma de comunicação");draw();
}
function fotoMeuPerfil(inp){
  var f=inp.files[0];if(!f)return;
  shrink(f,function(d){
    var u=userById(session.userId);if(!u)return;
    u.foto=d;saveUsers();logAcao("Alterou foto de perfil");draw();
  });
}
function renderBuscaColab(term){
  term=(term||"").trim().toLowerCase();
  var el=document.getElementById("rescol");if(!el)return;
  if(term.length<2){el.innerHTML='<div class="s">Digite pelo menos 2 letras.</div>';return}
  var jaUsados=USERS.filter(function(x){return x.colId}).map(function(x){return x.colId});
  var lista=COL.filter(function(c){
    return jaUsados.indexOf(c.id)<0&&((c.nome||"").toLowerCase().indexOf(term)>=0||(c.matricula||"").indexOf(term)>=0);
  }).slice(0,20);
  if(!lista.length){el.innerHTML='<div class="s">Nenhum colaborador encontrado.</div>';return}
  /* Por privacidade, a busca não exibe mais CPF de terceiros */
  el.innerHTML=lista.map(function(c){
    return '<div class="row" style="padding:8px 0;cursor:pointer" onclick="vincularMeuColab('+c.id+')">'
      +'<span style="flex:1"><b>'+esc(c.nome)+'</b>'
      +'<div class="s">'+esc(c.funcao||"")+(c.matricula?' · mat. '+esc(c.matricula):'')+'</div>'
      +'</span><button class="pill ok" style="border:0">Sou eu</button></div>';
  }).join("");
}
function vincularMeuColab(colId){
  var u=userById(session.userId);if(!u)return;
  var c=cget(colId);if(!c)return;
  if(USERS.some(function(x){return x.colId==colId&&x.id!=u.id})){alert("Este cadastro já está vinculado a outro usuário.");return}
  if(!confirm("Confirmar que você é "+c.nome+"?"))return;
  u.colId=colId;cid=colId;sv("rad_cid",cid);
  if(!u.mat&&c.matricula)u.mat=c.matricula;
  if(!u.cpf&&c.cpf)u.cpf=c.cpf;
  if(!u.endereco&&c.endereco){
    u.endereco=(c.endereco||"")+(c.numero?", "+c.numero:"")+(c.bairro?" · "+c.bairro:"")+(c.cidade?" · "+c.cidade+(c.uf?"/"+c.uf:""):"");
  }
  if(!u.nascimento&&c.nascimento)u.nascimento=c.nascimento;
  if(!u.telefone&&c.telefone)u.telefone=c.telefone;
  saveUsers();
  logAcao("Vinculou-se ao efetivo",c.nome);
  draw();
  alert("Cadastro vinculado! Confira seus dados e clique em Concluir.");
}
function desvincularMeuColab(){
  var u=userById(session.userId);if(!u)return;
  if(!confirm("Desvincular do cadastro do efetivo?"))return;
  u.colId=null;cid="";sv("rad_cid",cid);saveUsers();logAcao("Desvinculou do efetivo");draw();
}
function pularVinculo(){
  var u=userById(session.userId);if(!u)return;
  if(!confirm("Continuar sem vincular ao efetivo? Você poderá vincular depois."))return;
  var bc=document.getElementById("bcol");if(bc)bc.blur();
  var el=document.getElementById("rescol");if(el)el.innerHTML='<div class="s">Preencha seus dados abaixo e clique em Concluir.</div>';
}
/* CORREÇÃO: "pular vínculo" liberava o sistema sem exigir telefone/e-mail; agora sempre passa por Concluir */
function concluirPerfil(){
  var u=userById(session.userId);if(!u)return;
  if(u.pinTrocar){alert("Defina seu PIN pessoal primeiro.");return}
  if(!u.nome){alert("Informe seu nome.");return}
  if(!u.telefone&&!u.email){alert("Informe telefone ou e-mail para contato.");return}
  u.perfilCompleto=true;saveUsers();
  logAcao("Concluiu primeiro acesso");
  alert("Perfil completo! Bem-vindo.");
  startApp();
}
/* ===================================================================
   TELA: ESCALA (Configurações → Escala)
   =================================================================== */
function vEscala(){
  if(!canEdit())return '<div class="card s">Sem permissão.</div>';
  var h='<div class="card"><b>📅 Escalas da equipe</b><div class="s">'+COL.length+' colaborador(es)</div></div>';
  if(!COL.length)return h+'<div class="card s">Nenhum colaborador cadastrado.</div>';
  ESC.forEach(function(e){
    var lista=COL.filter(function(c){return (c.escala||"6X1")==e});
    if(!lista.length)return;
    h+='<div class="card"><b>'+esc(e)+'</b> <span class="s">· '+lista.length+'</span>';
    h+=lista.map(function(c){
      return '<div class="row"><span style="flex:1"><b>'+esc(c.nome)+'</b><div class="s">'+esc(c.funcao||"")+'</div></span>'
        +'<select style="width:auto;margin:0" onchange="setEscalaColab('+c.id+',this.value)">'
        +ESC.map(function(x){return '<option'+(x==(c.escala||"6X1")?' selected':'')+'>'+esc(x)+'</option>'}).join("")
        +'</select></div>';
    }).join("");
    h+='</div>';
  });
  return h;
}
function setEscalaColab(id,v){
  if(!canEdit()){alert("Sem permissão.");return}
  var c=cget(id);if(!c)return;
  c.escala=v;svCol();logAcao("Mudou escala",c.nome,v);draw();
}

/* ===================================================================
   TELA: ESTOQUE
   =================================================================== */
function podeEstoque(){return canEdit()||temAcao("gerenciar_contratos")}
function estAddUI(){
  if(!podeEstMov()){alert("Sem permissão.");return}
  var nome=$("#est_nome").value;
  if(estDoCt(esCt).some(function(x){return x.nome.toLowerCase()==nome.trim().toLowerCase()})){alert("Item já cadastrado neste contrato.");return}
  if(estAdd(nome,$("#est_unid").value,$("#est_qtd").value,$("#est_min").value,$("#est_local").value)){
    logAcao("Cadastrou item de estoque",nome);draw();
  }
}
function estEntradaUI(id){
  if(!podeEstMov()){alert("Sem permissão.");return}
  var q=prompt("Quantidade de entrada:");if(q===null)return;
  q=+String(q).replace(",",".");if(!q||q<=0){alert("Valor inválido.");return}
  var m=prompt("Motivo (ex: compra, devolução):","Compra");if(m===null)return;
  estMov(id,q,m||"Entrada");logAcao("Entrada de estoque","+"+q);draw();
}
function estBaixaUI(id){
  if(!podeEstMov()){alert("Sem permissão.");return}
  var it=EST.filter(function(x){return x.id==id})[0];if(!it)return;
  var q=prompt("Quantidade de baixa (disponível: "+it.qtd+"):");if(q===null)return;
  q=+String(q).replace(",",".");if(!q||q<=0){alert("Valor inválido.");return}
  var m=prompt("Motivo:","Consumo");if(m===null)return;
  if(estBaixa(id,q,m||"Consumo")){logAcao("Baixa de estoque",it.nome+" -"+q);draw()}
}
function estHistUI(id){
  var it=EST.filter(function(x){return x.id==id})[0];if(!it)return;
  var d=document.createElement("div");
  d.style.cssText="position:fixed;inset:20px;z-index:99;overflow-y:auto;background:var(--card);color:var(--ink);border:1px solid var(--line);border-radius:12px;padding:16px;cursor:pointer;box-shadow:0 10px 40px rgba(0,0,0,.3)";
  d.onclick=function(){d.remove()};
  d.innerHTML='<b>📜 '+esc(it.nome)+'</b><div class="s">Toque para fechar</div>'
    +(it.mov||[]).slice().reverse().map(function(m){
      return '<div class="row"><span class="s" style="flex:1">'+new Date(m.t).toLocaleString("pt-BR")+' · '+esc(m.m||"")+' · '+esc(m.u||"")+'</span>'
        +'<b style="color:'+(m.d>=0?'var(--ok)':'var(--warn)')+'">'+(m.d>=0?'+':'')+m.d+'</b></div>';
    }).join("");
  document.body.appendChild(d);
}

/* ===================================================================
   TELA: REQUISIÇÕES
   =================================================================== */
function vReq(){
  var podeGer=podeEstoque();
  var lista=REQ.slice().sort(function(a,b){return b.ini-a.ini});
  if(!podeGer)lista=lista.filter(function(r){return r.colab===(user||"")});

  var h='<div class="card"><b>🛒 '+T("menu_req")+'</b><div class="s">'+lista.filter(function(r){return r.st=="Pendente"}).length+' pendente(s) · '+lista.length+' total</div></div>';

  if(lista.length){
    h+='<div class="card">'+lista.map(function(r){
      var cl={"Pendente":"run","Aprovada":"ok","Rejeitada":"no","Entregue":"ok"}[r.st]||"pen";
      return '<div style="border-top:1px solid var(--line);padding:10px 0">'
        +'<div class="row" style="border:0;padding:0"><b style="flex:1">'+esc(r.itemNome)+'</b>'
        +'<span class="pill '+cl+'" style="cursor:default">'+esc(r.st)+'</span></div>'
        +'<div class="s">'+r.qtd+' '+esc(r.unidade)+' · por '+esc(r.colab||"?")+' · '+new Date(r.ini).toLocaleString("pt-BR")+'</div>'
        +(r.obs?'<div class="s">📝 '+esc(r.obs)+'</div>':'')
        +'<div class="s">Destino: '+esc(setor(r.setor).nome)+'</div>'
        +(r.hist&&r.hist.length?'<div class="s" style="margin-top:4px">'+r.hist.map(function(x){return new Date(x.t).toLocaleString("pt-BR")+' · '+esc(x.x)}).join("<br>")+'</div>':'')
        +(podeGer&&r.st=="Pendente"?'<div class="chips" style="margin-top:6px">'
          +'<button class="sel" onclick="reqAcao('+r.id+',\'Aprovada\')">✅ Aprovar</button>'
          +'<button onclick="reqAcao('+r.id+',\'Rejeitada\')">❌ Rejeitar</button></div>':'')
        +(podeGer&&r.st=="Aprovada"?'<div class="chips" style="margin-top:6px"><button onclick="reqAcao('+r.id+',\'Entregue\')">📦 Marcar entregue</button></div>':'')
        +'</div>';
    }).join("")+'</div>';
  }else h+='<div class="card s">Nenhuma requisição.</div>';

  /* CORREÇÃO: o formulário de nova requisição ficava inacessível quando não havia requisições */
  h+='<div class="card"><b>➕ Nova requisição</b>';
  if(!EST.length){h+='<div class="s">Cadastre itens no estoque primeiro.</div>'}
  else{
    h+='<select id="req_item">'+EST.map(function(it){return '<option value="'+it.id+'">'+esc(it.nome)+' ('+it.qtd+' '+esc(it.unidade)+')</option>'}).join("")+'</select>'
      +'<div class="s">Quantidade</div><input id="req_qtd" type="number" min="1" value="1">'
      +'<input id="req_obs" placeholder="Observação">'
      +'<select id="req_setor">'+SET.map(function(s){return '<option value="'+s.id+'"'+(s.id=="compras"?' selected':'')+'>'+esc(s.nome)+'</option>'}).join("")+'</select>'
      +'<button class="btn" onclick="reqNovaUI()">Enviar requisição</button>';
  }
  return h+'</div>';
}
function reqAcao(id,v){
  if(!podeEstoque()){alert("Sem permissão.");return}
  reqSt(id,v);logAcao("Requisição "+v.toLowerCase());draw();
}
function reqNovaUI(){
  if(reqNova(+$("#req_item").value,+$("#req_qtd").value,$("#req_obs").value,$("#req_setor").value)){
    logAcao("Criou requisição");draw();
  }
}

/* ===================================================================
   TELA: RELATÓRIO
   =================================================================== */

/* ===================================================================
   TIPOS DE SERVIÇO (cada contrato tem um tipo → checklists e modelos)
   =================================================================== */
var TIPOS_CT=[
  {id:"hosp",ico:"🏥",nome:"Hospitalar",desc:"hospitais, clínicas, medicina, odontologia, veterinária"},
  {id:"corp",ico:"🏢",nome:"Corporativo",desc:"escritórios, sedes administrativas"},
  {id:"edu",ico:"🎓",nome:"Educacional",desc:"escolas, faculdades, faculdades de medicina"},
  {id:"ind",ico:"🏭",nome:"Industrial",desc:"fábricas, galpões, plantas"},
  {id:"vig",ico:"🛡️",nome:"Vigilância / Portaria",desc:"postos, rondas, controle de acesso"},
  {id:"fac",ico:"🏗️",nome:"Facilities",desc:"multi-serviço, condomínios comerciais"},
  {id:"var",ico:"🏬",nome:"Varejo / Loja",desc:"lojas, shoppings, supermercados"},
  {id:"joa",ico:"💎",nome:"Joalheria",desc:"lojas de alto padrão, vitrines, cofres"},
  {id:"ban",ico:"🏦",nome:"Banco / Financeiro",desc:"agências, cofres, tesouraria"},
  {id:"out",ico:"📌",nome:"Outro",desc:"não classificado ainda"}
];
function tipoCt(id){return TIPOS_CT.filter(function(t){return t.id==id})[0]||TIPOS_CT[TIPOS_CT.length-1]}
function tipoCtNome(id){var t=tipoCt(id);return t.ico+" "+t.nome}
/* Um contrato pode ter vários tipos de serviço (ex.: hospital com limpeza + portaria) */
function tiposCt(c){if(!c)return [];var t=(c.tipos&&c.tipos.length)?c.tipos:[c.tipo||"out"];return t.filter(function(x,i){return t.indexOf(x)==i})}
function tiposCtNome(c){return tiposCt(c).map(tipoCtNome).join(" · ")}
function tiposCtIco(c){return tiposCt(c).map(function(t){return tipoCt(t).ico}).join("")}
function ctTemTipo(c,t){return tiposCt(c).indexOf(t)>=0}
function selTipoCt(c){
  var at=tiposCt(c);
  return '<div class="s">Toque para marcar todos os serviços prestados neste contrato:</div><div>'+TIPOS_CT.map(function(t){return chip(at.indexOf(t.id)>=0,"togTipoCt("+c.id+",\'"+t.id+"\')",t.ico+" "+t.nome)}).join("")+'</div>';
}
function togTipoCt(id,t){
  var c=ctById(id);if(!c)return;
  if(!podeAreasCt(id)&&!temAcao("gerenciar_contratos")){alert("Sem permissão.");draw();return}
  var at=tiposCt(c).slice(),p=at.indexOf(t);
  if(p>=0){
    if(at.length==1){alert("O contrato precisa de pelo menos um tipo de serviço.");return}
    var usando=areasDoContrato(id).filter(function(i){return LOC[i].tipo==t}).length;
    if(usando&&!confirm(usando+" área(s) usam "+tipoCt(t).nome+". Elas passarão para "+tipoCt(at.filter(function(x){return x!=t})[0]).nome+". Continuar?"))return;
    at.splice(p,1);
    areasDoContrato(id).forEach(function(i){if(LOC[i].tipo==t)LOC[i].tipo=at[0]});svLoc();
  }else{
    if(at.length==1&&at[0]=="out"&&!areasDoContrato(id).length)at=[];
    at.push(t);
  }
  c.tipos=at;c.tipo=at[0];svCt();logAcao("Tipos de serviço do contrato",c.codigo,at.map(function(x){return tipoCt(x).nome}).join(", "));
  buildNav();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function setTipoArea(i,t){if(!arChk(i))return;LOC[i].tipo=t;svLoc();draw()}
function tipoDaArea(i){var L=LOC[i];var c=L&&ctById(L.ct);if(!c)return "";var at=tiposCt(c);return at.indexOf(L.tipo)>=0?L.tipo:at[0]}

/* Modelos por tipo: áreas sugeridas (área → espaços), atividades e checklist próprio.
   c = criticidade do item: C crítico · S semicrítico · N não crítico */
var MODELOS={
  hosp:{
    areas:[["Recepção e espera",["Balcão de atendimento","Sala de espera","Banheiro público M","Banheiro público F","Banheiro PCD"]],["Consultórios",["Consultório 1","Consultório 2","Consultório 3"]],["Centro cirúrgico",["Sala cirúrgica 1","Sala cirúrgica 2","Recuperação","Expurgo","CME"]],["Internação / Enfermaria",["Enfermaria","Posto de enfermagem","Isolamento"]],["Apoio",["DML","Copa","Vestiário","Abrigo de resíduos"]]],
    ativ:["Limpeza concorrente","Limpeza terminal","Desinfecção de superfícies de contato","Recolhimento de resíduos infectantes","Recolhimento de resíduos comuns","Limpeza de sanitários","Reposição de insumos"],
    chk:null
  },
  corp:{
    areas:[["Recepção",["Balcão","Sala de espera"]],["Escritórios",["Estações de trabalho","Sala de reunião","Diretoria"]],["Copa",["Copa","Refeitório"]],["Sanitários",["Banheiro M","Banheiro F","Banheiro PCD"]],["Áreas comuns",["Corredores","Escadas","Elevadores","Hall"]]],
    ativ:["Limpeza de pisos","Limpeza de mesas e estações","Higienização de banheiros","Limpeza da copa","Recolhimento de lixo","Reposição de insumos","Limpeza de vidros"],
    chk:{nome:"Inspeção Corporativo",itens:[["Pisos","N"],["Mesas e estações de trabalho","S",["Tampo","Teclado e mouse","Telefone"]],["Salas de reunião","S"],["Banheiros","C",["Vaso e mictório","Pia e torneira","Espelho","Dispensadores"]],["Copa","S",["Bancada","Pia","Geladeira e micro-ondas"]],["Lixeiras","S"],["Vidros e portas","N"],["Reposição de insumos","S"]]}
  },
  edu:{
    areas:[["Salas de aula",["Sala 1","Sala 2","Sala 3"]],["Laboratórios",["Laboratório de informática","Laboratório de ciências","Laboratório de anatomia"]],["Biblioteca",["Salão de leitura","Acervo"]],["Sanitários",["Banheiro M","Banheiro F","Banheiro PCD"]],["Áreas comuns",["Pátio","Corredores","Cantina","Sala dos professores"]]],
    ativ:["Limpeza de salas de aula","Limpeza de carteiras e quadros","Limpeza de laboratórios","Higienização de banheiros","Limpeza de pátio","Recolhimento de lixo","Reposição de insumos"],
    chk:{nome:"Inspeção Educacional",itens:[["Pisos das salas","N"],["Carteiras e mesas","S",["Tampo","Assento","Estrutura"]],["Quadros e lousas","N"],["Laboratórios","C",["Bancadas","Pias","Descarte de resíduos"]],["Banheiros","C",["Vaso e mictório","Pia e torneira","Dispensadores"]],["Bebedouros","C"],["Cantina / refeitório","C"],["Lixeiras","S"],["Pátio e áreas externas","N"]]}
  },
  ind:{
    areas:[["Produção",["Linha de produção","Área de máquinas"]],["Armazém / Galpão",["Estoque","Doca de carga"]],["Refeitório",["Salão","Cozinha de apoio"]],["Vestiários e sanitários",["Vestiário M","Vestiário F","Banheiros"]],["Administrativo",["Escritório","Sala de reunião"]],["Área externa",["Pátio","Portaria"]]],
    ativ:["Limpeza de pisos industriais","Remoção de óleo e resíduos","Limpeza de vestiários","Higienização do refeitório","Coleta seletiva","Limpeza de escritórios"],
    chk:{nome:"Inspeção Industrial",itens:[["Pisos sem óleo ou resíduos","C"],["Faixas e sinalização visíveis","S"],["Uso de EPI pela equipe","C"],["Refeitório","C",["Mesas","Bancadas","Pisos"]],["Vestiários","S",["Armários","Chuveiros","Pisos"]],["Banheiros","C"],["Coleta seletiva e descarte","S"],["Áreas externas","N"]]}
  },
  vig:{
    areas:[["Portaria principal",["Guarita","Recepção de visitantes","Controle de veículos"]],["Rondas",["Perímetro","Estacionamento","Áreas internas"]],["Central / CFTV",["Sala de monitoramento"]]],
    ativ:["Controle de acesso","Registro de visitantes","Ronda perimetral","Ronda interna","Registro de ocorrências","Abertura e fechamento","Monitoramento CFTV"],
    chk:{nome:"Inspeção de Posto",itens:[["Vigilante presente no posto","C"],["Uniforme e identificação","S"],["Livro de ocorrências atualizado","C"],["Rádio / comunicação operante","C"],["Controle de acesso e registro de visitantes","C"],["Ronda realizada conforme roteiro","C"],["Equipamentos do posto (lanterna, detector)","S"],["Iluminação e limpeza do posto","N"],["CFTV gravando","S"]]}
  },
  fac:{
    areas:[["Hall e recepção",["Hall de entrada","Recepção"]],["Garagem",["Garagem","Rampa"]],["Circulação",["Elevadores","Escadas","Corredores"]],["Sanitários comuns",["Banheiro M","Banheiro F"]],["Áreas externas",["Jardins","Fachada","Lixeira central"]],["Manutenção",["Casa de máquinas","Sala de manutenção"]]],
    ativ:["Limpeza de áreas comuns","Limpeza de elevadores","Limpeza de garagem","Jardinagem","Manutenção predial","Recolhimento de lixo","Limpeza de vidros e fachada"],
    chk:{nome:"Inspeção Facilities",itens:[["Hall e recepção","S"],["Elevadores","S",["Botoeiras","Espelho","Piso","Portas"]],["Escadas e corrimãos","S"],["Garagem","N"],["Banheiros comuns","C"],["Jardins","N"],["Iluminação funcionando","S"],["Vazamentos / manutenção","C"],["Lixeira central","S"]]}
  },
  var:{
    areas:[["Salão de vendas",["Corredores","Gôndolas / araras","Entrada"]],["Vitrines",["Vitrine principal"]],["Provadores",["Provadores"]],["Caixas",["Frente de caixa","Balcão"]],["Estoque",["Estoque","Doca"]],["Clientes",["Banheiro clientes","Praça de alimentação"]]],
    ativ:["Limpeza do salão","Limpeza de vitrines","Limpeza de provadores","Limpeza dos caixas","Higienização de carrinhos e cestas","Limpeza de banheiros de clientes","Recolhimento de lixo"],
    chk:{nome:"Inspeção Varejo",itens:[["Entrada e fachada","S"],["Pisos do salão","S"],["Vitrines sem marcas","N"],["Provadores","S"],["Caixas e balcões","S",["Esteira","Balcão","Leitor / máquina de cartão"]],["Carrinhos e cestas","S"],["Banheiros de clientes","C"],["Áreas de alimentos (açougue, câmaras, praça)","C"],["Lixeiras","S"]]}
  },
  joa:{
    areas:[["Salão de exposição",["Vitrines","Expositores","Balcões de atendimento"]],["Atendimento reservado",["Sala VIP"]],["Área restrita",["Cofre","Sala técnica"]],["Apoio",["Copa","Banheiro","DML"]]],
    ativ:["Limpeza de vitrines e vidros","Limpeza de expositores","Limpeza de balcões","Limpeza da sala VIP","Limpeza de área restrita (acompanhada)","Limpeza de pisos"],
    chk:{nome:"Inspeção Joalheria",itens:[["Vitrines de vidro sem marcas","C",["Face frontal","Laterais","Tampo"]],["Expositores e bandejas","S"],["Balcões de atendimento","S"],["Espelhos","C"],["Iluminação das vitrines","S"],["Sala VIP","S"],["Área do cofre (limpeza acompanhada)","C"],["Pisos","N"],["Banheiro","S"]]}
  },
  ban:{
    areas:[["Autoatendimento",["Terminais (ATM)","Hall de autoatendimento"]],["Agência",["Atendimento / guichês","Espera","Gerência"]],["Área restrita",["Tesouraria","Cofre"]],["Apoio",["Copa","Banheiros","DML"]]],
    ativ:["Higienização de terminais de autoatendimento","Limpeza de guichês","Limpeza de porta giratória","Limpeza da área de espera","Limpeza de tesouraria (acompanhada)","Limpeza de banheiros","Recolhimento de lixo"],
    chk:{nome:"Inspeção Bancária",itens:[["Terminais de autoatendimento","C",["Tela","Teclado","Leitor de cartão","Bocal"]],["Porta giratória","S"],["Guichês e balcões","S"],["Área de espera (cadeiras, senha)","S"],["Tesouraria / cofre (acompanhada)","C"],["Banheiros","C"],["Pisos","N"],["Lixeiras","S"]]}
  },
  out:{
    areas:[],
    ativ:["Limpeza de pisos","Limpeza de superfícies","Limpeza de sanitários","Recolhimento de lixo","Reposição de insumos"],
    chk:{nome:"Inspeção Geral",itens:[["Pisos","N"],["Superfícies de contato","S"],["Banheiros","C"],["Lixeiras","S"],["Vidros","N"],["Reposição de insumos","S"]]}
  }
};
function ensureModelChecklists(){
  var mud=false;
  CHECKLISTS.forEach(function(c){
    if(!c.tipos){c.tipos=(c.id=="lz"||c.id=="rc016")?["hosp"]:["*"];mud=true}
  });
  TIPOS_CT.forEach(function(t){
    var m=MODELOS[t.id];if(!m||!m.chk)return;
    var id="mdl_"+t.id,itens=m.chk.itens.map(function(x){var o={n:x[0],c:x[1]};if(x[2])o.p=x[2];return o});
    var c=chklById(id);
    if(!c){CHECKLISTS.push({id:id,nome:m.chk.nome,cat:t.nome,ico:t.ico,tipos:[t.id],ativo:true,fixed:true,itens:itens});mud=true}
    else c.itens=itens;
  });
  if(mud)svChkl();
}
function chklServeTipo(c,tipo){var t=c.tipos||["*"];return t.indexOf("*")>=0||!tipo||t.indexOf(tipo)>=0}

/* ===================================================================
   ÁREAS POR CONTRATO
   =================================================================== */
var CT_ATUAL=ld("rad_ctAtual",null);
function areasDoContrato(ctId){var r=[];LOC.forEach(function(L,i){if(!L.del&&L.ct==ctId)r.push(i)});return r}
function locaisCt(c){return c?areasDoContrato(c.id):[]}
function areasSemContrato(){var r=[];LOC.forEach(function(L,i){if(!L.del&&(L.ct==null||!ctById(L.ct)))r.push(i)});return r}
function ehColabCargo(){var cg=cargoAtual();return cg&&cg.id=="colab"}
function meusContratos(){
  var u=session?userById(session.userId):null;if(!u)return [];
  var cg=cargoAtual();
  if(cg&&(cg.fixo||cg.id=="admin"))return CONTRATOS.slice();
  if(cg&&cg.id=="colab"){
    var col=u.colId?cget(u.colId):null;if(!col)return [];
    return (col.vinculos||[]).map(function(v){return ctById(v.contratoId)}).filter(Boolean);
  }
  return contratosDoUsuario(u);
}
function ctAtualId(){
  var m=meusContratos();
  if(m.some(function(c){return c.id==CT_ATUAL}))return CT_ATUAL;
  return m.length?m[0].id:null;
}
function setCtAtual(v){CT_ATUAL=v?+v:null;sv("rad_ctAtual",CT_ATUAL);loc=-1;draw()}
function vinculoMeu(ctId){
  var u=session?userById(session.userId):null;var col=u&&u.colId?cget(u.colId):null;
  return col?(col.vinculos||[]).filter(function(v){return v.contratoId==ctId})[0]||null:null;
}
function areasPermitidas(){
  var ct=ctAtualId();
  var cg=cargoAtual();
  var lista=ct?areasDoContrato(ct):((cg&&(cg.fixo||cg.id=="admin"))?areasSemContrato():[]);
  lista=lista.filter(function(i){return !LOC[i].off});
  if(ehColabCargo()){
    var v=ct?vinculoMeu(ct):null;
    var lib=v?(v.areas||[]):[];
    lista=lista.filter(function(i){return lib.indexOf(i)>=0});
  }
  return lista;
}
function ajustaLoc(){var p=areasPermitidas();if(p.indexOf(+loc)<0)loc=p.length?p[0]:-1}
/* Escolha de contrato com barra de busca: mostra o atual, e ao digitar lista os contratos encontrados */
var PICK_FN={};
function pickContrato(k,lista,atualId,fn){
  PICK_FN[k]=fn;
  var at=lista.filter(function(c){return c.id==atualId})[0];
  var h='<div class="s">Contrato'+(lista.length>1?' ('+lista.length+' disponíveis)':'')+'</div>'
    +'<div style="font-weight:700;padding:6px 0">'+(at?esc(tiposCtIco(at)+" "+at.codigo+" · "+at.nome):'—')+(at?'<div class="s" style="font-weight:400">👔 Supervisor: '+esc(nomesSupCt(at.id)||"não definido")+'</div>':'')+'</div>';
  if(lista.length<2)return h;
  h+=campoBusca(k,"Buscar contrato por código, nome ou supervisor");
  if(termo(k)){
    var r=lista.filter(function(c){return c.id!=atualId&&bate(k,c.codigo,c.nome,c.responsavel,nomesSupCt(c.id))});
    h+=(r.length?r.slice(0,15).map(function(c){return '<div class="row" style="cursor:pointer" onclick="escolherCt(\''+k+'\','+c.id+')"><span style="flex:1">'+esc(tiposCtIco(c))+' '+marca(c.codigo,k)+' · '+marca(c.nome,k)+'<div class="s">👔 '+(nomesSupCt(c.id)?marca(nomesSupCt(c.id),k):'sem supervisor')+' · '+areasDoContrato(c.id).length+' área(s)</div></span><span class="pill ok">Abrir</span></div>'}).join("")+(r.length>15?'<div class="s">+'+(r.length-15)+' resultado(s). Refine a busca.</div>':''):'<div class="s">Nenhum contrato encontrado.</div>');
  }
  return h;
}
function escolherCt(k,id){BUSCA[k]="";var f=PICK_FN[k];if(f)f(id)}
function ctSelCard(){
  var m=meusContratos();if(m.length<2)return "";
  var at=ctAtualId();
  return '<div class="card noprint">'+pickContrato("ctsel",m,at,setCtAtual)+(m.some(function(c){return c.geo})?'<button class="btn alt" onclick="geoMeuContrato()">📍 Usar minha localização</button>':'')+'</div>';
}
function semAreasCard(){
  if(ehColabCargo()){
    var u=userById(session.userId);
    if(!u||!u.colId)return '<div class="card"><b>Sem vínculo com o efetivo</b><div class="s">Vincule seu cadastro em 👤 Meu Perfil para ver suas áreas.</div></div>';
    return '<div class="card"><b>Nenhuma área liberada</b><div class="s">Seu supervisor ainda não liberou áreas de atividade para você neste contrato.</div></div>';
  }
  return '<div class="card"><b>Nenhuma área mapeada</b><div class="s">Este contrato ainda não tem áreas. Crie em 🗺️ Áreas do Contrato.</div>'+(acessosDoUsuario().indexOf("ar")>=0?'<button class="btn" onclick="go(\'ar\')">🗺️ Mapear áreas</button>':'')+'</div>';
}
function podeAreasCt(ctId){
  var cg=cargoAtual();if(!cg)return false;
  if(cg.fixo||cg.id=="admin")return true;
  if(!temAcao("editar_areas"))return false;
  var u=userById(session.userId);
  return contratosDoUsuario(u).some(function(c){return c.id==ctId});
}

/* ===================================================================
   TELA: ÁREAS DO CONTRATO (supervisor mapeia e libera acesso)
   =================================================================== */
var arCt=ld("rad_arCt",null),arAberta=-1;
function selArCt(id){arCt=id;sv("rad_arCt",arCt);arAberta=-1;copiaAberta=false;draw();window.scrollTo(0,0)}
function abrirAreasCt(id){arCt=id;sv("rad_arCt",arCt);go("ar")}
function vAreas(){
  var m=meusContratos().filter(function(c){return podeAreasCt(c.id)});
  if(!m.length)return '<div class="card"><b>🗺️ '+T("menu_ar")+'</b><div class="s">Nenhum contrato sob sua responsabilidade. Os contratos vêm dos departamentos do efetivo.</div></div>';
  if(!m.some(function(c){return c.id==arCt})){arCt=m[0].id}
  var c=ctById(arCt),tps=tiposCt(c);
  var areas=areasDoContrato(c.id);
  var h='<div class="card"><b>🗺️ '+T("menu_ar")+'</b><div class="s">Mapeie as áreas e espaços do contrato e libere para cada colaborador as áreas onde ele atua.</div>'
    +'<div style="margin-top:8px">'+pickContrato("arct",m,c.id,selArCt)+'</div>'
    +'<div class="s" style="margin-top:8px"><b>Tipos de serviço</b></div>'+selTipoCt(c)
    +tps.map(function(t){return '<div class="s">'+tipoCt(t).ico+' Checklists de '+esc(tipoCt(t).nome)+': '+(CHECKLISTS.filter(function(k){return k.ativo&&chklServeTipo(k,t)}).map(function(k){return esc(k.nome)}).join(" · ")||"nenhum")+'</div>'}).join("")+'</div>';

  /* modelo */
  var soPadrao=areas.length&&areas.every(function(i){return LOC[i].padrao});
  if(soPadrao)h+='<div class="card" style="border-left:4px solid var(--run)"><b>Áreas com nome padrão</b><div class="s">Este contrato recebeu '+areas.length+' áreas com nome padrão. Toque em ✏️ Editar em cada uma para colocar os nomes reais, ou use um modelo abaixo para substituí-las.</div></div>';
  var comModelo=tps.filter(function(t){return MODELOS[t]&&MODELOS[t].areas.length&&!areas.some(function(i){return LOC[i].modelo==t})});
  if(comModelo.length){
    h+='<div class="card" style="border-left:4px solid var(--brand)"><b>Modelos de áreas</b><div class="s">Cria as áreas, espaços e atividades sugeridos para cada serviço'+(soPadrao?', no lugar das áreas de nome padrão':'')+'. Depois é só ajustar ao contrato real.</div>'
      +comModelo.map(function(t){var tt=tipoCt(t);return '<div style="border-top:1px solid var(--line);padding:8px 0"><b>'+tt.ico+' '+esc(tt.nome)+'</b><div class="s">'+MODELOS[t].areas.map(function(a){return esc(a[0])}).join(" · ")+'</div><button class="btn alt" onclick="aplicarModelo('+c.id+',\''+t+'\')">Usar modelo '+esc(tt.nome)+'</button></div>'}).join("")+'</div>';
  }
  var orf=areasSemContrato();
  if(orf.length&&temAcao("gerenciar_contratos")){
    h+='<div class="card"><b>Áreas sem contrato ('+orf.length+')</b><div class="s">Áreas da versão anterior. Mova para este contrato as que pertencem a ele.</div>'
      +orf.map(function(i){return '<div class="row"><span>'+esc(LOC[i].n)+' <span class="s">· '+LOC[i].e.filter(function(_,e){return !espDel(i,e)}).length+' espaço(s)</span></span><button class="pill ok" style="border:0" onclick="moverArea('+i+','+c.id+')">Mover</button></div>'}).join("")
      +'<button class="btn alt" onclick="moverTodas('+c.id+')">Mover todas para este contrato</button></div>';
  }

  /* áreas */
  h+='<div class="card"><b>Áreas ('+areas.length+')</b>';
  if(!areas.length)h+='<div class="s">Nenhuma área ainda.</div>';
  var buscando=!!termo("ar"),nVis=0;
  areas.forEach(function(i){
    var L=LOC[i];
    var esps=L.e.map(function(_,e){return e}).filter(function(e){return !espDel(i,e)});
    var ativs=PG.filter(function(g){return g.l==i});
    if(buscando&&!bate("ar",L.n)&&!esps.some(function(e){return bate("ar",L.e[e])})&&!ativs.some(function(g){return bate("ar",g.nome,g.obs)}))return;
    nVis++;
    var ab=arAberta==i||buscando;
    h+='<div style="border-top:1px solid var(--line);padding:10px 0'+(L.off?';opacity:.6':'')+'">'
      +'<div class="row" style="border:0;padding:0;cursor:pointer" onclick="arAberta='+(ab?-1:i)+';draw()"><span style="flex:1"><b>'+marca(L.n,"ar")+'</b>'+(L.padrao?' <span class="pill run" style="cursor:default;font-size:11px">nome padrão</span>':'')+(L.off?' <span class="s">(desativada)</span>':'')+'<div class="s">'+(tps.length>1?tipoCtNome(tipoDaArea(i))+' · ':'')+esps.length+' espaço(s) · '+ativs.length+' atividade(s)</div></span><span class="pill '+(ab?'run':'pen')+'">'+(ab?'▲ Fechar':'✏️ Editar')+'</span></div>';
    if(ab){
      h+='<div class="s" style="margin-top:8px"><b>Nome da área</b></div><input value="'+esc(L.n)+'" onchange="arRenL('+i+',this.value)">'
        +(tps.length>1?'<div class="s">Serviço desta área (define os checklists)</div><select onchange="setTipoArea('+i+',this.value)">'+tps.map(function(t){return '<option value="'+t+'"'+(tipoDaArea(i)==t?' selected':'')+'>'+esc(tipoCtNome(t))+'</option>'}).join("")+'</select>':'')
        +'<div class="s"><b>Espaços</b> (nome · classificação de risco · pausar · excluir)</div><div class="s">A classificação vem pré-definida pelo nome do espaço e o serviço da área. Ajuste se necessário; "(auto)" indica que ainda é a sugestão do sistema.</div>'
        +esps.map(function(e){var off=espOff(i,e);return '<div class="esp-linha'+(off?' off':'')+'"'+(buscando&&bate("ar",L.e[e])?' style="outline:2px solid #ffe08a;border-radius:8px"':'')+'><input value="'+esc(L.e[e])+'" onchange="arRenE('+i+','+e+',this.value)">'+aclSel(i,e)+'<button class="pill '+(off?'pen':'ok')+' mini-btn" onclick="arTogE('+i+','+e+')">'+(off?'▶':'⏸')+'</button><button class="pill no mini-btn" onclick="arDelE('+i+','+e+')">✕</button></div><div class="s" style="margin:-2px 0 8px 4px;font-size:12px">🧤 '+esc(episTxt(i,e))+'</div>'}).join("")
        +'<div style="display:flex;gap:6px;margin-top:6px"><input id="ne_'+i+'" placeholder="Novo espaço" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px" onclick="arAddE('+i+')">Add</button></div>'
        +'<div class="s" style="margin-top:10px"><b>Atividades da área</b> (nome · frequência · como fazer)</div>'
        +ativs.map(function(g){return '<div class="ativ-ed"'+(buscando&&bate("ar",g.nome,g.obs)?' style="outline:2px solid #ffe08a"':'')+'>'
          +'<div class="esp-linha"><input value="'+esc(g.nome)+'" onchange="renPg('+g.id+',this.value)"><button class="pill no mini-btn" onclick="delPg('+g.id+')">✕</button></div>'
          +'<select onchange="setPgCampo('+g.id+',\'freq\',this.value)">'+FREQS.map(function(f){return '<option'+((g.freq||"Diária")==f?' selected':'')+'>'+f+'</option>'}).join("")+'</select>'
          +'<div class="s" style="margin-top:2px">📖 Como fazer'+(g.obsAuto&&g.obs?' <span class="pill ok" style="cursor:default;font-size:11px">padrão</span>':'')+(!g.obsAuto&&comoFazerPadrao(g.nome)?' <a style="cursor:pointer;color:var(--brand);font-size:12px" onclick="comoFazerRestaurar('+g.id+')">↺ voltar ao padrão</a>':'')+'</div>'
          +'<textarea rows="3" placeholder="Descreva como executar esta atividade" onchange="setPgCampo('+g.id+',\'obs\',this.value)" style="width:100%;padding:9px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit;font-size:14px">'+esc(g.obs||"")+'</textarea>'+cfTradEditor(g)+'</div>'}).join("")
        +'<div style="display:flex;gap:6px;margin-top:6px"><input id="na_'+i+'" placeholder="Nova atividade" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px" onclick="arAddAtiv('+i+')">Add</button></div>'
        +((MODELOS[tipoDaArea(i)]||{ativ:[]}).ativ.length?'<button class="btn alt" onclick="arAtivModelo('+i+')">Incluir atividades padrão '+tipoCt(tipoDaArea(i)).ico+'</button>':'')
        +'<button class="btn alt" onclick="reclassificarArea('+i+')">↺ Reclassificar espaços automaticamente</button>'
        +'<div class="grid-acoes"><button class="btn alt" onclick="arTogL('+i+')">'+(L.off?'▶ Reativar':'⏸ Desativar')+'</button><button class="btn red" onclick="arDelL('+i+')">Excluir área</button></div>';
    }
    h+='</div>';
  });
  if(buscando&&!nVis)h+='<div class="s">Nenhuma área, espaço ou atividade encontrada.</div>';
  h+='<div style="display:flex;gap:6px;margin-top:10px"><input id="nl_ar" placeholder="Nova área (ex: 2º andar, Bloco B)" style="margin:0;flex:1"><button class="pill ok" style="border:0;padding:10px 14px" onclick="arAddL('+c.id+')">Add</button></div></div>';

  h+=cardCopiaProg(c)+'<div class="card noprint"><button class="btn alt" onclick="imprimirPTG('+c.id+')">🖨️ Imprimir PTG deste contrato</button></div>';
  h+='<div class="card noprint"><details><summary style="cursor:pointer;font-weight:700">🛡️ Critérios da classificação de risco</summary>'
    +'<div class="s" style="margin-top:6px"><b>Hospitalar</b> (Manual ANVISA de limpeza e desinfecção de superfícies):</div>'
    +'<div class="s">🔴 <b>Crítica</b>: centro cirúrgico, CME, UTI, isolamento, hemodiálise, laboratórios, banco de sangue, lactário, nutrição/cozinha, farmácia, lavanderia (área suja), salas de procedimento, expurgo, resíduos infectantes.</div>'
    +'<div class="s">🟡 <b>Semicrítica</b>: enfermarias, apartamentos, consultórios, ambulatórios, salas de exame, banheiros, posto de enfermagem, recepção e espera de pacientes, corredores, elevadores, DML.</div>'
    +'<div class="s">🟢 <b>Não crítica</b>: áreas administrativas, copa, vestiários, almoxarifado e demais áreas sem pacientes.</div>'
    +'<div class="s" style="margin-top:6px"><b>Demais serviços</b>: crítica para manipulação de alimentos, laboratório de anatomia e resíduos infectantes; semicrítica para sanitários, vestiários, copa/refeitório, bebedouros, provadores, terminais de autoatendimento, elevadores e lixo; não crítica para o restante.</div>'
    +'</details></div>';
  /* acesso dos colaboradores */
  var colabs=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==c.id})&&!ehDeslig(x.situacao)}).sort(function(a,b){return (a.nome||"").localeCompare(b.nome||"")});
  var ativas=areas.filter(function(i){return !LOC[i].off});
  h+='<div class="card"><b>👥 Acesso dos colaboradores ('+colabs.length+')</b><div class="s">Toque nas áreas para liberar ou retirar. O colaborador só vê na Rotina as áreas liberadas.</div>'
    +'<div class="s" style="margin-top:6px">PIN inicial do app: <b>4 últimos dígitos do CPF</b> (ou da matrícula, se não houver CPF). A troca é obrigatória no 1º acesso.</div>';
  if(colabs.length&&ativas.length)h+='<div class="grid-acoes"><button class="btn alt" onclick="liberarTodasAreas('+c.id+')">Liberar todas as áreas para todos</button><button class="btn alt" onclick="liberarAcessoTodos('+c.id+')">📲 Criar acesso ao app para todos</button></div>';
  if(!colabs.length)h+='<div class="s" style="margin-top:8px">Nenhum colaborador vinculado a este contrato.</div>';
  if(buscando){var cf=colabs.filter(function(x){return bate("ar",x.nome,x.funcao,x.matricula)});if(cf.length||!nVis)colabs=cf}
  colabs.forEach(function(x){
    var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==c.id})[0];
    var us=userDoColab(x.id);
    var lib=(v&&v.areas)||[];
    h+='<div style="border-top:1px solid var(--line);padding:10px 0"><div class="row" style="border:0;padding:0;flex-wrap:wrap"><span style="flex:1"><b>'+marca(x.nome,"ar")+'</b> <span class="s">· '+esc(x.funcao||"")+'</span><div class="s">'+lib.filter(function(i){return ativas.indexOf(i)>=0}).length+'/'+ativas.length+' área(s) liberada(s) · '
      +(us?(us.bloqueado?'acesso bloqueado':us.pinTrocar?'acesso criado, aguardando 1º login':'usa o app'):'sem acesso ao app')+'</div></span>'
      +(us?'':'<button class="pill ok" style="border:0" onclick="liberarAcesso('+x.id+')">📲 Criar acesso</button>')+'</div>'
      +'<div>'+ativas.map(function(i){return chip(lib.indexOf(i)>=0,"togAreaColab("+x.id+","+c.id+","+i+")",LOC[i].n)}).join("")+'</div></div>';
  });
  return h+'</div>';
}
var FREQS=["Diária","Mais de uma vez ao dia","Semanal","Quinzenal","Mensal","Eventual / sob demanda"];
function setPgCampo(id,k,v){var g=PG.filter(function(x){return x.id==id})[0];if(!g)return;if(!arChk(g.l))return;g[k]=(v||"").trim();if(k=="obs")g.obsAuto=!g[k];svPg();if(k=="obs"&&!g[k]){var y=window.scrollY;draw();window.scrollTo(0,y)}}
var NOTA_RISCO={
  C:"Área crítica: pano e balde exclusivos desta área, técnica de dois baldes (solução e enxágue) ou MOP, desinfecção obrigatória após a limpeza e higiene das mãos antes e depois.",
  S:"Área semicrítica: limpeza com detergente seguida de desinfecção das superfícies de contato; panos separados por tipo de superfície.",
  N:"Área não crítica: limpeza com detergente neutro; desinfecção quando houver matéria orgânica."
};
function comoFazerCard(l,e){
  if(l<0||!LOC[l])return "";
  var ativs=PG.filter(function(g){return g.l==l});
  var cls=aclOf(l,e),vig=tipoDaArea(l)=="vig";
  if(!ativs.length&&vig)return "";
  var L=langUser(),ui=CF_UI[L],nota=L!="pt-BR"&&NOTA_RISCO_T[L]?NOTA_RISCO_T[L][cls]:NOTA_RISCO[cls];
  return '<details style="margin:8px 0;border:1px solid var(--line);border-radius:10px;padding:8px 10px;background:var(--bg)"'+(L!="pt-BR"?' open':'')+'><summary style="cursor:pointer;font-weight:700">📖 '+(ui?ui[0]:'Como fazer')+' ('+ativs.length+' '+(ui?(ativs.length==1?ui[1]:ui[2]):('atividade'+(ativs.length==1?'':'s')))+')'+(L!="pt-BR"?' '+(DOC_LANGS.filter(function(x){return x[0]==L})[0]||["","🌐"])[1]:'')+'</summary>'
    +(!vig&&nota?'<div class="s" style="margin:6px 0;padding:6px;border-left:4px solid '+acCor(cls)+'">'+esc(nota)+'</div>':'')
    +ativs.map(function(g){var o=obsAtivT(g,L);return '<div style="border-top:1px solid var(--line);padding:6px 0"><b>'+esc(nomeAtivT(g,L))+'</b> <span class="s">· '+esc(freqT(g.freq,L))+'</span>'+(o.txt?'<div class="s" style="white-space:pre-wrap;color:var(--ink)">'+esc(o.txt)+'</div>':'')+(!o.oficial&&o.txt?'<a class="s" href="'+gTradUrl(g.obs,L)+'" target="_blank" rel="noopener" style="color:var(--brand)">🌐 '+(L=="en"?"Automatic translation":L=="es"?"Traducción automática":"Traduction automatique")+'</a>':'')+(L!="pt-BR"&&nomeAtivT(g,L)!=g.nome?'<div class="s" style="opacity:.7">🇧🇷 '+esc(g.nome)+'</div>':'')+'</div>'}).join("")
    +'</details>';
}
function comoFazerRestaurar(id){var g=PG.filter(function(x){return x.id==id})[0];if(!g||!arChk(g.l))return;g.obs="";g.obsAuto=true;svPg();var y=window.scrollY;draw();window.scrollTo(0,y)}
function arChk(i){var L=LOC[i];if(!L||!podeAreasCt(L.ct)){alert("Sem permissão.");return false}return true}
var AREAS_PADRAO=[["Área 01 — Recepção / Entrada",["Espaço 01","Espaço 02"]],["Área 02 — Sanitários",["Banheiro masculino","Banheiro feminino"]],["Área 03 — Áreas comuns",["Corredores","Escadas"]],["Área 04 — Copa / Refeitório",["Copa"]],["Área 05 — Apoio (DML)",["DML"]]];
var ATIV_PADRAO_GERAL=["Limpeza de pisos","Limpeza de superfícies","Limpeza de sanitários","Recolhimento de lixo","Reposição de insumos"];
function criarAreasPadrao(ctId){
  var c=ctById(ctId);if(!c)return 0;
  var t=tiposCt(c)[0];
  AREAS_PADRAO.forEach(function(a){
    LOC.push({n:a[0],e:a[1].slice(),ct:ctId,tipo:t,padrao:true});
    var i=LOC.length-1;ATIV_PADRAO_GERAL.forEach(function(n){PG.push({id:proxPgId(),l:i,nome:n,freq:"Diária"})});
  });
  return AREAS_PADRAO.length;
}
/* Áreas da versão anterior pertencem a um único contrato (Sadalla): vão para ele automaticamente */
function autoAtribuirOrfas(){
  var orf=areasSemContrato();if(!orf.length)return 0;
  var alvo=CONTRATOS.filter(function(c){return /sadalla/i.test(normT(c.nome))});
  if(alvo.length!=1)return 0;
  var t=tiposCt(alvo[0]);
  if(!ctTemTipo(alvo[0],"hosp")){alvo[0].tipos=(t.length==1&&t[0]=="out")?["hosp"]:t.concat(["hosp"]);alvo[0].tipo=alvo[0].tipos[0]}
  orf.forEach(function(i){LOC[i].ct=alvo[0].id;LOC[i].tipo="hosp"});
  /* atividades dessas áreas ficam como estavam */
  return orf.length;
}
/* Todo contrato sem áreas recebe áreas com nome padrão, para o supervisor ajustar no mapeamento */
function garantirAreasPadrao(){
  var n=0;
  CONTRATOS.forEach(function(c){if(!areasDoContrato(c.id).length&&!c.padraoCriado){criarAreasPadrao(c.id);c.padraoCriado=true;n++}});
  if(n){svLoc();svPg();svCt()}
  return n;
}
function prepararAreas(){var o=autoAtribuirOrfas();if(o){svLoc();svCt()}garantirAreasPadrao()}
function arRenL(i,v){if(!arChk(i))return;var nv=(v||"").trim();if(nv&&nv!=LOC[i].n){LOC[i].n=nv;LOC[i].padrao=false}svLoc();draw()}
function arRenE(i,e,v){if(!arChk(i))return;LOC[i].e[e]=(v||"").trim()||LOC[i].e[e];LOC[i].padrao=false;svLoc()}
function arTogE(i,e){if(!arChk(i))return;var L=LOC[i];L.offE=L.offE||{};L.offE[e]=!L.offE[e];svLoc();draw()}
function arDelE(i,e){if(!arChk(i))return;var L=LOC[i];if(!confirm("Excluir o espaço "+L.e[e]+"?"))return;L.delE=L.delE||{};L.delE[e]=true;svLoc();logAcao("Excluiu espaço",L.n+" · "+L.e[e]);draw()}
function arAddE(i){if(!arChk(i))return;var el=document.getElementById("ne_"+i),v=(el&&el.value||"").trim();if(!v)return;LOC[i].e.push(v);LOC[i].padrao=false;svLoc();draw()}
function arTogL(i){if(!arChk(i))return;LOC[i].off=!LOC[i].off;svLoc();draw()}
function arDelL(i){
  if(!arChk(i))return;var L=LOC[i];
  if(!confirm("Excluir a área "+L.n+"? O histórico de registros é mantido."))return;
  L.del=true;svLoc();
  COL.forEach(function(c){(c.vinculos||[]).forEach(function(v){if(v.areas)v.areas=v.areas.filter(function(x){return x!=i})})});svCol();
  logAcao("Excluiu área",L.n);arAberta=-1;draw();
}
function arAddL(ctId){
  if(!podeAreasCt(ctId)){alert("Sem permissão.");return}
  var el=document.getElementById("nl_ar"),v=(el&&el.value||"").trim();if(!v)return;
  var at=tiposCt(ctById(ctId));
  LOC.push({n:v,e:[],ct:ctId,tipo:at[0]});arAberta=LOC.length-1;svLoc();logAcao("Criou área",v);draw();
}
function proxPgId(){return PG.reduce(function(m,g){return Math.max(m,g.id)},0)+1}
function arAddAtiv(i){if(!arChk(i))return;var el=document.getElementById("na_"+i),v=(el&&el.value||"").trim();if(!v)return;PG.push({id:proxPgId(),l:i,nome:v});svPg();draw()}
function arAtivModelo(i){
  if(!arChk(i))return;var mdl=MODELOS[tipoDaArea(i)||"out"]||{ativ:[]};
  var ex=PG.filter(function(g){return g.l==i}).map(function(g){return g.nome.toLowerCase()});
  mdl.ativ.forEach(function(n){if(ex.indexOf(n.toLowerCase())<0)PG.push({id:proxPgId(),l:i,nome:n})});svPg();draw();
}
function aplicarModelo(ctId,t){
  if(!podeAreasCt(ctId)){alert("Sem permissão.");return}
  var c=ctById(ctId);t=t||tiposCt(c)[0];var mdl=MODELOS[t];if(!mdl||!mdl.areas.length)return;
  var padr=areasDoContrato(ctId).filter(function(i){return LOC[i].padrao});
  var todasPadrao=padr.length&&padr.length==areasDoContrato(ctId).length;
  if(!confirm("Criar "+mdl.areas.length+" áreas do modelo "+tipoCt(t).nome+(todasPadrao?" no lugar das "+padr.length+" áreas de nome padrão":"")+"?"))return;
  if(todasPadrao){padr.forEach(function(i){LOC[i].del=true;PG=PG.filter(function(g){return g.l!=i})});COL.forEach(function(x){(x.vinculos||[]).forEach(function(v){if(v.areas)v.areas=v.areas.filter(function(a){return padr.indexOf(a)<0})})});svCol()}
  mdl.areas.forEach(function(a){
    LOC.push({n:a[0],e:a[1].slice(),ct:ctId,tipo:t,modelo:t});
    var i=LOC.length-1;mdl.ativ.forEach(function(n){PG.push({id:proxPgId(),l:i,nome:n})});
  });
  svLoc();svPg();logAcao("Aplicou modelo de áreas",c.codigo,tipoCt(t).nome);draw();
}
function moverArea(i,ctId){if(!temAcao("gerenciar_contratos"))return;LOC[i].ct=ctId;svLoc();logAcao("Moveu área para contrato",LOC[i].n);draw()}
function moverTodas(ctId){if(!temAcao("gerenciar_contratos"))return;if(!confirm("Mover todas as áreas sem contrato para este contrato?"))return;areasSemContrato().forEach(function(i){LOC[i].ct=ctId});svLoc();draw()}
function togAreaColab(colId,ctId,i){
  if(!podeAreasCt(ctId)){alert("Sem permissão.");return}
  var c=cget(colId);if(!c)return;
  var v=(c.vinculos||[]).filter(function(x){return x.contratoId==ctId})[0];if(!v)return;
  v.areas=v.areas||[];var p=v.areas.indexOf(i);
  if(p>=0)v.areas.splice(p,1);else v.areas.push(i);
  svCol();var y=window.scrollY;draw();window.scrollTo(0,y);
}
function liberarTodasAreas(ctId){
  if(!podeAreasCt(ctId)){alert("Sem permissão.");return}
  if(!confirm("Liberar todas as áreas ativas para todos os colaboradores deste contrato?"))return;
  var at=areasDoContrato(ctId).filter(function(i){return !LOC[i].off});
  COL.forEach(function(c){(c.vinculos||[]).forEach(function(v){if(v.contratoId==ctId){v.areas=v.areas||[];at.forEach(function(i){if(v.areas.indexOf(i)<0)v.areas.push(i)})}})});
  svCol();logAcao("Liberou todas as áreas",(ctById(ctId)||{}).codigo);draw();
}

/* ===================================================================
   PERFIS AUTOMÁTICOS (efetivo → usuários)
   =================================================================== */
function pinInicial(c){
  var d=String(c.cpf||"").replace(/\D/g,"");
  if(d.length<4)d=String(c.matricula||c.funcionarioCod||"").replace(/\D/g,"");
  if(d.length<4)d=("0000"+d);
  return d.slice(-4);
}
function userDoColab(colId){return USERS.filter(function(u){return u.colId==colId})[0]||null}
function cargoPorFuncao(f){
  f=String(f||"");
  if(/supervis/i.test(f))return "sup";
  if(/encarregad|l[ií]der/i.test(f))return "enc";
  return "colab";
}
function criarUserColab(c,cargoId){
  var ct=ctDoDepto(c);
  var u={id:Date.now()+Math.floor(Math.random()*100000),nome:c.nome,mat:c.matricula||"",cargoId:cargoId,pin:pinInicial(c),pinTrocar:true,colId:c.id,
    contratosIds:ct&&cargoId!="colab"?[ct.id]:[],ctAuto:ct&&cargoId!="colab"?ct.id:null,cpf:c.cpf||"",idiomaComun:c.idiomaComun||"pt-BR",
    perfilCompleto:false,origem:"efetivo",criado:Date.now()};
  USERS.push(u);return u;
}
/* Todo supervisor do efetivo ganha perfil automaticamente, ligado ao contrato do seu departamento */
function syncSupervisores(){
  var novos=0,atual=0;
  COL.forEach(function(c){
    if(cargoPorFuncao(c.funcao)!="sup")return;
    var u=userDoColab(c.id)||USERS.filter(function(x){return !x.colId&&((c.matricula&&x.mat==c.matricula)||(x.nome||"").toLowerCase()==(c.nome||"").toLowerCase())})[0];
    if(ehDeslig(c.situacao)){if(u&&u.origem=="efetivo"&&!u.bloqueado){u.bloqueado=true;atual++}return}
    if(!u){u=criarUserColab(c,"sup");novos++;}
    else{
      if(!u.colId){u.colId=c.id;atual++}
      if(u.bloqueado&&u.origem=="efetivo"){u.bloqueado=false;atual++}
      var ct=ctDoDepto(c);
      if(ct){
        u.contratosIds=u.contratosIds||[];
        if(u.ctAuto&&u.ctAuto!=ct.id)u.contratosIds=u.contratosIds.filter(function(x){return x!=u.ctAuto});
        if(u.contratosIds.indexOf(ct.id)<0&&u.ctAutoBloq!=ct.id){u.contratosIds.push(ct.id);atual++}
        u.ctAuto=ct.id;
      }
    }
  });
  /* supervisor do departamento vira supervisor dos vínculos sem supervisor */
  CONTRATOS.forEach(function(ct){
    var sups=USERS.filter(function(u){return u.cargoId=="sup"&&!u.bloqueado&&u.ctAuto==ct.id});
    if(sups.length!=1)return;
    COL.forEach(function(c){(c.vinculos||[]).forEach(function(v){if(v.contratoId==ct.id&&!v.supervisorId)v.supervisorId=sups[0].id})});
  });
  /* desligados perdem acesso */
  USERS.forEach(function(u){if(u.origem=="efetivo"&&u.colId){var c=cget(u.colId);if(c&&ehDeslig(c.situacao)&&!u.bloqueado)u.bloqueado=true}});
  saveUsers();svCol();
  return {novos:novos,atual:atual};
}
function liberarAcesso(colId,silencioso){
  var c=cget(colId);if(!c)return false;
  var ok=(c.vinculos||[]).some(function(v){return podeAreasCt(v.contratoId)});
  if(!ok){if(!silencioso)alert("Sem permissão.");return false}
  if(userDoColab(colId)){if(!silencioso)alert("Já tem acesso.");return false}
  var u=criarUserColab(c,cargoPorFuncao(c.funcao)=="sup"?"sup":cargoPorFuncao(c.funcao));
  saveUsers();logAcao("Criou acesso ao app",c.nome);
  if(!silencioso){draw();alert("Acesso criado para "+c.nome+".\nPIN inicial: 4 últimos dígitos do "+(String(c.cpf||"").replace(/\D/g,"").length>=4?"CPF":"número da matrícula")+". A troca será exigida no 1º acesso.")}
  return true;
}
function liberarAcessoTodos(ctId){
  if(!podeAreasCt(ctId)){alert("Sem permissão.");return}
  var lista=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==ctId})&&!ehDeslig(x.situacao)&&!userDoColab(x.id)});
  if(!lista.length){alert("Todos já têm acesso.");return}
  if(!confirm("Criar acesso ao app para "+lista.length+" colaborador(es)?"))return;
  var n=0;lista.forEach(function(x){if(liberarAcesso(x.id,true))n++});
  draw();alert(n+" acesso(s) criado(s). PIN inicial: 4 últimos dígitos do CPF (ou matrícula).");
}
function salvarPinInicial(){
  var u=userById(session.userId);if(!u)return;
  var a=($("#np1")||{}).value||"",b=($("#np2")||{}).value||"";
  if(!/^\d{4}$/.test(a)){alert("O PIN deve ter 4 dígitos.");return}
  if(a!==b){alert("Os PINs não conferem.");return}
  if(a===u.pin){alert("Escolha um PIN diferente do inicial.");return}
  if(/^(\d)\1{3}$/.test(a)||a=="1234"){alert("PIN muito fácil. Escolha outro.");return}
  u.pin=a;u.pinTrocar=false;saveUsers();logAcao("Definiu PIN no 1º acesso");
  if(u.perfilCompleto)startApp();else draw();
}

/* ===================================================================
   COPIAR PROGRAMAÇÃO DE LIMPEZA PARA OUTROS CONTRATOS
   (áreas, espaços, classificação crítica/semicrítica/não crítica e atividades)
   =================================================================== */
var copiaAberta=false,CP={a:{},d:{},modo:"mesclar",ct:null};
function cpSel(t,id,on){CP[t][id]=on}
function cpTodos(t,ids,on){ids.forEach(function(i){CP[t][i]=on});var y=window.scrollY;draw();window.scrollTo(0,y)}
function cardCopiaProg(c){
  var areas=areasDoContrato(c.id);
  if(!areas.length)return "";
  var destinos=meusContratos().filter(function(x){return x.id!=c.id&&podeAreasCt(x.id)});
  var h='<div class="card noprint"><div class="row" style="border:0;padding:0;cursor:pointer" onclick="copiaAberta=!copiaAberta;draw()"><span style="flex:1"><b>📋 Copiar esta programação para outros contratos</b><div class="s">Replica áreas, espaços, classificação e atividades (com frequência e "como fazer").</div></span><span class="pill pen">'+(copiaAberta?'▲':'▼')+'</span></div>';
  if(!copiaAberta)return h+'</div>';
  if(CP.ct!=c.id){CP={a:{},d:{},modo:"mesclar",ct:c.id};areas.forEach(function(i){CP.a[i]=true})}
  if(!destinos.length)return h+'<div class="s" style="margin-top:8px">Não há outro contrato sob sua responsabilidade para receber a cópia.</div></div>';
  h+='<div class="s" style="margin-top:10px"><b>1. Quais áreas copiar</b></div><div class="chips" style="margin:4px 0"><button onclick="cpTodos(\'a\',['+areas.join(",")+'],true)">Todas</button><button onclick="cpTodos(\'a\',['+areas.join(",")+'],false)">Nenhuma</button></div>'
    +areas.map(function(i){var L=LOC[i];return '<label class="s" style="display:block;padding:4px 0;cursor:pointer"><input type="checkbox" class="cp_area" value="'+i+'"'+(CP.a[i]?' checked':'')+' onchange="cpSel(\'a\','+i+',this.checked)" style="width:auto;margin-right:8px">'+esc(L.n)+' <span class="s">('+L.e.filter(function(_,e){return !espDel(i,e)}).length+' espaços · '+PG.filter(function(g){return g.l==i}).length+' atividades)</span></label>'}).join("")
    +'<div class="s" style="margin-top:10px"><b>2. Para quais contratos</b></div>'
    +campoBusca("cp","Buscar contrato de destino")
    +destinos.filter(function(x){return bate("cp",x.codigo,x.nome)}).map(function(x){return '<label class="s" style="display:block;padding:4px 0;cursor:pointer"><input type="checkbox" class="cp_dest" value="'+x.id+'"'+(CP.d[x.id]?' checked':'')+' onchange="cpSel(\'d\','+x.id+',this.checked)" style="width:auto;margin-right:8px">'+esc(tiposCtIco(x)+" "+x.codigo+" · "+x.nome)+' <span class="s">('+areasDoContrato(x.id).length+' áreas hoje)</span></label>'}).join("")
    +'<div class="s" style="margin-top:10px"><b>3. Se a área já existir no destino (mesmo nome)</b></div>'
    +'<select id="cp_modo" onchange="CP.modo=this.value"><option value="mesclar"'+(CP.modo=="mesclar"?' selected':'')+'>Mesclar: acrescenta só os espaços e atividades que faltam</option><option value="nova"'+(CP.modo=="nova"?' selected':'')+'>Criar outra área com o mesmo nome</option></select>'
    +'<button class="btn" onclick="copiarProgramacao('+c.id+')">Copiar programação</button>';
  return h+'</div>';
}
function marcarTodos(cls,on){$all("."+cls).forEach(function(x){x.checked=on})}
function copiarProgramacao(origId){
  var areas=areasDoContrato(origId).filter(function(i){return CP.a[i]});
  var dests=Object.keys(CP.d).filter(function(k){return CP.d[k]}).map(Number).filter(function(id){return id!=origId&&ctById(id)&&podeAreasCt(id)});
  var modo=CP.modo||"mesclar";
  if(!areas.length){alert("Marque ao menos uma área.");return}
  if(!dests.length){alert("Marque ao menos um contrato de destino.");return}
  if(!confirm("Copiar "+areas.length+" área(s) para "+dests.length+" contrato(s)?"))return;
  var tot={areas:0,esp:0,ativ:0};
  dests.forEach(function(did){
    var dc=ctById(did);if(!dc)return;
    areas.forEach(function(i){
      var L=LOC[i],tipo=tipoDaArea(i);
      /* o destino passa a ter o tipo de serviço da área copiada */
      if(tipo&&!ctTemTipo(dc,tipo)){var at=tiposCt(dc).filter(function(t){return !(t=="out"&&!areasDoContrato(did).length)});at.push(tipo);dc.tipos=at;dc.tipo=at[0]}
      var alvo=-1;
      if(modo=="mesclar")alvo=areasDoContrato(did).filter(function(j){return normT(LOC[j].n)==normT(L.n)})[0];
      if(alvo==null||alvo<0){LOC.push({n:L.n,e:[],ct:did,tipo:tipo||tiposCt(dc)[0],origemCopia:i});alvo=LOC.length-1;tot.areas++}
      var D=LOC[alvo];
      L.e.forEach(function(nome,e){
        if(espDel(i,e))return;
        var ex=D.e.map(function(n,k){return espDel(alvo,k)?null:normT(n)}).indexOf(normT(nome));
        if(ex<0){D.e.push(nome);ex=D.e.length-1;tot.esp++}
        var cls=ACL[i+"_"+e];if(cls&&!ACLA[i+"_"+e]){ACL[alvo+"_"+ex]=cls;delete ACLA[alvo+"_"+ex]}
      });
      var exA=PG.filter(function(g){return g.l==alvo}).map(function(g){return normT(g.nome)});
      PG.filter(function(g){return g.l==i}).forEach(function(g){
        if(exA.indexOf(normT(g.nome))>=0)return;
        PG.push({id:proxPgId(),l:alvo,nome:g.nome,freq:g.freq||"",obs:g.obs||""});tot.ativ++;
      });
    });
  });
  svLoc();svPg();sv("rad_acl",ACL);svCt();
  logAcao("Copiou programação de limpeza",(ctById(origId)||{}).codigo,dests.length+" contrato(s)");
  copiaAberta=false;CP.ct=null;draw();
  alert("Programação copiada para "+dests.length+" contrato(s):\n"+tot.areas+" área(s) nova(s), "+tot.esp+" espaço(s), "+tot.ativ+" atividade(s).\n\nAs áreas copiadas ainda precisam ser liberadas aos colaboradores de cada contrato.");
}

/* ===================================================================
   IMPRESSÃO: PTI (Plano de Trabalho Individual) e PTG (Plano de Trabalho Geral)
   =================================================================== */
var DOC_CSS='*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}body{font:11.5px/1.4 "Segoe UI",Arial,sans-serif;color:#14211e;margin:0;padding:18px}'
 +'.doc{max-width:190mm;margin:0 auto;page-break-after:always}.doc:last-child{page-break-after:auto}'
 +'.cab{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:3px solid #0f5c4d;padding-bottom:6px;margin-bottom:10px}'
 +'.cab .emp{font-weight:700;font-size:13px;color:#0f5c4d}.cab .tit{font-size:16px;font-weight:800;text-transform:uppercase;letter-spacing:.3px}.cab .cod{text-align:right;font-size:10px;color:#5d6e69}'
 +'.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:0;border:1px solid #9bb0a9;margin-bottom:10px}.grid div{padding:5px 7px;border:1px solid #d8e0dd}.grid b{display:block;font-size:9px;text-transform:uppercase;color:#5d6e69;font-weight:600}'
 +'h3{font-size:12px;background:#0f5c4d;color:#fff;padding:5px 8px;margin:12px 0 0;text-transform:uppercase;letter-spacing:.3px}'
 +'h4{font-size:11.5px;margin:8px 0 2px;color:#0f5c4d}'
 +'table{width:100%;border-collapse:collapse;margin:0 0 6px}th,td{border:1px solid #9bb0a9;padding:4px 6px;text-align:left;vertical-align:top}th{background:#e8efec;font-size:10px;text-transform:uppercase}'
 +'.cr{display:inline-block;padding:0 5px;border-radius:3px;color:#fff;font-size:9px;font-weight:700}'
 +'.ass{display:grid;grid-template-columns:repeat(3,1fr);gap:22px;margin-top:34px}.ass div{border-top:1px solid #14211e;text-align:center;padding-top:4px;font-size:10px}'
 +'.rod{margin-top:14px;font-size:9px;color:#5d6e69;border-top:1px solid #d8e0dd;padding-top:4px;display:flex;justify-content:space-between}'
 +'.marca{border:2px dashed #c62828;color:#c62828;text-align:center;font-weight:700;padding:4px;margin-bottom:8px;text-transform:uppercase}'
 +'.mut{color:#5d6e69}.nobr{break-inside:avoid}'
 +'.bar{position:sticky;top:0;background:#fff;padding:8px 0;margin-bottom:8px;text-align:right}.bar button{padding:8px 16px;font:inherit;font-weight:700;border:0;border-radius:6px;background:#0f5c4d;color:#fff;cursor:pointer;margin-left:6px}'
 +'@page{size:A4;margin:12mm}@media print{body{padding:0}.bar{display:none}}';
function abrirImpressao(titulo,corpo){
  var html='<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>'+esc(titulo)+'</title><style>'+DOC_CSS+'</style></head><body><div class="bar"><button onclick="window.print()">🖨️ Imprimir / Salvar PDF</button><button onclick="window.close()" style="background:#5d6e69">Fechar</button></div>'+corpo+'<script>setTimeout(function(){window.print()},400)<\/script></body></html>';
  var w=null;try{w=window.open("","_blank")}catch(e){}
  if(!w){
    /* pop-up bloqueado: baixa o documento para abrir e imprimir */
    baixar(titulo.replace(/[^\w\-]+/g,"_")+".html",html,"text/html");
    alert("O navegador bloqueou a nova janela. O documento foi baixado; abra-o para imprimir ou salvar em PDF.");
    return;
  }
  w.document.open();w.document.write(html);w.document.close();
}
function crTag(k){return k&&k!="X"?'<span class="cr" style="background:'+acCor(k)+'">'+esc(acNome(k))+'</span>':'<span class="mut">—</span>'}
function cabDoc(titulo,codigo){
  return '<div class="cab"><div>'+(typeof TEMA!="undefined"&&TEMA.logo?'<img src="'+TEMA.logo+'" alt="" style="max-height:38px;max-width:160px;display:block;margin-bottom:4px">':'')+'<div class="emp">'+esc(CONTRATO)+'</div><div class="tit">'+esc(titulo)+'</div></div><div class="cod">'+esc(codigo)+'<br>Emitido em '+new Date().toLocaleString("pt-BR")+'<br>por '+esc(user||"")+'</div></div>';
}
function rodDoc(){return '<div class="rod"><span>Documento gerado pelo RAD Digital — Registro de Atividades Diárias</span><span>Cópia controlada somente no sistema</span></div>'}
function nomeAprov(p,nivel){var a=(p.aprovacoes||[]).filter(function(x){return x.nivel==nivel&&x.status!="Rejeitado"}).pop();return a?esc(a.nome)+'<br><span class="mut">'+new Date(a.data).toLocaleDateString("pt-BR")+'</span>':""}
var PER_NOME={diaria:"Diária",semanal:"Semanal",mensal:"Mensal",eventual:"Eventual"};

function htmlPTI(p){
  var c=cget(p.colaboradorId)||{},ct=ctById(p.contratoId)||{},v=(c.vinculos||[]).filter(function(x){return x.contratoId==p.contratoId})[0]||{};
  var sup=v.supervisorId?userById(v.supervisorId):(p.supervisorId?userById(p.supervisorId):null);
  var uC=p.colaboradorId?userDoColab(p.colaboradorId):null,LC=uC&&LIX[uC.idiomaComun]!=null?uC.idiomaComun:null;
  var h='<div class="doc">'+cabDoc("Plano de Trabalho Individual — PTI","PTI nº "+String(p.id).slice(-6));
  if(p.status!="Aprovado")h+='<div class="marca">'+esc(p.status)+' — documento ainda não aprovado</div>';
  h+='<div class="grid">'
    +'<div><b>Colaborador</b>'+esc(c.nome||"?")+'</div><div><b>Matrícula</b>'+esc(c.matricula||"—")+'</div><div><b>Função</b>'+esc(v.funcao||c.funcao||"—")+'</div>'
    +'<div><b>Contrato</b>'+esc((ct.codigo||"")+" · "+(ct.nome||""))+'</div><div><b>Serviço</b>'+esc(tiposCtNome(ct).replace(/[^\wÀ-ÿ ·\/]/g,"").trim())+'</div><div><b>Escala</b>'+esc(v.escala||c.escala||"—")+'</div>'
    +'<div><b>Supervisor</b>'+esc(sup?sup.nome:"—")+'</div><div><b>Situação do PTI</b>'+esc(p.status)+'</div><div><b>Atualizado em</b>'+new Date(p.atualizado||p.criado).toLocaleDateString("pt-BR")+'</div>'
    +'</div>';
  h+='<h3>Áreas e atividades sob responsabilidade</h3>';
  if(!p.areas.length)h+='<p class="mut">Nenhuma área definida.</p>';
  p.areas.forEach(function(a){
    var L=LOC[a.locIdx];if(!L)return;
    var esps=L.e.map(function(n,e){return [n,e]}).filter(function(x){return !espDel(a.locIdx,x[1])&&!espOff(a.locIdx,x[1])});
    var ativs=a.atividades.map(function(id){return PG.filter(function(g){return g.id==id})[0]}).filter(Boolean);
    h+='<div class="nobr"><h4>'+esc(L.n)+(L.del?' (área excluída)':'')+(a.horario?' <span class="mut" style="font-weight:400">· horário '+esc(a.horario)+'</span>':'')+'</h4>'
      +'<table><tr><th style="width:30%">Atividade</th><th style="width:26%">Periodicidade · dias · horário</th><th>Como fazer / observação</th></tr>'
      +(ativs.length?ativs.map(function(g){var o=LC?obsAtivT(g,LC):null;return '<tr><td>'+esc(g.nome)+(LC&&nomeAtivT(g,LC)!=g.nome?'<br><i style="color:#555">'+esc(nomeAtivT(g,LC))+'</i>':'')+'</td><td>'+esc(ptiCfgTxt((a.cfg||{})[g.id])||g.freq||"Diária")+'</td><td>'+esc(g.obs||"")+(o&&o.oficial&&o.txt&&o.txt!=g.obs?'<div style="margin-top:4px;padding-top:4px;border-top:1px dashed #bbb;color:#333"><b>'+LC.toUpperCase()+':</b> '+esc(o.txt)+'</div>':'')+'</td></tr>'}).join(""):'<tr><td colspan="3" class="mut">Sem atividades marcadas.</td></tr>')
      +'</table>'
      +(esps.length?'<table><tr><th style="width:34%">Espaços da área</th><th style="width:16%">Classificação</th><th>EPI obrigatório</th></tr>'+esps.map(function(x){return '<tr><td>'+esc(x[0])+'</td><td>'+crTag(aclOf(a.locIdx,x[1]))+'</td><td>'+esc(episTxt(a.locIdx,x[1]))+'</td></tr>'}).join("")+'</table>':'')
      +'</div>';
  });
  if(p.metas)h+='<h3>Metas e observações</h3><p style="white-space:pre-wrap">'+esc(p.metas)+'</p>';
  var hist=(p.aprovacoes||[]);
  if(hist.length)h+='<h3>Histórico de aprovação</h3><table><tr><th>Etapa</th><th>Responsável</th><th>Data</th><th>Situação</th><th>Observação</th></tr>'+hist.map(function(a){return '<tr><td>'+esc(a.nivel)+'</td><td>'+esc(a.nome)+'</td><td>'+new Date(a.data).toLocaleString("pt-BR")+'</td><td>'+esc(a.status)+'</td><td>'+esc(a.obs||"")+'</td></tr>'}).join("")+'</table>';
  h+='<div class="ass"><div>'+esc(c.nome||"")+'<br>Colaborador — ciente</div><div>'+(nomeAprov(p,"Supervisor")||esc(sup?sup.nome:""))+'<br>Supervisor</div><div>'+nomeAprov(p,"Coordenação")+'<br>Coordenação</div></div>'
    +'<div class="ass" style="grid-template-columns:repeat(2,1fr)"><div>'+nomeAprov(p,"Gerência")+'<br>Gerência</div><div>'+nomeAprov(p,"Administrador")+'<br>Administração</div></div>';
  return h+rodDoc()+'</div>';
}
function imprimirPTI(id){var p=ptiById(id);if(!p)return;logAcao("Imprimiu PTI",(cget(p.colaboradorId)||{}).nome);abrirImpressao("PTI_"+((cget(p.colaboradorId)||{}).nome||id),htmlPTI(p))}
function imprimirPTIsContrato(ctId,soAprov){
  var l=PTIS.filter(function(p){return p.contratoId==ctId&&(!soAprov||p.status=="Aprovado")}).sort(function(a,b){return ((cget(a.colaboradorId)||{}).nome||"").localeCompare((cget(b.colaboradorId)||{}).nome||"")});
  if(!l.length){alert(soAprov?"Nenhum PTI aprovado neste contrato.":"Nenhum PTI neste contrato.");return}
  logAcao("Imprimiu PTIs do contrato",(ctById(ctId)||{}).codigo,l.length+" PTI(s)");
  abrirImpressao("PTIs_"+((ctById(ctId)||{}).codigo||ctId),l.map(htmlPTI).join(""));
}
function htmlPTG(ctId){
  var c=ctById(ctId);if(!c)return "";
  var pg=planoGeral(ctId),areas=locaisCt(c);
  var colabs=pg.colaboradores.slice().sort(function(a,b){return (a.nome||"").localeCompare(b.nome||"")});
  var sups=supervisoresDoContrato(ctId);
  var h='<div class="doc">'+cabDoc("Plano de Trabalho Geral — PTG","Contrato "+c.codigo);
  h+='<div class="grid">'
    +'<div><b>Contrato</b>'+esc(c.codigo+" · "+c.nome)+'</div><div><b>Serviços</b>'+esc(tiposCt(c).map(function(t){return tipoCt(t).nome}).join(", "))+'</div><div><b>Situação</b>'+esc(c.status||"")+'</div>'
    +'<div><b>Supervisão</b>'+esc(sups.map(function(u){return u.nome}).join(", ")||"—")+'</div><div><b>Responsável no cliente</b>'+esc(c.responsavel||"—")+'</div><div><b>Vigência</b>'+esc((c.inicio?fmtDataBR(c.inicio):"—")+" a "+(c.fim?fmtDataBR(c.fim):"—"))+'</div>'
    +'<div><b>Áreas</b>'+areas.length+'</div><div><b>Colaboradores</b>'+pg.totalColaboradores+'</div><div><b>PTIs aprovados / total</b>'+pg.aprovados+' / '+pg.totalPtis+'</div>'
    +'</div>';
  h+='<h3>Programação por área</h3>';
  if(!areas.length)h+='<p class="mut">Nenhuma área mapeada.</p>';
  areas.forEach(function(i){
    var L=LOC[i];
    var esps=L.e.map(function(n,e){return [n,e]}).filter(function(x){return !espDel(i,x[1])});
    var ativs=PG.filter(function(g){return g.l==i});
    /* quem atua: PTI aprovado na área ou área liberada ao colaborador */
    var quem=colabs.filter(function(x){
      var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==ctId})[0]||{};
      var pti=ptiDoColab(x.id,ctId);
      return (v.areas||[]).indexOf(i)>=0||(pti&&pti.status=="Aprovado"&&pti.areas.some(function(a){return a.locIdx==i}));
    });
    h+='<div class="nobr"><h4>'+esc(L.n)+(L.off?' (desativada)':'')+' <span class="mut" style="font-weight:400">· '+esc(tipoCt(tipoDaArea(i)).nome)+'</span></h4>'
      +'<table><tr><th style="width:34%">Espaços</th><th style="width:16%">Classificação</th><th>EPI obrigatório</th></tr>'
      +(esps.length?esps.map(function(x){return '<tr><td>'+esc(x[0])+(espOff(i,x[1])?' <span class="mut">(pausado)</span>':'')+'</td><td>'+crTag(aclOf(i,x[1]))+'</td><td>'+esc(episTxt(i,x[1]))+'</td></tr>'}).join(""):'<tr><td colspan="3" class="mut">Sem espaços.</td></tr>')+'</table>'
      +'<table><tr><th style="width:38%">Atividade</th><th style="width:18%">Frequência</th><th>Como fazer / observação</th></tr>'
      +(ativs.length?ativs.map(function(g){return '<tr><td>'+esc(g.nome)+'</td><td>'+esc(g.freq||"Diária")+'</td><td>'+esc(g.obs||"")+'</td></tr>'}).join(""):'<tr><td colspan="3" class="mut">Sem atividades.</td></tr>')+'</table>'
      +'<div style="margin-bottom:6px"><b>Equipe designada:</b> '+(quem.length?quem.map(function(x){return esc(x.nome)}).join(", "):'<span style="color:#c62828;font-weight:700">sem cobertura</span>')+'</div></div>';
  });
  h+='<h3>Equipe do contrato</h3><table><tr><th>Colaborador</th><th>Matrícula</th><th>Função</th><th>Escala</th><th>Áreas</th><th>PTI</th></tr>'
    +(colabs.length?colabs.map(function(x){
      var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==ctId})[0]||{};
      var pti=ptiDoColab(x.id,ctId);
      var ar=(v.areas||[]).filter(function(i){return LOC[i]&&!LOC[i].del}).map(function(i){return LOC[i].n});
      return '<tr><td>'+esc(x.nome)+'</td><td>'+esc(x.matricula||"")+'</td><td>'+esc(v.funcao||x.funcao||"")+'</td><td>'+esc(v.escala||x.escala||"")+'</td><td>'+esc(ar.join(", ")||"—")+'</td><td>'+esc(pti?pti.status:"Sem PTI")+'</td></tr>'}).join(""):'<tr><td colspan="6" class="mut">Nenhum colaborador.</td></tr>')
    +'</table>';
  var lac=areas.filter(function(i){return !colabs.some(function(x){var v=(x.vinculos||[]).filter(function(vv){return vv.contratoId==ctId})[0]||{};var pti=ptiDoColab(x.id,ctId);return (v.areas||[]).indexOf(i)>=0||(pti&&pti.status=="Aprovado"&&pti.areas.some(function(a){return a.locIdx==i}))})});
  if(lac.length)h+='<h3 style="background:#c62828">Áreas sem cobertura ('+lac.length+')</h3><p>'+lac.map(function(i){return esc(LOC[i].n)}).join(" · ")+'</p>';
  h+='<div class="ass"><div>Supervisor</div><div>Coordenação</div><div>Gerência</div></div><div class="ass" style="grid-template-columns:1fr 1fr"><div>Responsável do cliente</div><div>Qualidade / Compliance</div></div>';
  return h+rodDoc()+'</div>';
}
function imprimirPTG(ctId){var c=ctById(ctId);if(!c)return;logAcao("Imprimiu PTG",c.codigo);abrirImpressao("PTG_"+c.codigo,htmlPTG(ctId))}

/* ===================================================================
   AGENDA E AVISOS — recados programados, compromissos com prazo e reuniões
   Criam: Administrador, Coordenação, Gerência e Supervisor (ação "criar_avisos").
   Destinatários por pessoa, cargo ou contrato. Lembretes automáticos no app,
   notificação do navegador, WhatsApp e arquivo de agenda (.ics).
   =================================================================== */
var AGENDA=ld("rad_agenda",[]);
function svAg(){sv("rad_agenda",AGENDA)}
var AG_TIPOS={
  recado:{nome:"Recado / aviso",ico:"📣",desc:"Mensagem que sai na data e hora programadas"},
  compromisso:{nome:"Compromisso com prazo",ico:"⏰",desc:"Tarefa com prazo; alerta antes e se atrasar"},
  reuniao:{nome:"Reunião",ico:"👥",desc:"Encontro com participantes, local e confirmação de presença"}
};
var AG_LEMB=[[0,"No horário"],[15,"15 min antes"],[60,"1 h antes"],[240,"4 h antes"],[1440,"1 dia antes"],[2880,"2 dias antes"],[10080,"1 semana antes"],[21600,"15 dias antes"],[43200,"30 dias antes"]];
var AG_REC={nao:"Não repete",diaria:"Todo dia",uteis:"Dias úteis (seg a sex)",semanal:"Toda semana",quinzenal:"A cada 15 dias",mensal:"Todo mês",bimestral:"A cada 2 meses",trimestral:"A cada 3 meses",semestral:"A cada 6 meses",anual:"Todo ano"};
/* passo da repetição: dias ou meses */
var AG_PASSO={diaria:{d:1},semanal:{d:7},quinzenal:{d:15},mensal:{m:1},bimestral:{m:2},trimestral:{m:3},semestral:{m:6},anual:{m:12}};
/* soma meses sem "escorregar" (31/01 + 1 mês = 28 ou 29/02, não 03/03) */
function addMeses(t,n){var d=new Date(t),dd=d.getDate();d.setDate(1);d.setMonth(d.getMonth()+n);var ult=new Date(d.getFullYear(),d.getMonth()+1,0).getDate();d.setDate(Math.min(dd,ult));return d.getTime()}
var agTab="mim",agAberto=null;
var AGF=null;
function agNovoForm(){var d=new Date(Date.now()+3600000);AGF={tipo:"recado",titulo:"",msg:"",data:dia(d),hora:p2(d.getHours())+":00",horaFim:"",local:"",link:"",users:[],cargos:[],cts:[],lemb:[0],recorr:"nao",ate:"",ciente:true}}
function podeCriarAviso(){return temAcao("criar_avisos")}
function agTs(data,hora){if(!data)return 0;var t=new Date(data+"T"+(hora||"00:00")).getTime();return isNaN(t)?0:t}
function fmtDH(t){return new Date(t).toLocaleDateString("pt-BR",{weekday:"short",day:"2-digit",month:"2-digit",year:"numeric"})+" às "+hm(t)}

/* ---- destinatários ---- */
function agDestIds(it){
  var d=it.dest||{},ids={};
  (d.users||[]).forEach(function(id){ids[id]=1});
  USERS.forEach(function(u){
    if(u.bloqueado)return;
    if((d.cargos||[]).indexOf(u.cargoId)>=0)ids[u.id]=1;
    if((d.cts||[]).length){
      var cts=(u.contratosIds||[]).slice();
      var col=u.colId?cget(u.colId):null;
      if(col)(col.vinculos||[]).forEach(function(v){cts.push(v.contratoId)});
      if(cts.some(function(c){return d.cts.indexOf(c)>=0}))ids[u.id]=1;
    }
  });
  return Object.keys(ids).map(Number).filter(function(id){var u=userById(id);return u&&!u.bloqueado});
}
function agDestResumo(d){
  var p=[];
  if((d.users||[]).length)p.push((d.users||[]).map(function(id){var u=userById(id);return u?u.nome.split(" ")[0]:"?"}).slice(0,4).join(", ")+((d.users.length>4)?" +"+(d.users.length-4):""));
  if((d.cargos||[]).length)p.push("cargo: "+d.cargos.map(function(c){return (cargoById(c)||{}).nome||c}).join(", "));
  if((d.cts||[]).length)p.push("contrato: "+d.cts.map(function(c){return (ctById(c)||{}).codigo||"?"}).join(", "));
  return p.join(" · ")||"ninguém";
}

/* ---- ocorrências (recorrência) ---- */
/* k-ésima ocorrência calculada a partir da data base (evita acumular erro mês a mês) */
function agOcorrK(it,k){
  if(k==0)return it.quando;
  var rec=it.recorr,ps=AG_PASSO[rec];
  if(ps&&ps.m)return addMeses(it.quando,ps.m*k);
  if(ps&&ps.d){var d=new Date(it.quando);d.setDate(d.getDate()+ps.d*k);return d.getTime()}
  return null;
}
function agOcorrencias(it,de,ate){
  var r=[],lim=it.ate||Infinity;
  if(!it.recorr||it.recorr=="nao"){if(it.quando>=de&&it.quando<=ate)r.push(it.quando);return r}
  if(it.recorr=="uteis"){
    var d=new Date(it.quando),n=0;
    /* pula direto para perto do início da janela */
    if(de>it.quando){var dias=Math.floor((de-it.quando)/86400000)-1;if(dias>0)d.setDate(d.getDate()+dias)}
    while(d.getTime()<=ate&&d.getTime()<=lim&&n<3000){var w=d.getDay(),t=d.getTime();if(w!=0&&w!=6&&t>=de&&t>=it.quando)r.push(t);d.setDate(d.getDate()+1);n++}
    return r;
  }
  var ps=AG_PASSO[it.recorr];if(!ps)return r;
  /* estimativa do primeiro índice dentro da janela, para recorrências longas */
  var passoMs=ps.m?ps.m*30.44*86400000:ps.d*86400000;
  var k=Math.max(0,Math.floor((de-it.quando)/passoMs)-2),n2=0;
  while(n2<3000){var t2=agOcorrK(it,k);if(t2==null||t2>ate||t2>lim)break;if(t2>=de)r.push(t2);k++;n2++}
  return r;
}
function agDur(it){return it.tipo=="reuniao"&&it.fim?Math.max(0,it.fim-it.quando):3600000}
/* ocorrência que interessa agora para o usuário */
function agOcorrAtual(it,uid){
  var agora=Date.now();
  if(it.tipo=="recado"){var l=agOcorrencias(it,it.quando,agora);return l.length?l[l.length-1]:null}
  if(it.tipo=="reuniao"){var n=agOcorrencias(it,agora-agDur(it),agora+800*86400000);return n.length?n[0]:null}
  /* compromisso: a primeira ocorrência ainda não concluída (até 30 dias para trás) */
  var c=agOcorrencias(it,Math.max(it.quando,agora-60*86400000),agora+800*86400000);
  for(var i=0;i<c.length;i++){if(!(it.concl||{})[uid+"|"+c[i]])return c[i]}
  return null;
}
function agK(uid,occ){return uid+"|"+occ}
function agCiente(it,uid,occ){return !!(it.lido||{})[agK(uid,occ)]}
function agConcluido(it,uid,occ){return !!(it.concl||{})[agK(uid,occ)]}
function agResp(it,uid,occ){return (it.resp||{})[agK(uid,occ)]||""}

/* ---- itens do usuário logado ---- */
function agMeus(){
  var uid=session?session.userId:null;if(!uid)return [];
  return AGENDA.filter(function(it){return !it.cancelado&&agDestIds(it).indexOf(uid)>=0}).map(function(it){return {it:it,occ:agOcorrAtual(it,uid)}}).filter(function(x){return x.occ});
}
function agStatus(x,uid){
  var it=x.it,occ=x.occ,agora=Date.now();
  if(it.tipo=="compromisso"){if(agConcluido(it,uid,occ))return "ok";return occ<agora?"atrasado":"pendente"}
  if(it.tipo=="reuniao"){if(!agResp(it,uid,occ))return "confirmar";return agResp(it,uid,occ)=="sim"?"ok":"recusou"}
  return (it.ciente&&!agCiente(it,uid,occ))?"pendente":"ok";
}
function agPendentes(){var uid=session?session.userId:null;return agMeus().filter(function(x){var s=agStatus(x,uid);return s=="pendente"||s=="atrasado"||s=="confirmar"})}

/* ---- ações do destinatário ---- */
function agAcharIt(id){return AGENDA.filter(function(x){return x.id==id})[0]}
function agMarcar(id,occ,campo,valor){
  var it=agAcharIt(id);if(!it)return;var uid=session.userId;
  it[campo]=it[campo]||{};it[campo][agK(uid,occ)]=valor;
  if(campo!="lido"){it.lido=it.lido||{};it.lido[agK(uid,occ)]=Date.now()}
  svAg();logAcao({lido:"Deu ciente em aviso",concl:"Concluiu compromisso",resp:"Respondeu reunião"}[campo]||campo,it.titulo,valor===true?"":String(valor));
  fecharToast(id);var y=window.scrollY;draw();window.scrollTo(0,y);
}

/* ---- alertas automáticos (lembretes) ---- */
var TOASTS={};
function agVerificar(){
  if(!session)return;
  var uid=session.userId,agora=Date.now(),mud=false;
  agMeus().forEach(function(x){
    var it=x.it,occ=x.occ;it.disp=it.disp||{};
    var pend=agStatus(x,uid);
    if(pend=="ok"||pend=="recusou")return;
    var gatilhos=it.tipo=="recado"?[0]:(it.lemb&&it.lemb.length?it.lemb:[0]);
    gatilhos.forEach(function(m){
      var t=occ-m*60000,k=uid+"|"+occ+"|"+m;
      if(agora>=t&&agora<occ+agDur(it)+86400000&&!it.disp[k]){it.disp[k]=agora;mud=true;mostrarToast(it,occ,m)}
    });
    if(it.tipo=="compromisso"&&occ<agora){var k2=uid+"|"+occ+"|late";if(!it.disp[k2]){it.disp[k2]=agora;mud=true;mostrarToast(it,occ,"late")}}
  });
  if(mud)svAg();
}
function textoAlerta(it,occ,m){
  if(m=="late")return "⚠️ Prazo vencido em "+fmtDH(occ);
  if(it.tipo=="recado")return it.msg||"";
  if(m===0)return (it.tipo=="reuniao"?"Começa agora":"Prazo agora")+" — "+fmtDH(occ);
  var lb=(AG_LEMB.filter(function(l){return l[0]==m})[0]||[0,""])[1];
  if(m>=1440){var dias=Math.max(0,Math.ceil((occ-Date.now())/86400000));lb="faltam "+dias+" dia(s)"}
  return (it.tipo=="reuniao"?"Reunião ":"Prazo ")+fmtDH(occ)+" ("+lb.toLowerCase()+")";
}
function mostrarToast(it,occ,m){
  var box=document.getElementById("toasts");
  if(!box){box=document.createElement("div");box.id="toasts";box.style.cssText="position:fixed;right:12px;bottom:12px;z-index:200;display:flex;flex-direction:column;gap:8px;max-width:360px;width:calc(100% - 24px)";document.body.appendChild(box)}
  var tp=AG_TIPOS[it.tipo]||{ico:"🔔"};
  var d=document.createElement("div");d.id="toast_"+it.id;
  d.style.cssText="background:var(--card);color:var(--ink);border:1px solid var(--line);border-left:5px solid "+(m=="late"?"var(--warn)":"var(--brand)")+";border-radius:10px;padding:12px;box-shadow:0 8px 30px rgba(0,0,0,.25)";
  d.innerHTML='<div style="display:flex;gap:8px;align-items:flex-start"><span style="font-size:22px">'+tp.ico+'</span><div style="flex:1"><b>'+esc(it.titulo)+'</b><div class="s" style="white-space:pre-wrap">'+esc(textoAlerta(it,occ,m))+'</div><div class="s">de '+esc(it.autorNome||"")+'</div></div><button onclick="fecharToast('+it.id+')" style="border:0;background:none;font-size:20px;color:var(--mut);cursor:pointer">×</button></div>'
    +'<div style="display:flex;gap:6px;margin-top:8px"><button class="pill ok" style="border:0;cursor:pointer;padding:8px 12px" onclick="agTab=\'mim\';agAberto='+it.id+';go(\'ag\')">Abrir</button>'
    +(it.tipo=="recado"&&it.ciente?'<button class="pill pen" style="border:0;cursor:pointer;padding:8px 12px" onclick="agMarcar('+it.id+','+occ+',\'lido\',Date.now())">Ciente</button>':'')+'</div>';
  var old=document.getElementById("toast_"+it.id);if(old)old.remove();
  box.appendChild(d);
  try{if(window.Notification&&Notification.permission=="granted")new Notification(tp.ico+" "+it.titulo,{body:textoAlerta(it,occ,m),tag:"rad"+it.id})}catch(e){}
}
function fecharToast(id){var d=document.getElementById("toast_"+id);if(d)d.remove()}
function ativarNotif(){
  if(!window.Notification){alert("Este navegador não suporta notificações.");return}
  Notification.requestPermission().then(function(p){alert(p=="granted"?"Notificações ativadas neste aparelho.":"Notificações não autorizadas.");draw()});
}
setInterval(agVerificar,30000);

/* ---- WhatsApp e agenda (.ics) ---- */
function agTextoZap(it,occ){
  var tp=AG_TIPOS[it.tipo];
  var L=["*"+tp.ico+" "+tp.nome.toUpperCase()+": "+it.titulo+"*",""];
  if(it.tipo!="recado")L.push((it.tipo=="reuniao"?"📅 ":"⏰ Prazo: ")+fmtDH(occ)+(it.tipo=="reuniao"&&it.fim?" até "+hm(occ+agDur(it)):""));
  if(it.local)L.push("📍 "+it.local);
  if(it.link)L.push("🔗 "+it.link);
  if(it.recorr&&it.recorr!="nao")L.push("🔁 "+AG_REC[it.recorr]);
  if(it.msg){L.push("");L.push(it.msg)}
  L.push("");L.push("_Enviado por "+(it.autorNome||"")+" · RAD Digital_");
  return L.join("\n");
}
function foneZap(u){var n=String((u&&u.telefone)||"").replace(/\D/g,"");if(n.length<10)return "";return n.length<=11?"55"+n:n}
function agZap(id,uid){var it=agAcharIt(id);if(!it)return;var u=uid?userById(uid):null;var occ=agOcorrAtual(it,uid||0)||it.quando;var f=u?foneZap(u):"";window.open("https://wa.me/"+(f||"")+"?text="+encodeURIComponent(agTextoZap(it,occ)),"_blank")}
function icsEsc(s){return String(s||"").replace(/\\/g,"\\\\").replace(/;/g,"\;").replace(/,/g,"\\,").replace(/\r?\n/g,"\\n")}
function icsDt(t){var d=new Date(t);return d.getFullYear()+p2(d.getMonth()+1)+p2(d.getDate())+"T"+p2(d.getHours())+p2(d.getMinutes())+"00"}
function agIcs(id){
  var it=agAcharIt(id);if(!it)return;
  var L=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//RAD Digital//PT-BR","CALSCALE:GREGORIAN","BEGIN:VEVENT","UID:rad-"+it.id+"@raddigital","DTSTAMP:"+icsDt(Date.now()),"DTSTART:"+icsDt(it.quando),"DTEND:"+icsDt(it.quando+agDur(it)),"SUMMARY:"+icsEsc((AG_TIPOS[it.tipo]||{}).ico+" "+it.titulo),"DESCRIPTION:"+icsEsc((it.msg||"")+"\nOrganizador: "+(it.autorNome||""))];
  if(it.local)L.push("LOCATION:"+icsEsc(it.local));
  if(it.link)L.push("URL:"+it.link);
  if(it.recorr&&it.recorr!="nao"){
    var rr={diaria:"FREQ=DAILY",uteis:"FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR",semanal:"FREQ=WEEKLY",quinzenal:"FREQ=DAILY;INTERVAL=15",mensal:"FREQ=MONTHLY",bimestral:"FREQ=MONTHLY;INTERVAL=2",trimestral:"FREQ=MONTHLY;INTERVAL=3",semestral:"FREQ=MONTHLY;INTERVAL=6",anual:"FREQ=YEARLY"}[it.recorr];
    if(rr)L.push("RRULE:"+rr+(it.ate?";UNTIL="+icsDt(it.ate).slice(0,8)+"T235900":""));
  }
  (it.lemb||[]).forEach(function(m){L.push("BEGIN:VALARM","ACTION:DISPLAY","DESCRIPTION:"+icsEsc(it.titulo),"TRIGGER:-"+(m%1440==0&&m>0?"P"+(m/1440)+"D":"PT"+m+"M"),"END:VALARM")});
  L.push("END:VEVENT","END:VCALENDAR");
  try{var b=new Blob([L.join("\r\n")],{type:"text/calendar;charset=utf-8"}),u=URL.createObjectURL(b),a=document.createElement("a");a.href=u;a.download=(it.titulo||"evento").replace(/[^\wÀ-ÿ\- ]+/g,"").slice(0,40)+".ics";document.body.appendChild(a);a.click();setTimeout(function(){URL.revokeObjectURL(u);a.remove()},1200)}catch(e){alert("Erro: "+e.message)}
}
function agListaPresenca(id){
  var it=agAcharIt(id);if(!it)return;var occ=agOcorrAtual(it,0)||it.quando;
  var ids=agDestIds(it);
  var h='<div class="doc">'+cabDoc("Lista de presença — Reunião","Reunião nº "+String(it.id).slice(-6))
    +'<div class="grid"><div><b>Assunto</b>'+esc(it.titulo)+'</div><div><b>Data e hora</b>'+esc(fmtDH(occ))+(it.fim?' até '+hm(occ+agDur(it)):'')+'</div><div><b>Local</b>'+esc(it.local||it.link||"—")+'</div>'
    +'<div><b>Organizador</b>'+esc(it.autorNome||"")+'</div><div><b>Convidados</b>'+ids.length+'</div><div><b>Confirmados</b>'+ids.filter(function(u){return agResp(it,u,occ)=="sim"}).length+'</div></div>'
    +(it.msg?'<h3>Pauta</h3><p style="white-space:pre-wrap">'+esc(it.msg)+'</p>':'')
    +'<h3>Participantes</h3><table><tr><th style="width:5%">#</th><th style="width:35%">Nome</th><th style="width:18%">Cargo</th><th style="width:14%">Confirmação</th><th>Assinatura</th></tr>'
    +ids.map(function(uid,i){var u=userById(uid)||{};var cg=cargoById(u.cargoId)||{};var r=agResp(it,uid,occ);return '<tr style="height:30px"><td>'+(i+1)+'</td><td>'+esc(u.nome||"")+'</td><td>'+esc(cg.nome||"")+'</td><td>'+(r=="sim"?"Confirmou":r=="nao"?"Não poderá":"—")+'</td><td></td></tr>'}).join("")
    +[1,2,3].map(function(){return '<tr style="height:30px"><td></td><td></td><td></td><td></td><td></td></tr>'}).join("")
    +'</table><h3>Registro / encaminhamentos</h3><div style="border:1px solid #9bb0a9;height:140px"></div>'
    +'<div class="ass"><div>'+esc(it.autorNome||"")+'<br>Organizador</div><div>Qualidade / Compliance</div><div>Gerência</div></div>'+rodDoc()+'</div>';
  abrirImpressao("Presenca_"+it.titulo,h);
}

/* ---- criar ---- */
function agSalvar(){
  if(!podeCriarAviso()){alert("Sem permissão.");return}
  var f=AGF;
  if(!f.titulo.trim()){alert("Informe o título.");return}
  var q=agTs(f.data,f.hora);if(!q){alert("Informe data e hora.");return}
  if(f.tipo=="recado"&&!f.msg.trim()){alert("Escreva a mensagem do recado.");return}
  var fim=0;if(f.tipo=="reuniao"&&f.horaFim){fim=agTs(f.data,f.horaFim);if(fim<=q){alert("O término deve ser depois do início.");return}}
  var dest={users:f.users.slice(),cargos:f.cargos.slice(),cts:f.cts.slice()};
  var it={id:Date.now(),tipo:f.tipo,titulo:f.titulo.trim(),msg:f.msg.trim(),quando:q,fim:fim,local:f.local.trim(),link:f.link.trim(),
    dest:dest,lemb:f.tipo=="recado"?[0]:f.lemb.slice().sort(function(a,b){return b-a}),recorr:f.recorr,ate:f.recorr!="nao"&&f.ate?agTs(f.ate,"23:59"):0,
    ciente:f.tipo=="recado"?f.ciente:true,autorId:session.userId,autorNome:user,criado:Date.now(),lido:{},concl:{},resp:{},disp:{}};
  /* o organizador participa da própria reunião */
  if(it.tipo=="reuniao"&&dest.users.indexOf(session.userId)<0)dest.users.push(session.userId);
  var n=agDestIds(it).length;
  if(!n){alert("Escolha pelo menos um destinatário.");return}
  AGENDA.push(it);svAg();logAcao("Criou "+AG_TIPOS[it.tipo].nome.toLowerCase(),it.titulo,n+" destinatário(s)");
  agNovoForm();agTab="env";agAberto=it.id;BUSCA.agu="";BUSCA.agc="";
  draw();window.scrollTo(0,0);agVerificar();
  alert(AG_TIPOS[it.tipo].nome+" criado para "+n+" pessoa(s).\nUse 💬 WhatsApp para avisar quem não abre o app com frequência.");
}
function agCancelar(id){var it=agAcharIt(id);if(!it)return;if(it.autorId!=session.userId&&!temAcao("gerenciar_usuarios")){alert("Só quem criou pode cancelar.");return}if(!confirm("Cancelar \""+it.titulo+"\"?"))return;it.cancelado=Date.now();svAg();logAcao("Cancelou aviso/agenda",it.titulo);draw()}
function agPrevia(){var f=AGF,q=agTs(f.data,f.hora);if(!q)return '';var it={quando:q,recorr:f.recorr,ate:f.ate?agTs(f.ate,'23:59'):0};var l=agOcorrencias(it,q,q+3*366*86400000).slice(0,5);return '<div class="s" style="margin:4px 0 8px">📆 Próximas datas: '+l.map(function(t){return new Date(t).toLocaleDateString('pt-BR')}).join(' · ')+(l.length==5?' …':'')+'</div>'}
function agTog(lista,v){var a=AGF[lista],p=a.indexOf(v);if(p>=0)a.splice(p,1);else a.push(v);var y=window.scrollY;draw();window.scrollTo(0,y)}
function agAddUser(id){if(AGF.users.indexOf(id)<0)AGF.users.push(id);BUSCA.agu="";draw()}
function agAddCt(id){if(AGF.cts.indexOf(id)<0)AGF.cts.push(id);BUSCA.agc="";draw()}

/* ---- tela ---- */
function vAgenda(){
  if(!AGF)agNovoForm();
  var uid=session.userId,pode=podeCriarAviso();
  var abas=[["mim","📥 Para mim ("+agPendentes().length+")"]];
  if(pode)abas.push(["env","📤 Enviados"],["novo","➕ Novo"]);
  if(abas.map(function(a){return a[0]}).indexOf(agTab)<0)agTab="mim";
  var h='<div class="card"><b>📅 '+T("menu_ag")+'</b><div class="s">Recados programados, compromissos com prazo e reuniões.</div>'
    +'<div class="chips">'+abas.map(function(a){return '<button class="'+(agTab==a[0]?'sel':'')+'" onclick="agTab=\''+a[0]+'\';draw()">'+a[1]+'</button>'}).join("")+'</div>'
    +(window.Notification&&Notification.permission!="granted"?'<button class="btn alt" onclick="ativarNotif()">🔔 Ativar notificações neste aparelho</button>':'')+'</div>';
  if(agTab=="novo")return h+agFormHtml();
  if(agTab=="env")return h+agEnviadosHtml();
  return h+agParaMimHtml(uid);
}
function agParaMimHtml(uid){
  var l=agMeus().sort(function(a,b){return a.occ-b.occ});
  if(!l.length)return '<div class="card s">Nenhum recado, compromisso ou reunião para você.</div>';
  var grupos={pend:[],prox:[],ok:[]};
  l.forEach(function(x){var s=agStatus(x,uid);(s=="ok"||s=="recusou"?grupos.ok:(s=="pendente"&&x.it.tipo=="compromisso"&&x.occ-Date.now()>2*86400000?grupos.prox:grupos.pend)).push(x)});
  var h="";
  [["pend","Precisa da sua atenção"],["prox","Próximos"],["ok","Resolvidos"]].forEach(function(g){
    var it=grupos[g[0]];if(!it.length)return;
    h+='<div class="card"><b>'+g[1]+' ('+it.length+')</b>'+it.map(function(x){return agItemMim(x,uid)}).join("")+'</div>';
  });
  return h;
}
function agItemMim(x,uid){
  var it=x.it,occ=x.occ,tp=AG_TIPOS[it.tipo],s=agStatus(x,uid);
  var cor={pendente:"run",atrasado:"no",confirmar:"run",ok:"ok",recusou:"pen"}[s];
  var rot={pendente:it.tipo=="compromisso"?"Pendente":"Dar ciente",atrasado:"Atrasado",confirmar:"Confirmar presença",ok:it.tipo=="compromisso"?"Concluído":it.tipo=="reuniao"?"Presença confirmada":"Ciente",recusou:"Não irei"}[s];
  var ab=agAberto==it.id;
  var h='<div style="border-top:1px solid var(--line);padding:10px 0"><div class="row" style="border:0;padding:0;cursor:pointer" onclick="agAberto='+(ab?'null':it.id)+';draw()"><span style="flex:1">'+tp.ico+' <b>'+esc(it.titulo)+'</b><div class="s">'+(it.tipo=="recado"?'Enviado ':it.tipo=="reuniao"?'':'Prazo: ')+esc(fmtDH(occ))+' · de '+esc(it.autorNome||"")+(it.recorr&&it.recorr!="nao"?' · 🔁 '+AG_REC[it.recorr]:'')+'</div></span><span class="pill '+cor+'" style="cursor:pointer">'+rot+'</span></div>';
  if(ab){
    h+=(it.msg?'<div style="white-space:pre-wrap;margin:8px 0">'+esc(it.msg)+'</div>':'')
      +(it.local?'<div class="s">📍 '+esc(it.local)+'</div>':'')
      +(it.link?'<div class="s">🔗 <a href="'+esc(urlSegura(it.link))+'" target="_blank" rel="noopener">'+esc(it.link)+'</a></div>':'')
      +(it.tipo=="reuniao"?'<div class="s">👥 '+agDestIds(it).length+' participante(s)'+(it.fim?' · até '+hm(occ+agDur(it)):'')+'</div>':'')
      +'<div class="chips" style="margin-top:8px">';
    if(it.tipo=="recado"&&it.ciente&&s!="ok")h+='<button class="sel" onclick="agMarcar('+it.id+','+occ+',\'lido\',Date.now())">✔ Ciente</button>';
    if(it.tipo=="compromisso"&&s!="ok")h+='<button class="sel" onclick="agMarcar('+it.id+','+occ+',\'concl\',Date.now())">✔ Concluí</button>';
    if(it.tipo=="reuniao")h+='<button class="'+(agResp(it,uid,occ)=="sim"?'sel':'')+'" onclick="agMarcar('+it.id+','+occ+',\'resp\',\'sim\')">👍 Vou participar</button><button class="'+(agResp(it,uid,occ)=="nao"?'sel':'')+'" onclick="agMarcar('+it.id+','+occ+',\'resp\',\'nao\')">🚫 Não poderei</button>';
    if(it.tipo!="recado")h+='<button onclick="agIcs('+it.id+')">📆 Adicionar à minha agenda</button>';
    h+='</div>';
  }
  return h+'</div>';
}
function agEnviadosHtml(){
  var todos=temAcao("gerenciar_usuarios");
  var l=AGENDA.filter(function(it){return todos||it.autorId==session.userId}).sort(function(a,b){return b.criado-a.criado});
  if(!l.length)return '<div class="card s">Nada enviado ainda. Use ➕ Novo.</div>';
  return '<div class="card"><b>'+(todos?'Todos os envios':'Enviados por você')+' ('+l.length+')</b>'+campoBusca("agenv","Buscar por título ou destinatário")+'</div>'
    +l.filter(function(it){return bate("agenv",it.titulo,it.msg,agDestResumo(it.dest||{}),it.autorNome)}).slice(0,60).map(function(it){
    var tp=AG_TIPOS[it.tipo],ids=agDestIds(it),occ=agOcorrAtual(it,0)||it.quando;
    var lidos=ids.filter(function(u){return agCiente(it,u,occ)}).length;
    var conc=ids.filter(function(u){return agConcluido(it,u,occ)}).length;
    var sim=ids.filter(function(u){return agResp(it,u,occ)=="sim"}).length,nao=ids.filter(function(u){return agResp(it,u,occ)=="nao"}).length;
    var resumo=it.tipo=="reuniao"?sim+" confirmaram · "+nao+" não poderão · "+(ids.length-sim-nao)+" sem resposta":it.tipo=="compromisso"?conc+"/"+ids.length+" concluíram":(it.ciente?lidos+"/"+ids.length+" deram ciente":ids.length+" destinatário(s)");
    var ab=agAberto==it.id;
    var h='<div class="card'+(it.cancelado?'" style="opacity:.55':'')+'"><div class="row" style="border:0;padding:0;cursor:pointer" onclick="agAberto='+(ab?'null':it.id)+';draw()"><span style="flex:1">'+tp.ico+' <b>'+esc(it.titulo)+'</b>'+(it.cancelado?' <span class="pill no" style="cursor:default">cancelado</span>':'')+'<div class="s">'+esc(fmtDH(occ))+(it.recorr&&it.recorr!="nao"?' · 🔁 '+AG_REC[it.recorr]:'')+' · para '+esc(agDestResumo(it.dest||{}))+'</div><div class="s"><b>'+esc(resumo)+'</b>'+(todos&&it.autorId!=session.userId?' · por '+esc(it.autorNome):'')+'</div></span><span class="pill pen">'+(ab?'▲':'▼')+'</span></div>';
    if(ab){
      h+=(it.msg?'<div style="white-space:pre-wrap;margin:8px 0">'+esc(it.msg)+'</div>':'')
        +'<div class="s" style="margin-top:6px"><b>Destinatários</b></div>'
        +ids.map(function(uid){var u=userById(uid)||{};var st=it.tipo=="reuniao"?({sim:"👍 vai",nao:"🚫 não vai"}[agResp(it,uid,occ)]||"⏳ sem resposta"):it.tipo=="compromisso"?(agConcluido(it,uid,occ)?"✅ concluiu":(occ<Date.now()?"⚠️ atrasado":"⏳ pendente")):(agCiente(it,uid,occ)?"✅ ciente "+new Date(it.lido[agK(uid,occ)]).toLocaleString("pt-BR"):"⏳ ainda não viu");
          return '<div class="row" style="padding:6px 0"><span style="flex:1">'+esc(u.nome||"?")+'<div class="s">'+st+'</div></span>'+(it.cancelado?'':'<button class="pill ok" style="border:0;cursor:pointer" onclick="agZap('+it.id+','+uid+')" title="'+(foneZap(u)?'Enviar por WhatsApp':'Sem telefone no perfil — abre o WhatsApp para escolher o contato')+'">💬</button>')+'</div>'}).join("");
      if(!it.cancelado)h+='<div class="chips" style="margin-top:8px"><button onclick="agZap('+it.id+')">💬 WhatsApp (grupo)</button>'+(it.tipo!="recado"?'<button onclick="agIcs('+it.id+')">📆 Arquivo de agenda</button>':'')+(it.tipo=="reuniao"?'<button onclick="agListaPresenca('+it.id+')">🖨️ Lista de presença</button>':'')+'<button onclick="agCancelar('+it.id+')">✕ Cancelar</button></div>';
    }
    return h+'</div>';
  }).join("");
}
function agFormHtml(){
  var f=AGF,tp=AG_TIPOS[f.tipo];
  var h='<div class="card"><b>➕ Novo</b><div class="chips">'+Object.keys(AG_TIPOS).map(function(k){return '<button class="'+(f.tipo==k?'sel':'')+'" onclick="AGF.tipo=\''+k+'\';draw()">'+AG_TIPOS[k].ico+' '+AG_TIPOS[k].nome+'</button>'}).join("")+'</div><div class="s">'+esc(tp.desc)+'</div>'
    +'<input placeholder="Título (ex.: '+(f.tipo=="reuniao"?'Reunião mensal de supervisores':f.tipo=="compromisso"?'Entregar mapa de frequência':'Mudança de horário do refeitório')+')" value="'+esc(f.titulo)+'" oninput="AGF.titulo=this.value">'
    +'<textarea rows="4" placeholder="'+(f.tipo=="reuniao"?'Pauta':'Mensagem')+'" oninput="AGF.msg=this.value" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin-bottom:8px">'+esc(f.msg)+'</textarea>'
    +'<div class="s">'+(f.tipo=="recado"?'Enviar em':f.tipo=="reuniao"?'Data e início':'Prazo')+'</div>'
    +'<div style="display:flex;gap:6px"><input type="date" value="'+esc(f.data)+'" onchange="AGF.data=this.value"><input type="time" value="'+esc(f.hora)+'" onchange="AGF.hora=this.value"></div>';
  if(f.tipo=="reuniao")h+='<div class="s">Término (opcional)</div><input type="time" value="'+esc(f.horaFim)+'" onchange="AGF.horaFim=this.value"><input placeholder="Local (ex.: Sala de reunião — 4º andar)" value="'+esc(f.local)+'" oninput="AGF.local=this.value"><input placeholder="Link da reunião online (opcional)" value="'+esc(f.link)+'" oninput="AGF.link=this.value">';
  if(f.tipo!="recado")h+='<div class="s">Lembretes automáticos</div><div>'+AG_LEMB.map(function(l){return chip(f.lemb.indexOf(l[0])>=0,"agTog(\'lemb\',"+l[0]+")",l[1])}).join("")+'</div>';
  else h+='<label class="s" style="display:flex;align-items:center;gap:8px;margin:6px 0"><input type="checkbox" '+(f.ciente?'checked':'')+' onchange="AGF.ciente=this.checked" style="width:auto;margin:0"> Pedir "ciente" de quem receber</label>';
  h+='<div class="s" style="margin-top:8px">Repetição</div><select onchange="AGF.recorr=this.value;draw()">'+Object.keys(AG_REC).map(function(k){return '<option value="'+k+'"'+(f.recorr==k?' selected':'')+'>'+AG_REC[k]+'</option>'}).join("")+'</select>'
    +(f.recorr!="nao"?'<div class="s">Repetir até (opcional)</div><input type="date" value="'+esc(f.ate)+'" onchange="AGF.ate=this.value;draw()">'+agPrevia():'')
    +'</div>';

  /* destinatários */
  var prev={dest:{users:f.users,cargos:f.cargos,cts:f.cts}};
  var n=agDestIds(prev).length;
  h+='<div class="card"><b>'+(f.tipo=="reuniao"?'👥 Participantes':'🎯 Para quem')+'</b> <span class="pill '+(n?'ok':'pen')+'" style="cursor:default">'+n+' pessoa(s)</span>'
    +'<div class="s" style="margin-top:8px"><b>Pessoas</b></div><div>'+f.users.map(function(id){var u=userById(id);return u?chip(true,"agTog(\'users\',"+id+")",u.nome+" ✕"):""}).join("")+'</div>'
    +campoBusca("agu","Buscar pessoa por nome");
  if(termo("agu")){
    var r=USERS.filter(function(u){return !u.bloqueado&&f.users.indexOf(u.id)<0&&bate("agu",u.nome,u.mat,(cargoById(u.cargoId)||{}).nome)});
    h+=r.length?r.slice(0,10).map(function(u){return '<div class="row" style="padding:6px 0;cursor:pointer" onclick="agAddUser('+u.id+')"><span style="flex:1">'+marca(u.nome,"agu")+'<div class="s">'+esc((cargoById(u.cargoId)||{}).nome||"")+'</div></span><span class="pill ok">Incluir</span></div>'}).join(""):'<div class="s">Ninguém encontrado. Só aparecem pessoas com acesso ao app.</div>';
  }
  h+='<div class="s" style="margin-top:10px"><b>Por cargo</b> (todos com o cargo)</div><div>'+CARGOS.map(function(c){return chip(f.cargos.indexOf(c.id)>=0,"agTog(\'cargos\',\'"+c.id+"\')",c.nome)}).join("")+'</div>'
    +'<div class="s" style="margin-top:10px"><b>Por contrato</b> (supervisão e equipe do contrato)</div><div>'+f.cts.map(function(id){var c=ctById(id);return c?chip(true,"agTog(\'cts\',"+id+")",c.codigo+" · "+c.nome+" ✕"):""}).join("")+'</div>'
    +campoBusca("agc","Buscar contrato por código, nome ou supervisor");
  if(termo("agc")){
    var rc=meusContratos().filter(function(c){return f.cts.indexOf(c.id)<0&&bate("agc",c.codigo,c.nome,nomesSupCt(c.id))});
    h+=rc.length?rc.slice(0,10).map(function(c){return '<div class="row" style="padding:6px 0;cursor:pointer" onclick="agAddCt('+c.id+')"><span style="flex:1">'+esc(tiposCtIco(c))+' '+marca(c.codigo+" · "+c.nome,"agc")+'</span><span class="pill ok">Incluir</span></div>'}).join(""):'<div class="s">Nenhum contrato encontrado.</div>';
  }
  h+='<div class="s" style="margin-top:8px">Só recebem no app as pessoas com acesso ao RAD Digital. Depois de salvar, use 💬 WhatsApp em 📤 Enviados para avisar também fora do app.</div>'
    +'<button class="btn" onclick="agSalvar()">'+(f.tipo=="reuniao"?'Marcar reunião':f.tipo=="compromisso"?'Criar compromisso':'Programar recado')+'</button></div>';
  return h;
}
function agBanner(){
  if(!session||acessosDoUsuario().indexOf("ag")<0)return "";
  var p=agPendentes();if(!p.length)return "";
  var at=p.filter(function(x){return agStatus(x,session.userId)=="atrasado"}).length;
  var hoje_=p.filter(function(x){return x.it.tipo=="reuniao"&&dia(x.occ)==dia(Date.now())}).length;
  return '<div class="card noprint" style="border-left:4px solid '+(at?'var(--warn)':'var(--brand)')+';cursor:pointer" onclick="agTab=\'mim\';go(\'ag\')"><div style="display:flex;align-items:center;gap:10px"><span style="font-size:24px">📅</span><span style="flex:1"><b>'+p.length+' item(ns) na sua agenda precisam de atenção</b><div class="s">'+(at?at+' compromisso(s) atrasado(s) · ':'')+(hoje_?hoje_+' reunião(ões) hoje · ':'')+'Toque para ver.</div></span></div></div>';
}

/* ===================================================================
   ESCOPO POR CONTRATO (cliente e equipes veem só o que é deles)
   =================================================================== */
function verTudo(){var cg=cargoAtual();return !!cg&&(cg.fixo||cg.id=="admin"||(cg.acoes||[]).indexOf("ver_todos_contratos")>=0)}
function meusCtIds(){return meusContratos().map(function(c){return c.id})}
function ctDaArea(l){return LOC[l]?LOC[l].ct:null}
function regDoMeuEscopo(r){if(verTudo())return true;if(r.origem=="giro"||r.l==null||r.l<0)return !ehClienteCargo();return meusCtIds().indexOf(ctDaArea(r.l))>=0}
function colDoMeuEscopo(c){if(verTudo())return true;var ids=meusCtIds();return (c.vinculos||[]).some(function(v){return ids.indexOf(v.contratoId)>=0})}
function ehClienteCargo(){var cg=cargoAtual();return !!cg&&cg.id=="cli"}

/* ===================================================================
   INDICADORES MENSAIS (automáticos + registro dos meses anteriores)
   =================================================================== */
var IND_DEF=[
  ["conf","Conformidade geral","%",90],
  ["confC","Conformidade — áreas críticas","%",95],
  ["confS","Conformidade — áreas semicríticas","%",90],
  ["confN","Conformidade — áreas não críticas","%",85],
  ["insp","Inspeções realizadas","",null],
  ["limp","Limpezas realizadas","",null],
  ["rec","Recusas / não realizadas","",null],
  ["epi","Uso de EPI confirmado","%",100],
  ["ncA","Não conformidades abertas","",null],
  ["ncF","Não conformidades encerradas","",null],
  ["chA","Chamados abertos","",null],
  ["recl","Reclamações do cliente","",null],
  ["elog","Elogios do cliente","",null]
];
var IND_DEF_BASE=IND_DEF;
function indDefCt(ctId){var c=ctById(ctId),m=(c&&c.metas)||{};return IND_DEF_BASE.map(function(d){return [d[0],d[1],d[2],m[d[0]]!=null&&m[d[0]]!==""?+m[d[0]]:d[3]]})}
var IND_MAN=ld("rad_indman",{});
function svIndMan(){sv("rad_indman",IND_MAN)}
function mesIni(m){var p=m.split("-");return new Date(+p[0],+p[1]-1,1).getTime()}
function mesFim(m){var p=m.split("-");return new Date(+p[0],+p[1],1).getTime()-1}
function mesNome(m){var p=m.split("-");return ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"][+p[1]-1]+"/"+p[0].slice(2)}
function mesesAte(n){var r=[],d=new Date();d.setDate(1);for(var i=0;i<n;i++){r.unshift(d.getFullYear()+"-"+p2(d.getMonth()+1));d.setMonth(d.getMonth()-1)}return r}
function media(a){return a.length?Math.round(a.reduce(function(s,x){return s+x},0)/a.length*10)/10:null}
/* números calculados pelo sistema para um contrato num intervalo */
function indCalc(ctId,de,ate){
  var areas=areasDoContrato(ctId).concat(LOC.map(function(L,i){return L.del&&L.ct==ctId?i:-1}).filter(function(i){return i>=0}));
  var naArea=function(l){return areas.indexOf(l)>=0};
  var av=AVALS.filter(function(a){return a.ini>=de&&a.ini<=ate&&naArea(a.l)});
  var porCls={C:[],S:[],N:[]};av.forEach(function(a){var k=aclOf(a.l,a.e);if(porCls[k])porCls[k].push(a.s.p||0)});
  var rg=regs.filter(function(r){return r.ini>=de&&r.ini<=ate&&naArea(r.l)});
  var ok=rg.filter(function(r){return r.st=="ok"&&!r.x});
  var comEpi=ok.filter(function(r){return r.epi=="sim"||r.epi=="não"});
  var pas=pa.filter(function(p){return naArea(p.l)});
  var ch=chs.filter(function(c){return c.ini>=de&&c.ini<=ate&&c.l!=null&&naArea(c.l)});
  var r={
    conf:media(av.map(function(a){return a.s.p||0})),confC:media(porCls.C),confS:media(porCls.S),confN:media(porCls.N),
    insp:av.length,limp:ok.length,rec:rg.filter(function(r){return r.st=="nao"}).length,
    epi:comEpi.length?Math.round(comEpi.filter(function(r){return r.epi=="sim"}).length/comEpi.length*1000)/10:null,
    ncA:pas.filter(function(p){return p.ini>=de&&p.ini<=ate}).length,ncF:pas.filter(function(p){return p.fim&&p.fim>=de&&p.fim<=ate}).length,
    chA:ch.length,recl:ch.filter(function(c){return c.canal=="CLI"&&/reclama/i.test(c.tipo)}).length,elog:ch.filter(function(c){return c.canal=="CLI"&&/elogio/i.test(c.tipo)}).length
  };
  var temDados=av.length||rg.length||ch.length||pas.some(function(p){return p.ini>=de&&p.ini<=ate});
  return temDados?r:null;
}
function indValor(ctId,mes){
  var man=IND_MAN[ctId+"|"+mes],auto=indCalc(ctId,mesIni(mes),mesFim(mes));
  var v={},orig={};
  IND_DEF.forEach(function(d){var k=d[0];
    if(man&&man[k]!==undefined&&man[k]!==""&&man[k]!==null){v[k]=+man[k];orig[k]="manual"}
    else if(auto&&auto[k]!=null){v[k]=auto[k];orig[k]="auto"}
    else{v[k]=null;orig[k]=""}
  });
  return {v:v,orig:orig,man:man,temAuto:!!auto};
}
function fmtInd(v,un){return v==null?'—':(String(Math.round(v*10)/10).replace(".",",")+(un=="%"?"%":""))}
function corMeta(v,d){if(v==null||d[3]==null)return "";return v>=d[3]?"var(--ok)":(v>=d[3]-10?"var(--run)":"var(--warn)")}
var indCt=null,indMesEd="",indAba="ind";
function podeRegistrarInd(){return temAcao("registrar_indicadores")}
function indSalvar(ctId){
  if(!podeRegistrarInd()){alert("Sem permissão.");return}
  var m=($("#ind_mes")||{}).value;if(!m){alert("Escolha o mês.");return}
  if(mesIni(m)>Date.now()){alert("Não é possível registrar mês futuro.");return}
  var o={};IND_DEF.forEach(function(d){var el=document.getElementById("ind_"+d[0]);var x=el?String(el.value).replace(",",".").trim():"";if(x!==""){var n=+x;if(isNaN(n)||n<0||(d[2]=="%"&&n>100)){o._erro=d[1]}else o[d[0]]=n}});
  if(o._erro){alert("Valor inválido em: "+o._erro);return}
  if(!Object.keys(o).length){alert("Preencha pelo menos um indicador.");return}
  o.obs=(($("#ind_obs")||{}).value||"").trim();o.por=user;o.em=Date.now();
  IND_MAN[ctId+"|"+m]=o;svIndMan();logAcao("Registrou indicadores do mês",(ctById(ctId)||{}).codigo,m);
  indMesEd="";draw();alert("Indicadores de "+mesNome(m)+" registrados.");
}
function indExcluir(ctId,m){if(!podeRegistrarInd())return;if(!confirm("Excluir o registro manual de "+mesNome(m)+"? Os números calculados pelo sistema continuam."))return;delete IND_MAN[ctId+"|"+m];svIndMan();logAcao("Excluiu indicadores manuais",(ctById(ctId)||{}).codigo,m);draw()}
function indCSV(ctId){
  var c=ctById(ctId);var ms=indMeses(ctId);var IND_DEF=indDefCt(ctId);
  var L=[["Indicador","Meta"].concat(ms.map(mesNome))];
  IND_DEF.forEach(function(d){L.push([d[1],d[3]==null?"":d[3]+(d[2]=="%"?"%":"")].concat(ms.map(function(m){var x=indValor(ctId,m);return x.v[d[0]]==null?"":fmtInd(x.v[d[0]],d[2])+(x.orig[d[0]]=="manual"?" (manual)":"")})))});
  baixar("Indicadores_"+c.codigo+".csv",L.map(function(l){return l.map(function(v){return '"'+String(v).replace(/"/g,'""')+'"'}).join(";")}).join("\n"),"text/csv");
}
function indMeses(ctId){
  var ms=mesesAte(12);
  Object.keys(IND_MAN).forEach(function(k){var p=k.split("|");if(p[0]==String(ctId)&&ms.indexOf(p[1])<0)ms.push(p[1])});
  return ms.sort();
}
function vI(){
  var mc=meusContratos();
  if(!mc.length)return '<div class="card s">Nenhum contrato disponível para você.</div>';
  if(!mc.some(function(c){return c.id==indCt}))indCt=mc[0].id;
  var c=ctById(indCt);var IND_DEF=indDefCt(c.id);
  var h='<div class="card"><b>📈 '+T("menu_ind")+'</b>'+pickContrato("indct",mc,c.id,function(id){indCt=id;indMesEd="";draw()})
    +(podeVerDesempenho()?'<div class="chips" style="margin-top:6px"><button class="'+(indAba!="des"?'sel':'')+'" onclick="indAba=\'ind\';draw()">📈 Indicadores</button><button class="'+(indAba=="des"?'sel':'')+'" onclick="indAba=\'des\';draw()">🏆 Desempenho e ranking</button></div>':'')+'</div>';
  if(indAba=="des"&&podeVerDesempenho())return h+vDesempenho(c);
  /* equipe do contrato */
  var col=COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==c.id})});
  var ativos=col.filter(function(x){return ehAtivo(x.situacao)}).length,afast=col.filter(function(x){return ehAfast(x.situacao)}).length;
  h+='<div class="card"><b>👥 Equipe do contrato</b><div class="row"><span>Total</span><b>'+col.length+'</b></div><div class="row"><span>🟢 Ativos</span><b style="color:var(--ok)">'+ativos+'</b></div><div class="row"><span>🔴 Afastados</span><b style="color:var(--warn)">'+afast+'</b></div></div>';
  /* histórico mensal */
  var ms=indMeses(c.id),vals=ms.map(function(m){return indValor(c.id,m)});
  h+='<div class="card tbl"><b>📊 Indicadores mês a mês</b><div class="s">✍️ = registrado manualmente · demais calculados pelo sistema · cor conforme a meta.</div>'
    +'<table style="margin-top:8px"><tr><th>Indicador</th><th>Meta</th>'+ms.map(function(m){return '<th>'+mesNome(m)+'</th>'}).join("")+'</tr>'
    +IND_DEF.map(function(d){return '<tr><td style="white-space:normal;min-width:160px">'+esc(d[1])+'</td><td>'+(d[3]==null?'—':d[3]+(d[2]=="%"?"%":""))+'</td>'+vals.map(function(x){var v=x.v[d[0]];return '<td style="font-weight:600;color:'+(corMeta(v,d)||"inherit")+'">'+fmtInd(v,d[2])+(x.orig[d[0]]=="manual"?' ✍️':'')+'</td>'}).join("")+'</tr>'}).join("")
    +'</table>'
    +'<div class="grid-acoes"><button class="btn alt" onclick="indCSV('+c.id+')">⬇️ Exportar CSV</button><button class="btn alt" onclick="go(\'r\')">📄 Relatório por criticidade</button></div></div>';
  h+=paCardIndicadores(c.id);
  /* registro de meses anteriores */
  if(podeRegistrarInd()){
    var regs_=ms.filter(function(m){return IND_MAN[c.id+"|"+m]});
    var ed=indMesEd&&IND_MAN[c.id+"|"+indMesEd]||{};
    h+='<div class="card"><b>✍️ Registrar indicadores de meses anteriores</b><div class="s">Use para os meses em que o RAD Digital ainda não era usado, ou para corrigir um mês. Campos em branco usam o valor calculado pelo sistema, quando houver.</div>'
      +'<div class="s" style="margin-top:8px">Mês</div><input id="ind_mes" type="month" max="'+dia(Date.now()).slice(0,7)+'" value="'+esc(indMesEd)+'" onchange="indMesEd=this.value;draw()">'
      +IND_DEF.map(function(d){return '<div class="row" style="padding:4px 0;border:0"><span style="flex:1" class="s">'+esc(d[1])+(d[2]=="%"?' (%)':'')+'</span><input id="ind_'+d[0]+'" inputmode="decimal" style="width:110px;margin:0" value="'+(ed[d[0]]!=null?esc(String(ed[d[0]]).replace(".",",")):'')+'"></div>'}).join("")
      +'<textarea id="ind_obs" rows="2" placeholder="Observação / fonte dos dados (ex.: relatório mensal enviado ao cliente)" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:8px 0">'+esc(ed.obs||"")+'</textarea>'
      +'<button class="btn" onclick="indSalvar('+c.id+')">Salvar mês</button>'
      +(regs_.length?'<div class="s" style="margin-top:10px"><b>Meses registrados manualmente</b></div>'+regs_.map(function(m){var o=IND_MAN[c.id+"|"+m];return '<div class="row" style="padding:6px 0"><span style="flex:1">'+mesNome(m)+'<div class="s">por '+esc(o.por||"")+' em '+new Date(o.em).toLocaleDateString("pt-BR")+(o.obs?' · '+esc(o.obs):'')+'</div></span><button class="pill pen" style="border:0;cursor:pointer" onclick="indMesEd=\''+m+'\';draw()">✏️</button><button class="pill no" style="border:0;cursor:pointer" onclick="indExcluir('+c.id+',\''+m+'\')">✕</button></div>'}).join(""):'')
      +'</div>';
  }
  return h;
}

/* ===================================================================
   RELATÓRIO DETALHADO POR ÁREA E GRAU DE CRITICIDADE
   =================================================================== */
var RP={ct:null,per:"mes",de:"",ate:""};
function rpPeriodo(){
  var d=new Date(),ini,fim=Date.now();
  if(RP.per=="hoje"){d.setHours(0,0,0,0);ini=d.getTime()}
  else if(RP.per=="7d"){ini=Date.now()-7*86400000}
  else if(RP.per=="mesAnt"){var a=new Date(d.getFullYear(),d.getMonth()-1,1);ini=a.getTime();fim=new Date(d.getFullYear(),d.getMonth(),1).getTime()-1}
  else if(RP.per=="custom"&&RP.de){ini=agTs(RP.de,"00:00");fim=RP.ate?agTs(RP.ate,"23:59")+59999:Date.now()}
  else{ini=new Date(d.getFullYear(),d.getMonth(),1).getTime()}
  return [ini,fim];
}
function rpDados(ctId,de,ate){
  var areas=areasDoContrato(ctId);
  var D={C:{areas:[],esp:0,limp:0,rec:0,conf:[],nc:0,semReg:0},S:{areas:[],esp:0,limp:0,rec:0,conf:[],nc:0,semReg:0},N:{areas:[],esp:0,limp:0,rec:0,conf:[],nc:0,semReg:0}};
  var rgAll=regs.filter(function(r){return r.ini>=de&&r.ini<=ate});
  var avAll=AVALS.filter(function(a){return a.ini>=de&&a.ini<=ate});
  var porColab={};
  areas.forEach(function(l){
    var L=LOC[l];var linhas={C:[],S:[],N:[]};
    L.e.forEach(function(nome,e){
      if(espDel(l,e))return;
      var k=aclOf(l,e);if(!D[k])k="N";
      var rg=rgAll.filter(function(r){return r.l==l&&r.e==e&&r.origem!="giro"});
      var ok=rg.filter(function(r){return r.st=="ok"});var nao=rg.filter(function(r){return r.st=="nao"});
      var av=avAll.filter(function(a){return a.l==l&&a.e==e});
      var ncs=pa.filter(function(p){return p.l==l&&p.e==e&&p.st!="Concluído"});
      var epiC=ok.filter(function(r){return r.epi=="sim"||r.epi=="não"});
      var tm=ok.length?Math.round(ok.reduce(function(s,r){return s+mins(r)},0)/ok.length):null;
      var ult=ok.length?Math.max.apply(null,ok.map(function(r){return r.fim||r.ini})):null;
      var cf=av.length?media(av.map(function(a){return a.s.p||0})):null;
      ok.forEach(function(r){var n=r.colab||"?";porColab[n]=(porColab[n]||0)+1});
      linhas[k].push({nome:nome,pausado:espOff(l,e),limp:ok.length,rec:nao.length,tm:tm,ult:ult,conf:cf,insp:av.length,nc:ncs.length,epi:epiC.length?Math.round(epiC.filter(function(r){return r.epi=="sim"}).length/epiC.length*100):null});
      D[k].esp++;D[k].limp+=ok.length;D[k].rec+=nao.length;if(cf!=null)D[k].conf.push(cf);D[k].nc+=ncs.length;if(!ok.length&&!espOff(l,e))D[k].semReg++;
    });
    ["C","S","N"].forEach(function(k){if(linhas[k].length)D[k].areas.push({l:l,nome:L.n,tipo:tipoDaArea(l),linhas:linhas[k]})});
  });
  return {D:D,porColab:porColab,giros:rgAll.length};
}
function rpHtml(ctId,de,ate,impressao){
  var c=ctById(ctId),R=rpDados(ctId,de,ate),D=R.D;
  var tot={esp:0,limp:0,rec:0,nc:0,semReg:0};["C","S","N"].forEach(function(k){["esp","limp","rec","nc","semReg"].forEach(function(f){tot[f]+=D[k][f]})});
  var confG=media([].concat(D.C.conf,D.S.conf,D.N.conf));
  var h='';
  h+=(impressao?'<h3>Resumo do período</h3><table>':'<div class="card"><b>📊 Resumo</b><div class="s">'+esc(c.codigo+" · "+c.nome)+' · '+new Date(de).toLocaleDateString("pt-BR")+' a '+new Date(ate).toLocaleDateString("pt-BR")+'</div><div class="tbl"><table style="margin-top:8px">')
    +'<tr><th>Criticidade</th><th>Espaços</th><th>Limpezas</th><th>Recusas</th><th>Sem registro</th><th>Conformidade</th><th>NC abertas</th></tr>'
    +["C","S","N"].map(function(k){var x=D[k],cf=media(x.conf);return '<tr><td>'+(impressao?crTag(k):'<b style="color:'+acCor(k)+'">'+AL[k]+'</b>')+'</td><td>'+x.esp+'</td><td>'+x.limp+'</td><td>'+x.rec+'</td><td'+(x.semReg&&k=="C"?' style="color:#c62828;font-weight:700"':'')+'>'+x.semReg+'</td><td>'+fmtInd(cf,"%")+'</td><td>'+x.nc+'</td></tr>'}).join("")
    +'<tr><td><b>Total</b></td><td><b>'+tot.esp+'</b></td><td><b>'+tot.limp+'</b></td><td><b>'+tot.rec+'</b></td><td><b>'+tot.semReg+'</b></td><td><b>'+fmtInd(confG,"%")+'</b></td><td><b>'+tot.nc+'</b></td></tr>'
    +(impressao?'</table>':'</table></div></div>');
  ["C","S","N"].forEach(function(k){
    var x=D[k];if(!x.areas.length)return;
    h+=impressao?'<h3 style="background:'+acCor(k)+'">Áreas '+AL[k].toLowerCase()+'s — '+x.esp+' espaço(s)</h3>':'<div class="card" style="border-left:5px solid '+acCor(k)+'"><b style="color:'+acCor(k)+'">'+AL[k]+'</b> <span class="s">· '+x.esp+' espaço(s) · '+x.limp+' limpeza(s) · conformidade '+fmtInd(media(x.conf),"%")+'</span>';
    x.areas.forEach(function(a){
      h+=(impressao?'<div class="nobr"><h4>'+esc(a.nome)+'</h4>':'<div style="margin-top:10px"><b>'+esc(a.nome)+'</b>')
        +(impressao?'<table>':'<div class="tbl"><table style="margin-top:4px">')
        +'<tr><th>Espaço</th><th>Limpezas</th><th>Recusas</th><th>Tempo médio</th><th>Última limpeza</th><th>Inspeções</th><th>Conformidade</th><th>NC abertas</th><th>EPI</th></tr>'
        +a.linhas.map(function(r){var alerta=!r.limp&&!r.pausado;return '<tr'+(alerta?' style="background:'+(k=="C"?'rgba(198,40,40,.12)':'rgba(224,161,0,.12)')+'"':'')+'><td style="white-space:normal">'+esc(r.nome)+(r.pausado?' <span class="s">(pausado)</span>':'')+'</td><td>'+(r.limp||(alerta?'<b style="color:#c62828">0</b>':'0'))+'</td><td>'+r.rec+'</td><td>'+(r.tm==null?'—':r.tm+' min')+'</td><td>'+(r.ult?new Date(r.ult).toLocaleDateString("pt-BR")+' '+hm(r.ult):'—')+'</td><td>'+r.insp+'</td><td>'+fmtInd(r.conf,"%")+'</td><td>'+(r.nc?'<b style="color:#c62828">'+r.nc+'</b>':'0')+'</td><td>'+(r.epi==null?'—':r.epi+'%')+'</td></tr>'}).join("")
        +(impressao?'</table></div>':'</table></div></div>');
    });
    if(!impressao)h+='</div>';
  });
  if(!ehClienteCargo()){
    var nomes=Object.keys(R.porColab).sort(function(a,b){return R.porColab[b]-R.porColab[a]});
    if(nomes.length)h+=(impressao?'<h3>Produtividade por colaborador</h3><table>':'<div class="card"><b>👥 Produtividade por colaborador</b><div class="tbl"><table style="margin-top:6px">')+'<tr><th>Colaborador</th><th>Limpezas</th></tr>'+nomes.map(function(n){return '<tr><td>'+esc(n)+'</td><td>'+R.porColab[n]+'</td></tr>'}).join("")+(impressao?'</table>':'</table></div></div>');
  }
  return h;
}
function vR(){
  var mc=meusContratos();
  if(!mc.length)return '<div class="card s">Nenhum contrato disponível para você.</div>';
  if(!mc.some(function(c){return c.id==RP.ct}))RP.ct=mc[0].id;
  var pr=rpPeriodo();
  var per=[["hoje","Hoje"],["7d","7 dias"],["mes","Este mês"],["mesAnt","Mês anterior"],["custom","Período"]];
  var h='<div class="card noprint"><b>📄 '+T("menu_rel")+'</b><div class="s">Detalhado por área e espaço, separado pelo grau de criticidade.</div>'
    +pickContrato("rpct",mc,RP.ct,function(id){RP.ct=id;draw()})
    +'<div class="chips">'+per.map(function(p){return '<button class="'+(RP.per==p[0]?'sel':'')+'" onclick="RP.per=\''+p[0]+'\';draw()">'+p[1]+'</button>'}).join("")+'</div>'
    +(RP.per=="custom"?'<div style="display:flex;gap:6px"><input type="date" value="'+esc(RP.de)+'" onchange="RP.de=this.value;draw()"><input type="date" value="'+esc(RP.ate)+'" onchange="RP.ate=this.value;draw()"></div>':'')
    +'<div class="grid-acoes"><button class="btn" onclick="imprimirRelatorio()">🖨️ '+T("btn_print")+' / PDF</button><button class="btn alt" onclick="exportarRelatorioCSV()">⬇️ CSV</button></div></div>';
  if(!areasDoContrato(RP.ct).length)return h+'<div class="card s">Este contrato ainda não tem áreas mapeadas.</div>';
  return h+rpHtml(RP.ct,pr[0],pr[1],false)+'<div class="card s">Linhas destacadas: espaços ativos sem nenhuma limpeza registrada no período (vermelho nas áreas críticas).</div>';
}
function imprimirRelatorio(){
  var c=ctById(RP.ct);if(!c)return;var pr=rpPeriodo();
  var h='<div class="doc">'+cabDoc("Relatório de execução por criticidade","Contrato "+c.codigo)
    +'<div class="grid"><div><b>Contrato</b>'+esc(c.codigo+" · "+c.nome)+'</div><div><b>Período</b>'+new Date(pr[0]).toLocaleDateString("pt-BR")+' a '+new Date(pr[1]).toLocaleDateString("pt-BR")+'</div><div><b>Serviços</b>'+esc(tiposCt(c).map(function(t){return tipoCt(t).nome}).join(", "))+'</div></div>'
    +rpHtml(RP.ct,pr[0],pr[1],true)
    +'<p class="mut" style="margin-top:8px">Linhas destacadas: espaços ativos sem limpeza registrada no período. Conformidade = média das inspeções (Luz Negra / checklists) do período.</p>'
    +'<div class="ass"><div>Supervisor</div><div>Coordenação</div><div>Responsável do cliente</div></div>'+rodDoc()+'</div>';
  logAcao("Imprimiu relatório por criticidade",c.codigo);
  abrirImpressao("Relatorio_"+c.codigo,h);
}
function exportarRelatorioCSV(){
  var c=ctById(RP.ct);if(!c)return;var pr=rpPeriodo(),R=rpDados(RP.ct,pr[0],pr[1]);
  var L=[["Criticidade","Área","Espaço","Limpezas","Recusas","Tempo médio (min)","Última limpeza","Inspeções","Conformidade (%)","NC abertas","EPI confirmado (%)"]];
  ["C","S","N"].forEach(function(k){R.D[k].areas.forEach(function(a){a.linhas.forEach(function(r){L.push([AL[k],a.nome,r.nome,r.limp,r.rec,r.tm==null?"":r.tm,r.ult?new Date(r.ult).toLocaleString("pt-BR"):"",r.insp,r.conf==null?"":String(r.conf).replace(".",","),r.nc,r.epi==null?"":r.epi])})})});
  baixar("Relatorio_"+c.codigo+"_"+dia(pr[0])+"_a_"+dia(pr[1])+".csv",L.map(function(l){return l.map(function(x){return '"'+String(x==null?"":x).replace(/"/g,'""')+'"'}).join(";")}).join("\n"),"text/csv");
  logAcao("Exportou relatório por criticidade",c.codigo);
}

/* ===================================================================
   CONTATOS E LINKS: quem pode ver
   =================================================================== */
function contatoVisivel(c){
  var cg=cargoAtual();if(!cg)return false;
  if(cg.fixo||cg.id=="admin")return true;
  var vis=c.vis||"interno";
  var ok=vis=="todos"||(vis=="interno"&&cg.id!="cli")||(Array.isArray(vis)&&vis.indexOf(cg.id)>=0);
  if(!ok)return false;
  if(c.ct&&meusCtIds().indexOf(+c.ct)<0)return false;
  return true;
}
function visTxt(c){var v=c.vis||"interno";if(v=="todos")return "Todos (inclusive cliente)";if(v=="interno")return "Equipe interna (sem cliente)";return "Só: "+v.map(function(id){return (cargoById(id)||{}).nome||id}).join(", ")}
function setVisContato(id,v){if(!temAcao("gerenciar_usuarios"))return;var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c)return;c.vis=v=="cargos"?[]:v;svContatos();var y=window.scrollY;draw();window.scrollTo(0,y)}
function togVisCargo(id,cg){var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c||!Array.isArray(c.vis))return;var p=c.vis.indexOf(cg);if(p>=0)c.vis.splice(p,1);else c.vis.push(cg);svContatos();var y=window.scrollY;draw();window.scrollTo(0,y)}
function setCtContato(id,v){var c=CONTATOS.filter(function(x){return x.id==id})[0];if(!c)return;c.ct=v?+v:null;svContatos();draw()}

/* ===================================================================
   PLANO DE AÇÃO — padrão REG-060 (SGI)
   Junta: NCs das inspeções, ações sugeridas pelos indicadores abaixo da meta,
   ações manuais e planilhas REG-060 importadas.
   =================================================================== */
var REG060={cod:"REG-060",rev:"01"};
function paCt(p){return p.ct!=null?p.ct:ctDaArea(p.l)}
function paPrazo(p){return p.novoPrazo||przTs(p)}
function paSituacao(p){
  if(p.st=="Concluído")return "CONCLUÍDO";
  var t=paPrazo(p);
  return t&&Date.now()>t?"ATRASADA":"EM ANDAMENTO";
}
function paNoPrazo(p){return p.st=="Concluído"&&p.conclusao&&(!paPrazo(p)||p.conclusao<=paPrazo(p)+86399999)}
function paSetor(p){if(p.setor)return p.setor;if(p.l!=null&&LOC[p.l])return nomeLoc(p.l)+(p.e!=null&&p.e>=0?" – "+nomeEsp(p.l,p.e):"");return "—"}
function paItem(p){return p.itemTxt||String(p.item||"").replace(/\s*\[[^\]]*\]$/,"")}
function paOque(p){return p.oque||(p.pts&&p.pts.length?"Corrigir não conformidade: "+p.pts.join(", "):"")}
function paComo(p){return p.como||p.acao||""}
function paQuem(p){return p.quem||p.resp||""}
function paCausador(p){return p.causador||""}
function verCausador(){return !ehClienteCargo()}
function paDoContrato(ctId){return pa.filter(function(p){return paCt(p)==ctId})}
var PAF=null,paFiltro="abertos",paEdit=null;
function paGerenc(lista){
  var r={c:0,a:0,e:0,rep:0,np:0};
  lista.forEach(function(p){var s=paSituacao(p);if(s=="CONCLUÍDO"){r.c++;if(paNoPrazo(p))r.np++}else if(s=="ATRASADA")r.a++;else r.e++;if(p.novoPrazo)r.rep++});
  r.t=r.c+r.a+r.e;return r;
}
function pct1(a,t){return t?Math.round(a/t*1000)/10:0}
/* sugestões de ação para indicador abaixo da meta */
var PA_SUG={
  conf:["Elevar a conformidade geral das inspeções para a meta","Reforço do check-list de limpeza com a equipe, nova inspeção nos espaços reprovados e acompanhamento semanal pela supervisão"],
  confC:["Elevar a conformidade das áreas críticas para a meta","Revisar técnica de limpeza e desinfecção (dois baldes, pano exclusivo, tempo de contato do desinfetante), treinamento prático e inspeção com Luz Negra após a correção"],
  confS:["Elevar a conformidade das áreas semicríticas para a meta","Reforçar a desinfecção das superfícies de contato e reinspecionar os espaços reprovados"],
  confN:["Elevar a conformidade das áreas não críticas para a meta","Ajustar a rotina de limpeza e conferir a execução pelo encarregado"],
  epi:["Garantir 100% de uso de EPI","Diálogo de segurança (DDS) sobre EPI por risco, conferência na entrada do turno e reposição dos EPIs em falta"],
  rec:["Reduzir as atividades não realizadas","Analisar os motivos das recusas com o cliente e reprogramar os horários de limpeza"],
  recl:["Tratar as reclamações do cliente","Analisar a causa de cada reclamação, responder ao cliente e verificar a eficácia da correção"]
};
function paSugestoes(ctId,mes){
  var IND_DEF=indDefCt(ctId);
  var x=indValor(ctId,mes),r=[];
  IND_DEF.forEach(function(d){var v=x.v[d[0]];if(d[3]!=null&&v!=null&&v<d[3]&&PA_SUG[d[0]])r.push({k:d[0],nome:d[1],v:v,meta:d[3],un:d[2]})});
  if(x.v.recl)r.push({k:"recl",nome:"Reclamações do cliente",v:x.v.recl,meta:0,un:""});
  if(x.v.rec>0&&!r.some(function(s){return s.k=="rec"}))r.push({k:"rec",nome:"Recusas / não realizadas",v:x.v.rec,meta:0,un:""});
  /* espaços críticos sem limpeza no mês */
  var D=rpDados(ctId,mesIni(mes),Math.min(mesFim(mes),Date.now())).D;
  if(D.C.semReg)r.push({k:"semC",nome:"Espaços críticos sem limpeza registrada",v:D.C.semReg,meta:0,un:""});
  return r.filter(function(s){return !pa.some(function(p){return paCt(p)==ctId&&p.indicador==s.k&&p.mes==mes&&p.st!="Concluído"})});
}
function paNovoDeSug(ctId,mes,k){
  var s=PA_SUG[k]||(k=="semC"?["Garantir a limpeza de todos os espaços críticos","Revisar a escala e o PTI para cobrir todos os espaços críticos e registrar cada limpeza no RAD Digital"]:["",""]);
  var nome=(IND_DEF.filter(function(d){return d[0]==k})[0]||[0,"Espaços críticos sem limpeza"])[1];
  var pz=new Date();pz.setDate(pz.getDate()+15);
  PAF={ct:ctId,causador:"",setor:/C$/.test(k)||k=="semC"?"Áreas críticas":"Todos os setores do contrato",item:nome+" ("+mesNome(mes)+")",oque:s[0],como:s[1],prazo:dia(pz),quem:nomesSupCt(ctId)||"Supervisão Operacional",obs:"",indicador:k,mes:mes};
  draw();setTimeout(function(){var e=document.getElementById("paf_card");if(e&&e.scrollIntoView)e.scrollIntoView()},50);
}
function paNovo(ctId){var pz=new Date();pz.setDate(pz.getDate()+15);PAF={ct:ctId,causador:"",setor:"",item:"",oque:"",como:"",prazo:dia(pz),quem:"",obs:""};paEdit=null;draw();setTimeout(function(){var e=document.getElementById("paf_card");if(e&&e.scrollIntoView)e.scrollIntoView()},50)}
function paEditar(id){var p=pa.filter(function(x){return x.id==id})[0];if(!p)return;PAF={ct:paCt(p),causador:paCausador(p),setor:paSetor(p),item:paItem(p),oque:paOque(p),como:paComo(p),prazo:p.prazo?dia(przTs(p)):"",quem:paQuem(p),obs:p.obs||""};paEdit=id;draw();setTimeout(function(){var e=document.getElementById("paf_card");if(e&&e.scrollIntoView)e.scrollIntoView()},50)}
function paSalvar(){
  if(!podeRegistrarInd()){alert("Sem permissão.");return}
  var f=PAF;if(!f.oque.trim()){alert("Informe o que precisa ser feito.");return}
  if(!f.prazo){alert("Informe o prazo de conclusão.");return}
  if(!f.quem.trim()){alert("Informe quem vai resolver.");return}
  var cz=(f.causador||"").trim(),czCol=COL.filter(function(c){return normT(c.nome)==normT(cz)})[0];
  var dados={ct:f.ct,causador:cz,causadorId:czCol?czCol.id:null,setor:f.setor.trim(),itemTxt:f.item.trim(),oque:f.oque.trim(),como:f.como.trim(),prazo:agTs(f.prazo,"23:59"),quem:f.quem.trim(),obs:f.obs.trim()};
  if(paEdit){var p=pa.filter(function(x){return x.id==paEdit})[0];if(p){Object.keys(dados).forEach(function(k){p[k]=dados[k]});p.resp=dados.quem;p.acao=dados.como}logAcao("Editou ação do plano",dados.oque)}
  else{dados.id=Date.now();dados.ini=Date.now();dados.st="Aberto";dados.origem=f.indicador?"indicador":"manual";dados.indicador=f.indicador||null;dados.mes=f.mes||null;dados.l=null;dados.e=null;dados.resp=dados.quem;dados.acao=dados.como;dados.item=dados.itemTxt;pa.push(dados);logAcao("Criou ação no plano (REG-060)",dados.oque)}
  sv("rad_pa",pa);PAF=null;paEdit=null;draw();
}
function paConcluir(id){
  if(!podeRegistrarInd()){alert("Sem permissão.");return}
  var p=pa.filter(function(x){return x.id==id})[0];if(!p)return;
  var d=prompt("Data de conclusão (dd/mm/aaaa):",new Date().toLocaleDateString("pt-BR"));if(d===null)return;
  var m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(d.trim());if(!m){alert("Data inválida.");return}
  var obs=prompt("Observação / verificação de eficácia (opcional):",p.obs||"");
  p.st="Concluído";p.conclusao=new Date(+m[3],+m[2]-1,+m[1],12).getTime();p.fim=p.conclusao;if(obs!==null)p.obs=obs.trim();
  sv("rad_pa",pa);logAcao("Concluiu ação do plano",paOque(p));draw();
}
function paReprogramar(id){
  if(!podeRegistrarInd()){alert("Sem permissão.");return}
  var p=pa.filter(function(x){return x.id==id})[0];if(!p)return;
  var d=prompt("Novo prazo de conclusão (dd/mm/aaaa):","");if(d===null)return;
  var m=/^(\d{2})\/(\d{2})\/(\d{4})$/.exec(d.trim());if(!m){alert("Data inválida.");return}
  var mot=prompt("Motivo da reprogramação:","");if(!mot){alert("Informe o motivo.");return}
  p.novoPrazo=new Date(+m[3],+m[2]-1,+m[1],23,59).getTime();p.obs=((p.obs?p.obs+" | ":"")+"Reprogramado: "+mot.trim());
  sv("rad_pa",pa);logAcao("Reprogramou ação do plano",paOque(p));draw();
}
function paReabrir(id){var p=pa.filter(function(x){return x.id==id})[0];if(!p||!podeRegistrarInd())return;if(!confirm("Reabrir esta ação?"))return;p.st="Aberto";p.conclusao=null;sv("rad_pa",pa);draw()}
function paExcluir(id){var p=pa.filter(function(x){return x.id==id})[0];if(!p||!podeRegistrarInd())return;if(!confirm("Excluir esta ação do plano?"))return;pa=pa.filter(function(x){return x.id!=id});sv("rad_pa",pa);logAcao("Excluiu ação do plano",paOque(p));draw()}
function dt(t){return t?new Date(t).toLocaleDateString("pt-BR"):""}
function corSit(s){return s=="CONCLUÍDO"?"ok":s=="ATRASADA"?"no":"run"}

function paCardIndicadores(ctId){
  var lista=paDoContrato(ctId).sort(function(a,b){return (paPrazo(a)||0)-(paPrazo(b)||0)});
  var g=paGerenc(lista),pode=podeRegistrarInd();
  var h='<div class="card"><b>📋 Plano de ação ('+REG060.cod+')</b><div class="s">Não conformidades das inspeções, ações dos indicadores abaixo da meta e ações registradas.</div>'
    +'<div class="tbl"><table style="margin-top:8px"><tr><th>Gerenciamento</th><th>Qtd.</th><th>%</th></tr>'
    +'<tr><td><span class="pill ok" style="cursor:default">CONCLUÍDO</span></td><td>'+g.c+'</td><td>'+String(pct1(g.c,g.t)).replace(".",",")+'%</td></tr>'
    +'<tr><td><span class="pill no" style="cursor:default">ATRASADA</span></td><td>'+g.a+'</td><td>'+String(pct1(g.a,g.t)).replace(".",",")+'%</td></tr>'
    +'<tr><td><span class="pill run" style="cursor:default">EM ANDAMENTO</span></td><td>'+g.e+'</td><td>'+String(pct1(g.e,g.t)).replace(".",",")+'%</td></tr>'
    +'<tr><td>Concluídas no prazo</td><td>'+g.np+'</td><td>'+String(pct1(g.np,g.c)).replace(".",",")+'% das concluídas</td></tr>'
    +'<tr><td>Reprogramadas</td><td>'+g.rep+'</td><td></td></tr></table></div>';
  if(verCausador()){
    var rk={};lista.forEach(function(p){var n=paCausador(p);if(!n||n=="Não identificado")return;rk[n]=rk[n]||{t:0,a:0};rk[n].t++;if(paSituacao(p)!="CONCLUÍDO")rk[n].a++});
    var nm=Object.keys(rk).sort(function(a,b){return rk[b].t-rk[a].t}).slice(0,8);
    if(nm.length)h+='<div class="s" style="margin-top:10px"><b>👤 Não conformidades por colaborador</b> (para orientação e treinamento)</div><div class="tbl"><table><tr><th>Colaborador</th><th>Total</th><th>Em aberto</th></tr>'+nm.map(function(n){return '<tr><td>'+esc(n)+'</td><td>'+rk[n].t+'</td><td>'+rk[n].a+'</td></tr>'}).join("")+'</table></div>';
  }
  /* sugestões do mês atual e do anterior */
  if(pode){
    var ms=mesesAte(2),sug=[];
    ms.forEach(function(m){paSugestoes(ctId,m).forEach(function(s){s.mes=m;sug.push(s)})});
    if(sug.length)h+='<div style="border:2px dashed var(--run);border-radius:10px;padding:10px;margin-top:10px"><b>⚠️ Indicadores pedindo ação</b>'
      +sug.map(function(s){return '<div class="row" style="padding:6px 0"><span style="flex:1">'+esc(s.nome)+' <span class="s">· '+mesNome(s.mes)+' · '+fmtInd(s.v,s.un)+(s.meta?' (meta '+s.meta+s.un+')':'')+'</span></span><button class="pill run" style="border:0;cursor:pointer" onclick="paNovoDeSug('+ctId+',\''+s.mes+'\',\''+s.k+'\')">Gerar ação</button></div>'}).join("")+'</div>';
  }
  h+='<div class="chips" style="margin-top:10px">'+[["abertos","Em aberto"],["atrasadas","Atrasadas"],["concluidas","Concluídas"],["todas","Todas"]].map(function(f){return '<button class="'+(paFiltro==f[0]?'sel':'')+'" onclick="paFiltro=\''+f[0]+'\';draw()">'+f[1]+'</button>'}).join("")+'</div>'
    +campoBusca("pa","Buscar por setor, item, causador ou responsável");
  var vis=lista.filter(function(p){var s=paSituacao(p);return (paFiltro=="todas"||(paFiltro=="abertos"&&s!="CONCLUÍDO")||(paFiltro=="atrasadas"&&s=="ATRASADA")||(paFiltro=="concluidas"&&s=="CONCLUÍDO"))&&bate("pa",paSetor(p),paItem(p),paOque(p),paQuem(p),verCausador()?paCausador(p):"")});
  if(!vis.length)h+='<div class="s">Nenhuma ação neste filtro.</div>';
  h+=vis.slice(0,80).map(function(p){var s=paSituacao(p);return '<div style="border-top:1px solid var(--line);padding:8px 0">'
    +'<div class="row" style="border:0;padding:0"><span style="flex:1"><b>'+esc(paSetor(p))+'</b> <span class="s">· '+esc(paItem(p))+'</span></span><span class="pill '+corSit(s)+'" style="cursor:default">'+s+'</span></div>'
    +'<div style="margin-top:4px"><b class="s">O que:</b> '+esc(paOque(p))+'</div>'
    +(paComo(p)?'<div><b class="s">Como:</b> '+esc(paComo(p))+'</div>':'')
    +(verCausador()&&paCausador(p)?'<div class="s">👤 Causador da NC: <b>'+esc(paCausador(p))+'</b></div>':'')
    +'<div class="s">Prazo '+dt(przTs(p))+(p.novoPrazo?' → novo prazo <b>'+dt(p.novoPrazo)+'</b>':'')+' · 🛠️ Quem vai resolver: <b>'+esc(paQuem(p)||"—")+'</b>'+(p.conclusao?' · Concluída em '+dt(p.conclusao)+(paNoPrazo(p)?' (no prazo)':' (fora do prazo)'):'')+' · Origem: '+esc(p.origem||"inspeção")+'</div>'
    +(p.obs?'<div class="s">Obs.: '+esc(p.obs)+'</div>':'')
    +(pode?'<div class="chips" style="margin-top:4px">'+(s!="CONCLUÍDO"?'<button class="sel" onclick="paConcluir('+p.id+')">✔ Concluir</button><button onclick="paReprogramar('+p.id+')">📅 Reprogramar</button>':'<button onclick="paReabrir('+p.id+')">↺ Reabrir</button>')+'<button onclick="paEditar('+p.id+')">✏️</button><button onclick="paExcluir('+p.id+')">✕</button></div>':'')
    +'</div>'}).join("");
  if(vis.length>80)h+='<div class="s">Mostrando 80 de '+vis.length+'. Use a busca.</div>';
  if(pode){
    h+='<div class="grid-acoes" style="margin-top:8px"><button class="btn alt" onclick="paNovo('+ctId+')">➕ Nova ação</button><label class="btn alt" style="text-align:center;margin:0">📥 Importar REG-060 (.xlsx)<input type="file" accept=".xlsx,.xls" style="display:none" onchange="paImportar(this,'+ctId+')"></label></div>';
  }
  h+='<div class="grid-acoes" style="margin-top:8px"><button class="btn" onclick="imprimirIndicadores('+ctId+')">🖨️ Relatório de indicadores + plano</button><button class="btn alt" onclick="paExportarXlsx('+ctId+')">⬇️ REG-060 em Excel</button></div></div>';
  if(PAF&&PAF.ct==ctId&&pode){
    var areasDl=locaisCt(ctById(ctId)).map(function(i){return '<option value="'+esc(LOC[i].n)+'">'}).join("");
    h+='<div class="card" id="paf_card" style="border:2px solid var(--brand)"><b>'+(paEdit?'✏️ Editar ação':'➕ Nova ação')+' — '+REG060.cod+'</b>'
      +'<div class="s">Setor</div><input list="dl_setor" value="'+esc(PAF.setor)+'" oninput="PAF.setor=this.value" placeholder="Ex.: Centro Cirúrgico – Sala 03"><datalist id="dl_setor">'+areasDl+'<option value="Todos os setores do contrato"><option value="Áreas críticas"></datalist>'
      +'<div class="s">Item</div><input value="'+esc(PAF.item)+'" oninput="PAF.item=this.value" placeholder="Ex.: Vaso sanitário / Indicador de conformidade">'
      +'<div class="s">O que precisa ser feito</div><textarea rows="2" oninput="PAF.oque=this.value" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin-bottom:8px">'+esc(PAF.oque)+'</textarea>'
      +'<div class="s">Como será feito</div><textarea rows="3" oninput="PAF.como=this.value" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin-bottom:8px">'+esc(PAF.como)+'</textarea>'
      +'<div class="s">Prazo de conclusão</div><input type="date" value="'+esc(PAF.prazo)+'" onchange="PAF.prazo=this.value">'
      +'<div class="s">Colaborador que causou a não conformidade</div><input list="dl_causa" value="'+esc(PAF.causador||"")+'" oninput="PAF.causador=this.value" placeholder="Busque pelo nome (deixe em branco se não se aplica)"><datalist id="dl_causa">'+COL.filter(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==ctId})}).map(function(x){return '<option value="'+esc(x.nome)+'">'}).join("")+'</datalist>'
      +'<div class="s">Quem vai resolver (quem irá realizar)</div><input list="dl_quem" value="'+esc(PAF.quem)+'" oninput="PAF.quem=this.value" placeholder="Nome ou equipe"><datalist id="dl_quem"><option value="Equipe de Higienização"><option value="Supervisão Operacional"><option value="Coordenação e Supervisão">'+supervisoresDoContrato(ctId).map(function(u){return '<option value="'+esc(u.nome)+'">'}).join("")+'</datalist>'
      +'<div class="s">Observação</div><input value="'+esc(PAF.obs)+'" oninput="PAF.obs=this.value">'
      +'<div class="grid-acoes"><button class="btn" onclick="paSalvar()">Salvar ação</button><button class="btn alt" onclick="PAF=null;paEdit=null;draw()">Cancelar</button></div></div>';
  }
  return h;
}

/* ---- importar a planilha REG-060 já usada ---- */
function paImportar(inp,ctId){
  if(!podeRegistrarInd()){alert("Sem permissão.");inp.value="";return}
  var f=inp.files[0];if(!f)return;
  if(typeof XLSX==="undefined"){alert("Biblioteca de Excel não carregou (verifique a internet).");inp.value="";return}
  var r=new FileReader();
  r.onload=function(){
    try{
      var wb=XLSX.read(r.result,{type:"array",cellDates:true});
      var nomeAba=wb.SheetNames.filter(function(n){return /060|plano/i.test(n)})[0]||wb.SheetNames[0];
      var rows=XLSX.utils.sheet_to_json(wb.Sheets[nomeAba],{header:1,defval:"",raw:true});
      var hi=-1;for(var i=0;i<Math.min(rows.length,40);i++){if(rows[i].some(function(x){return /o que precisa/i.test(String(x))})){hi=i;break}}
      if(hi<0){alert("Não encontrei o cabeçalho do REG-060 (coluna \"O QUE PRECISA SER FEITO\").");return}
      var H=rows[hi].map(function(x){return normT(x)});
      var col=function(re){for(var j=0;j<H.length;j++)if(re.test(H[j]))return j;return -1};
      var C={causa:col(/causa|causador|colaborador/),setor:col(/^setor/),item:col(/^item/),oque:col(/o que precisa/),como:col(/como sera/),prazo:col(/^prazo/),novo:col(/novo prazo/),quem:col(/quem/),conc:col(/data de conclus/),st:col(/^status/),obs:col(/observa/)};
      var paraTs=function(v){if(!v)return 0;if(v instanceof Date)return v.getTime();var m=/^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(String(v));if(m)return new Date(+m[3],+m[2]-1,+m[1],12).getTime();if(typeof v=="number"&&v>20000&&v<80000)return Math.round((v-25569)*86400000);return 0};
      var add=0,dup=0;
      rows.slice(hi+1).forEach(function(rw,k){
        var g=function(j){return j>=0&&rw[j]!=null?rw[j]:""};
        var oque=String(g(C.oque)).trim();if(!oque)return;
        var setorT=String(g(C.setor)).trim(),itemT=String(g(C.item)).trim();
        if(pa.some(function(p){return paCt(p)==ctId&&normT(paOque(p))==normT(oque)&&normT(paSetor(p))==normT(setorT)&&normT(paItem(p))==normT(itemT)})){dup++;return}
        var novo=g(C.novo),novoTs=paraTs(novo);
        /* algumas planilhas trazem o responsável na coluna "novo prazo": junta ao "quem" */
        var quem=[novoTs?"":String(novo||"").trim(),String(g(C.quem)).trim()].filter(Boolean).join(" / ");
        var st=normT(g(C.st)),conc=paraTs(g(C.conc));
        pa.push({id:Date.now()+k,ct:ctId,l:null,e:null,causador:String(g(C.causa)).trim(),setor:setorT,itemTxt:itemT,item:itemT,oque:oque,como:String(g(C.como)).trim(),acao:String(g(C.como)).trim(),
          prazo:paraTs(g(C.prazo))||0,novoPrazo:novoTs||null,quem:quem,resp:quem,conclusao:/conclu/.test(st)?(conc||paraTs(g(C.prazo))):null,
          st:/conclu/.test(st)?"Concluído":"Aberto",obs:String(g(C.obs)).trim(),origem:"importado REG-060",ini:Date.now()});
        add++;
      });
      sv("rad_pa",pa);logAcao("Importou REG-060",(ctById(ctId)||{}).codigo,add+" ação(ões)");
      alert(add+" ação(ões) importada(s)"+(dup?" · "+dup+" já existia(m)":"")+".");draw();
    }catch(e){alert("Erro ao ler a planilha: "+e.message)}
  };
  r.readAsArrayBuffer(f);inp.value="";
}
/* ---- exportar no layout REG-060 ---- */
function paExportarXlsx(ctId){
  if(typeof XLSX==="undefined"){alert("Biblioteca de Excel não carregou (verifique a internet).");return}
  var c=ctById(ctId),lista=paDoContrato(ctId).sort(function(a,b){return (paPrazo(a)||0)-(paPrazo(b)||0)}),g=paGerenc(lista);
  var aoa=[
    ["","","","SISTEMA DE GESTÃO INTEGRADO","","","","","","","Nº.: "+REG060.cod],
    ["","","","","","","","","","","REV.: "+REG060.rev],
    ["","","","PLANO DE AÇÃO — "+c.codigo+" · "+c.nome,"","","","","","","Data: "+new Date().toLocaleDateString("pt-BR")],
    ["","","","GERENCIAMENTO"],
    ["","","CONCLUÍDO",pct1(g.c,g.t)/100,g.c],
    ["","","ATRASADA",pct1(g.a,g.t)/100,g.a],
    ["","","EM ANDAMENTO",pct1(g.e,g.t)/100,g.e],
    [],
    ["SETOR","ITEM","O QUE PRECISA SER FEITO","COMO SERÁ FEITO","PRAZO DE CONCLUSÃO","NOVO PRAZO DE CONCLUSÃO","COLABORADOR (CAUSA DA NC)","QUEM IRÁ REALIZAR (NOME)","DATA DE CONCLUSÃO","STATUS","OBSERVAÇÃO"]
  ];
  lista.forEach(function(p){aoa.push([paSetor(p),paItem(p),paOque(p),paComo(p),przTs(p)?new Date(przTs(p)):"",p.novoPrazo?new Date(p.novoPrazo):"",verCausador()?paCausador(p):"",paQuem(p),p.conclusao?new Date(p.conclusao):"",paSituacao(p),p.obs||""])});
  var ws=XLSX.utils.aoa_to_sheet(aoa,{cellDates:true});
  ws["!cols"]=[{wch:32},{wch:26},{wch:40},{wch:48},{wch:14},{wch:14},{wch:28},{wch:26},{wch:14},{wch:14},{wch:34}];
  ws["!merges"]=[{s:{r:0,c:3},e:{r:1,c:9}},{s:{r:2,c:3},e:{r:2,c:9}},{s:{r:3,c:3},e:{r:3,c:9}}];
  ["D5","D6","D7"].forEach(function(a){if(ws[a])ws[a].z="0.0%"});
  for(var r=9;r<aoa.length;r++)["E","F","I"].forEach(function(cl){var a=cl+(r+1);if(ws[a]&&ws[a].t=="d")ws[a].z="dd/mm/yyyy"});
  var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,REG060.cod);
  XLSX.writeFile(wb,REG060.cod+"_PLANO_DE_ACAO_"+c.codigo+".xlsx");
  logAcao("Exportou REG-060",c.codigo);
}
/* ---- impressão: indicadores + plano de ação ---- */
function imprimirIndicadores(ctId){
  var c=ctById(ctId);if(!c)return;var IND_DEF=indDefCt(ctId);
  var ms=indMeses(ctId).slice(-12),vals=ms.map(function(m){return indValor(ctId,m)});
  var lista=paDoContrato(ctId).sort(function(a,b){return (paPrazo(a)||0)-(paPrazo(b)||0)}),g=paGerenc(lista);
  var h='<div class="doc">'+cabDoc("Relatório de indicadores e plano de ação","Contrato "+c.codigo)
    +'<div class="grid"><div><b>Contrato</b>'+esc(c.codigo+" · "+c.nome)+'</div><div><b>Serviços</b>'+esc(tiposCt(c).map(function(t){return tipoCt(t).nome}).join(", "))+'</div><div><b>Período</b>'+mesNome(ms[0])+' a '+mesNome(ms[ms.length-1])+'</div></div>'
    +'<h3>Indicadores mês a mês</h3><table><tr><th>Indicador</th><th>Meta</th>'+ms.map(function(m){return '<th>'+mesNome(m)+'</th>'}).join("")+'</tr>'
    +IND_DEF.map(function(d){return '<tr><td>'+esc(d[1])+'</td><td>'+(d[3]==null?'—':d[3]+(d[2]=="%"?"%":""))+'</td>'+vals.map(function(x){var v=x.v[d[0]];var cor=v==null||d[3]==null?"":(v>=d[3]?"#2e8b57":"#c62828");return '<td style="font-weight:600;color:'+(cor||"inherit")+'">'+fmtInd(v,d[2])+(x.orig[d[0]]=="manual"?'*':'')+'</td>'}).join("")+'</tr>'}).join("")
    +'</table><p class="mut">* valor registrado manualmente. Vermelho: abaixo da meta.</p>';
  var mAt=ms[ms.length-1],pr=[mesIni(mAt),Math.min(mesFim(mAt),Date.now())];
  h+='<h3>Execução por criticidade — '+mesNome(mAt)+'</h3>';
  var R=rpDados(ctId,pr[0],pr[1]).D;
  h+='<table><tr><th>Criticidade</th><th>Espaços</th><th>Limpezas</th><th>Recusas</th><th>Sem registro</th><th>Conformidade</th><th>NC abertas</th></tr>'+["C","S","N"].map(function(k){var x=R[k];return '<tr><td>'+crTag(k)+'</td><td>'+x.esp+'</td><td>'+x.limp+'</td><td>'+x.rec+'</td><td>'+x.semReg+'</td><td>'+fmtInd(media(x.conf),"%")+'</td><td>'+x.nc+'</td></tr>'}).join("")+'</table>';
  h+='</div><div class="doc">'
    +'<div class="cab"><div><div class="emp">'+esc(CONTRATO)+'</div><div class="tit">Sistema de Gestão Integrado — Plano de ação</div><div class="mut">'+esc(c.codigo+" · "+c.nome)+'</div></div><div class="cod">Nº.: '+REG060.cod+'<br>REV.: '+REG060.rev+'<br>Data: '+new Date().toLocaleDateString("pt-BR")+'</div></div>'
    +'<table style="width:55%"><tr><th colspan="3">Gerenciamento</th></tr><tr><td><span class="cr" style="background:#2e8b57">CONCLUÍDO</span></td><td>'+String(pct1(g.c,g.t)).replace(".",",")+'%</td><td>'+g.c+'</td></tr><tr><td><span class="cr" style="background:#c62828">ATRASADA</span></td><td>'+String(pct1(g.a,g.t)).replace(".",",")+'%</td><td>'+g.a+'</td></tr><tr><td><span class="cr" style="background:#e0a100">EM ANDAMENTO</span></td><td>'+String(pct1(g.e,g.t)).replace(".",",")+'%</td><td>'+g.e+'</td></tr><tr><td>Concluídas no prazo</td><td colspan="2">'+g.np+' de '+g.c+'</td></tr></table>'
    +'<table style="font-size:9px;table-layout:fixed;width:100%;word-break:break-word"><colgroup><col style="width:10%"><col style="width:8.5%"><col style="width:12%"><col style="width:14%"><col style="width:8.5%"><col style="width:8%"><col style="width:8.5%"><col style="width:8.5%"><col style="width:8.5%"><col style="width:7%"><col style="width:6.5%"></colgroup><tr><th>Setor</th><th>Item</th><th>O que precisa ser feito</th><th>Como será feito</th><th>Prazo</th><th>Novo prazo</th><th>Causador da NC</th><th>Quem irá realizar</th><th>Conclu&shy;são</th><th>Status</th><th>Observa&shy;ção</th></tr>'
    +(lista.length?lista.map(function(p){var s=paSituacao(p);return '<tr><td>'+esc(paSetor(p))+'</td><td>'+esc(paItem(p))+'</td><td>'+esc(paOque(p))+'</td><td>'+esc(paComo(p))+'</td><td>'+dt(przTs(p))+'</td><td>'+dt(p.novoPrazo)+'</td><td>'+esc(verCausador()?paCausador(p):"")+'</td><td>'+esc(paQuem(p))+'</td><td>'+dt(p.conclusao)+'</td><td><span class="cr" style="white-space:normal;background:'+(s=="CONCLUÍDO"?"#2e8b57":s=="ATRASADA"?"#c62828":"#e0a100")+'">'+s+'</span></td><td>'+esc(p.obs||"")+'</td></tr>'}).join(""):'<tr><td colspan="11" class="mut">Nenhuma ação registrada.</td></tr>')
    +'</table>'
    +'<div class="ass"><div>Supervisão Operacional</div><div>Qualidade / Compliance</div><div>Responsável do cliente</div></div>'+rodDoc()+'</div>';
  logAcao("Imprimiu indicadores + plano de ação",c.codigo);
  abrirImpressao("Indicadores_REG060_"+c.codigo,h);
}

/* ===================================================================
   ESTOQUE POR CONTRATO + INSPEÇÃO DE RECEBIMENTO + OCORRÊNCIAS
   =================================================================== */
var RECEB=ld("rad_receb",[]),OCORR=ld("rad_ocorr",[]);
function svReceb(){sv("rad_receb",RECEB)}function svOcorr(){sv("rad_ocorr",OCORR)}
var esCt=null,esAba="itens";
function podeEstMov(){var cg=cargoAtual();return !!cg&&(cg.fixo||["admin","ger","coord","sup","enc"].indexOf(cg.id)>=0||podeEstoque())}
function podeOcorrencia(){var cg=cargoAtual();return !!cg&&(cg.fixo||["admin","ger","coord","sup"].indexOf(cg.id)>=0)}
function estDoCt(ctId){return EST.filter(function(it){return (it.ct==null?null:it.ct)==ctId})}
function estResumo(ctId){var l=estDoCt(ctId);return {itens:l.length,crit:l.filter(estCritico).length,zero:l.filter(function(i){return i.qtd<=0}).length,ocAb:OCORR.filter(function(o){return o.ct==ctId&&o.st!="Resolvida"}).length}}
function estCtsVisiveis(){var l=meusContratos().slice();if(verTudo()&&EST.some(function(i){return i.ct==null}))l.unshift({id:null,codigo:"—",nome:"Estoque geral (sem contrato)",tipo:"out",tipos:["out"]});return l}
function resumoEstoqueCard(titulo){
  var cts=estCtsVisiveis().filter(function(c){return estDoCt(c.id).length||OCORR.some(function(o){return o.ct==c.id})});
  if(!cts.length)return "";
  var tot={i:0,c:0,z:0,o:0};
  var linhas=cts.map(function(c){var r=estResumo(c.id);tot.i+=r.itens;tot.c+=r.crit;tot.z+=r.zero;tot.o+=r.ocAb;
    return '<tr style="cursor:pointer" onclick="esCt='+(c.id==null?'null':c.id)+';esAba=\'itens\';go(\'es\')"><td style="white-space:normal">'+esc((c.codigo!="—"?c.codigo+" · ":"")+c.nome)+'</td><td>'+r.itens+'</td><td'+(r.crit?' style="color:var(--warn);font-weight:700"':'')+'>'+r.crit+'</td><td'+(r.zero?' style="color:var(--warn);font-weight:700"':'')+'>'+r.zero+'</td><td'+(r.ocAb?' style="color:var(--run);font-weight:700"':'')+'>'+r.ocAb+'</td></tr>'}).join("");
  return '<div class="card"><b>📦 '+(titulo||"Resumo do estoque dos seus contratos")+'</b><div class="tbl"><table style="margin-top:6px"><tr><th>Contrato</th><th>Itens</th><th title="Abaixo do mínimo">↓Mín</th><th>Zero</th><th title="Ocorrências de recebimento abertas">Ocorr.</th></tr>'+linhas
    +(cts.length>1?'<tr><td><b>Total</b></td><td><b>'+tot.i+'</b></td><td><b>'+tot.c+'</b></td><td><b>'+tot.z+'</b></td><td><b>'+tot.o+'</b></td></tr>':'')+'</table></div><div class="s">Toque no contrato para abrir.</div></div>';
}
function vEstoque(){
  var cts=estCtsVisiveis();
  if(!cts.length)return '<div class="card s">Nenhum contrato disponível para você.</div>';
  if(!cts.some(function(c){return c.id===esCt}))esCt=cts[0].id;
  var c=cts.filter(function(x){return x.id===esCt})[0];
  var h=resumoEstoqueCard();
  h+='<div class="card"><b>📦 '+T("menu_estoque")+'</b>'+(cts.length>1?pickContrato("esct",cts.filter(function(x){return x.id!=null}),esCt,function(id){esCt=id;draw()})+(verTudo()&&cts[0].id==null?'<button class="btn alt" onclick="esCt=null;draw()">Estoque geral (sem contrato)</button>':''):'<div class="s">'+esc(c.codigo+" · "+c.nome)+'</div>')
    +'<div class="chips">'+[["itens","📦 Itens"],["receb","📥 Recebimentos"],["ocorr","⚠️ Ocorrências ("+OCORR.filter(function(o){return o.ct==esCt&&o.st!="Resolvida"}).length+")"]].map(function(a){return '<button class="'+(esAba==a[0]?'sel':'')+'" onclick="esAba=\''+a[0]+'\';draw()">'+a[1]+'</button>'}).join("")+'</div></div>';
  if(esAba=="receb")return h+vRecebimentos(c);
  if(esAba=="ocorr")return h+vOcorrencias(c);
  return h+vItensEstoque(c);
}
function vItensEstoque(c){
  var pode=podeEstMov(),l=estDoCt(c.id).sort(function(a,b){return a.nome.localeCompare(b.nome)});
  var h='<div class="card">'+campoBusca("est","Buscar item")+'<div class="s">'+l.length+' item(ns) · '+l.filter(estCritico).length+' abaixo do mínimo</div>';
  var vis=l.filter(function(it){return bate("est",it.nome,it.local)});
  if(!vis.length)h+='<div class="s">Nenhum item.</div>';
  h+=vis.map(function(it){var cls=it.qtd<=0?'no':estCritico(it)?'run':'ok';
    return '<div style="border-top:1px solid var(--line);padding:10px 0"><div class="row" style="border:0;padding:0"><b style="flex:1">'+marca(it.nome,"est")+'</b><span class="pill '+cls+'" style="cursor:default">'+it.qtd+' '+esc(it.unidade)+'</span></div>'
      +'<div class="s">Mín: '+it.min+' '+esc(it.unidade)+(it.local?' · '+esc(it.local):'')+'</div>'
      +'<div class="chips" style="margin-top:6px">'+(pode?'<button onclick="estEntradaUI('+it.id+')">➕ Entrada</button><button onclick="estBaixaUI('+it.id+')">➖ Baixa</button>':'')+'<button onclick="estHistUI('+it.id+')">📜 Histórico</button></div></div>'}).join("");
  h+='</div>';
  if(pode)h+='<div class="card"><b>➕ Novo item</b><input id="est_nome" placeholder="Nome do item"><input id="est_unid" placeholder="Unidade (un, L, kg, cx, fardo)" value="un"><div class="s">Quantidade inicial</div><input id="est_qtd" type="number" min="0" value="0"><div class="s">Estoque mínimo</div><input id="est_min" type="number" min="0" value="0"><input id="est_local" placeholder="Localização (ex.: DML 1º andar)"><button class="btn" onclick="estAddUI()">Cadastrar item</button></div>';
  return h;
}

/* ---- inspeção de recebimento ---- */
var RCF=null;
var RC_CHECK=[["correto","Produto correto (marca/especificação)"],["embalagem","Embalagem íntegra"],["validade","Dentro da validade"],["qtd","Quantidade conferida"]];
var RC_PROB=["Falta / quantidade menor","Produto avariado","Produto errado","Fora da validade","Embalagem violada","Outro"];
function rcNovo(){RCF={nf:"",forn:"",obs:"",linhas:[]};rcAddLinha();draw()}
function rcAddLinha(){RCF.linhas.push({estId:"",nome:"",unidade:"un",ped:"",rec:"",chk:{correto:true,embalagem:true,validade:true,qtd:true},probs:[],foto:""});var y=window.scrollY;draw();window.scrollTo(0,y)}
function rcSet(i,k,v){var L=RCF.linhas[i];if(!L)return;L[k]=v;if(k=="estId"){var it=EST.filter(function(x){return x.id==v})[0];if(it){L.nome=it.nome;L.unidade=it.unidade}}var y=window.scrollY;if(k=="estId"||k=="rec"||k=="ped")draw();window.scrollTo(0,y)}
function rcChk(i,k){var L=RCF.linhas[i];L.chk[k]=!L.chk[k];var y=window.scrollY;draw();window.scrollTo(0,y)}
function rcProb(i,pb){var L=RCF.linhas[i],p=L.probs.indexOf(pb);if(p>=0)L.probs.splice(p,1);else L.probs.push(pb);var y=window.scrollY;draw();window.scrollTo(0,y)}
function rcFoto(i,inp){var f=inp.files[0];if(!f)return;shrinkMax(f,700,function(d){RCF.linhas[i].foto=d;var y=window.scrollY;draw();window.scrollTo(0,y)})}
function shrinkMax(f,max,cb){var im=new Image();var url=URL.createObjectURL(f);im.onload=function(){var c=document.createElement("canvas"),s=Math.min(1,max/Math.max(im.width,im.height));c.width=Math.round(im.width*s);c.height=Math.round(im.height*s);c.getContext("2d").drawImage(im,0,0,c.width,c.height);URL.revokeObjectURL(url);cb(c.toDataURL("image/jpeg",.65))};im.onerror=function(){URL.revokeObjectURL(url);alert("Imagem inválida.")};im.src=url}
function rcLinhaNC(L){var ped=+String(L.ped).replace(",","."),rec=+String(L.rec).replace(",",".");var falta=L.ped!==""&&rec<ped;return falta||RC_CHECK.some(function(c){return !L.chk[c[0]]})||L.probs.length>0}
function rcRemover(i){RCF.linhas.splice(i,1);if(!RCF.linhas.length)rcAddLinha();else draw()}
function vRecebimentos(c){
  var pode=podeEstMov(),h="";
  if(pode&&!RCF)h+='<div class="card"><button class="btn" onclick="rcNovo()">📥 Nova inspeção de recebimento</button><div class="s">Confira cada produto entregue: tire a foto, informe a quantidade pedida e a recebida e marque o que estiver errado. O que estiver conforme entra no estoque.</div></div>';
  if(pode&&RCF){
    var its=estDoCt(c.id).sort(function(a,b){return a.nome.localeCompare(b.nome)});
    h+='<div class="card" style="border:2px solid var(--brand)"><b>📥 Inspeção de recebimento</b>'
      +'<input placeholder="Nº da nota fiscal / pedido (opcional)" value="'+esc(RCF.nf)+'" oninput="RCF.nf=this.value"><input placeholder="Fornecedor / quem entregou (opcional)" value="'+esc(RCF.forn)+'" oninput="RCF.forn=this.value">'
      +RCF.linhas.map(function(L,i){var nc=rcLinhaNC(L);return '<div style="border:1px solid '+(nc?'var(--warn)':'var(--line)')+';border-radius:10px;padding:10px;margin:8px 0;background:var(--bg)">'
        +'<div class="row" style="border:0;padding:0"><b>Produto '+(i+1)+'</b>'+(nc?'<span class="pill no" style="cursor:default">não conforme</span>':'<span class="pill ok" style="cursor:default">conforme</span>')+'</div>'
        +'<select onchange="rcSet('+i+',\'estId\',this.value)"><option value="">— produto do estoque —</option>'+its.map(function(it){return '<option value="'+it.id+'"'+(L.estId==it.id?' selected':'')+'>'+esc(it.nome)+' ('+esc(it.unidade)+')</option>'}).join("")+'<option value="novo"'+(L.estId=="novo"?' selected':'')+'>➕ Produto novo (não cadastrado)</option></select>'
        +(L.estId=="novo"?'<div style="display:flex;gap:6px"><input placeholder="Nome do produto" value="'+esc(L.nome)+'" oninput="RCF.linhas['+i+'].nome=this.value" style="flex:2"><input placeholder="Unid." value="'+esc(L.unidade)+'" oninput="RCF.linhas['+i+'].unidade=this.value" style="flex:1"></div>':'')
        +'<div style="display:flex;gap:6px"><div style="flex:1"><div class="s">Qtd. pedida</div><input inputmode="decimal" value="'+esc(L.ped)+'" onchange="rcSet('+i+',\'ped\',this.value)"></div><div style="flex:1"><div class="s">Qtd. recebida</div><input inputmode="decimal" value="'+esc(L.rec)+'" onchange="rcSet('+i+',\'rec\',this.value)"></div></div>'
        +'<div>'+RC_CHECK.map(function(k){return chip(L.chk[k[0]],"rcChk("+i+",\'"+k[0]+"\')",(L.chk[k[0]]?"✔ ":"✖ ")+k[1])}).join("")+'</div>'
        +(nc?'<div class="s" style="margin-top:6px">Problema encontrado</div><div>'+RC_PROB.map(function(pb,pi){return chip(L.probs.indexOf(pb)>=0,"rcProb("+i+",RC_PROB["+pi+"])",pb)}).join("")+'</div>':'')
        +'<div class="s" style="margin-top:6px">📷 Foto do produto recebido</div>'+(L.foto?'<img src="'+L.foto+'" alt="" style="max-width:100%;border-radius:8px;margin:4px 0">':'')+'<input type="file" accept="image/*" capture="environment" onchange="rcFoto('+i+',this)">'
        +(RCF.linhas.length>1?'<button class="pill no" style="border:0;cursor:pointer;margin-top:6px" onclick="rcRemover('+i+')">Remover produto</button>':'')+'</div>'}).join("")
      +'<button class="btn alt" onclick="rcAddLinha()">➕ Outro produto</button>'
      +'<textarea rows="2" placeholder="Observações do recebimento" oninput="RCF.obs=this.value" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:8px 0">'+esc(RCF.obs)+'</textarea>'
      +'<div class="grid-acoes"><button class="btn" onclick="rcSalvar('+(c.id==null?'null':c.id)+')">Registrar recebimento</button><button class="btn alt" onclick="RCF=null;draw()">Cancelar</button></div></div>';
  }
  var l=RECEB.filter(function(r){return r.ct==c.id}).sort(function(a,b){return b.data-a.data});
  h+='<div class="card"><b>Recebimentos registrados ('+l.length+')</b>'+(l.length?l.slice(0,40).map(function(r){var nc=r.linhas.filter(function(x){return x.nc}).length,oc=OCORR.filter(function(o){return o.recId==r.id})[0];
    return '<div style="border-top:1px solid var(--line);padding:8px 0"><div class="row" style="border:0;padding:0"><span style="flex:1"><b>'+new Date(r.data).toLocaleString("pt-BR")+'</b><div class="s">'+r.linhas.length+' produto(s)'+(r.nf?' · NF '+esc(r.nf):'')+(r.forn?' · '+esc(r.forn):'')+' · conferido por '+esc(r.por)+'</div></span><span class="pill '+(nc?'no':'ok')+'" style="cursor:default">'+(nc?nc+' NC':'Conforme')+'</span></div>'
      +r.linhas.map(function(x){return '<div class="s">• '+esc(x.nome)+': pedido '+esc(x.ped||"—")+' / recebido '+esc(x.rec||"0")+' '+esc(x.unidade)+(x.nc?' — <b style="color:var(--warn)">'+esc((x.probs.length?x.probs:["não conforme"]).join(", "))+'</b>':'')+'</div>'}).join("")
      +(nc&&podeOcorrencia()?(oc?'<div class="s">Ocorrência: <b>'+esc(oc.st)+'</b> <a style="cursor:pointer;color:var(--brand)" onclick="esAba=\'ocorr\';draw()">ver</a></div>':'<button class="btn alt" style="margin-top:6px" onclick="ocAbrir('+r.id+')">⚠️ Abrir ocorrência de recebimento</button>'):'')
      +'</div>'}).join(""):'<div class="s">Nenhum recebimento.</div>')+'</div>';
  return h;
}
function rcSalvar(ctId){
  if(!podeEstMov()){alert("Sem permissão.");return}
  var ls=RCF.linhas;
  for(var i=0;i<ls.length;i++){var L=ls[i];
    if(!L.estId||(L.estId=="novo"&&!L.nome.trim())){alert("Produto "+(i+1)+": escolha o produto.");return}
    if(L.rec===""){alert("Produto "+(i+1)+": informe a quantidade recebida (0 se não veio).");return}
    if(!L.foto&&!confirm("Produto "+(i+1)+" está sem foto. Registrar mesmo assim?"))return;
    if(rcLinhaNC(L)&&!L.probs.length){alert("Produto "+(i+1)+" está não conforme: marque o problema encontrado.");return}
  }
  var rec={id:Date.now(),ct:ctId,data:Date.now(),nf:RCF.nf.trim(),forn:RCF.forn.trim(),obs:RCF.obs.trim(),por:user,porId:session.userId,linhas:[]};
  ls.forEach(function(L){
    var it=L.estId=="novo"?null:EST.filter(function(x){return x.id==L.estId})[0];
    if(!it){EST.push({id:Date.now()+Math.floor(Math.random()*1e5),ct:ctId,nome:L.nome.trim(),unidade:(L.unidade||"un").trim(),qtd:0,min:0,local:"",mov:[]});it=EST[EST.length-1]}
    var q=Math.max(0,+String(L.rec).replace(",",".")||0),nc=rcLinhaNC(L);
    /* entra no estoque o que foi recebido; produto errado/vencido/avariado não entra */
    var bloqueia=L.probs.some(function(p){return /errado|validade|avariado|violada/i.test(p)});
    if(q>0&&!bloqueia)estMov(it.id,q,"Recebimento"+(rec.nf?" NF "+rec.nf:"")+(nc?" (parcial/NC)":""));
    rec.linhas.push({estId:it.id,nome:it.nome,unidade:it.unidade,ped:L.ped,rec:L.rec,chk:L.chk,probs:L.probs.slice(),foto:L.foto,nc:nc,entrou:q>0&&!bloqueia?q:0});
  });
  svEst();RECEB.push(rec);svReceb();RCF=null;
  var nNc=rec.linhas.filter(function(x){return x.nc}).length;
  logAcao("Registrou recebimento de materiais",(ctById(ctId)||{codigo:"geral"}).codigo,nNc?nNc+" NC":"conforme");
  draw();
  if(nNc&&podeOcorrencia()&&confirm(nNc+" produto(s) com não conformidade. Abrir ocorrência de recebimento agora?"))ocAbrir(rec.id);
  else if(nNc)alert("Recebimento registrado com "+nNc+" não conformidade(s). Avise o supervisor para abrir a ocorrência.");
}

/* ---- ocorrência de recebimento + e-mail para Compras ---- */
function ocAbrir(recId){
  if(!podeOcorrencia()){alert("Somente supervisão, coordenação ou gerência abre ocorrência.");return}
  var r=RECEB.filter(function(x){return x.id==recId})[0];if(!r)return;
  if(OCORR.some(function(o){return o.recId==recId})){esAba="ocorr";draw();return}
  var pedido=prompt("O que a Gávea solicita a Compras/fornecedor?\n(ex.: reposição do faltante em 48 h; troca dos itens avariados)","Reposição/troca dos itens não conformes em até 48 horas");
  if(pedido===null)return;
  var oc={id:Date.now(),recId:recId,ct:r.ct,num:"OC-"+new Date().getFullYear()+"-"+String(OCORR.length+1).padStart(3,"0"),aberta:Date.now(),por:user,porId:session.userId,pedido:pedido.trim(),st:"Aberta",hist:[{t:Date.now(),x:"Ocorrência aberta por "+user}]};
  OCORR.push(oc);svOcorr();logAcao("Abriu ocorrência de recebimento",oc.num);esAba="ocorr";draw();
}
function ocTexto(oc){
  var r=RECEB.filter(function(x){return x.id==oc.recId})[0]||{linhas:[]},c=ctById(oc.ct);
  var nc=r.linhas.filter(function(x){return x.nc});
  var L=["Prezados(as) do setor de Compras,","","Registramos uma não conformidade no recebimento de materiais do contrato "+(c?c.codigo+" – "+c.nome:"(estoque geral)")+".","",
    "Ocorrência: "+oc.num,"Data do recebimento: "+new Date(r.data).toLocaleString("pt-BR"),(r.nf?"Nota fiscal / pedido: "+r.nf:""),(r.forn?"Fornecedor / entregador: "+r.forn:""),"Conferido por: "+(r.por||""),"",
    "Itens com problema:"];
  nc.forEach(function(x,i){var ped=x.ped===""?"—":x.ped,rec=x.rec===""?"0":x.rec;var falta=x.ped!==""?Math.max(0,(+String(x.ped).replace(",","."))-(+String(x.rec).replace(",",".")||0)):0;
    L.push((i+1)+". "+x.nome+" — pedido: "+ped+" "+x.unidade+" · recebido: "+rec+" "+x.unidade+(falta?" · faltam: "+String(falta).replace(".",",")+" "+x.unidade:"")+" · problema: "+(x.probs.join(", ")||"não conforme"))});
  var ok=r.linhas.filter(function(x){return !x.nc});
  if(ok.length)L.push("","Itens recebidos conformes: "+ok.map(function(x){return x.nome+" ("+x.rec+" "+x.unidade+")"}).join("; ")+".");
  if(r.obs)L.push("","Observações da conferência: "+r.obs);
  L.push("","Solicitação: "+(oc.pedido||"providenciar a regularização dos itens acima."),"","As fotos da conferência seguem em anexo.","","Fico no aguardo do retorno com a previsão de regularização.","","Atenciosamente,",user,(cargoAtual()||{}).nome||"","Gávea Facilities – "+(c?c.codigo+" "+c.nome:""));
  return L.filter(function(x,i,a){return !(x===""&&a[i-1]==="")}).join("\n");
}
function ocAssunto(oc){var c=ctById(oc.ct);return "Ocorrência de recebimento "+oc.num+" – "+(c?c.codigo+" "+c.nome:"Estoque geral")}
function ocEmail(id){
  var oc=OCORR.filter(function(o){return o.id==id})[0];if(!oc)return;
  var mail=(setor("compras").mail||"").trim();
  if(!mail){var m=prompt("E-mail do setor de Compras (fica salvo em Setores):","");if(!m)return;setSetor("compras","mail",m);mail=m.trim()}
  var corpo=ocTexto(oc);
  if(corpo.length>1800)corpo=corpo.slice(0,1800)+"\n[…texto completo copiado para a área de transferência]";
  try{navigator.clipboard&&navigator.clipboard.writeText(ocTexto(oc))}catch(e){}
  oc.email=Date.now();oc.hist.push({t:Date.now(),x:"E-mail para Compras ("+mail+") preparado por "+user});if(oc.st=="Aberta")oc.st="Enviada a Compras";svOcorr();
  logAcao("Enviou ocorrência a Compras",oc.num);
  window.location.href="mailto:"+mail.replace(/[^\w.@+;,-]/g,"")+"?subject="+encodeURIComponent(ocAssunto(oc))+"&body="+encodeURIComponent(corpo);
  setTimeout(function(){alert("O seu aplicativo de e-mail foi aberto com o texto pronto.\n\nAnexe as fotos: use o botão \"⬇️ Baixar fotos\" da ocorrência e anexe-as ao e-mail antes de enviar.");draw()},600);
}
function ocCopiar(id){var oc=OCORR.filter(function(o){return o.id==id})[0];if(!oc)return;var t=ocAssunto(oc)+"\n\n"+ocTexto(oc);if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(t).then(function(){alert("Texto copiado.")},function(){prompt("Copie o texto:",t)});else prompt("Copie o texto:",t)}
function ocFotos(id){
  var oc=OCORR.filter(function(o){return o.id==id})[0];var r=oc&&RECEB.filter(function(x){return x.id==oc.recId})[0];if(!r)return;
  var n=0;r.linhas.forEach(function(x,i){if(!x.foto)return;n++;setTimeout(function(){var a=document.createElement("a");a.href=x.foto;a.download=oc.num+"_"+(i+1)+"_"+x.nome.replace(/[^\wÀ-ÿ]+/g,"_").slice(0,30)+".jpg";document.body.appendChild(a);a.click();a.remove()},n*400)});
  if(!n)alert("Este recebimento não tem fotos.");
}
function ocStatus(id,st){
  if(!podeOcorrencia())return;var oc=OCORR.filter(function(o){return o.id==id})[0];if(!oc)return;
  var obs=prompt(st=="Resolvida"?"Como foi resolvida? (ex.: itens repostos em 12/10, NF 1234)":"Resposta de Compras:","");if(obs===null)return;
  oc.st=st;oc.hist.push({t:Date.now(),x:st+(obs?": "+obs.trim():"")+" ("+user+")"});if(st=="Resolvida")oc.fim=Date.now();svOcorr();logAcao("Ocorrência de recebimento "+st.toLowerCase(),oc.num);draw();
}
function vOcorrencias(c){
  var l=OCORR.filter(function(o){return o.ct==c.id}).sort(function(a,b){return b.aberta-a.aberta});
  if(!l.length)return '<div class="card s">Nenhuma ocorrência de recebimento. Elas são abertas a partir de um recebimento não conforme.</div>';
  var pode=podeOcorrencia();
  return l.map(function(oc){var r=RECEB.filter(function(x){return x.id==oc.recId})[0]||{linhas:[]};var cl={"Aberta":"no","Enviada a Compras":"run","Respondida":"run","Resolvida":"ok"}[oc.st]||"pen";
    return '<div class="card"><div class="row" style="border:0;padding:0"><span style="flex:1"><b>⚠️ '+esc(oc.num)+'</b><div class="s">Aberta em '+new Date(oc.aberta).toLocaleString("pt-BR")+' por '+esc(oc.por)+(oc.email?' · e-mail em '+new Date(oc.email).toLocaleString("pt-BR"):'')+'</div></span><span class="pill '+cl+'" style="cursor:default">'+esc(oc.st)+'</span></div>'
      +r.linhas.filter(function(x){return x.nc}).map(function(x){return '<div style="display:flex;gap:8px;margin-top:6px">'+(x.foto?'<img src="'+x.foto+'" alt="" style="width:64px;height:64px;object-fit:cover;border-radius:6px">':'')+'<div class="s"><b>'+esc(x.nome)+'</b><br>pedido '+esc(x.ped||"—")+' · recebido '+esc(x.rec)+' '+esc(x.unidade)+'<br>'+esc(x.probs.join(", "))+'</div></div>'}).join("")
      +'<div class="s" style="margin-top:6px">Solicitação: '+esc(oc.pedido)+'</div>'
      +'<details style="margin-top:6px"><summary class="s" style="cursor:pointer">📄 Ver texto do e-mail</summary><pre style="white-space:pre-wrap;font:13px/1.4 inherit;background:var(--bg);padding:8px;border-radius:8px">'+esc(ocAssunto(oc)+"\n\n"+ocTexto(oc))+'</pre></details>'
      +'<div class="s" style="margin-top:6px">'+oc.hist.map(function(x){return new Date(x.t).toLocaleString("pt-BR")+' · '+esc(x.x)}).join("<br>")+'</div>'
      +(pode?'<div class="chips" style="margin-top:6px"><button class="sel" onclick="ocEmail('+oc.id+')">📧 Enviar e-mail para Compras</button><button onclick="ocCopiar('+oc.id+')">📋 Copiar texto</button><button onclick="ocFotos('+oc.id+')">⬇️ Baixar fotos</button>'+(oc.st!="Resolvida"?'<button onclick="ocStatus('+oc.id+',\'Respondida\')">💬 Registrar resposta</button><button onclick="ocStatus('+oc.id+',\'Resolvida\')">✔ Resolvida</button>':'')+'</div>':'')
      +'</div>'}).join("");
}

/* ===================================================================
   PORTAL DO CONTRATO (QR) + GEOLOCALIZAÇÃO
   =================================================================== */
function ptToken(c){if(!c.portalTk){c.portalTk=Math.random().toString(36).slice(2,8)+Date.now().toString(36).slice(-4);svCt()}return c.portalTk}
function ptUrl(c){return (sysUrlBase()||location.href.replace(/[?#].*$/,""))+"?portal="+ptToken(c)}
function ptCtPorToken(tk){return CONTRATOS.filter(function(c){return c.portalTk&&c.portalTk==tk})[0]||null}
function geoDist(a,b,c,d){var R=6371000,r=Math.PI/180,x=Math.sin((c-a)*r/2),y=Math.sin((d-b)*r/2);var h=x*x+Math.cos(a*r)*Math.cos(c*r)*y*y;return 2*R*Math.asin(Math.sqrt(h))}
function geoAqui(ok,erro){
  if(!navigator.geolocation){erro("Este aparelho não informa a localização.");return}
  navigator.geolocation.getCurrentPosition(function(p){ok(p.coords.latitude,p.coords.longitude,p.coords.accuracy)},function(e){erro(e.code==1?"Permissão de localização negada. Libere a localização para este site nas configurações do navegador.":"Não foi possível obter a localização ("+(e.message||e.code)+").")},{enableHighAccuracy:true,timeout:15000,maximumAge:60000});
}
/* contrato mais próximo dentro do raio (lista = contratos candidatos) */
function geoContrato(lat,lng,lista){
  var best=null;(lista||CONTRATOS).forEach(function(c){if(!c.geo||c.status=="Encerrado")return;var d=geoDist(lat,lng,c.geo.lat,c.geo.lng);if(d<=(c.geo.raio||300)+100&&(!best||d<best.d))best={c:c,d:d}});return best;
}
function geoDefinirCt(id){
  var c=ctById(id);if(!c)return;
  if(c.geo&&!confirm("Substituir a localização já cadastrada pela sua posição atual?"))return;
  toastMsg("📍 Obtendo localização…");
  geoAqui(function(lat,lng,acc){
    var raio=prompt("Localização capturada (precisão ±"+Math.round(acc)+" m).\n\nRaio do contrato em metros (área do prédio/terreno):",String((c.geo&&c.geo.raio)||300));
    if(raio===null)return;
    c.geo={lat:+lat.toFixed(6),lng:+lng.toFixed(6),raio:Math.max(50,+raio||300),em:Date.now(),por:user};svCt();
    logAcao("Definiu localização do contrato",c.codigo);draw();
  },function(m){alert(m)});
}
function geoLimparCt(id){var c=ctById(id);if(!c||!confirm("Remover a localização do contrato?"))return;delete c.geo;svCt();draw()}
function toastMsg(t){try{var d=document.createElement("div");d.textContent=t;d.style.cssText="position:fixed;left:50%;bottom:24px;transform:translateX(-50%);background:#222;color:#fff;padding:10px 16px;border-radius:99px;z-index:9999;font-size:14px";document.body.appendChild(d);setTimeout(function(){d.remove()},2500)}catch(e){}}
/* no app: escolhe o contrato pela localização */
function geoMeuContrato(){
  var m=meusContratos();
  geoAqui(function(lat,lng){
    var b=geoContrato(lat,lng,m);
    if(!b){alert(m.some(function(c){return c.geo})?"Você não está dentro de nenhum dos seus contratos com localização cadastrada.":"Nenhum dos seus contratos tem localização cadastrada ainda (Contratos → abrir → 📍 Localização).");return}
    setCtAtual(b.c.id);toastMsg("📍 "+b.c.codigo+" · "+b.c.nome+" ("+Math.round(b.d)+" m)");
  },function(e){alert(e)});
}

/* ---- bloco no detalhe do contrato ---- */
function portalCtBloco(c,pode){
  var gere=pode||podeAreasCt(c.id);
  var abertos=chs.filter(function(x){return x.ct==c.id&&x.origem=="portal QR"&&x.st!="Resolvido"&&x.st!="Fechado"}).length;
  var h='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin:12px 0 6px">📱 Portal do contrato (QR para os funcionários do cliente)</div>'
    +'<div class="s">Pelo QR, os funcionários do cliente veem as limpezas feitas hoje por área (sem nomes) e abrem solicitação de limpeza, chamado, reclamação ou elogio — sem login e sem acesso ao efetivo.'+(abertos?' <b style="color:var(--run)">'+abertos+' pedido(s) do portal em aberto</b> (em 💬 Cliente).':'')+'</div>';
  if(gere)h+='<div class="chips" style="margin-top:6px"><button onclick="ptCartaz('+c.id+')">🖨️ Imprimir cartaz com QR</button><button onclick="ptCopiar('+c.id+')">🔗 Copiar link</button><button onclick="ptAbrir('+c.id+')">👁️ Ver portal</button>'+(c.portalTk?'<button onclick="ptNovoToken('+c.id+')">♻️ Gerar novo QR</button>':'')+'</div>';
  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin:12px 0 6px">📍 Localização do contrato</div>'
    +(c.geo?'<div class="s">'+c.geo.lat+', '+c.geo.lng+' · raio '+c.geo.raio+' m · <a href="https://maps.google.com/?q='+c.geo.lat+','+c.geo.lng+'" target="_blank" rel="noopener" style="color:var(--brand)">ver no mapa</a></div>':'<div class="s">Não cadastrada. Vá até o local e toque em “usar minha posição”: o sistema passa a reconhecer o contrato pela geolocalização (no app e no portal).</div>')
    +(gere?'<div class="chips" style="margin-top:6px"><button onclick="geoDefinirCt('+c.id+')">📍 '+(c.geo?'Atualizar':'Usar minha posição atual')+'</button>'+(c.geo?'<button onclick="geoLimparCt('+c.id+')">Remover</button>':'')+'</div>':'');
  return h;
}
function ptCopiar(id){var c=ctById(id);if(!c)return;var u=ptUrl(c);if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(u).then(function(){alert("Link copiado:\n"+u)},function(){prompt("Copie o link:",u)});else prompt("Copie o link:",u)}
function ptAbrir(id){var c=ctById(id);if(c)window.open(ptUrl(c),"_blank")}
function ptNovoToken(id){var c=ctById(id);if(!c||!confirm("Gerar novo QR? Os cartazes já impressos deixam de funcionar."))return;c.portalTk=null;ptToken(c);logAcao("Gerou novo QR do portal",c.codigo);draw()}
function qrDataUrl(txt,tam){
  try{var d=document.createElement("div");d.style.cssText="position:fixed;left:-9999px";document.body.appendChild(d);new QRCode(d,{text:txt,width:tam,height:tam,correctLevel:QRCode.CorrectLevel?QRCode.CorrectLevel.M:undefined});
    var cv=d.querySelector("canvas"),src=cv?cv.toDataURL("image/png"):((d.querySelector("img")||{}).src||"");d.remove();return src}catch(e){return ""}
}
function ptCartaz(id){
  var c=ctById(id);if(!c)return;var url=ptUrl(c),q=qrDataUrl(url,520);
  if(!q){alert("Biblioteca de QR indisponível (sem internet?). Use “Copiar link”.");return}
  var html='<div style="text-align:center;padding:30px 20px;font-family:Arial,sans-serif">'
    +'<div style="font-size:15px;letter-spacing:2px;color:#555">GÁVEA FACILITIES</div>'
    +'<h1 style="font-size:34px;margin:10px 0 4px">Limpeza deste local</h1>'
    +'<div style="font-size:18px;color:#333">'+esc(c.nome)+'</div>'
    +'<img src="'+q+'" alt="QR" style="width:330px;height:330px;margin:26px auto;display:block">'
    +'<div style="font-size:20px;font-weight:700">Aponte a câmera do celular</div>'
    +'<table style="margin:18px auto;font-size:17px;text-align:left;border-collapse:collapse"><tr><td style="padding:5px 10px">🧹</td><td>Veja as limpezas realizadas hoje</td></tr><tr><td style="padding:5px 10px">🎯</td><td>Solicite uma limpeza</td></tr><tr><td style="padding:5px 10px">💬</td><td>Abra um chamado ou reclamação</td></tr><tr><td style="padding:5px 10px">🌟</td><td>Deixe um elogio à equipe</td></tr></table>'
    +'<div style="font-size:11px;color:#888;margin-top:20px;word-break:break-all">'+esc(url)+'</div></div>';
  logAcao("Imprimiu cartaz do portal",c.codigo);
  abrirImpressao("QR_Portal_"+c.codigo,html);
}

/* ---- página pública ---- */
var PT={ct:null,aba:"hoje",tipo:"",area:"",nome:"",setor:"",txt:"",foto:"",msg:""};
function ptMeus(){return ld("rad_portal_meus",[])}
function ptRegsHoje(ct){var h0=new Date();h0.setHours(0,0,0,0);return regs.filter(function(r){return r.st=="ok"&&r.fim>=h0.getTime()&&ctDaArea(r.l)==ct})}
function ptRender(){
  var box=document.getElementById("portal");if(!box)return;var c=PT.ct;
  if(!c){box.innerHTML='<div class="card" style="text-align:center"><b>📍 Portal da limpeza</b><div class="s" style="margin:8px 0">'+esc(PT.msg||"Identificando o contrato pela sua localização…")+'</div><button class="btn" onclick="ptGeo()">📍 Usar minha localização</button></div>';return}
  var h='<div class="card" style="text-align:center"><div class="s" style="letter-spacing:2px">GÁVEA FACILITIES</div><b style="font-size:20px">'+esc(c.nome)+'</b><div class="s">Portal da limpeza · '+new Date().toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long"})+'</div></div>'
    +'<div class="card"><div class="chips">'+[["hoje","🧹 Limpezas de hoje"],["pedir","➕ Abrir pedido"],["meus","📄 Meus pedidos"]].map(function(a){return '<button class="'+(PT.aba==a[0]?'sel':'')+'" onclick="PT.aba=\''+a[0]+'\';ptRender()">'+a[1]+'</button>'}).join("")+'</div></div>';
  if(PT.aba=="hoje"){
    var rs=ptRegsHoje(c.id),porA={};rs.forEach(function(r){(porA[r.l]=porA[r.l]||[]).push(r.fim)});
    var areas=areasDoContrato(c.id).filter(function(i){return !LOC[i].off});
    h+='<div class="card"><b>🧹 Limpezas realizadas hoje ('+rs.length+')</b>'
      +(areas.length?areas.map(function(i){var t=(porA[i]||[]).sort(),ult=t[t.length-1];
        return '<div class="row"><span style="flex:1">'+esc(LOC[i].n)+'<div class="s">'+(t.length?t.map(function(x){return new Date(x).toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}).join(" · "):"Ainda não registrada hoje")+'</div></span><span class="pill '+(t.length?'ok':'pen')+'" style="cursor:default">'+(t.length?'✔ '+new Date(ult).toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"}):'—')+'</span></div>'}).join(""):'<div class="s">Áreas ainda não cadastradas.</div>')
      +'</div>';
  }else if(PT.aba=="pedir"){
    var tipos=[["limpeza","🎯 Solicitar limpeza","🎯 Solicitação de serviço"],["chamado","🛠️ Chamado / problema","❓ Dúvida / Informação"],["reclamacao","⚠️ Reclamação","⚠️ Reclamação"],["elogio","🌟 Elogio","🌟 Elogio"],["urgencia","🚨 Urgência","🚨 Urgência"]];
    h+='<div class="card"><b>➕ Novo pedido</b><div class="chips" style="margin:6px 0">'+tipos.map(function(t){return '<button class="'+(PT.tipo==t[0]?'sel':'')+'" onclick="PT.tipo=\''+t[0]+'\';ptRender()">'+t[1]+'</button>'}).join("")+'</div>'
      +'<div class="s">Local</div><select onchange="PT.area=this.value"><option value="">— onde? —</option>'+areasDoContrato(c.id).filter(function(i){return !LOC[i].off}).map(function(i){return '<option value="'+i+'"'+(PT.area==String(i)?' selected':'')+'>'+esc(LOC[i].n)+'</option>'}).join("")+'<option value="-1"'+(PT.area=="-1"?' selected':'')+'>Outro local</option></select>'
      +'<textarea id="pt_txt" rows="3" placeholder="'+(PT.tipo=="elogio"?"Conte o que gostou (pode citar a equipe ou o local)":"Descreva o que precisa")+'" oninput="PT.txt=this.value" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit;margin:6px 0">'+esc(PT.txt)+'</textarea>'
      +'<input placeholder="Seu nome (opcional)" value="'+esc(PT.nome)+'" oninput="PT.nome=this.value"><input placeholder="Setor / ramal (opcional, para retorno)" value="'+esc(PT.setor)+'" oninput="PT.setor=this.value">'
      +'<div class="s">📷 Foto (opcional)</div>'+(PT.foto?'<img src="'+PT.foto+'" alt="" style="max-width:100%;border-radius:8px">':'')+'<input type="file" accept="image/*" capture="environment" onchange="ptFoto(this)">'
      +'<button class="btn" onclick="ptEnviar()" style="margin-top:8px">Enviar</button><div class="s">Seus dados são usados apenas para retorno deste pedido (LGPD).</div></div>';
  }else{
    var meus=ptMeus().filter(function(m){return m.ct==c.id}).map(function(m){return chs.filter(function(x){return x.id==m.id})[0]}).filter(Boolean).sort(function(a,b){return b.ini-a.ini});
    h+='<div class="card"><b>📄 Meus pedidos ('+meus.length+')</b>'+(meus.length?meus.map(function(x){var resp=(x.msgs||[]).filter(function(m){return m.de!="c"});
      return '<div style="border-top:1px solid var(--line);padding:8px 0"><div class="row" style="border:0;padding:0"><span style="flex:1"><b>'+esc(x.tipo)+'</b> <span class="s">· '+esc(x.prot||"")+'</span><div class="s">'+new Date(x.ini).toLocaleString("pt-BR")+' · '+esc(x.onde)+'</div></span><span class="pill '+(/Resol|Fech/.test(x.st)?'ok':'run')+'" style="cursor:default">'+esc(x.st)+'</span></div><div class="s">'+esc((x.msgs[0]||{}).x||"")+'</div>'
        +resp.map(function(m){return '<div class="s" style="background:var(--bg);padding:6px;border-radius:6px;margin-top:4px">💬 Equipe: '+esc(m.x)+'</div>'}).join("")+'</div>'}).join(""):'<div class="s">Nenhum pedido feito neste aparelho.</div>')+'</div>';
  }
  h+='<div class="s" style="text-align:center;margin:10px 0">RAD Digital · Gávea Facilities</div>';
  box.innerHTML=h;
}
function ptFoto(inp){var f=inp.files[0];if(f)shrinkMax(f,700,function(d){PT.foto=d;ptRender()})}
function ptEnviar(){
  var c=PT.ct;if(!c)return;
  var mapa={limpeza:"🎯 Solicitação de serviço",chamado:"❓ Dúvida / Informação",reclamacao:"⚠️ Reclamação",elogio:"🌟 Elogio",urgencia:"🚨 Urgência"};
  if(!PT.tipo){alert("Escolha o tipo de pedido.");return}
  if(PT.area===""){alert("Informe o local.");return}
  var x=(PT.txt||"").trim();if(!x&&PT.tipo!="limpeza"){alert("Descreva o pedido.");return}
  var l=+PT.area,tp=mapa[PT.tipo],dest=chDestinoCan("CLI",tp),prot="P"+String(Date.now()).slice(-6);
  var quem=(PT.nome.trim()||"Funcionário do cliente")+(PT.setor.trim()?" ("+PT.setor.trim()+")":"");
  var ch={id:Date.now(),canal:"CLI",tipo:tp,onde:(l>=0?nomeLoc(l):"Outro local")+" · "+c.nome,l:l,e:-1,ct:c.id,ini:Date.now(),colab:"Portal QR – "+quem,st:"Aberto",setor:dest,origem:"portal QR",prot:prot,foto:PT.foto||"",msgs:[{de:"c",t:Date.now(),x:x||"Solicitação de limpeza"}]};
  chs.push(ch);sv("rad_chs",chs);
  try{var hist=ld("rad_notif",[]);hist.push({id:Date.now(),setor:dest,canal:"CLI",titulo:"Portal QR: "+tp,texto:ch.onde+"\n"+quem+"\n\n"+(x||""),ts:Date.now()});sv("rad_notif",hist.slice(-200))}catch(e){}
  var m=ptMeus();m.push({id:ch.id,ct:c.id});sv("rad_portal_meus",m);
  var eraElogio=PT.tipo=="elogio";PT.tipo="";PT.txt="";PT.foto="";PT.area="";PT.aba="meus";ptRender();
  alert((eraElogio?"Obrigado pelo elogio!":"Pedido enviado.")+"\nProtocolo: "+prot);
}
function ptGeo(){
  PT.msg="Obtendo localização…";ptRender();
  geoAqui(function(lat,lng){var b=geoContrato(lat,lng);if(b){PT.ct=b.c;PT.msg=""}else PT.msg="Não encontramos um contrato da Gávea Facilities no local onde você está. Use o QR code afixado no local.";ptRender()},function(e){PT.msg=e;ptRender()});
}
function ptIniciar(){
  var m=/[?&]portal=([\w-]+)/.exec(location.search);if(!m)return false;
  document.getElementById("login").classList.add("hide");document.getElementById("shell").classList.add("hide");
  var box=document.createElement("div");box.id="portal";box.style.cssText="max-width:560px;margin:0 auto;padding:12px";document.body.appendChild(box);
  if(m[1]=="geo"){ptRender();ptGeo()}
  else{PT.ct=ptCtPorToken(m[1]);if(!PT.ct)PT.msg="QR code não reconhecido neste aparelho. Tente pela localização.";ptRender()}
  return true;
}

/* ===================================================================
   ACESSOS (usuários) — menu próprio, vínculo a contratos com busca,
   acesso do cliente pelo contrato e mensagem de envio do acesso
   =================================================================== */
var UN={cargo:"",cts:[],filtro:"todos"},UED=null;
function pinAleatorio(){var p;do{p=String(Math.floor(1000+Math.random()*9000))}while(/^(\d)\1{3}$/.test(p)||p=="1234");return p}
function ctsNomesUser(u){var ids=u.contratosIds||[];return ids.map(function(i){var c=ctById(i);return c?c.codigo+" · "+c.nome:null}).filter(Boolean)}
function selCtsHtml(k,sel,fnTog){
  var l=CONTRATOS.slice().sort(function(a,b){return (a.codigo||"").localeCompare(b.codigo||"")});
  var marcados=l.filter(function(c){return sel.indexOf(c.id)>=0});
  var h='<div class="s">Contratos vinculados ('+sel.length+')</div>'
    +(marcados.length?'<div>'+marcados.map(function(c){return chip(true,fnTog+"("+c.id+")",c.codigo+" · "+c.nome+" ✕")}).join("")+'</div>':'<div class="s" style="color:var(--warn)">Nenhum contrato escolhido.</div>')
    +campoBusca(k,"Buscar contrato para vincular (código, nome ou supervisor)");
  if(termo(k)){var r=l.filter(function(c){return sel.indexOf(c.id)<0&&bate(k,c.codigo,c.nome,nomesSupCt(c.id))});
    h+=r.length?r.slice(0,10).map(function(c){return '<div class="row" style="padding:6px 0"><span style="flex:1">'+marca(c.codigo+" · "+c.nome,k)+'<div class="s">👔 '+esc(nomesSupCt(c.id)||"sem supervisor")+'</div></span><button class="pill ok" style="border:0;cursor:pointer" onclick="'+fnTog+'('+c.id+')">+ Vincular</button></div>'}).join(""):'<div class="s">Nenhum contrato encontrado.</div>'}
  return h;
}
function unTogCt(id){var i=UN.cts.indexOf(id);if(i>=0)UN.cts.splice(i,1);else{if(UN.cargo=="cli")UN.cts=[];UN.cts.push(id);BUSCA["unct"]=""}var y=window.scrollY;draw();window.scrollTo(0,y)}
function uedTogCt(id){var i=UED.cts.indexOf(id);if(i>=0)UED.cts.splice(i,1);else{UED.cts.push(id);BUSCA["uedct"]=""}var y=window.scrollY;draw();window.scrollTo(0,y)}
function editarContratosUsuario(id){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var u=userById(id);if(!u)return;UED={id:id,cts:(u.contratosIds||[]).slice()};if(tab!="ac"&&tab!="k")tab="ac";draw();window.scrollTo(0,0);
}
function uedSalvar(){var u=userById(UED.id);if(!u)return;var cg=cargoById(u.cargoId);
  if(cg&&cg.id=="cli"&&UED.cts.length!=1&&!confirm("Acesso de cliente normalmente fica vinculado a um único contrato. Salvar com "+UED.cts.length+"?"))return;
  u.contratosIds=UED.cts.slice();saveUsers();logAcao("Editou contratos do usuário",u.nome,UED.cts.length+" contrato(s)");UED=null;draw()}
function primeiroNome(n){return (String(n||"").split(" ").filter(function(w){return w&&!/\.$/.test(w)})[0])||n}
function msgAcesso(u,pin,novo){
  var c=ctsNomesUser(u),cg=cargoById(u.cargoId)||{},link=sysUrlBase()||"(link do RAD Digital — configure em Configuração do Sistema)";
  var oculto=ld("rad_sys",{loginOculto:true}).loginOculto!==false;
  return "*RAD Digital – Gávea Facilities*\n\nOlá, "+primeiroNome(u.nome)+"! "+(novo===false?"Seguem seus dados de acesso ao sistema.":"Sua conta no RAD Digital foi aberta.")+"\n\n"
    +(c.length?"📁 Contrato: "+c.join(", ")+"\n":"")+"🎯 Perfil: "+(cg.nome||"—")+"\n\n"
    +"🔗 Link: "+link+"\n"
    +"👤 Login: "+u.nome+(oculto?"\n   (digite as 3 primeiras letras do nome e toque nele)":"\n   (toque no seu nome)")+"\n"
    +"🔑 Senha (PIN): "+(pin||"use o seu PIN atual")+"\n\n"
    +(pin?"No primeiro acesso o sistema pede para você criar uma senha (PIN) própria e completar o perfil.\n\n":"")
    +"Não compartilhe sua senha. Dúvidas: fale com o seu gestor"+(user?" ("+user+")":"")+".";
}
var ENV=null;
function enviarAcesso(id,pin){
  var u=userById(id);if(!u)return;
  if(!pin&&u.pinTrocar)pin=u.pin;
  ENV={id:id,pin:pin||"",tel:u.telefone||(/@/.test(u.contato||"")?"":u.contato)||"",mail:u.email||(/@/.test(u.contato||"")?u.contato:"")||""};
  ENV.txt=msgAcesso(u,ENV.pin);envRender();
}
function envRender(){
  var u=userById(ENV.id);var box=document.getElementById("envBox");
  if(!box){box=document.createElement("div");box.id="envBox";document.body.appendChild(box)}
  box.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;display:flex;align-items:flex-end;justify-content:center;padding:12px";
  box.innerHTML='<div class="card" style="max-width:520px;width:100%;max-height:92vh;overflow:auto;margin:0">'
    +'<div class="row" style="border:0;padding:0"><b style="flex:1">📤 Enviar acesso — '+esc(u.nome)+'</b><button class="pill pen" style="border:0;cursor:pointer" onclick="envFechar()">✕</button></div>'
    +(ENV.pin?'':'<div class="s" style="background:rgba(224,161,0,.12);padding:8px;border-radius:8px;margin:6px 0">Esta pessoa já criou a própria senha, que o sistema não mostra. <a style="color:var(--brand);cursor:pointer;font-weight:700" onclick="envNovoPin()">Gerar nova senha provisória e enviar</a></div>')
    +'<div class="s">WhatsApp</div><input id="env_tel" inputmode="tel" placeholder="(21) 99999-9999" value="'+esc(ENV.tel)+'" oninput="ENV.tel=this.value">'
    +'<div class="s">E-mail</div><input id="env_mail" type="email" placeholder="nome@empresa.com.br" value="'+esc(ENV.mail)+'" oninput="ENV.mail=this.value">'
    +'<div class="s">Mensagem (pode editar)</div><textarea id="env_txt" rows="13" oninput="ENV.txt=this.value" style="width:100%;padding:10px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:13px/1.45 inherit">'+esc(ENV.txt)+'</textarea>'
    +'<div class="grid-acoes" style="margin-top:8px"><button class="btn" style="background:#25d366;border-color:#25d366;color:#fff" onclick="envVia(\'zap\')">📲 WhatsApp</button><button class="btn" onclick="envVia(\'mail\')">✉️ E-mail</button></div>'
    +'<button class="btn alt" onclick="envVia(\'copia\')">📋 Copiar mensagem</button>'
    +(u.acessoEnviado?'<div class="s">Último envio: '+new Date(u.acessoEnviado.t).toLocaleString("pt-BR")+' por '+esc(u.acessoEnviado.via)+'</div>':'')
    +'</div>';
}
function envFechar(){var b=document.getElementById("envBox");if(b)b.remove();ENV=null;draw()}
function envNovoPin(){
  var u=userById(ENV.id);if(!u||!confirm("Gerar uma nova senha provisória para "+u.nome+"? A senha atual deixa de valer."))return;
  var p=pinAleatorio();u.pin=p;u.pinTrocar=true;saveUsers();PINLOCK[u.id]={tentativas:0,ate:0};sv("rad_pinlock",PINLOCK);logAcao("Gerou senha provisória",u.nome);
  ENV.pin=p;ENV.txt=msgAcesso(u,p,false);envRender();
}
function envVia(via){
  var u=userById(ENV.id);if(!u)return;var t=ENV.txt;
  if(via=="zap"){var tel=String(ENV.tel||"").replace(/\D/g,"");if(tel.length<10){alert("Informe o WhatsApp com DDD.");return}if(tel.length<=11)tel="55"+tel;u.telefone=u.telefone||ENV.tel.trim();window.open("https://wa.me/"+tel+"?text="+encodeURIComponent(t),"_blank")}
  else if(via=="mail"){var m=String(ENV.mail||"").trim();if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(m)){alert("Informe um e-mail válido.");return}u.email=u.email||m;window.location.href="mailto:"+m+"?subject="+encodeURIComponent("Sua conta no RAD Digital – Gávea Facilities")+"&body="+encodeURIComponent(t.replace(/\*/g,""))}
  else{if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(t).then(function(){toastMsg("📋 Mensagem copiada")},function(){prompt("Copie:",t)});else prompt("Copie:",t)}
  if(!u.contato)u.contato=via=="mail"?ENV.mail.trim():ENV.tel.trim();
  u.acessoEnviado={t:Date.now(),via:via=="zap"?"WhatsApp":via=="mail"?"e-mail":"texto copiado",por:user};saveUsers();
  logAcao("Enviou dados de acesso",u.nome,u.acessoEnviado.via);envRender();
}
function usersCard(){
  if(!temAcao("gerenciar_usuarios"))return '<div class="card s">Sem permissão.</div>';
  var h="";
  if(UED){var ue=userById(UED.id);if(ue)h+='<div class="card" style="border:2px solid var(--brand)"><b>📁 Contratos de '+esc(ue.nome)+'</b><div class="s">'+esc((cargoById(ue.cargoId)||{}).nome||"")+'</div>'+selCtsHtml("uedct",UED.cts,"uedTogCt")+'<div class="grid-acoes" style="margin-top:8px"><button class="btn" onclick="uedSalvar()">Salvar vínculos</button><button class="btn alt" onclick="UED=null;draw()">Cancelar</button></div></div>';}
  var filtros=[["todos","Todos"],["cli","🤝 Clientes"],["lid","👔 Liderança"],["col","🧹 Operação"]];
  var lista=USERS.slice().sort(function(a,b){return(a.nome||"").localeCompare(b.nome||"")}).filter(function(u){var cg=cargoById(u.cargoId)||{};
    if(UN.filtro=="cli"&&cg.id!="cli")return false;
    if(UN.filtro=="lid"&&["admin","ger","coord","sup","enc"].indexOf(cg.id)<0&&!cg.fixo)return false;
    if(UN.filtro=="col"&&cg.id!="colab")return false;
    return bate("us",u.nome,u.mat,cg.nome,ctsNomesUser(u).join(" "))});
  h+='<div class="card"><b>🔑 Acessos ao sistema</b><div class="s">Cada pessoa entra com nome + PIN. O <b>cargo</b> define o que ela vê e os <b>contratos</b> definem de onde ela vê. '+USERS.length+' acesso(s).</div>'
    +'<div class="chips">'+filtros.map(function(f){return '<button class="'+(UN.filtro==f[0]?'sel':'')+'" onclick="UN.filtro=\''+f[0]+'\';draw()">'+f[1]+'</button>'}).join("")+'</div>'
    +campoBusca("us","Buscar por nome, matrícula, cargo ou contrato")+'<button class="btn alt" onclick="document.getElementById(\'un_nome\').scrollIntoView({behavior:\'smooth\'});document.getElementById(\'un_nome\').focus()">➕ Criar novo acesso</button></div>';
  h+='<div class="card">'+(lista.length?'':'<div class="s">Nenhum acesso encontrado.</div>')+lista.slice(0,150).map(function(u){
    var eu=session&&session.userId==u.id,cg=cargoById(u.cargoId)||cargoById("colab"),cts=ctsNomesUser(u);
    return '<div class="row" style="flex-wrap:wrap"><span class="uav">'+esc(iniciais(u.nome))+'</span>'
      +'<span class="un" style="flex:1;min-width:150px"><b>'+marca(u.nome,"us")+(eu?' (você)':'')+'</b><span>'+esc(cg?cg.nome:"—")+(u.mat?' · mat. '+esc(u.mat):'')+(u.contato?' · '+esc(u.contato):'')+(!u.perfilCompleto?' · <b style="color:var(--run)">ainda não entrou</b>':'')+(u.acessoEnviado?' · 📤 enviado '+new Date(u.acessoEnviado.t).toLocaleDateString("pt-BR"):(!u.perfilCompleto?' · <b style="color:var(--warn)">acesso não enviado</b>':''))+(u.pin==="0000"?' · <b style="color:var(--warn)">PIN padrão</b>':'')+'</span>'
      +'<span>📁 '+(cts.length?esc(cts.length>3?cts.slice(0,3).join(" · ")+" +"+(cts.length-3):cts.join(" · ")):(cg&&(cg.fixo||cg.id=="admin")?"todos os contratos":'<b style="color:var(--warn)">sem contrato</b>'))+'</span></span>'
      +'<span style="display:flex;gap:4px;flex-wrap:wrap">'
        +'<button class="pill run" style="border:0;cursor:pointer" onclick="editarCargoUsuario('+u.id+')">🎯 Cargo</button>'
        +'<button class="pill pen" style="border:0;cursor:pointer" onclick="editarContratosUsuario('+u.id+')">📁 Contratos</button>'
        +'<button class="pill pen" style="border:0;cursor:pointer" onclick="resetPinUser('+u.id+')">🔑 PIN</button>'
        +'<button class="pill pen" style="border:0;cursor:pointer" onclick="enviarAcesso('+u.id+')">📤 Enviar</button>'
        +'<button class="pill pen" style="border:0;cursor:pointer" onclick="editarIdiomaComunicacao('+u.id+')">🗣️</button>'
        +(eu?'':'<button class="pill no" style="border:0;cursor:pointer" onclick="delUser('+u.id+')">X</button>')
      +'</span></div>'}).join("")+(lista.length>150?'<div class="s">Mostrando 150 de '+lista.length+'. Use a busca.</div>':'')+'</div>';
  if(!UN.cargo)UN.cargo="cli";
  var ehCli=UN.cargo=="cli";
  h+='<div class="card"><b>➕ Novo acesso</b>'
    +'<div class="s">Cargo</div><select id="un_cargo" onchange="UN.cargo=this.value;if(UN.cargo==\'cli\'&&UN.cts.length>1)UN.cts=UN.cts.slice(0,1);draw()">'+CARGOS.map(function(c){return '<option value="'+c.id+'"'+(UN.cargo==c.id?' selected':'')+'>'+esc(c.nome)+'</option>'}).join("")+'</select>'
    +(ehCli?'<div class="s" style="background:var(--bg);padding:8px;border-radius:8px;margin:6px 0">🤝 <b>Acesso do cliente:</b> vê só o contrato vinculado — indicadores, relatório e chamados do cliente. Não vê efetivo, outros contratos nem contatos internos.</div>':'')
    +'<input id="un_nome" placeholder="Nome completo">'
    +'<input id="un_contato" placeholder="'+(ehCli?'WhatsApp ou e-mail (para enviar o acesso)':'WhatsApp ou e-mail (opcional)')+'">'
    +(ehCli?'':'<input id="un_mat" placeholder="Matrícula (opcional)">')
    +'<div class="s">PIN inicial (a pessoa troca no primeiro acesso)</div><input id="un_pin" maxlength="4" inputmode="numeric" value="'+pinAleatorio()+'">'
    +selCtsHtml("unct",UN.cts,"unTogCt")
    +'<button class="btn" onclick="addUserUI()" style="margin-top:10px">Criar acesso</button></div>';
  return h;
}
function addUserUI(nomeArg,contatoArg,cargoArg,ctsArg){
  if(!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}
  var nome=nomeArg!=null?nomeArg:$("#un_nome").value,contato=contatoArg!=null?contatoArg:(($("#un_contato")||{}).value||"");
  var cargoId=cargoArg||UN.cargo||$("#un_cargo").value,cts=(ctsArg||UN.cts).slice();
  var mat=nomeArg!=null?"":(($("#un_mat")||{}).value||"");
  var pin=nomeArg!=null?pinAleatorio():String(($("#un_pin")||{}).value||"").replace(/\D/g,"");
  if(!nome||!nome.trim()){alert("Informe o nome.");return}
  if(!/^\d{4}$/.test(pin)){alert("PIN de 4 dígitos.");return}
  if(userByName(nome)){alert("Já existe um acesso com esse nome.");return}
  var cg=cargoById(cargoId);
  if(cg&&cg.id=="cli"&&cts.length!=1){alert("Vincule o acesso do cliente a um contrato.");return}
  if(cg&&!cg.fixo&&cg.id!="admin"&&!cts.length&&!confirm("Nenhum contrato vinculado. Criar mesmo assim?"))return;
  var u={id:Date.now(),nome:nome.trim(),mat:(mat||"").trim(),contato:(contato||"").trim(),cargoId:cargoId,pin:pin,pinTrocar:true,contratosIds:cts,idiomaComun:"pt-BR",perfilCompleto:false,criado:Date.now(),criadoPor:user};
  if(u.contato){if(/@/.test(u.contato))u.email=u.contato;else u.telefone=u.contato}
  USERS.push(u);saveUsers();logAcao("Criou acesso",u.nome,(cg||{}).nome+" · "+cts.length+" contrato(s)");
  UN.cts=[];BUSCA["unct"]="";draw();
  enviarAcesso(u.id,pin);
  return u;
}
/* ---- bloco no detalhe do contrato ---- */
function acessosCtBloco(c,pode){
  var gere=temAcao("gerenciar_usuarios");
  var cli=USERS.filter(function(u){var cg=cargoById(u.cargoId);return cg&&cg.id=="cli"&&(u.contratosIds||[]).indexOf(c.id)>=0});
  var h='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin:12px 0 6px">🤝 Acessos do cliente ('+cli.length+')</div>';
  h+=cli.length?cli.map(function(u){return '<div class="row" style="padding:6px 0;flex-wrap:wrap"><span style="flex:1"><b>'+esc(u.nome)+'</b><div class="s">'+esc(u.contato||"sem contato")+(u.perfilCompleto?' · já acessou':' · <b style="color:var(--run)">ainda não entrou</b>')+(u.acessoEnviado?' · 📤 '+esc(u.acessoEnviado.via)+' '+new Date(u.acessoEnviado.t).toLocaleDateString("pt-BR"):'')+'</div></span>'
    +(gere?'<button class="pill pen" style="border:0;cursor:pointer" onclick="enviarAcesso('+u.id+')">📤 Enviar</button><button class="pill pen" style="border:0;cursor:pointer" onclick="resetPinUser('+u.id+')">🔑 PIN</button><button class="pill no" style="border:0;cursor:pointer" onclick="delUser('+u.id+')">X</button>':'')+'</div>'}).join(""):'<div class="s">Nenhum acesso de cliente neste contrato.</div>';
  if(gere)h+='<div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:6px"><input id="cli_nome_'+c.id+'" placeholder="Nome do cliente / gestor do contrato" style="flex:2;min-width:160px;margin:0"><input id="cli_ct_'+c.id+'" placeholder="WhatsApp ou e-mail" style="flex:1;min-width:130px;margin:0"></div><button class="btn alt" onclick="criarAcessoCliente('+c.id+')">➕ Criar acesso do cliente neste contrato</button>';
  else h+='<div class="s">Peça ao administrador para criar o acesso do cliente.</div>';
  return h;
}
function criarAcessoCliente(ctId){
  var n=$("#cli_nome_"+ctId),k=$("#cli_ct_"+ctId);
  var cg=CARGOS.filter(function(x){return x.id=="cli"})[0];if(!cg){alert("Cargo Cliente não existe. Crie em Configurações → Cargos.");return}
  addUserUI(n?n.value:"",k?k.value:"","cli",[ctId]);
}
/* ---- administrar o contrato: atalhos ---- */
function ctIrPara(ctId,t){
  if(meusContratos().some(function(c){return c.id==ctId})){CT_ATUAL=ctId;sv("rad_ctAtual",CT_ATUAL)}
  esCt=ctId;sv("rad_planoCt",ctId);RP.ct=ctId;indCt=ctId;arCt=ctId;sv("rad_arCt",ctId);
  go(t);
}
function atalhosCtBloco(c){
  var ac=acessosDoUsuario(),itens=[["ar","🗺️","Áreas e atividades"],["pti","📋","PTI"],["plano","📊","PTG"],["eq","👥","Equipe"],["es","📦","Estoque"],["l","🔍","Avaliação"],["i","🎯","Indicadores"],["r","📊","Relatório"],["CLI","💬","Chamados do cliente"]].filter(function(x){return ac.indexOf(x[0])>=0});
  if(!itens.length)return "";
  return '<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-bottom:6px">🧭 Administrar este contrato</div><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:6px;margin-bottom:6px">'
    +itens.map(function(x){return '<button class="btn alt" style="margin:0;padding:10px 6px;font-size:13px" onclick="ctIrPara('+c.id+',\''+x[0]+'\')">'+x[1]+' '+x[2]+'</button>'}).join("")
    +'<button class="btn alt" style="margin:0;padding:10px 6px;font-size:13px" onclick="FC.ct='+c.id+';sv(\'rad_fcCt\','+c.id+');FC.passo=\'\';go(\'fc\')">🏭 Fábrica</button>'
    +(temAcao("gerenciar_usuarios")?'<button class="btn alt" style="margin:0;padding:10px 6px;font-size:13px" onclick="document.getElementById(\'cli_nome_'+c.id+'\').scrollIntoView({behavior:\'smooth\'})">🤝 Acesso do cliente</button>':'')
    +'<button class="btn alt" style="margin:0;padding:10px 6px;font-size:13px" onclick="ptCartaz('+c.id+')">📱 QR do contrato</button></div>';
}
/* ---- menu agrupado ---- */
var NAV_GRUPOS=[
  ["",["perfil","ag"]],
  ["Operação",["c","g","l","q"]],
  ["Chamados",["hOP","hAD","CLI"]],
  ["Gestão do contrato",["fc","ct","ar","pti","plano","eq","es","rq"]],
  ["Resultados",["p","i","r"]],
  ["Administração",["ac","imp","st","t","k"]],
  ["",["dc","cont"]]
];

function supCtBloco(c,pode){
  var h='';
  h+='<div class="s" style="text-transform:uppercase;font-weight:700;font-size:11px;margin-top:12px;margin-bottom:6px">📋 Supervisores / Coordenação / Gerência</div>';
  var sups=supervisoresDoContrato(c.id);
  if(sups.length){
    h+=sups.map(function(u){var cg=cargoById(u.cargoId);return '<div class="row" style="padding:6px 0"><span style="flex:1"><b>'+esc(u.nome)+'</b> <span class="s">· '+esc(cg?cg.nome:"")+(u.ctAuto==c.id?' · do departamento':'')+'</span></span>'+(pode?'<button class="pill no" style="border:0;cursor:pointer" onclick="removerSupCt('+u.id+','+c.id+')">✕</button>':'')+'</div>'}).join("");
  } else {
    h+='<div class="s" style="color:var(--warn)">Nenhum supervisor neste contrato.</div>';
  }
  if(pode){
    h+='<div class="s" style="font-weight:700;margin-top:10px">➕ Adicionar supervisor</div>'
      +campoBusca("supct_"+c.id,"Digite o nome do supervisor");
    if(termo("supct_"+c.id)){
      var cands=candidatosSup(c.id).filter(function(x){return bate("supct_"+c.id,x.nome,x.funcao,x.dep,x.mat)});
      h+=cands.length?cands.slice(0,12).map(function(x){return '<div class="row" style="padding:6px 0"><span style="flex:1"><b>'+marca(x.nome,"supct_"+c.id)+'</b><div class="s">'+esc(x.funcao||"")+(x.dep?' · '+esc(x.dep):'')+(x.uid?'':' · ainda sem acesso ao app')+'</div></span><button class="pill ok" style="border:0;cursor:pointer" onclick="adicionarSupCt(\''+x.key+'\','+c.id+')">Adicionar</button></div>'}).join("")+(cands.length>12?'<div class="s">+'+(cands.length-12)+' resultado(s). Continue digitando.</div>':'')
        :'<div class="s">Nenhum supervisor encontrado com esse nome.</div>';
    }
  }

  return h;
}
/* ===================================================================
   FÁBRICA DE CONTRATOS — configuração do contrato passo a passo
   =================================================================== */
var FC={ct:ld("rad_fcCt",null),passo:"",modo:""};
var EST_PADRAO={
  base:[["Papel higiênico","fardo",2],["Papel toalha","fardo",2],["Sabonete líquido","L",5],["Saco de lixo 100 L","pct",3],["Saco de lixo 50 L","pct",3],["Saco de lixo 15 L","pct",3],["Detergente neutro","L",5],["Desinfetante","L",5],["Álcool 70%","L",5],["Limpa-vidros","L",2],["Luva de látex","par",10],["Pano multiuso","un",10],["Fibra / esponja","un",10],["Refil de MOP","un",4]],
  hosp:[["Saco infectante (branco leitoso)","pct",3],["Quaternário de amônio","L",5],["Hipoclorito 1%","L",5],["Máscara descartável","cx",2],["Avental descartável","un",20]],
  ind:[["Desengraxante","L",5],["Luva nitrílica","par",10]],
  edu:[["Papel higiênico rolão","fardo",2]],
  corp:[],fac:[],varejo:[],joa:[["Flanela microfibra","un",10]],banco:[],vig:[],out:[]
};
function fcPode(c){return temAcao("gerenciar_contratos")||canEdit()||(c&&podeAreasCt(c.id))}
function fcSet(id,k,v){var c=ctById(id);if(!c||!fcPode(c))return;v=String(v||"").trim();if(k=="codigo"&&!v){alert("O código não pode ficar vazio.");draw();return}if(k=="codigo"&&CONTRATOS.some(function(x){return x.id!=id&&x.codigo===v})){alert("Já existe contrato com esse código.");draw();return}c[k]=v;svCt();logAcao("Fábrica: editou "+k,c.codigo);if(k=="codigo"||k=="nome")draw()}
function fcMeta(id,k,v){var c=ctById(id);if(!c)return;c.metas=c.metas||{};v=String(v).replace(",",".").trim();if(v===""){delete c.metas[k]}else{var n=+v;if(isNaN(n)||n<0||n>100){alert("Meta entre 0 e 100.");draw();return}c.metas[k]=n}c.metasOk=true;svCt()}
function fcMetasPadrao(id){var c=ctById(id);if(!c)return;c.metas={};c.metasOk=true;svCt();draw()}
function fcEstPadrao(id){
  var c=ctById(id);if(!c)return;var ts=["base"].concat(tiposCt(c)),n=0,ex=estDoCt(id).map(function(i){return normT(i.nome)});
  ts.forEach(function(t){(EST_PADRAO[t]||[]).forEach(function(x){if(ex.indexOf(normT(x[0]))>=0)return;ex.push(normT(x[0]));EST.push({id:Date.now()+Math.floor(Math.random()*1e6),ct:id,nome:x[0],unidade:x[1],qtd:0,min:x[2],local:"",mov:[]});n++})});
  svEst();logAcao("Fábrica: carregou materiais padrão",c.codigo,n+" item(ns)");draw();alert(n+" material(is) padrão cadastrado(s) com quantidade 0. Ajuste o estoque mínimo e lance a primeira entrada (ou faça a inspeção de recebimento).");
}
function fcAddContato(id){
  var c=ctById(id);if(!c)return;var n=$("#fcc_nome").value.trim(),v=$("#fcc_valor").value.trim(),vis=$("#fcc_vis").value;
  if(!n||!v){alert("Informe nome e número/link.");return}
  var tipo=/^https?:/i.test(v)?"link":/@/.test(v)?"mail":"tel";
  CONTATOS.push({id:Date.now(),tipo:tipo,nome:n,valor:v,ico:tipo=="link"?"🔗":tipo=="mail"?"✉️":"📞",cat:vis=="todos"?"Telefones":"Telefones",vis:vis,ct:id});svContatos();logAcao("Fábrica: adicionou contato",c.codigo+" · "+n);draw();
}
function fcDelContato(cid){if(!confirm("Remover este contato?"))return;CONTATOS=CONTATOS.filter(function(x){return x.id!=cid});svContatos();draw()}
function fcPassos(c){
  var ar=areasDoContrato(c.id),pad=ar.filter(function(i){return LOC[i].padrao}).length;
  var colabs=COL.filter(function(x){return !ehDeslig(x.situacao)&&(x.vinculos||[]).some(function(v){return v.contratoId==c.id})});
  var ptis=PTIS.filter(function(p){return p.contratoId==c.id}),aprov=ptis.filter(function(p){return p.status=="Aprovado"}).length;
  var ativ=PG.filter(function(g){return ar.indexOf(g.l)>=0}).length;
  var cli=USERS.filter(function(u){var g=cargoById(u.cargoId);return g&&g.id=="cli"&&(u.contratosIds||[]).indexOf(c.id)>=0}).length;
  var cts=CONTATOS.filter(function(x){return x.ct==c.id}).length;
  return [
    {id:"dados",ico:"📝",nome:"Dados do contrato",ok:!!(c.codigo&&c.nome&&c.responsavel&&(c.endereco||c.cnpj)),res:c.responsavel?esc(c.responsavel)+(c.inicio?" · desde "+esc(fmtDataBR(c.inicio)):""):"Falta responsável/endereço"},
    {id:"tipos",ico:"🧭",nome:"Tipos de serviço",ok:tiposCt(c).some(function(t){return t!="out"}),res:esc(tiposCtNome(c))},
    {id:"sup",ico:"👔",nome:"Supervisão",ok:supervisoresDoContrato(c.id).length>0,res:esc(nomesSupCt(c.id)||"Sem supervisor")},
    {id:"areas",ico:"🗺️",nome:"Áreas e atividades",ok:ar.length>0&&pad<ar.length&&ativ>0,res:ar.length+" área(s) · "+ativ+" atividade(s)"+(pad?" · "+pad+" com nome padrão":"")},
    {id:"equipe",ico:"👥",nome:"Equipe",ok:colabs.length>0,res:colabs.length+" colaborador(es)"},
    {id:"pti",ico:"📋",nome:"PTI e PTG",ok:aprov>0,res:ptis.length+" PTI · "+aprov+" aprovado(s)"},
    {id:"estoque",ico:"📦",nome:"Estoque inicial",ok:estDoCt(c.id).length>0,res:estDoCt(c.id).length+" item(ns)"},
    {id:"metas",ico:"🎯",nome:"Metas dos indicadores",ok:!!c.metasOk,res:c.metasOk?(Object.keys(c.metas||{}).length?Object.keys(c.metas).length+" meta(s) própria(s)":"metas padrão"):"Não revisadas"},
    {id:"contatos",ico:"📞",nome:"Contatos do contrato",ok:cts>0,res:cts+" contato(s)"},
    {id:"cliente",ico:"🤝",nome:"Acesso do cliente",ok:cli>0,res:cli+" acesso(s)"},
    {id:"portal",ico:"📱",nome:"QR do contrato e localização",ok:!!(c.portalTk&&c.geo),res:(c.portalTk?"QR gerado":"QR não gerado")+" · "+(c.geo?"localização ok":"sem localização")}
  ];
}
function fcPassoHtml(c,p){
  var pode=temAcao("gerenciar_contratos")||canEdit();
  var inp=function(k,ph,tp){return '<div class="s">'+ph+'</div><input'+(tp?' type="'+tp+'"':'')+' value="'+esc(c[k]||"")+'" onchange="fcSet('+c.id+',\''+k+'\',this.value)"'+(fcPode(c)?'':' disabled')+'>'};
  if(p=="dados")return inp("codigo","Código")+inp("nome","Nome do contrato")+inp("cnpj","CNPJ do cliente")+inp("endereco","Endereço")+inp("responsavel","Responsável do cliente (gestor do contrato)")+inp("telefone","Telefone do cliente")+inp("email","E-mail do cliente")+'<div style="display:flex;gap:6px"><div style="flex:1">'+inp("inicio","Início da vigência","date")+'</div><div style="flex:1">'+inp("fim","Fim da vigência","date")+'</div></div>'+inp("obs","Observações (escopo, horários, particularidades)");
  if(p=="tipos")return selTipoCt(c)+'<div class="s">Os tipos definem checklists, classificação de risco, EPI, modelos de áreas e materiais padrão.</div>';
  if(p=="sup")return supCtBloco(c,pode);
  if(p=="areas"){var ar=areasDoContrato(c.id),cnt={C:0,S:0,N:0};ar.forEach(function(i){LOC[i].e.forEach(function(n,e){if(!espDel(i,e)){var k=aclOf(i,e);if(cnt[k]!=null)cnt[k]++}})});
    return '<div class="s">'+(ar.map(function(i){return esc(LOC[i].n)+(LOC[i].padrao?' <b style="color:var(--run)">(renomear)</b>':'')}).join(" · ")||"Nenhuma área.")+'</div><div class="s" style="margin-top:4px">Espaços: '+cnt.C+' críticos · '+cnt.S+' semicríticos · '+cnt.N+' não críticos</div>'
      +'<div class="chips" style="margin-top:6px"><button class="sel" onclick="abrirAreasCt('+c.id+')">🗺️ Mapear áreas, espaços e atividades</button></div><div class="s">Dica: para aproveitar a estrutura de outro contrato, use “Clonar contrato” no topo da Fábrica ou “Copiar programação” na tela de Áreas.</div>'}
  if(p=="equipe")return '<div class="s">Os colaboradores vêm do efetivo pelo departamento. Libere as áreas de cada um em Áreas e atividades.</div><div class="chips"><button onclick="ctIrPara('+c.id+',\'eq\')">👥 Abrir equipe</button>'+(acessosDoUsuario().indexOf("imp")>=0?'<button onclick="go(\'imp\')">📥 Importar efetivo</button>':'')+'</div>';
  if(p=="pti")return '<div class="chips"><button onclick="ctIrPara('+c.id+',\'pti\')">📋 Montar PTI</button><button onclick="ctIrPara('+c.id+',\'plano\')">📊 Ver PTG</button><button onclick="imprimirPTG('+c.id+')">🖨️ Imprimir PTG</button></div>';
  if(p=="estoque"){var l=estDoCt(c.id);return '<div class="s">'+(l.length?l.slice(0,12).map(function(i){return esc(i.nome)+" ("+i.qtd+"/"+i.min+" "+esc(i.unidade)+")"}).join(" · ")+(l.length>12?" …":""):"Sem itens.")+'</div><div class="chips" style="margin-top:6px"><button class="sel" onclick="fcEstPadrao('+c.id+')">⚡ Carregar materiais padrão ('+esc(tiposCtNome(c))+')</button><button onclick="ctIrPara('+c.id+',\'es\')">📦 Abrir estoque</button></div>'}
  if(p=="metas"){var D=indDefCt(c.id).filter(function(d){return d[2]=="%"});var m=c.metas||{};
    return '<div class="s">Deixe em branco para usar a meta padrão da Gávea.</div>'+D.map(function(d){var base=IND_DEF_BASE.filter(function(x){return x[0]==d[0]})[0][3];return '<div class="row" style="padding:4px 0"><span style="flex:1" class="s">'+esc(d[1])+' <span style="opacity:.7">(padrão '+base+'%)</span></span><input inputmode="decimal" style="width:80px;margin:0" value="'+(m[d[0]]!=null?m[d[0]]:"")+'" placeholder="'+base+'" onchange="fcMeta('+c.id+',\''+d[0]+'\',this.value)"></div>'}).join("")+(c.metasOk?'':'<button class="btn alt" onclick="fcMetasPadrao('+c.id+')">✔ Usar metas padrão</button>')}
  if(p=="contatos"){var l2=CONTATOS.filter(function(x){return x.ct==c.id});
    return (l2.length?l2.map(function(x){return '<div class="row" style="padding:6px 0"><span style="flex:1"><b>'+esc(x.ico||"📌")+' '+esc(x.nome)+'</b><div class="s">'+esc(x.valor)+' · 👁️ '+esc(visTxt(x))+'</div></span>'+(temAcao("gerenciar_usuarios")?'<button class="pill no" style="border:0;cursor:pointer" onclick="fcDelContato('+x.id+')">X</button>':'')+'</div>'}).join(""):'<div class="s">Nenhum contato só deste contrato (ramal da manutenção, portaria, gestor do cliente, link do sistema do cliente…).</div>')
      +(temAcao("gerenciar_usuarios")?'<input id="fcc_nome" placeholder="Nome (ex.: Manutenção do cliente)"><input id="fcc_valor" placeholder="Telefone, ramal, e-mail ou link"><select id="fcc_vis"><option value="interno">Só equipe interna</option><option value="todos">Todos, inclusive o cliente</option></select><button class="btn alt" onclick="fcAddContato('+c.id+')">➕ Adicionar contato</button>':'')}
  if(p=="cliente")return acessosCtBloco(c,pode);
  if(p=="portal")return portalCtBloco(c,pode);
  return "";
}
function fcNovo(){
  if(!temAcao("gerenciar_contratos")&&!canEdit()){alert("Sem permissão.");return}
  var cod=$("#fcn_cod").value.trim(),nome=$("#fcn_nome").value.trim(),base=+($("#fcn_base")||{}).value||0;
  var tipos=$all(".fcn_tp:checked").map(function(x){return x.value});
  if(!cod||!nome){alert("Informe código e nome.");return}
  if(CONTRATOS.some(function(x){return x.codigo===cod})){alert("Código já existe.");return}
  var b=base?ctById(base):null;
  if(b&&!tipos.length)tipos=tiposCt(b).slice();
  var c=novoContrato({codigo:cod,nome:nome,tipo:tipos[0]||"out",tipos:tipos.length?tipos:["out"]});
  CONTRATOS.push(c);svCt();
  if(b)fcClonar(b,c);
  prepararAreas();
  logAcao("Fábrica: criou contrato",cod+" · "+nome,b?"clonado de "+b.codigo:"");
  FC.ct=c.id;sv("rad_fcCt",c.id);FC.modo="";FC.passo="dados";
  /* quem cria passa a ter o contrato no escopo */
  var u=session?userById(session.userId):null;if(u&&!verTudo()){u.contratosIds=u.contratosIds||[];if(u.contratosIds.indexOf(c.id)<0){u.contratosIds.push(c.id);saveUsers()}}
  draw();window.scrollTo(0,0);
}
function fcClonar(b,c){
  c.metas=JSON.parse(JSON.stringify(b.metas||{}));c.metasOk=!!b.metasOk;
  estDoCt(b.id).forEach(function(i){EST.push({id:Date.now()+Math.floor(Math.random()*1e6),ct:c.id,nome:i.nome,unidade:i.unidade,qtd:0,min:i.min,local:"",mov:[]})});svEst();
  var ar=areasDoContrato(b.id);if(!ar.length){svCt();return}
  CP={a:{},d:{},modo:"nova",ct:null};ar.forEach(function(i){CP.a[i]=true});CP.d[c.id]=true;
  var oc=window.confirm,oa=window.alert;window.confirm=function(){return true};window.alert=function(){};
  try{copiarProgramacao(b.id)}finally{window.confirm=oc;window.alert=oa}
  /* remove áreas de nome padrão que sobraram */
  areasDoContrato(c.id).forEach(function(i){if(LOC[i].padrao)LOC[i].del=true});svLoc();svCt();
}
function fcExcluirVazio(id){
  var c=ctById(id);if(!c)return;
  if(PTIS.some(function(p){return p.contratoId==id})||regs.some(function(r){return ctDaArea(r.l)==id})||COL.some(function(x){return (x.vinculos||[]).some(function(v){return v.contratoId==id})})){alert("Este contrato já tem equipe, PTI ou registros. Use Encerrar em Contratos.");return}
  if(!confirm("Excluir o contrato "+c.codigo+" (ainda sem uso)?"))return;
  areasDoContrato(id).forEach(function(i){LOC[i].del=true});EST=EST.filter(function(i){return i.ct!=id});svEst();svLoc();
  CONTRATOS=CONTRATOS.filter(function(x){return x.id!=id});svCt();FC.ct=null;sv("rad_fcCt",null);logAcao("Fábrica: excluiu contrato vazio",c.codigo);draw();
}
function vFabrica(){
  var mc=meusContratos().slice().sort(function(a,b){return (a.codigo||"").localeCompare(b.codigo||"")});
  var podeCriar=temAcao("gerenciar_contratos")||canEdit();
  var h='<div class="card"><b>🏭 Fábrica de Contratos</b><div class="s">Configure cada contrato de ponta a ponta: dados, serviços, supervisão, áreas, equipe, PTI, estoque, metas, contatos, acesso do cliente e QR. Cada etapa concluída fica ✔.</div>'
    +(podeCriar?'<div class="chips" style="margin-top:6px"><button class="'+(FC.modo=="novo"?'sel':'')+'" onclick="FC.modo=FC.modo==\'novo\'?\'\':\'novo\';draw()">➕ Novo contrato</button><button class="'+(FC.modo=="clone"?'sel':'')+'" onclick="FC.modo=FC.modo==\'clone\'?\'\':\'clone\';draw()">🧬 Clonar contrato existente</button></div>':'')+'</div>';
  if(podeCriar&&FC.modo){
    h+='<div class="card" style="border:2px solid var(--brand)"><b>'+(FC.modo=="clone"?'🧬 Novo contrato a partir de outro':'➕ Novo contrato')+'</b>'
      +(FC.modo=="clone"?'<div class="s">Copia tipos de serviço, áreas, espaços, classificação de risco, atividades, metas e a lista de materiais (com quantidade zero). Equipe, PTI e acessos não são copiados.</div><div class="s">Contrato modelo</div><select id="fcn_base">'+CONTRATOS.map(function(x){return '<option value="'+x.id+'">'+esc(x.codigo+" · "+x.nome)+'</option>'}).join("")+'</select>':'')
      +'<input id="fcn_cod" placeholder="Código (ex.: 0123 ou CT-2026-001)"><input id="fcn_nome" placeholder="Nome do contrato / cliente">'
      +(FC.modo=="novo"?'<div class="s">Tipos de serviço</div><div>'+TIPOS_CT.map(function(t){return '<label class="s" style="display:inline-block;margin:3px 8px 3px 0;cursor:pointer"><input type="checkbox" class="fcn_tp" value="'+t.id+'" style="width:auto;margin-right:4px">'+t.ico+' '+esc(t.nome)+'</label>'}).join("")+'</div>':'')
      +'<button class="btn" onclick="fcNovo()">Criar e configurar</button><div class="s">Contratos que vêm do efetivo (departamentos) já aparecem na lista abaixo; é só configurar.</div></div>';
  }
  if(!mc.length)return h+'<div class="card s">Nenhum contrato disponível.</div>';
  /* visão geral */
  var resumo=mc.map(function(c){var ps=fcPassos(c),ok=ps.filter(function(p){return p.ok}).length;return {c:c,ok:ok,n:ps.length}});
  if(!mc.some(function(c){return c.id==FC.ct}))FC.ct=resumo.slice().sort(function(a,b){return a.ok-b.ok})[0].c.id;
  var c=ctById(FC.ct);var ps=fcPassos(c),ok=ps.filter(function(p){return p.ok}).length,pct=Math.round(ok/ps.length*100);
  h+='<div class="card">'+pickContrato("fcct",mc,c.id,function(id){FC.ct=id;sv("rad_fcCt",id);FC.passo="";draw()})
    +'<div style="display:flex;align-items:center;gap:10px;margin-top:8px"><div style="flex:1;height:10px;background:var(--line);border-radius:99px;overflow:hidden"><div style="width:'+pct+'%;height:100%;background:'+(pct==100?'var(--ok)':pct>=60?'var(--run)':'var(--warn)')+'"></div></div><b>'+ok+'/'+ps.length+'</b></div>'
    +(pct==100?'<div class="s" style="color:var(--ok);font-weight:700;margin-top:4px">✔ Contrato configurado</div>':'<div class="s" style="margin-top:4px">Próxima etapa: <b>'+esc((ps.filter(function(p){return !p.ok})[0]||{}).nome||"")+'</b></div>')+'</div>';
  if(!FC.passo){var pend=ps.filter(function(p){return !p.ok})[0];FC.passo=pend?pend.id:"dados"}
  h+=ps.map(function(p,i){var aberto=FC.passo==p.id;
    return '<div class="card"'+(aberto?' style="border:2px solid var(--brand)"':'')+'><div class="row" style="border:0;padding:0;cursor:pointer" onclick="FC.passo=FC.passo==\''+p.id+'\'?\'-\':\''+p.id+'\';draw()"><span style="width:28px;height:28px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;background:'+(p.ok?'var(--ok)':'var(--line)')+';color:'+(p.ok?'#fff':'var(--ink)')+';margin-right:8px">'+(p.ok?'✔':(i+1))+'</span><span style="flex:1"><b>'+p.ico+' '+esc(p.nome)+'</b><div class="s">'+p.res+'</div></span><span class="s">'+(aberto?'▲':'▼')+'</span></div>'
      +(aberto?'<div style="border-top:1px dashed var(--line);margin-top:8px;padding-top:8px">'+fcPassoHtml(c,p.id)+(i<ps.length-1?'<button class="btn alt" style="margin-top:10px" onclick="FC.passo=\''+ps[i+1].id+'\';draw()">Próxima etapa →</button>':'')+'</div>':'')+'</div>'}).join("");
  if(mc.length>1){
    h+='<div class="card"><b>📋 Situação de todos os contratos</b>'+campoBusca("fcl","Buscar contrato")+resumo.filter(function(r){return bate("fcl",r.c.codigo,r.c.nome,nomesSupCt(r.c.id))}).sort(function(a,b){return a.ok/a.n-b.ok/b.n}).slice(0,60).map(function(r){var p=Math.round(r.ok/r.n*100);
      return '<div class="row" style="cursor:pointer;padding:6px 0" onclick="FC.ct='+r.c.id+';sv(\'rad_fcCt\','+r.c.id+');FC.passo=\'\';draw();window.scrollTo(0,0)"><span style="flex:1">'+marca(r.c.codigo+" · "+r.c.nome,"fcl")+'</span><span class="pill '+(p==100?'ok':p>=60?'run':'no')+'" style="cursor:pointer">'+r.ok+'/'+r.n+'</span></div>'}).join("")+'</div>';
  }
  if(podeCriar)h+='<div class="card"><button class="btn alt" onclick="fcExcluirVazio('+c.id+')">🗑️ Excluir este contrato (só se ainda não tiver uso)</button></div>';
  return h;
}

/* ===================================================================
   CONFIGURAÇÃO DO SISTEMA
   =================================================================== */
var SYS=ld("rad_sys",{url:"",cnpj:"",loginOculto:true});
function svSys(){sv("rad_sys",SYS)}
function sysUrlBase(){var u=(SYS.url||"").trim();if(u)return u.replace(/[?#].*$/,"");return location.origin&&location.origin!="null"&&!/^file:/.test(location.href)?location.origin+location.pathname:""}
function setSys(k,v){if(!canEdit()&&!temAcao("gerenciar_usuarios")){alert("Sem permissão.");return}SYS[k]=typeof v=="string"?v.trim():v;svSys();logAcao("Configuração do sistema",k);draw()}
function sistemaCard(){
  var base=sysUrlBase();
  return '<div class="card"><b>🏢 Empresa e sistema</b>'
    +'<div class="s">Nome da empresa (aparece nos documentos impressos)</div><input value="'+esc(CONTRATO)+'" onchange="setCt(this.value);logAcao(\'Configuração do sistema\',\'empresa\')">'
    +'<div class="s">CNPJ da empresa</div><input value="'+esc(SYS.cnpj||"")+'" onchange="setSys(\'cnpj\',this.value)">'
    +'<div class="s">Endereço (link) do sistema na internet</div><input placeholder="https://..." value="'+esc(SYS.url||"")+'" onchange="setSys(\'url\',this.value)">'
    +'<div class="s">'+(base?'Usado nos QR codes, no portal do contrato e na mensagem de acesso: <b>'+esc(base)+'</b>':'<b style="color:var(--warn)">Sem link.</b> Aberto como arquivo, os QR do portal e as mensagens de acesso não funcionam em outros celulares — informe o link onde o sistema está publicado.')+'</div></div>'
    +'<div class="card"><b>🔐 Login e privacidade</b>'
    +'<div class="row"><span style="flex:1">Esconder a lista de nomes na tela de login<div class="s">A pessoa digita as 3 primeiras letras do nome para aparecer. Evita que cliente ou visitante veja os nomes da equipe.</div></span><button class="pill '+(SYS.loginOculto?'ok':'pen')+'" style="border:0;cursor:pointer" onclick="setSys(\'loginOculto\','+(!SYS.loginOculto)+')">'+(SYS.loginOculto?'Ativado':'Desativado')+'</button></div></div>'
    +'<div class="card"><b>📣 Notificações automáticas</b><div class="s">E-mail e WhatsApp de Compras, Manutenção, RH e demais setores ficam em Setores.</div>'+(acessosDoUsuario().indexOf("st")>=0?'<button class="btn alt" onclick="go(\'st\')">Abrir Setores</button>':'')+'</div>';
}
var CFG_GRUPOS=[
  ["🏢 Sistema",[["Sistema","🏢","Empresa, link e login"],["Aparência","🎨","Logo, cores e fotos"],["Fábrica","🏭","Fábrica de Contratos"]]],
  ["👥 Pessoas e permissões",[["Acessos","🔑","Acessos e clientes"],["Cargos","🎯","Cargos e o que cada um vê"],["Permissões","🛡️","Atalho de edição"]]],
  ["🧹 Padrões operacionais",[["Checklists","📋","Checklists"],["EPI por risco","🦺","EPI por risco"],["Tempos","⏱️","Tempos e prazos"],["Escala","🗓️","Escalas"],["Equipe","👥","Colaboradores"]]],
  ["📞 Comunicação",[["Contatos","📞","Contatos e links"]]],
  ["🔒 Segurança e dados",[["Auditoria","🔍","Log de auditoria"],["Backup","💾","Backup e restauração"]]]
];

/* ===================================================================
   APARÊNCIA — logo, cores, fotos institucionais (login e interno)
   =================================================================== */
var TEMA=ld("rad_tema",{});
var TEMA_PRESETS=[["Gávea (verde)","#0f5c4d"],["Azul institucional","#1d4f91"],["Azul petróleo","#0b6477"],["Grafite","#37474f"],["Vinho","#7b1e3a"],["Laranja","#c4561a"],["Roxo","#5b3f8c"],["Verde-limão","#4f7d1c"]];
function svTema(){try{localStorage.setItem("rad_tema",JSON.stringify(TEMA))}catch(e){alert("Sem espaço para salvar as imagens neste aparelho. Use fotos menores ou remova alguma.");throw e}}
function hexRgb(h){h=String(h||"").replace("#","");if(h.length==3)h=h.split("").map(function(x){return x+x}).join("");var n=parseInt(h,16);return isNaN(n)?null:[n>>16&255,n>>8&255,n&255]}
function mixHex(h,alvo,f){var a=hexRgb(h),b=hexRgb(alvo);if(!a||!b)return h;return "#"+a.map(function(v,i){return ("0"+Math.round(v+(b[i]-v)*f).toString(16)).slice(-2)}).join("")}
function lumin(h){var c=hexRgb(h);if(!c)return 0;c=c.map(function(v){v/=255;return v<=.03928?v/12.92:Math.pow((v+.055)/1.055,2.4)});return .2126*c[0]+.7152*c[1]+.0722*c[2]}
var _fotoTimer=null;
function aplicarTema(){
  try{aplicarWallpaper()}catch(e){}
  var st=document.getElementById("temaCss");if(!st){st=document.createElement("style");st.id="temaCss";document.head.appendChild(st)}
  var css="";
  if(TEMA.cor&&hexRgb(TEMA.cor)){var on=lumin(TEMA.cor)>.45?"#14211e":"#ffffff",dk=mixHex(TEMA.cor,"#ffffff",.45);
    css+=':root:not([data-theme="dark"]){--brand:'+TEMA.cor+';--on:'+on+'}@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--brand:'+dk+';--on:#0f1715}}';}
  var fotos=TEMA.fotosLogin||[];
  if(fotos.length)css+='#login{background:#0f1715 center/cover no-repeat;transition:background-image .8s}#login::before{content:"";position:fixed;inset:0;background:linear-gradient(180deg,rgba(0,0,0,.25),rgba(0,0,0,.55));pointer-events:none}#login .login-card{position:relative;background:color-mix(in srgb,var(--card) 92%,transparent);backdrop-filter:blur(6px)}#loginLegenda{position:fixed;left:0;right:0;bottom:14px;text-align:center;color:#fff;font-size:13px;text-shadow:0 1px 4px rgba(0,0,0,.6);pointer-events:none}';
  st.textContent=css;
  var lg=document.getElementById("login");
  if(lg){clearInterval(_fotoTimer);
    if(fotos.length){var i=0,set=function(){lg.style.backgroundImage='url("'+fotos[i%fotos.length]+'")';i++};set();if(fotos.length>1)_fotoTimer=setInterval(set,6000)}else lg.style.backgroundImage="";
    var leg=document.getElementById("loginLegenda");if(!leg){leg=document.createElement("div");leg.id="loginLegenda";lg.appendChild(leg)}leg.textContent=TEMA.msgLogin||"";}
  var card=document.querySelector(".login-card");
  if(card){var im=document.getElementById("loginLogo");if(TEMA.logo){if(!im){im=document.createElement("img");im.id="loginLogo";im.alt="";im.style.cssText="display:block;max-height:64px;max-width:70%;margin:6px auto 10px;object-fit:contain";var t=document.getElementById("loginTitle");card.insertBefore(im,t)}im.src=TEMA.logo}else if(im)im.remove()}
  if(TEMA.titulo){var a=document.getElementById("loginTitle");if(a)a.textContent=TEMA.titulo}
  if(TEMA.sub){var b=document.getElementById("loginSub");if(b)b.textContent=TEMA.sub}
  var nome=TEMA.titulo||"RAD Digital",logoH=TEMA.logo?'<img src="'+TEMA.logo+'" alt="" style="max-height:28px;max-width:120px;object-fit:contain;vertical-align:middle;margin-right:8px;border-radius:4px">':'';
  var sh=document.querySelector(".sidebar-head span");if(sh)sh.innerHTML=logoH+esc(nome);
  var tt=document.querySelector("#topbar .ttl");if(tt)tt.innerHTML=(TEMA.logo?'<img src="'+TEMA.logo+'" alt="" style="max-height:24px;max-width:90px;object-fit:contain;vertical-align:middle;margin-right:6px;background:#fff;border-radius:4px;padding:1px">':'')+esc(nome);
  if(TEMA.titulo)document.title=TEMA.titulo;
}
function temaPodeEditar(){return canEdit()||temAcao("gerenciar_usuarios")}
function temaSet(k,v){if(!temaPodeEditar())return;TEMA[k]=typeof v=="string"?v.trim():v;svTema();logAcao("Aparência",k);aplicarTema();draw()}
function temaImg(k,inp,max){var f=inp.files&&inp.files[0];if(!f)return;shrinkMax(f,max||1600,function(d){
  if(k=="fotosLogin"){TEMA.fotosLogin=TEMA.fotosLogin||[];if(TEMA.fotosLogin.length>=5){alert("Máximo de 5 fotos no login.");return}TEMA.fotosLogin.push(d)}
  else if(k=="fotosInternas"){TEMA.fotosInternas=TEMA.fotosInternas||[];if(TEMA.fotosInternas.length>=5){alert("Máximo de 5 fotos.");return}TEMA.fotosInternas.push(d)}
  else TEMA[k]=d;
  try{svTema()}catch(e){if(Array.isArray(TEMA[k]))TEMA[k].pop();else delete TEMA[k];return}
  logAcao("Aparência: imagem",k);aplicarTema();draw()})}
function logoImg(inp){var f=inp.files&&inp.files[0];if(!f)return;
  var r=new FileReader();r.onload=function(){var im=new Image();im.onload=function(){var c=document.createElement("canvas"),s=Math.min(1,400/Math.max(im.width,im.height));c.width=Math.round(im.width*s);c.height=Math.round(im.height*s);c.getContext("2d").drawImage(im,0,0,c.width,c.height);TEMA.logo=c.toDataURL("image/png");try{svTema()}catch(e){delete TEMA.logo;return}aplicarTema();draw()};im.src=r.result};r.readAsDataURL(f)}
function temaDelFoto(k,i){(TEMA[k]||[]).splice(i,1);if(k=="fotosInternas"&&TEMA.wpPadrao!=null&&TEMA.wpPadrao>=i)TEMA.wpPadrao=Math.max(-1,TEMA.wpPadrao-1);svTema();aplicarTema();draw()}
function temaReset(){if(!confirm("Voltar ao visual padrão (remove logo, fotos e cor)?"))return;TEMA={};svTema();aplicarTema();draw()}
function aparenciaCard(){
  var cor=TEMA.cor||"#0f5c4d";
  var fotoGrid=function(k){return '<div style="display:flex;gap:6px;flex-wrap:wrap;margin:6px 0">'+(TEMA[k]||[]).map(function(f,i){return '<div style="position:relative"><img src="'+f+'" alt="" style="width:110px;height:72px;object-fit:cover;border-radius:8px;border:1px solid var(--line)"><button onclick="temaDelFoto(\''+k+'\','+i+')" style="position:absolute;top:2px;right:2px;border:0;border-radius:50%;width:22px;height:22px;background:rgba(0,0,0,.6);color:#fff;cursor:pointer">✕</button></div>'}).join("")+'</div>'};
  return '<div class="card"><b>🎨 Aparência</b><div class="s">Logo, cores e fotos institucionais. As imagens ficam salvas neste aparelho e entram no backup.</div></div>'
    +'<div class="card"><b>🖼️ Logo</b>'+(TEMA.logo?'<div style="background:var(--bg);padding:10px;border-radius:8px;margin:6px 0;text-align:center"><img src="'+TEMA.logo+'" alt="" style="max-height:70px;max-width:80%"></div>':'<div class="s">Sem logo.</div>')
    +'<input type="file" accept="image/png,image/jpeg,image/svg+xml,image/webp" onchange="logoImg(this)">'+(TEMA.logo?'<button class="btn alt" onclick="temaSet(\'logo\',\'\')">Remover logo</button>':'')+'<div class="s">Aparece no login, no menu e nos documentos impressos. Prefira PNG com fundo transparente.</div></div>'
    +'<div class="card"><b>🎨 Cor principal</b><div style="display:flex;gap:8px;flex-wrap:wrap;margin:6px 0">'+TEMA_PRESETS.map(function(p){return '<button title="'+esc(p[0])+'" onclick="temaSet(\'cor\',\''+p[1]+'\')" style="width:38px;height:38px;border-radius:50%;border:3px solid '+(cor.toLowerCase()==p[1]?'var(--ink)':'transparent')+';background:'+p[1]+';cursor:pointer"></button>'}).join("")
    +'<label style="display:flex;align-items:center;gap:6px" class="s">Outra: <input type="color" value="'+esc(cor)+'" onchange="temaSet(\'cor\',this.value)" style="width:44px;height:38px;padding:2px;margin:0"></label></div>'
    +'<div style="display:flex;gap:8px;align-items:center;margin-top:4px"><span style="background:var(--brand);color:var(--on);padding:8px 14px;border-radius:8px;font-weight:700">Botão</span><span class="pill ok" style="cursor:default">Exemplo</span></div></div>'
    +'<div class="card"><b>✏️ Nome e textos</b><div class="s">Nome do sistema</div><input placeholder="RAD Digital" value="'+esc(TEMA.titulo||"")+'" onchange="temaSet(\'titulo\',this.value)">'
    +'<div class="s">Subtítulo do login</div><input placeholder="Registro de Atividades Diárias" value="'+esc(TEMA.sub||"")+'" onchange="temaSet(\'sub\',this.value)">'
    +'<div class="s">Frase na tela de login (sobre as fotos)</div><input placeholder="Ex.: Cuidar bem é o nosso trabalho." value="'+esc(TEMA.msgLogin||"")+'" onchange="temaSet(\'msgLogin\',this.value)">'
    +'<div class="s">Mensagem institucional no Painel</div><input placeholder="Ex.: Segurança e qualidade em primeiro lugar." value="'+esc(TEMA.msgInterna||"")+'" onchange="temaSet(\'msgInterna\',this.value)"></div>'
    +'<div class="card"><b>📷 Fotos do login</b> <span class="s">(até 5 — trocam a cada 6 s)</span>'+fotoGrid("fotosLogin")+'<input type="file" accept="image/*" onchange="temaImg(\'fotosLogin\',this,1600)"></div>'
    +'<div class="card"><b>🖼️ Papéis de parede</b> <span class="s">(fundo das telas internas — até 5)</span>'+fotoGrid("fotosInternas")+'<input type="file" accept="image/*" onchange="temaImg(\'fotosInternas\',this,1600)">'
    +((TEMA.fotosInternas||[]).length?'<div class="s" style="margin-top:6px">Papel de parede padrão</div><select onchange="temaSet(\'wpPadrao\',+this.value)"><option value="-1"'+(TEMA.wpPadrao==-1?' selected':'')+'>Nenhum</option>'+(TEMA.fotosInternas||[]).map(function(f,i){return '<option value="'+i+'"'+((TEMA.wpPadrao==null?0:TEMA.wpPadrao)==i?' selected':'')+'>Foto '+(i+1)+'</option>'}).join("")+'</select>':'')
    +'<div class="row"><span style="flex:1">Cada pessoa pode escolher o próprio papel de parede<div class="s">Na galeria acima ou com uma foto própria, em Meu Perfil.</div></span><button class="pill '+(TEMA.wpLivre!==false?'ok':'pen')+'" style="border:0;cursor:pointer" onclick="temaSet(\'wpLivre\','+(TEMA.wpLivre===false)+')">'+(TEMA.wpLivre!==false?'Liberado':'Bloqueado')+'</button></div></div>'
    +'<div class="card"><button class="btn alt" onclick="temaReset()">↩️ Voltar ao visual padrão</button></div>';
}
function bannerInterno(){
  var f=[];if(!TEMA.msgInterna)return "";
  var u=session?userById(session.userId):null,h=new Date().getHours(),saud=h<12?"Bom dia":h<18?"Boa tarde":"Boa noite";
  var foto=f.length?f[Math.floor(Date.now()/60000)%f.length]:"";
  return '<div class="card noprint" style="padding:0;overflow:hidden;position:relative;min-height:'+(foto?'150px':'auto')+';'+(foto?'background:#000 url(\''+foto+'\') center/cover':'background:var(--brand)')+'"><div style="'+(foto?'position:absolute;inset:0;background:linear-gradient(90deg,rgba(0,0,0,.65),rgba(0,0,0,.1));':'')+'padding:18px;color:#fff"><div style="font-size:20px;font-weight:700">'+saud+(u?', '+esc(primeiroNome(u.nome)):'')+'!</div>'+(TEMA.msgInterna?'<div style="margin-top:6px;font-size:14px;max-width:420px">'+esc(TEMA.msgInterna)+'</div>':'')+'</div></div>';
}
/* ===================================================================
   DESEMPENHO — ranking de colaboradores e unidades, destaques e elogios
   =================================================================== */
var ELOG=ld("rad_elog",[]);function svElog(){sv("rad_elog",ELOG)}
var DS={mes:"",esc:"meus",aba:"col"};
function podeVerDesempenho(){var cg=cargoAtual();return !!cg&&!ehClienteCargo()&&!ehColabCargo()}
function mesAtual(){var d=new Date();return d.getFullYear()+"-"+("0"+(d.getMonth()+1)).slice(-2)}
function mesesOpc(){var r=[],d=new Date();for(var i=0;i<12;i++){var x=new Date(d.getFullYear(),d.getMonth()-i,1);r.push(x.getFullYear()+"-"+("0"+(x.getMonth()+1)).slice(-2))}return r}
function colPorNome(n){var k=normT(n||"");if(!k)return null;var u=USERS.filter(function(x){return normT(x.nome)==k})[0];if(u&&u.colId)return cget(u.colId);return COL.filter(function(c){return normT(c.nome)==k})[0]||null}
function desempColab(ctIds,de,ate){
  var naCt=function(l){return ctIds.indexOf(ctDaArea(l))>=0};
  var M={};
  var get=function(colId,nome){var key=colId?"c"+colId:"n"+normT(nome);if(!M[key]){var c=colId?cget(colId):null;M[key]={key:key,colId:colId||null,nome:c?c.nome:nome,funcao:c?c.funcao||"":"",ct:null,limp:0,rec:0,notas:[],nc:0,elog:0}}return M[key]};
  regs.forEach(function(r){if(r.ini<de||r.ini>ate||r.x||!naCt(r.l))return;var c=colPorNome(r.colab);var cg=null;
    var u=USERS.filter(function(x){return normT(x.nome)==normT(r.colab)})[0];if(u){cg=cargoById(u.cargoId);if(cg&&cg.id!="colab"&&!c)return}
    var m=get(c?c.id:null,r.colab);m.ct=m.ct||ctDaArea(r.l);if(r.st=="ok")m.limp++;else if(r.st=="nao")m.rec++});
  AVALS.forEach(function(a){if(a.ini<de||a.ini>ate||!naCt(a.l)||!a.limpou||(!a.limpou.colId&&!a.limpou.nome)||/não identificado/i.test(a.limpou.nome||""))return;var m=get(a.limpou.colId,a.limpou.nome);m.ct=m.ct||ctDaArea(a.l);m.notas.push(a.s.p||0)});
  pa.forEach(function(p){if(!p.causadorId&&!p.causador)return;if((p.ini||0)<de||(p.ini||0)>ate||!naCt(p.l))return;var m=get(p.causadorId,p.causador);m.nc++});
  ELOG.forEach(function(e){if(e.t<de||e.t>ate||ctIds.indexOf(e.ct)<0)return;var m=get(e.colId,e.nome);m.elog++});
  var L=Object.keys(M).map(function(k){return M[k]});
  var maxL=Math.max.apply(null,[1].concat(L.map(function(m){return m.limp})));
  L.forEach(function(m){m.nota=media(m.notas);var prod=m.limp/maxL*100;
    var base=m.nota!=null?m.nota*.6+prod*.4:prod*.7;
    m.pts=Math.max(0,Math.min(100,Math.round(base+5*m.elog-10*m.nc-3*m.rec)));
    m.elegivel=m.limp>=5||m.notas.length>=2});
  return L.sort(function(a,b){return b.pts-a.pts||b.limp-a.limp});
}
function desempUnidades(cts,mes){
  var de=mesIni(mes),ate=mesFim(mes);
  return cts.map(function(c){var x=indValor(c.id,mes).v,cal=indCalc(c.id,de,ate)||{};var meta=indDefCt(c.id)[0][3]||90;
    var conf=x.conf;var pts=conf==null?null:Math.max(0,Math.min(100,Math.round(conf-2*(x.recl||0)+1*(x.elog||0)-1*(x.ncA||0))));
    return {c:c,conf:conf,meta:meta,limp:x.limp||0,insp:x.insp||0,nc:x.ncA||0,recl:x.recl||0,elog:x.elog||0,epi:x.epi,pts:pts}}).sort(function(a,b){return (b.pts==null?-1:b.pts)-(a.pts==null?-1:a.pts)});
}
function medalha(i){return ["🥇","🥈","🥉"][i]||'<span style="display:inline-block;width:22px;text-align:center;font-weight:700">'+(i+1)+'º</span>'}
function barra(v,cor){return '<div style="height:6px;background:var(--line);border-radius:99px;overflow:hidden;margin-top:3px"><div style="width:'+Math.max(0,Math.min(100,v||0))+'%;height:100%;background:'+(cor||'var(--brand)')+'"></div></div>'}
function vDesempenho(ctAtual){
  if(!podeVerDesempenho())return '<div class="card s">Disponível para supervisão, coordenação, gerência e administração.</div>';
  if(!DS.mes)DS.mes=mesAtual();
  var mc=meusContratos(),cts=DS.esc=="ct"?[ctAtual]:mc,ids=cts.map(function(c){return c.id});
  var de=mesIni(DS.mes),ate=mesFim(DS.mes);
  var h='<div class="card"><b>🏆 Desempenho e ranking</b><div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:6px"><select style="flex:1;min-width:130px;margin:0" onchange="DS.mes=this.value;draw()">'+mesesOpc().map(function(m){return '<option value="'+m+'"'+(m==DS.mes?' selected':'')+'>'+mesNome(m)+'</option>'}).join("")+'</select>'
    +'<select style="flex:2;min-width:170px;margin:0" onchange="DS.esc=this.value;draw()"><option value="meus"'+(DS.esc=="meus"?' selected':'')+'>Todos os meus contratos ('+mc.length+')</option><option value="ct"'+(DS.esc=="ct"?' selected':'')+'>Só '+esc(ctAtual.codigo+" · "+ctAtual.nome)+'</option></select></div>'
    +'<div class="chips" style="margin-top:6px">'+[["col","👥 Colaboradores"],["uni","🏢 Unidades"],["elo","🌟 Elogios"]].map(function(a){return '<button class="'+(DS.aba==a[0]?'sel':'')+'" onclick="DS.aba=\''+a[0]+'\';draw()">'+a[1]+'</button>'}).join("")+'</div></div>';
  var L=desempColab(ids,de,ate),top=L.filter(function(m){return m.elegivel}).slice(0,3);
  if(top.length)h+='<div class="card" style="background:linear-gradient(135deg,rgba(224,161,0,.14),transparent)"><b>🏆 Destaques de '+mesNome(DS.mes)+'</b><div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px;margin-top:8px">'
    +top.map(function(m,i){var ct=ctById(m.ct);return '<div style="background:var(--card);border:1px solid var(--line);border-radius:12px;padding:10px;text-align:center"><div style="font-size:28px">'+["🥇","🥈","🥉"][i]+'</div><b>'+esc(m.nome)+'</b><div class="s">'+esc(m.funcao||"")+(ct?'<br>'+esc(ct.nome):'')+'</div><div style="font-size:22px;font-weight:800;color:var(--brand)">'+m.pts+'<span class="s"> pts</span></div><button class="btn alt" style="margin:6px 0 0;padding:8px;font-size:13px" onclick="elogiarUI(\''+m.key+'\')">🌟 Elogiar</button></div>'}).join("")+'</div></div>';
  if(DS.aba=="col"){
    h+='<div class="card"><b>👥 Ranking de colaboradores</b><div class="s">Pontos (0–100) = 60% nota média nas avaliações em que limpou + 40% produtividade (limpezas registradas) + 5 por elogio − 10 por não conformidade causada − 3 por limpeza não realizada. Destaque exige 5 limpezas ou 2 avaliações no mês.</div>'
      +(L.length?'<div class="tbl"><table style="margin-top:6px"><tr><th></th><th>Colaborador</th><th>Pts</th><th>Nota</th><th>Limp.</th><th>NC</th><th>🌟</th><th></th></tr>'+L.slice(0,100).map(function(m,i){return '<tr><td>'+medalha(i)+'</td><td style="white-space:normal;min-width:140px"><b>'+esc(m.nome)+'</b><div class="s">'+esc(m.funcao||"")+(m.ct&&ctById(m.ct)&&ids.length>1?' · '+esc(ctById(m.ct).codigo):'')+(m.elegivel?'':' · poucos dados')+'</div>'+barra(m.pts,m.pts>=85?'var(--ok)':m.pts>=70?'var(--run)':'var(--warn)')+'</td><td><b>'+m.pts+'</b></td><td>'+(m.nota==null?'—':m.nota+'%')+'</td><td>'+m.limp+'</td><td'+(m.nc?' style="color:var(--warn);font-weight:700"':'')+'>'+m.nc+'</td><td>'+m.elog+'</td><td><button class="pill ok" style="border:0;cursor:pointer" onclick="elogiarUI(\''+m.key+'\')">🌟</button></td></tr>'}).join("")+'</table></div>':'<div class="s">Sem registros de limpeza ou avaliação em '+mesNome(DS.mes)+'.</div>')+'</div>';
  }else if(DS.aba=="uni"){
    var U=desempUnidades(cts,DS.mes);
    h+='<div class="card"><b>🏢 Ranking das unidades</b><div class="s">Pontos = conformidade geral − 2 por reclamação + 1 por elogio do cliente − 1 por NC aberta no mês.</div>'
      +'<div class="tbl"><table style="margin-top:6px"><tr><th></th><th>Unidade</th><th>Pts</th><th>Conf.</th><th>Meta</th><th>Limp.</th><th>NC</th><th>Recl.</th><th>Elog.</th></tr>'+U.map(function(u,i){return '<tr><td>'+(u.pts==null?'—':medalha(i))+'</td><td style="white-space:normal;min-width:140px"><b>'+esc(u.c.codigo+" · "+u.c.nome)+'</b><div class="s">👔 '+esc(nomesSupCt(u.c.id)||"—")+'</div>'+(u.pts!=null?barra(u.pts,u.conf>=u.meta?'var(--ok)':'var(--warn)'):'')+'</td><td><b>'+(u.pts==null?'—':u.pts)+'</b></td><td style="color:'+(u.conf==null?'inherit':u.conf>=u.meta?'var(--ok)':'var(--warn)')+';font-weight:700">'+(u.conf==null?'—':u.conf+'%')+'</td><td>'+u.meta+'%</td><td>'+u.limp+'</td><td>'+u.nc+'</td><td>'+u.recl+'</td><td>'+u.elog+'</td></tr>'}).join("")+'</table></div><div class="s">Unidades sem avaliação no mês ficam sem pontuação.</div></div>';
  }else{
    var E=ELOG.filter(function(e){return ids.indexOf(e.ct)>=0}).sort(function(a,b){return b.t-a.t});
    h+='<div class="card"><b>🌟 Elogios registrados ('+E.length+')</b>'+(E.length?E.slice(0,80).map(function(e){var c=ctById(e.ct);return '<div style="border-top:1px solid var(--line);padding:8px 0"><b>'+esc(e.nome)+'</b> <span class="s">· '+new Date(e.t).toLocaleDateString("pt-BR")+(c?' · '+esc(c.codigo):'')+' · por '+esc(e.por)+'</span><div class="s">'+esc(e.txt)+'</div><div class="chips" style="margin-top:4px"><button onclick="certificadoElogio('+e.id+')">🖨️ Certificado</button></div></div>'}).join(""):'<div class="s">Nenhum elogio ainda. Use 🌟 no ranking.</div>')+'</div>';
  }
  return h;
}
function elogiarUI(key){
  if(!podeVerDesempenho())return;
  var de=mesIni(DS.mes||mesAtual()),ate=mesFim(DS.mes||mesAtual());
  var m=desempColab(meusCtIds(),de,ate).filter(function(x){return x.key==key})[0];if(!m)return;
  var sug="Parabéns, "+primeiroNome(m.nome)+"! Seu trabalho em "+mesNome(DS.mes||mesAtual())+" foi destaque: "+(m.nota!=null?"nota média "+m.nota+"% nas avaliações, ":"")+m.limp+" limpezas registradas"+(m.nc?"":" e nenhuma não conformidade")+". Obrigado pelo cuidado e dedicação!";
  var txt=prompt("Elogio para "+m.nome+" (pode editar):",sug);if(txt===null||!txt.trim())return;
  var e={id:Date.now(),colId:m.colId,nome:m.nome,ct:m.ct||(meusCtIds()[0]||null),t:Date.now(),por:user,txt:txt.trim(),pts:m.pts};
  ELOG.push(e);svElog();logAcao("Elogiou colaborador",m.nome);
  var col=m.colId?cget(m.colId):null,u=m.colId?userDoColab(m.colId):null,tel=String((u&&u.telefone)||(col&&col.telefone)||"").replace(/\D/g,"");
  var op=prompt("Elogio registrado! Como compartilhar?\n\n1 - WhatsApp"+(tel?" do colaborador":"")+"\n2 - Imprimir certificado\n3 - Só registrar","1");
  if(op=="1"){var msg="🌟 *Elogio – "+(TEMA.titulo||"Gávea Facilities")+"*\n\n"+e.txt+"\n\n— "+user;if(tel&&tel.length<=11)tel="55"+tel;window.open("https://wa.me/"+(tel||"")+"?text="+encodeURIComponent(msg),"_blank")}
  else if(op=="2")certificadoElogio(e.id);
  draw();
}
function certificadoElogio(id){
  var e=ELOG.filter(function(x){return x.id==id})[0];if(!e)return;var c=ctById(e.ct);
  var html='<div style="border:10px double '+(TEMA.cor||"#0f5c4d")+';padding:40px 36px;text-align:center;font-family:Georgia,serif;min-height:620px">'
    +(TEMA.logo?'<img src="'+TEMA.logo+'" alt="" style="max-height:80px;max-width:240px;margin-bottom:10px">':'<div style="letter-spacing:3px;font-size:14px;color:#555">'+esc(CONTRATO)+'</div>')
    +'<div style="font-size:44px;margin:16px 0 0">🌟</div><h1 style="font-size:36px;margin:6px 0;color:'+(TEMA.cor||"#0f5c4d")+'">Reconhecimento</h1>'
    +'<div style="font-size:16px;color:#555">Conferido a</div><div style="font-size:30px;font-weight:700;margin:10px 0 16px;border-bottom:1px solid #999;display:inline-block;padding:0 30px 6px">'+esc(e.nome)+'</div>'
    +'<p style="font-size:17px;line-height:1.6;max-width:560px;margin:10px auto">'+esc(e.txt)+'</p>'
    +(c?'<div style="font-size:14px;color:#555;margin-top:10px">'+esc(c.nome)+'</div>':'')
    +'<div style="display:flex;justify-content:space-around;margin-top:60px;font-size:13px"><div style="border-top:1px solid #333;padding-top:6px;min-width:200px">'+esc(e.por)+'</div><div style="border-top:1px solid #333;padding-top:6px;min-width:200px">'+new Date(e.t).toLocaleDateString("pt-BR",{day:"2-digit",month:"long",year:"numeric"})+'</div></div></div>';
  abrirImpressao("Reconhecimento_"+e.nome,html);
}
function destaquesPainel(){
  if(!podeVerDesempenho())return "";var m=mesAtual(),L=desempColab(meusCtIds(),mesIni(m),mesFim(m)).filter(function(x){return x.elegivel}).slice(0,3);
  if(!L.length)return "";
  return '<div class="card"><b>🏆 Destaques do mês</b>'+L.map(function(x,i){return '<div class="row"><span>'+["🥇","🥈","🥉"][i]+' <b>'+esc(x.nome)+'</b> <span class="s">'+esc(x.funcao||"")+'</span></span><b>'+x.pts+' pts</b></div>'}).join("")+'<button class="btn alt" onclick="indAba=\'des\';go(\'i\')">Ver ranking completo</button></div>';
}

/* ===================================================================
   PAPEL DE PAREDE — padrão da empresa + escolha de cada usuário
   =================================================================== */
function wpPref(){var u=session?session.userId:null;return u?ld("rad_wp_"+u,{modo:"padrao",op:82}):{modo:"padrao",op:82}}
function wpImgAtual(){
  var g=TEMA.fotosInternas||[],p=wpPref(),livre=TEMA.wpLivre!==false;
  if(!livre||p.modo=="padrao")return TEMA.wpPadrao==null||TEMA.wpPadrao<0?(g.length&&TEMA.wpPadrao==null?g[0]:""):(g[TEMA.wpPadrao]||"");
  if(p.modo=="gal")return g[p.i]||"";
  if(p.modo=="meu")return p.img||"";
  return "";
}
function aplicarWallpaper(){
  if(!document.body||!document.body.style)return;
  var img=session?wpImgAtual():"",p=wpPref();
  if(!img){document.body.style.background="";return}
  var dark=window.matchMedia&&matchMedia("(prefers-color-scheme: dark)").matches&&document.documentElement.getAttribute("data-theme")!="light";
  var op=Math.max(30,Math.min(95,(TEMA.wpLivre!==false&&p.op)||82))/100,c=dark?"15,23,21":"242,245,244";
  document.body.style.background="linear-gradient(rgba("+c+","+op+"),rgba("+c+","+op+")),url('"+img+"') center/cover fixed no-repeat";
}
function wpSet(o){var u=session&&session.userId;if(!u)return;var p=wpPref();for(var k in o)p[k]=o[k];try{localStorage.setItem("rad_wp_"+u,JSON.stringify(p))}catch(e){alert("Imagem grande demais para salvar neste aparelho. Escolha uma foto menor.");return}aplicarWallpaper();draw()}
function wpMinhaFoto(inp){var f=inp.files&&inp.files[0];if(f)shrinkMax(f,1400,function(d){wpSet({modo:"meu",img:d})})}
function wpCard(){
  var g=TEMA.fotosInternas||[],p=wpPref();
  if(TEMA.wpLivre===false)return g.length?'<div class="card"><b>🖼️ Papel de parede</b><div class="s">Definido pela empresa.</div></div>':"";
  var th=function(on,fn,inner,tit){return '<button title="'+esc(tit)+'" onclick="'+fn+'" style="width:96px;height:64px;border-radius:10px;border:3px solid '+(on?'var(--brand)':'var(--line)')+';padding:0;overflow:hidden;cursor:pointer;background:var(--bg);font-size:11px;color:var(--ink)">'+inner+'</button>'};
  var im=function(src){return '<img src="'+src+'" alt="" style="width:100%;height:100%;object-fit:cover">'};
  return '<div class="card"><b>🖼️ Papel de parede</b><div class="s">Escolha o fundo das suas telas. Só muda no seu acesso.</div><div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">'
    +th(p.modo=="padrao","wpSet({modo:\'padrao\'})",(TEMA.wpPadrao!=null&&TEMA.wpPadrao>=0&&g[TEMA.wpPadrao])||(TEMA.wpPadrao==null&&g[0])?im(g[TEMA.wpPadrao==null?0:TEMA.wpPadrao]):"Padrão",'Padrão da empresa')
    +g.map(function(src,i){return th(p.modo=="gal"&&p.i==i,"wpSet({modo:\'gal\',i:"+i+"})",im(src),"Galeria "+(i+1))}).join("")
    +(p.img?th(p.modo=="meu","wpSet({modo:\'meu\'})",im(p.img),"Minha foto"):'')
    +th(p.modo=="nenhum","wpSet({modo:\'nenhum\'})","Sem fundo","Sem papel de parede")
    +'</div><div class="s" style="margin-top:8px">📷 Usar uma foto minha</div><input type="file" accept="image/*" onchange="wpMinhaFoto(this)">'
    +'<div class="s">Transparência do fundo (mais alto = mais suave)</div><input type="range" min="40" max="95" value="'+(p.op||82)+'" onchange="wpSet({op:+this.value})" style="padding:0"></div>';
}
/* ===================================================================
   BIBLIOTECA: POPs, documentos e vídeos — com tradução e ciência de leitura
   =================================================================== */
var DOCS=ld("rad_docs",[]);function svDocs(){sv("rad_docs",DOCS)}
var DOC_LANGS=[["pt-BR","🇧🇷","Português"],["en","🇺🇸","English"],["es","🇪🇸","Español"],["fr","🇫🇷","Français"]];
var DOC_TIPOS=["POP","Instrução de trabalho","Treinamento","Vídeo","Política","Formulário","Documento"];
var DC={sel:null,ed:null,lang:"",tipo:""};
/* arquivos grandes (vídeo/PDF) ficam no IndexedDB do aparelho */
var _idb=null;
function idb(cb){if(_idb){cb(_idb);return}try{var r=indexedDB.open("rad_files",1);r.onupgradeneeded=function(){r.result.createObjectStore("f")};r.onsuccess=function(){_idb=r.result;cb(_idb)};r.onerror=function(){cb(null)}}catch(e){cb(null)}}
function idbPut(k,blob,cb){idb(function(db){if(!db){cb(false);return}var t=db.transaction("f","readwrite");t.objectStore("f").put(blob,k);t.oncomplete=function(){cb(true)};t.onerror=function(){cb(false)}})}
function idbGet(k,cb){idb(function(db){if(!db){cb(null);return}var q=db.transaction("f").objectStore("f").get(k);q.onsuccess=function(){cb(q.result||null)};q.onerror=function(){cb(null)}})}
function idbDel(k){idb(function(db){if(db)db.transaction("f","readwrite").objectStore("f").delete(k)})}
function podeDocs(){return canEdit()||temAcao("gerenciar_cargos")||temAcao("gerenciar_contratos")||["sup","coord","ger"].indexOf((cargoAtual()||{}).id)>=0}
function langUser(){var u=session?userById(session.userId):null;var l=u&&u.idiomaComun;return DOC_LANGS.some(function(x){return x[0]==l})?l:(DOC_LANGS.some(function(x){return x[0]==IDIOMA_UI})?IDIOMA_UI:"pt-BR")}
function ytId(u){var m=/(?:youtube\.com\/(?:watch\?v=|embed\/|shorts\/|live\/)|youtu\.be\/)([\w-]{11})/.exec(u||"");return m?m[1]:null}
function docVisivel(d){
  if(verTudo())return true;var cg=cargoAtual()||{};
  if(cg.id=="cli"&&!d.cli)return false;
  if(d.ct&&meusCtIds().indexOf(d.ct)<0)return false;
  if(d.tipos&&d.tipos.length&&!meusContratos().some(function(c){return d.tipos.some(function(t){return ctTemTipo(c,t)})}))return false;
  return true;
}
function docTemLang(d,l){return l=="pt-BR"?!!((d.txt||{})["pt-BR"]||d.arq||d.video||d.link):!!((d.txt||{})[l]||(d.arqTrad||{})[l])}
function idiomasEquipe(){var c={};USERS.forEach(function(u){var cg=cargoById(u.cargoId);if(u.idiomaComun&&u.idiomaComun!="pt-BR"&&DOC_LANGS.some(function(x){return x[0]==u.idiomaComun}))c[u.idiomaComun]=(c[u.idiomaComun]||0)+1});return c}
function gTradUrl(txt,l){return "https://translate.google.com/?sl=pt&tl="+l.slice(0,2)+"&op=translate&text="+encodeURIComponent(String(txt||"").slice(0,4800))}
function gTradPagina(url,l){return "https://translate.google.com/translate?sl=pt&tl="+l.slice(0,2)+"&u="+encodeURIComponent(url)}
function vDocs(){
  var pode=podeDocs(),L=langUser();if(!DC.lang)DC.lang=L;
  if(DC.ed)return vDocEditor();
  if(DC.sel)return vDocVer(DOCS.filter(function(d){return d.id==DC.sel})[0]);
  var vis=DOCS.filter(docVisivel).filter(function(d){return (!DC.tipo||d.tipo==DC.tipo)&&bate("doc",d.codigo,d.titulo,(d.titulos||{})[L],d.tipo)}).sort(function(a,b){return (a.codigo||a.titulo).localeCompare(b.codigo||b.titulo)});
  var u=session?userById(session.userId):null;
  var h='<div class="card"><b>📚 Documentos, POPs e vídeos</b><div class="s">Conteúdo no seu idioma: <b>'+(DOC_LANGS.filter(function(x){return x[0]==L})[0]||DOC_LANGS[0])[1]+' '+(DOC_LANGS.filter(function(x){return x[0]==L})[0]||DOC_LANGS[0])[2]+'</b> (definido no seu perfil).</div>'
    +campoBusca("doc","Buscar por código, título ou tipo")+'<div class="chips">'+[""].concat(DOC_TIPOS).map(function(t){return '<button class="'+(DC.tipo==t?'sel':'')+'" onclick="DC.tipo=\''+t+'\';draw()">'+(t||"Todos")+'</button>'}).join("")+'</div>'
    +(pode?'<button class="btn" onclick="docNovo()">➕ Novo documento, POP ou vídeo</button>':'')+'</div>';
  if(pode){var eq=idiomasEquipe(),ks=Object.keys(eq);if(ks.length){var falta=ks.map(function(l){var n=DOCS.filter(function(d){return !docTemLang(d,l)}).length;return n?'<b>'+n+'</b> sem '+DOC_LANGS.filter(function(x){return x[0]==l})[0][2]+' ('+eq[l]+' pessoa(s) na equipe)':''}).filter(Boolean);
    if(falta.length)h+='<div class="card" style="border-left:4px solid var(--run)"><b>🌐 Equidade de idioma</b><div class="s">'+falta.join(" · ")+'. Exigência de clientes: colaborador estrangeiro deve receber POPs e treinamentos no seu idioma.</div></div>'}}
  h+='<div class="card">'+(vis.length?vis.map(function(d){var lido=u&&(d.leituras||[]).some(function(x){return x.u==u.id&&x.rev==d.rev});
    return '<div class="row" style="cursor:pointer" onclick="DC.sel='+d.id+';DC.lang=\'\';draw();window.scrollTo(0,0)"><span style="flex:1"><b>'+(d.video?'🎬 ':d.tipo=="POP"?'📘 ':'📄 ')+marca((d.codigo?d.codigo+" · ":"")+((d.titulos||{})[L]||d.titulo),"doc")+'</b><div class="s">'+esc(d.tipo)+(d.rev?' · rev. '+esc(d.rev):'')+(d.ct&&ctById(d.ct)?' · '+esc(ctById(d.ct).codigo):'')+' · '+DOC_LANGS.map(function(x){return docTemLang(d,x[0])?x[1]:''}).join("")+'</div></span>'+(lido?'<span class="pill ok" style="cursor:default">✔ lido</span>':'<span class="pill pen" style="cursor:default">ler</span>')+'</div>'}).join(""):'<div class="s">Nenhum documento'+(pode?'. Cadastre o primeiro POP ou vídeo.':'.')+'</div>')+'</div>';
  return h;
}
function docMidia(d,l,cb){
  var h="";
  if(d.video){var y=ytId(d.video);
    if(y)h+='<div style="position:relative;padding-top:56.25%;border-radius:10px;overflow:hidden;background:#000;margin:8px 0"><iframe src="https://www.youtube-nocookie.com/embed/'+y+'?rel=0&cc_load_policy=1&cc_lang_pref='+l.slice(0,2)+'&hl='+l.slice(0,2)+'" title="'+esc(d.titulo)+'" allow="accelerometer;encrypted-media;picture-in-picture;fullscreen" allowfullscreen style="position:absolute;inset:0;width:100%;height:100%;border:0"></iframe></div>'+(l!="pt-BR"?'<div class="s">Legendas automáticas no seu idioma ativadas (se o vídeo tiver legenda). Use ⚙️ no vídeo → Legendas → Traduzir automaticamente.</div>':'');
    else if(/\.(mp4|webm|ogg)(\?|$)/i.test(d.video))h+='<video src="'+esc(d.video)+'" controls style="width:100%;border-radius:10px;margin:8px 0"></video>';
    else h+='<a class="btn" href="'+esc(d.video)+'" target="_blank" rel="noopener" style="display:block;text-align:center;text-decoration:none">🎬 Abrir vídeo</a>';}
  if(d.vidArq)h+='<video id="docVid" controls style="width:100%;border-radius:10px;margin:8px 0;background:#000"></video>';
  return h;
}
function vDocVer(d){
  if(!d){DC.sel=null;return vDocs()}
  var L=DC.lang||langUser(),txt=(d.txt||{})[L],pt=(d.txt||{})["pt-BR"]||"",u=session?userById(session.userId):null,pode=podeDocs();
  var tit=(d.titulos||{})[L]||d.titulo,arq=L!="pt-BR"&&(d.arqTrad||{})[L]?(d.arqTrad||{})[L]:d.arq;
  var lido=u&&(d.leituras||[]).filter(function(x){return x.u==u.id&&x.rev==d.rev})[0];
  var h='<div class="card"><button class="btn alt" onclick="DC.sel=null;draw()">← Voltar à biblioteca</button>'
    +'<div class="s" style="margin-top:6px">'+esc(d.tipo)+(d.codigo?' · '+esc(d.codigo):'')+(d.rev?' · rev. '+esc(d.rev):'')+'</div><b style="font-size:19px">'+esc(tit)+'</b>'
    +'<div class="chips" style="margin-top:6px">'+DOC_LANGS.map(function(x){var tem=docTemLang(d,x[0]);return '<button class="'+(L==x[0]?'sel':'')+'" onclick="DC.lang=\''+x[0]+'\';draw()" style="'+(tem?'':'opacity:.55')+'">'+x[1]+' '+x[2]+(tem?'':' ✖')+'</button>'}).join("")+'</div>'
    +docMidia(d,L);
  if(txt)h+='<div style="white-space:pre-wrap;line-height:1.55;margin-top:8px">'+esc(txt)+'</div>';
  else if(pt){h+='<div class="s" style="background:rgba(224,161,0,.12);padding:8px;border-radius:8px;margin-top:8px">'+(L=="pt-BR"?'':'Tradução oficial para '+DOC_LANGS.filter(function(x){return x[0]==L})[0][2]+' ainda não cadastrada. <a href="'+gTradUrl(pt,L)+'" target="_blank" rel="noopener" style="color:var(--brand);font-weight:700">🌐 Ver tradução automática</a> · abaixo, o original:')+'</div><div style="white-space:pre-wrap;line-height:1.55;margin-top:8px">'+esc(pt)+'</div>'}
  if(arq)h+='<button class="btn alt" style="margin-top:8px" onclick="docAbrirArq('+d.id+',\''+(arq===d.arq?'pt-BR':L)+'\')">📎 Abrir '+esc(arq.nome)+'</button>'+(L!="pt-BR"&&arq===d.arq?'<div class="s">Arquivo apenas em português. Peça a versão traduzida ao gestor.</div>':'');
  if(d.link)h+='<a class="btn alt" href="'+esc(d.link)+'" target="_blank" rel="noopener" style="display:block;text-align:center;text-decoration:none;margin-top:8px">🔗 Abrir documento</a>'+(L!="pt-BR"?'<a href="'+gTradPagina(d.link,L)+'" target="_blank" rel="noopener" class="s" style="color:var(--brand);display:block;text-align:center">🌐 Abrir traduzido automaticamente</a>':'');
  h+='</div><div class="card">'+(lido?'<div class="s" style="color:var(--ok);font-weight:700">✔ Você confirmou a leitura em '+new Date(lido.t).toLocaleString("pt-BR")+' ('+esc(lido.lang)+').</div>':'<button class="btn" onclick="docLido('+d.id+')">✔ Li e entendi</button><div class="s">Sua confirmação fica registrada como evidência de treinamento.</div>')+'</div>';
  if(pode){var lei=(d.leituras||[]).filter(function(x){return x.rev==d.rev});
    h+='<div class="card"><b>📋 Confirmações de leitura — rev. '+esc(d.rev||"—")+' ('+lei.length+')</b>'+(lei.length?lei.slice().sort(function(a,b){return b.t-a.t}).slice(0,60).map(function(x){return '<div class="row" style="padding:4px 0"><span style="flex:1">'+esc(x.nome)+'</span><span class="s">'+esc(x.lang)+' · '+new Date(x.t).toLocaleDateString("pt-BR")+'</span></div>'}).join(""):'<div class="s">Ninguém confirmou ainda.</div>')
      +'<div class="chips" style="margin-top:6px"><button onclick="DC.ed=JSON.parse(JSON.stringify(DOCS.filter(function(x){return x.id=='+d.id+'})[0]));draw()">✏️ Editar e traduzir</button><button onclick="docImprimir('+d.id+',\''+L+'\')">🖨️ Imprimir '+(L!="pt-BR"?'bilíngue':'')+'</button><button onclick="docLeiturasCSV('+d.id+')">⬇️ Lista de leitura</button></div></div>';}
  if(d.vidArq)setTimeout(function(){idbGet("vid_"+d.id,function(b){var v=document.getElementById("docVid");if(v&&b)v.src=URL.createObjectURL(b)})},30);
  return h;
}
function docLido(id){var d=DOCS.filter(function(x){return x.id==id})[0],u=session?userById(session.userId):null;if(!d||!u)return;d.leituras=d.leituras||[];d.leituras.push({u:u.id,nome:u.nome,t:Date.now(),lang:DC.lang||langUser(),rev:d.rev});svDocs();logAcao("Confirmou leitura",(d.codigo||"")+" "+d.titulo,DC.lang||langUser());draw()}
function docAbrirArq(id,l){var d=DOCS.filter(function(x){return x.id==id})[0];if(!d)return;var k=l=="pt-BR"?"arq_"+id:"arq_"+id+"_"+l;idbGet(k,function(b){if(!b){alert("Arquivo não encontrado neste aparelho.");return}var url=URL.createObjectURL(b);var w=window.open(url,"_blank");if(!w){var a=document.createElement("a");a.href=url;a.download=(l=="pt-BR"?d.arq:(d.arqTrad||{})[l]).nome;document.body.appendChild(a);a.click();a.remove()}})}
function docLeiturasCSV(id){var d=DOCS.filter(function(x){return x.id==id})[0];if(!d)return;var L=[["Documento","Revisão","Colaborador","Idioma","Data/hora"]].concat((d.leituras||[]).map(function(x){return [(d.codigo||"")+" "+d.titulo,x.rev||"",x.nome,x.lang,new Date(x.t).toLocaleString("pt-BR")]}));baixar("Leituras_"+(d.codigo||d.id)+".csv",L.map(function(l){return l.map(function(v){return '"'+String(v).replace(/"/g,'""')+'"'}).join(";")}).join("\n"),"text/csv")}
function docImprimir(id,l){var d=DOCS.filter(function(x){return x.id==id})[0];if(!d)return;var pt=(d.txt||{})["pt-BR"]||"",tr=l!="pt-BR"?(d.txt||{})[l]||"":"";var ln=DOC_LANGS.filter(function(x){return x[0]==l})[0];
  var corpo=cabDoc((d.codigo?d.codigo+" — ":"")+d.titulo+(tr?" / "+((d.titulos||{})[l]||""):""),(d.codigo||d.tipo)+(d.rev?" rev. "+d.rev:""))
    +(tr?'<table style="width:100%;border-collapse:collapse;table-layout:fixed"><tr><th style="border:1px solid #999;padding:6px;background:#eee">🇧🇷 Português</th><th style="border:1px solid #999;padding:6px;background:#eee">'+ln[1]+' '+ln[2]+'</th></tr><tr><td style="border:1px solid #999;padding:8px;vertical-align:top;white-space:pre-wrap">'+esc(pt)+'</td><td style="border:1px solid #999;padding:8px;vertical-align:top;white-space:pre-wrap">'+esc(tr)+'</td></tr></table>':'<div style="white-space:pre-wrap">'+esc(pt)+'</div>')
    +(d.video?'<p>Vídeo: '+esc(d.video)+'</p>':'')+'<h3>Ciência</h3><table style="width:100%;border-collapse:collapse">'+[0,1,2,3,4,5,6,7].map(function(){return '<tr><td style="border:1px solid #999;height:26px;width:55%"></td><td style="border:1px solid #999"></td><td style="border:1px solid #999;width:18%"></td></tr>'}).join("").replace('<tr>','<tr><th style="border:1px solid #999">Nome</th><th style="border:1px solid #999">Assinatura</th><th style="border:1px solid #999">Data</th></tr><tr>')+'</table>'+rodDoc();
  abrirImpressao((d.codigo||"DOC")+"_"+l,corpo)}
/* ---- editor ---- */
function docNovo(){DC.ed={id:0,codigo:"",titulo:"",tipo:"POP",rev:"00",ct:null,tipos:[],cli:false,txt:{},titulos:{},video:"",link:""};DC.lang="pt-BR";draw()}
function docEdSet(k,v){DC.ed[k]=v}
function docEdTxt(l,v){DC.ed.txt=DC.ed.txt||{};DC.ed.txt[l]=v}
function docEdTit(l,v){DC.ed.titulos=DC.ed.titulos||{};DC.ed.titulos[l]=v}
var _docPend={};
function docEdArq(inp,l){var f=inp.files&&inp.files[0];if(!f)return;if(f.size>60*1024*1024){alert("Arquivo acima de 60 MB. Use um link (Drive, SharePoint, YouTube).");return}_docPend[l]=f;var y=window.scrollY;if(l=="pt-BR")DC.ed.arq={nome:f.name,tam:f.size};else{DC.ed.arqTrad=DC.ed.arqTrad||{};DC.ed.arqTrad[l]={nome:f.name,tam:f.size}}draw();window.scrollTo(0,y)}
function docEdVid(inp){var f=inp.files&&inp.files[0];if(!f)return;if(f.size>200*1024*1024){alert("Vídeo acima de 200 MB. Publique no YouTube (pode ser “não listado”) e cole o link.");return}_docPend.vid=f;DC.ed.vidArq={nome:f.name,tam:f.size};var y=window.scrollY;draw();window.scrollTo(0,y)}
function vDocEditor(){
  var d=DC.ed,l=DC.lang||"pt-BR",pt=(d.txt||{})["pt-BR"]||"";
  var ta=function(v,fn,rows,ph){return '<textarea rows="'+rows+'" placeholder="'+esc(ph||"")+'" oninput="'+fn+'" style="width:100%;padding:11px;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink);font:inherit">'+esc(v||"")+'</textarea>'};
  var h='<div class="card" style="border:2px solid var(--brand)"><b>'+(d.id?'✏️ Editar':'➕ Novo')+' documento</b>'
    +'<div style="display:flex;gap:6px"><div style="flex:1"><div class="s">Código</div><input value="'+esc(d.codigo)+'" placeholder="POP-081" oninput="docEdSet(\'codigo\',this.value)"></div><div style="width:90px"><div class="s">Revisão</div><input value="'+esc(d.rev||"")+'" oninput="docEdSet(\'rev\',this.value)"></div></div>'
    +'<div class="s">Tipo</div><select onchange="docEdSet(\'tipo\',this.value)">'+DOC_TIPOS.map(function(t){return '<option'+(d.tipo==t?' selected':'')+'>'+t+'</option>'}).join("")+'</select>'
    +'<div class="s">Disponível para</div><select onchange="docEdSet(\'ct\',this.value?+this.value:null)"><option value="">Todos os contratos</option>'+meusContratos().map(function(c){return '<option value="'+c.id+'"'+(d.ct==c.id?' selected':'')+'>'+esc(c.codigo+" · "+c.nome)+'</option>'}).join("")+'</select>'
    +'<label class="s" style="display:block;margin:6px 0"><input type="checkbox" style="width:auto;margin-right:6px"'+(d.cli?' checked':'')+' onchange="docEdSet(\'cli\',this.checked)">Cliente também pode ver</label>'
    +'<div class="s">🎬 Vídeo — link do YouTube, Drive, Vimeo ou arquivo .mp4</div><input value="'+esc(d.video||"")+'" placeholder="https://youtu.be/..." oninput="docEdSet(\'video\',this.value)">'
    +'<div class="s">ou enviar arquivo de vídeo (fica neste aparelho)'+(d.vidArq?': <b>'+esc(d.vidArq.nome)+'</b>':'')+'</div><input type="file" accept="video/*" onchange="docEdVid(this)">'
    +'<div class="s">🔗 Link do documento (SharePoint, Drive, site)</div><input value="'+esc(d.link||"")+'" placeholder="https://..." oninput="docEdSet(\'link\',this.value)"></div>';
  h+='<div class="card"><b>🌐 Conteúdo por idioma</b><div class="chips">'+DOC_LANGS.map(function(x){var tem=x[0]=="pt-BR"?!!pt:!!((d.txt||{})[x[0]]||(d.arqTrad||{})[x[0]]);return '<button class="'+(l==x[0]?'sel':'')+'" onclick="DC.lang=\''+x[0]+'\';draw()">'+x[1]+' '+x[2]+(tem?' ✔':'')+'</button>'}).join("")+'</div>'
    +'<div class="s">Título ('+l+')</div><input value="'+esc(l=="pt-BR"?d.titulo:(d.titulos||{})[l]||"")+'" oninput="'+(l=="pt-BR"?"docEdSet(\'titulo\',this.value)":"docEdTit(\'"+l+"\',this.value)")+'">'
    +'<div class="s">Texto do POP / instruções ('+l+')</div>'+ta((d.txt||{})[l],"docEdTxt(\'"+l+"\',this.value)",12,l=="pt-BR"?"Objetivo, materiais, EPI, passo a passo, cuidados…":"Cole aqui a tradução revisada")
    +(l!="pt-BR"&&pt?'<a class="btn alt" href="'+gTradUrl((d.titulo?d.titulo+"\n\n":"")+pt,l)+'" target="_blank" rel="noopener" style="display:block;text-align:center;text-decoration:none">🌐 Traduzir o texto em português com o Google</a><div class="s">Abre o tradutor com o texto pronto: copie o resultado, revise termos técnicos e cole acima. Para PDF/Word, use translate.google.com → Documentos e anexe o arquivo traduzido abaixo.</div>':'')
    +'<div class="s" style="margin-top:6px">📎 Arquivo ('+l+') — PDF, Word, imagem'+(l=="pt-BR"&&d.arq?': <b>'+esc(d.arq.nome)+'</b>':(l!="pt-BR"&&(d.arqTrad||{})[l]?': <b>'+esc(d.arqTrad[l].nome)+'</b>':''))+'</div><input type="file" accept=".pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,image/*" onchange="docEdArq(this,\''+l+'\')"></div>';
  h+='<div class="card"><div class="grid-acoes"><button class="btn" onclick="docSalvar()">💾 Salvar</button><button class="btn alt" onclick="DC.ed=null;_docPend={};draw()">Cancelar</button></div>'+(d.id?'<button class="btn red" style="margin-top:6px" onclick="docExcluir('+d.id+')">Excluir documento</button>':'')+'<div class="s">Ao mudar a revisão, as confirmações de leitura começam de novo.</div></div>';
  return h;
}
function docSalvar(){
  var d=DC.ed;if(!d.titulo||!d.titulo.trim()){DC.lang="pt-BR";alert("Informe o título em português.");draw();return}
  if(!((d.txt||{})["pt-BR"]||d.video||d.link||d.arq||d.vidArq)){alert("Inclua o texto, um vídeo, um link ou um arquivo.");return}
  var novo=!d.id;if(novo){d.id=Date.now();d.criado=Date.now();d.por=user;d.leituras=[]}
  d.alt=Date.now();var pend=_docPend;_docPend={};
  var ix=DOCS.map(function(x){return x.id}).indexOf(d.id);if(ix>=0)DOCS[ix]=d;else DOCS.push(d);svDocs();
  var ks=Object.keys(pend),n=ks.length,fim=function(){logAcao(novo?"Cadastrou documento":"Editou documento",(d.codigo||"")+" "+d.titulo);DC.ed=null;DC.sel=d.id;draw()};
  if(!n){fim();return}
  ks.forEach(function(k){var key=k=="vid"?"vid_"+d.id:k=="pt-BR"?"arq_"+d.id:"arq_"+d.id+"_"+k;idbPut(key,pend[k],function(ok){if(!ok)alert("Não foi possível guardar "+pend[k].name+" neste aparelho.");if(--n==0)fim()})});
}
function docExcluir(id){if(!confirm("Excluir este documento e as confirmações de leitura?"))return;var d=DOCS.filter(function(x){return x.id==id})[0];DOCS=DOCS.filter(function(x){return x.id!=id});svDocs();idbDel("vid_"+id);idbDel("arq_"+id);DOC_LANGS.forEach(function(x){idbDel("arq_"+id+"_"+x[0])});logAcao("Excluiu documento",d&&d.titulo);DC.ed=null;DC.sel=null;draw()}

/* ===================================================================
   TRADUÇÃO DO "COMO FAZER" (en / es / fr) — padrão revisado + personalizado
   =================================================================== */
var CF_T={
en:[
"Unoccupied room (scheduled or after discharge/transfer). 1) Wear all PPE required for the area. 2) Remove waste and linen. 3) Clean from top to bottom: ceiling, light fixtures, air vents, walls, windows and doors. 4) Clean and disinfect furniture and equipment, including the undersides. 5) Wash the floor with two buckets (solution and rinse) or a machine, from the back towards the exit. 6) Apply the standard disinfectant to high-touch surfaces. 7) Restock supplies and record the task.",
"Room in use, every day. 1) Wear the PPE required for the area. 2) Collect the trash and replace the bags. 3) Clean high-touch surfaces (door handles, switches, countertops, tables, equipment) with a damp cloth and detergent; then apply the standard disinfectant. Work from the cleanest to the dirtiest area and from top to bottom. 4) Clean the floor with a damp mop/cloth, from the back towards the exit. 5) Restock supplies. 6) Place wet-floor signs.",
"Always after cleaning with detergent. Apply the standard disinfectant (e.g., 70% alcohol by rubbing or quaternary ammonium, according to the product data sheet) to door handles, handrails, switches, bed rails, countertops and buttons. Respect the product's contact time. Use a dedicated cloth and change it when moving to another room.",
"Long-cuff rubber gloves, waterproof apron and boots. Close the milky-white bag when it is at most 2/3 full, without squeezing it or pushing the air out with your hands; never transfer waste from one bag to another. Fill the sharps box only up to the dotted line. Take it to the waste shelter in the closed collection cart, along the defined route. Sanitize the cart and gloves at the end.",
"Separate paper, plastic, metal and glass into the labelled bins; do not mix in organic waste. Close the bags when at most 2/3 full and take them to the recycling collection point.",
"Close the bag when it is at most 2/3 full, without squeezing it. Replace the bag, clean the bin on the outside (and inside during terminal cleaning) and put the lid back. Separate recyclables according to selective collection. Take it to the waste room along the defined route.",
"1) Place signs in the area. 2) Collect the trash. 3) Apply the product to the toilet and urinal and let it act. 4) Clean the sink, faucet, mirror and dispensers. 5) Scrub the toilet inside, then clean the outside, seat and flush. 6) Clean walls, partitions and doors. 7) Wash or clean the floor last, from the back towards the door. 8) Restock toilet paper, soap and hand sanitizer. Use cloths of a color reserved for restrooms.",
"Check and restock toilet paper, paper towels, liquid soap and hand sanitizer. Replace the soap refill; never top it up. Missing supplies: open an operational ticket.",
"Use glass cleaner and a squeegee, from top to bottom, removing the excess with a microfiber cloth. Above 2 meters high, only with NR-35 training and proper equipment.",
"Buttons with a cloth lightly dampened with 70% alcohol (never spray directly). Mirror and walls with glass cleaner or a damp cloth. Floor with a damp mop, and clean the door track.",
"Sweep or use the scrubber. Oil stains: absorbent first, then degreaser and rinse. Empty the bins and keep the wet-floor signs in place.",
"Contain the spill with absorbent material, collect it and dispose of it as contaminated waste. Apply degreaser, scrub, rinse and dry. Keep the area signposted until dry.",
"In healthcare areas do not dry-sweep: use a damp mop. Elsewhere, sweep or vacuum first. Wash with detergent and rinse, or use a damp mop, from the back towards the exit. Place wet-floor signs and let it dry.",
"Do not move documents without permission. Desktop with a damp cloth and neutral detergent. Keyboard, mouse and phone with a cloth lightly dampened with 70% alcohol (never spray). Clean the chairs and tidy up the room.",
"Erase the board, clean it with a damp cloth or the product suited to the type of board, and clean the eraser tray.",
"Collect the trash, clean desks and tables, erase and clean the boards, clean door handles and switches, and finally the floor, from the back towards the door.",
"Countertops and sinks with detergent and then disinfectant. Do not touch reagents, samples or running equipment. Chemical and biological waste only as instructed by the person in charge of the laboratory.",
"Clean countertops, sink, tables, microwave and coffee maker. Refrigerator outside daily and inside according to the schedule. Use a cloth reserved for the kitchen. Discard organic waste and wash the floor last.",
"Microfiber cloth and ammonia-free glass cleaner, without spraying on the items. Never open showcases or display cases without a store representative present.",
"Only accompanied by a client representative. Do not move cash, documents or equipment. Record the time and the name of the person who accompanied you.",
"Cloth lightly dampened with 70% alcohol on the screen, keypad and card reader; never spray liquid. Do not touch the cash or envelope compartments.",
"Glass with glass cleaner, from top to bottom; frame and handles with a damp cloth and 70% alcohol. Do not force the mechanism.",
"Damp cloth with neutral detergent on the counter and service surfaces; 70% alcohol on pens, intercoms and touch points. Do not move documents.",
"Handles and grips with 70% alcohol or disinfectant; frame with a damp cloth and detergent.",
"Collect hangers and forgotten items, clean the mirror, bench and door handle, and mop the floor.",
"Conveyor belt, counter and card reader/machine with a cloth lightly dampened with 70% alcohol, without spraying on equipment.",
"Watering, pruning, leaf removal and fertilizing according to the schedule. Collect green waste. Wear gloves, goggles and shin guards (ear protection with the brush cutter) and isolate the area while using the brush cutter.",
"Inspect and record leaks, lighting, doors and locks. Fix what is within scope; for the rest, open an operational ticket with a photo.",
"Follow the defined route at the scheduled times, check doors, windows, lighting and unknown persons. Record the time and the points checked. Report incidents to the supervisor immediately.",
"Identify and register every visitor (name, ID document, destination and time), confirm authorization with the department, hand over the badge and record the exit.",
"Record in the logbook or system: date, time, place, people involved, an objective description of the event and actions taken. Inform the supervisor.",
"Follow the post checklist: doors, windows, alarm, lights and keys. Record the time and any abnormality.",
"Monitor the cameras, confirm they are recording, record failures and call the patrol in suspicious situations.",
"Cloth dampened with neutral detergent, from top to bottom and from the cleanest to the dirtiest; dry. On high-touch surfaces, then apply the standard disinfectant.",
"Sweep, collect leaves and trash, empty outdoor bins and wash when necessary. Place wet-floor signs.",
"1) Wear the PPE required for the area. 2) Collect the trash. 3) Clean from top to bottom and from the cleanest to the dirtiest: surfaces, furniture and touch points with a damp cloth and detergent; then disinfect the touch points. 4) Clean the floor last, from the back towards the exit. 5) Place wet-floor signs and restock supplies."
],
es:[
"Ambiente desocupado (programada o tras alta/traslado). 1) Use todos los EPP del área. 2) Retire residuos y ropa. 3) Limpie de arriba hacia abajo: techo, luminarias, salidas de aire, paredes, ventanas y puertas. 4) Limpie y desinfecte mobiliario y equipos, incluso las partes de abajo. 5) Lave el piso con dos baldes (solución y enjuague) o máquina, desde el fondo hacia la salida. 6) Aplique el desinfectante estandarizado en las superficies de contacto. 7) Reponga insumos y registre.",
"Con el ambiente en uso, todos los días. 1) Use los EPP del área. 2) Recoja la basura y cambie las bolsas. 3) Limpie las superficies de contacto (manijas, interruptores, mesadas, mesas, equipos) con paño húmedo y detergente; después aplique el desinfectante estandarizado. Trabaje de lo más limpio a lo más sucio y de arriba hacia abajo. 4) Limpie el piso con mopa/paño húmedo, desde el fondo hacia la salida. 5) Reponga insumos. 6) Señalice el piso mojado.",
"Siempre después de la limpieza con detergente. Aplique el desinfectante estandarizado (ej.: alcohol al 70% por fricción o amonio cuaternario, según la ficha técnica) en manijas, pasamanos, interruptores, barandas de cama, mesadas y botones. Respete el tiempo de contacto del producto. Use un paño exclusivo y cámbielo al pasar a otro ambiente.",
"Guante de goma de caña larga, delantal impermeable y botas. Cierre la bolsa blanca lechosa con hasta 2/3 de su capacidad, sin apretar ni sacar el aire con las manos; nunca pase residuos de una bolsa a otra. La caja de cortopunzantes solo hasta la línea punteada. Llévela al depósito en el carro de recolección cerrado, por la ruta definida. Higienice el carro y los guantes al final.",
"Separe papel, plástico, metal y vidrio en los recipientes identificados; no mezcle residuos orgánicos. Cierre las bolsas con hasta 2/3 de su capacidad y llévelas al punto de recolección selectiva.",
"Cierre la bolsa con hasta 2/3 de su capacidad, sin apretar. Cambie la bolsa, limpie el basurero por fuera (y por dentro en la limpieza terminal) y tápelo. Separe los reciclables según la recolección selectiva. Llévela al depósito de basura por el trayecto definido.",
"1) Señalice el área. 2) Recoja la basura. 3) Aplique el producto en el inodoro y el urinario y déjelo actuar. 4) Limpie lavabo, grifo, espejo y dispensadores. 5) Cepille el inodoro por dentro y límpielo por fuera, asiento y descarga. 6) Limpie paredes, divisiones y puertas. 7) Lave o limpie el piso al final, desde el fondo hacia la puerta. 8) Reponga papel, jabón y alcohol en gel. Use paños de color exclusivo para el baño.",
"Revise y reponga papel higiénico, toallas de papel, jabón líquido y alcohol en gel. Cambie el repuesto de jabón, nunca lo complete. Falta de material: abra un llamado operativo.",
"Use limpiavidrios y secador (rodo), de arriba hacia abajo, retirando el exceso con paño de microfibra. A más de 2 metros de altura, solo con capacitación NR-35 y equipo adecuado.",
"Botoneras con paño levemente humedecido en alcohol al 70% (nunca rocíe directamente). Espejo y paredes con limpiavidrios o paño húmedo. Piso con mopa húmeda y limpieza del riel de la puerta.",
"Barra o use la fregadora. Manchas de aceite: absorbente, luego desengrasante y enjuague. Vacíe los basureros y mantenga la señalización de piso mojado.",
"Contenga el derrame con material absorbente, recójalo y deséchelo como residuo contaminado. Aplique desengrasante, friegue, enjuague y seque. Señalice el área hasta que se seque.",
"En áreas de salud no barra en seco: use mopa húmeda. En las demás, barra o aspire antes. Lave con detergente y enjuague, o pase mopa húmeda, desde el fondo hacia la salida. Señalice el piso mojado y déjelo secar.",
"No mueva documentos sin autorización. Superficie de la mesa con paño húmedo y detergente neutro. Teclado, mouse y teléfono con paño levemente humedecido en alcohol al 70% (nunca rocíe). Limpie las sillas y ordene el ambiente.",
"Borre la pizarra, límpiela con paño húmedo o producto adecuado para el tipo de pizarra y limpie la bandeja del borrador.",
"Recoja la basura, limpie pupitres y mesas, borre y limpie las pizarras, limpie manijas e interruptores y, por último, el piso, desde el fondo hacia la puerta.",
"Mesadas y lavabos con detergente y después desinfectante. No toque reactivos, muestras ni equipos encendidos. Residuos químicos y biológicos solo según la orientación del responsable del laboratorio.",
"Limpie mesadas, fregadero, mesas, microondas y cafetera. Heladera por fuera todos los días y por dentro según el cronograma. Use un paño exclusivo de la cocina. Deseche los residuos orgánicos y lave el piso al final.",
"Paño de microfibra y limpiavidrios sin amoníaco, sin rociar sobre las piezas. Nunca abra vitrinas o exhibidores sin la presencia de un responsable de la tienda.",
"Solo acompañado por un responsable del cliente. No mueva valores, documentos ni equipos. Registre el horario y el nombre de quien lo acompañó.",
"Paño levemente humedecido en alcohol al 70% en la pantalla, el teclado y el lector de tarjetas; nunca rocíe líquido. No toque los compartimientos de billetes y sobres.",
"Vidrios con limpiavidrios, de arriba hacia abajo; estructura y tiradores con paño húmedo y alcohol al 70%. No fuerce el mecanismo.",
"Paño húmedo con detergente neutro en el mostrador y las superficies de atención; alcohol al 70% en bolígrafos, interfonos y puntos de contacto. No mueva documentos.",
"Asas y empuñaduras con alcohol al 70% o desinfectante; estructura con paño húmedo y detergente.",
"Recoja perchas y prendas olvidadas, limpie espejo, banco y manija, y pase la mopa por el piso.",
"Cinta, mostrador y lector/máquina de tarjetas con paño levemente humedecido en alcohol al 70%, sin rociar sobre los equipos.",
"Riego, poda, retiro de hojas y abonado según el cronograma. Recoja los residuos verdes. Use guantes, gafas y polainas (protector auditivo con la desbrozadora) y aísle el área durante el uso de la desbrozadora.",
"Inspeccione y registre filtraciones, iluminación, puertas y cerraduras. Resuelva lo que esté dentro del alcance; para lo demás, abra un llamado operativo con foto.",
"Recorra la ruta definida en los horarios previstos, verifique puertas, ventanas, iluminación y personas extrañas. Registre el horario y los puntos verificados. Comunique las novedades de inmediato al supervisor.",
"Identifique y registre a todo visitante (nombre, documento, destino y horario), confirme la autorización con el sector, entregue la credencial y registre la salida.",
"Registre en el libro o en el sistema: fecha, hora, lugar, involucrados, descripción objetiva del hecho y medidas tomadas. Comunique al supervisor.",
"Siga el checklist del puesto: puertas, ventanas, alarma, luces y llaves. Registre el horario y cualquier anomalía.",
"Siga las cámaras, confirme que estén grabando, registre fallas y active la ronda ante una situación sospechosa.",
"Paño humedecido en detergente neutro, de arriba hacia abajo y de lo más limpio a lo más sucio; seque. En las superficies de contacto, aplique después el desinfectante estandarizado.",
"Barra, recoja hojas y basura, vacíe los basureros externos y lave cuando sea necesario. Señalice el piso mojado.",
"1) Use los EPP del área. 2) Recoja la basura. 3) Limpie de arriba hacia abajo y de lo más limpio a lo más sucio: superficies, mobiliario y puntos de contacto con paño húmedo y detergente; después desinfecte los puntos de contacto. 4) Limpie el piso al final, desde el fondo hacia la salida. 5) Señalice el piso mojado y reponga insumos."
],
fr:[
"Local inoccupé (programmé ou après sortie/transfert). 1) Portez tous les EPI de la zone. 2) Retirez les déchets et le linge. 3) Nettoyez de haut en bas : plafond, luminaires, bouches d'aération, murs, fenêtres et portes. 4) Nettoyez et désinfectez le mobilier et les équipements, y compris le dessous. 5) Lavez le sol avec deux seaux (solution et rinçage) ou à la machine, du fond vers la sortie. 6) Appliquez le désinfectant standard sur les surfaces de contact. 7) Réapprovisionnez les consommables et enregistrez.",
"Local occupé, tous les jours. 1) Portez les EPI de la zone. 2) Ramassez les déchets et changez les sacs. 3) Nettoyez les surfaces de contact (poignées, interrupteurs, plans de travail, tables, équipements) avec un chiffon humide et du détergent ; appliquez ensuite le désinfectant standard. Travaillez du plus propre au plus sale et de haut en bas. 4) Nettoyez le sol avec un balai humide/chiffon humide, du fond vers la sortie. 5) Réapprovisionnez les consommables. 6) Signalez le sol mouillé.",
"Toujours après le nettoyage au détergent. Appliquez le désinfectant standard (ex. : alcool à 70 % par friction ou ammonium quaternaire, selon la fiche technique) sur les poignées, rampes, interrupteurs, barrières de lit, plans de travail et boutons. Respectez le temps de contact du produit. Utilisez un chiffon dédié et changez-le en changeant de local.",
"Gants en caoutchouc à manchette longue, tablier imperméable et bottes. Fermez le sac blanc laiteux rempli au maximum aux 2/3, sans le comprimer ni chasser l'air avec les mains ; ne transvasez jamais les déchets d'un sac à l'autre. Boîte à objets piquants/tranchants seulement jusqu'à la ligne pointillée. Apportez au local à déchets dans le chariot de collecte fermé, par le circuit défini. Désinfectez le chariot et les gants à la fin.",
"Séparez papier, plastique, métal et verre dans les bacs identifiés ; ne mélangez pas les déchets organiques. Fermez les sacs remplis au maximum aux 2/3 et apportez-les au point de tri sélectif.",
"Fermez le sac rempli au maximum aux 2/3, sans le comprimer. Changez le sac, nettoyez la poubelle à l'extérieur (et à l'intérieur lors du bionettoyage complet) et refermez-la. Séparez les recyclables selon le tri sélectif. Apportez au local à déchets par le trajet défini.",
"1) Signalez la zone. 2) Ramassez les déchets. 3) Appliquez le produit dans la cuvette et l'urinoir et laissez agir. 4) Nettoyez lavabo, robinet, miroir et distributeurs. 5) Brossez l'intérieur de la cuvette, nettoyez l'extérieur, l'abattant et la chasse d'eau. 6) Nettoyez murs, cloisons et portes. 7) Lavez ou nettoyez le sol en dernier, du fond vers la porte. 8) Réapprovisionnez papier, savon et gel hydroalcoolique. Utilisez des chiffons d'une couleur réservée aux sanitaires.",
"Vérifiez et réapprovisionnez papier toilette, essuie-mains, savon liquide et gel hydroalcoolique. Remplacez la recharge de savon, ne la complétez jamais. Manque de produits : ouvrez un ticket opérationnel.",
"Utilisez du nettoyant vitres et une raclette, de haut en bas, en retirant l'excédent avec un chiffon microfibre. Au-delà de 2 mètres de hauteur, uniquement avec la formation NR-35 et un équipement adapté.",
"Boutons avec un chiffon légèrement imbibé d'alcool à 70 % (ne jamais pulvériser directement). Miroir et parois avec du nettoyant vitres ou un chiffon humide. Sol au balai humide et nettoyage du rail de la porte.",
"Balayez ou utilisez l'autolaveuse. Taches d'huile : absorbant, puis dégraissant et rinçage. Videz les poubelles et maintenez la signalisation de sol mouillé.",
"Contenez le déversement avec un absorbant, ramassez-le et éliminez-le comme déchet contaminé. Appliquez du dégraissant, frottez, rincez et séchez. Signalez la zone jusqu'à séchage.",
"Dans les zones de soins, ne balayez pas à sec : utilisez un balai humide. Ailleurs, balayez ou aspirez d'abord. Lavez au détergent et rincez, ou passez le balai humide, du fond vers la sortie. Signalez le sol mouillé et laissez sécher.",
"Ne déplacez pas de documents sans autorisation. Plateau avec un chiffon humide et du détergent neutre. Clavier, souris et téléphone avec un chiffon légèrement imbibé d'alcool à 70 % (ne jamais pulvériser). Nettoyez les chaises et rangez la pièce.",
"Effacez le tableau, nettoyez-le avec un chiffon humide ou un produit adapté au type de tableau et nettoyez le support de la brosse.",
"Ramassez les déchets, nettoyez pupitres et tables, effacez et nettoyez les tableaux, nettoyez poignées et interrupteurs, et enfin le sol, du fond vers la porte.",
"Paillasses et éviers au détergent puis au désinfectant. Ne touchez pas aux réactifs, échantillons ou équipements en marche. Déchets chimiques et biologiques uniquement selon les consignes du responsable du laboratoire.",
"Nettoyez plans de travail, évier, tables, micro-ondes et cafetière. Réfrigérateur à l'extérieur tous les jours et à l'intérieur selon le planning. Utilisez un chiffon réservé à la cuisine. Jetez les déchets organiques et lavez le sol en dernier.",
"Chiffon microfibre et nettoyant vitres sans ammoniaque, sans pulvériser sur les articles. N'ouvrez jamais les vitrines ou présentoirs sans la présence d'un responsable du magasin.",
"Uniquement accompagné d'un responsable du client. Ne déplacez ni valeurs, ni documents, ni équipements. Notez l'heure et le nom de la personne qui vous a accompagné.",
"Chiffon légèrement imbibé d'alcool à 70 % sur l'écran, le clavier et le lecteur de carte ; ne pulvérisez jamais de liquide. Ne touchez pas aux compartiments à billets et à enveloppes.",
"Vitres avec du nettoyant vitres, de haut en bas ; structure et poignées avec un chiffon humide et de l'alcool à 70 %. Ne forcez pas le mécanisme.",
"Chiffon humide avec du détergent neutre sur le comptoir et les surfaces d'accueil ; alcool à 70 % sur les stylos, interphones et points de contact. Ne déplacez pas de documents.",
"Poignées et prises avec de l'alcool à 70 % ou du désinfectant ; structure avec un chiffon humide et du détergent.",
"Ramassez les cintres et les articles oubliés, nettoyez le miroir, le banc et la poignée, et passez le balai sur le sol.",
"Tapis, comptoir et lecteur/terminal de carte avec un chiffon légèrement imbibé d'alcool à 70 %, sans pulvériser sur les équipements.",
"Arrosage, taille, ramassage des feuilles et fertilisation selon le planning. Ramassez les déchets verts. Portez gants, lunettes et jambières (protection auditive avec la débroussailleuse) et isolez la zone pendant l'utilisation de la débroussailleuse.",
"Inspectez et notez les fuites, l'éclairage, les portes et les serrures. Réglez ce qui relève de votre périmètre ; pour le reste, ouvrez un ticket opérationnel avec photo.",
"Parcourez l'itinéraire défini aux horaires prévus, vérifiez portes, fenêtres, éclairage et personnes étrangères. Notez l'heure et les points vérifiés. Signalez immédiatement tout incident au superviseur.",
"Identifiez et enregistrez chaque visiteur (nom, pièce d'identité, destination et heure), confirmez l'autorisation auprès du service, remettez le badge et enregistrez la sortie.",
"Consignez dans le registre ou le système : date, heure, lieu, personnes impliquées, description objective des faits et mesures prises. Informez le superviseur.",
"Suivez la check-list du poste : portes, fenêtres, alarme, lumières et clés. Notez l'heure et toute anomalie.",
"Surveillez les caméras, vérifiez qu'elles enregistrent, notez les pannes et déclenchez la ronde en cas de situation suspecte.",
"Chiffon imbibé de détergent neutre, de haut en bas et du plus propre au plus sale ; séchez. Sur les surfaces de contact, appliquez ensuite le désinfectant standard.",
"Balayez, ramassez feuilles et déchets, videz les poubelles extérieures et lavez si nécessaire. Signalez le sol mouillé.",
"1) Portez les EPI de la zone. 2) Ramassez les déchets. 3) Nettoyez de haut en bas et du plus propre au plus sale : surfaces, mobilier et points de contact avec un chiffon humide et du détergent ; désinfectez ensuite les points de contact. 4) Nettoyez le sol en dernier, du fond vers la sortie. 5) Signalez le sol mouillé et réapprovisionnez les consommables."
]};
var NOTA_RISCO_T={
 en:{C:"Critical area: cloth and bucket reserved for this area, two-bucket technique (solution and rinse) or mop, mandatory disinfection after cleaning and hand hygiene before and after.",S:"Semi-critical area: cleaning with detergent followed by disinfection of high-touch surfaces; separate cloths for each type of surface.",N:"Non-critical area: cleaning with neutral detergent; disinfection when there is organic matter."},
 es:{C:"Área crítica: paño y balde exclusivos de esta área, técnica de dos baldes (solución y enjuague) o mopa, desinfección obligatoria después de la limpieza e higiene de manos antes y después.",S:"Área semicrítica: limpieza con detergente seguida de desinfección de las superficies de contacto; paños separados por tipo de superficie.",N:"Área no crítica: limpieza con detergente neutro; desinfección cuando haya materia orgánica."},
 fr:{C:"Zone critique : chiffon et seau réservés à cette zone, technique des deux seaux (solution et rinçage) ou balai, désinfection obligatoire après le nettoyage et hygiène des mains avant et après.",S:"Zone semi-critique : nettoyage au détergent suivi de la désinfection des surfaces de contact ; chiffons séparés par type de surface.",N:"Zone non critique : nettoyage au détergent neutre ; désinfection en présence de matière organique."}
};
var ATIV_T={
"limpeza concorrente":["Daily (concurrent) cleaning","Limpieza concurrente","Nettoyage quotidien (concomitant)"],
"limpeza terminal":["Terminal cleaning","Limpieza terminal","Bionettoyage complet (terminal)"],
"desinfeccao de superficies de contato":["Disinfection of high-touch surfaces","Desinfección de superficies de contacto","Désinfection des surfaces de contact"],
"recolhimento de residuos infectantes":["Collection of infectious waste","Recolección de residuos infecciosos","Collecte des déchets infectieux"],
"recolhimento de residuos comuns":["Collection of general waste","Recolección de residuos comunes","Collecte des déchets ordinaires"],
"recolhimento de residuos":["Waste collection","Recolección de residuos","Collecte des déchets"],
"limpeza de sanitarios":["Restroom cleaning","Limpieza de sanitarios","Nettoyage des sanitaires"],
"reposicao de insumos":["Restocking supplies","Reposición de insumos","Réapprovisionnement des consommables"],
"limpeza de pisos":["Floor cleaning","Limpieza de pisos","Nettoyage des sols"],
"limpeza de mesas e estacoes":["Desk and workstation cleaning","Limpieza de mesas y estaciones","Nettoyage des bureaux et postes de travail"],
"higienizacao de banheiros":["Bathroom sanitization","Higienización de baños","Désinfection des salles de bain"],
"limpeza da copa":["Kitchenette cleaning","Limpieza de la cocina","Nettoyage de l'office"],
"recolhimento de lixo":["Trash collection","Recolección de basura","Ramassage des déchets"],
"limpeza de vidros":["Glass cleaning","Limpieza de vidrios","Nettoyage des vitres"],
"limpeza de salas de aula":["Classroom cleaning","Limpieza de aulas","Nettoyage des salles de classe"],
"limpeza de carteiras e quadros":["Desk and board cleaning","Limpieza de pupitres y pizarras","Nettoyage des pupitres et tableaux"],
"limpeza de laboratorios":["Laboratory cleaning","Limpieza de laboratorios","Nettoyage des laboratoires"],
"limpeza de patio":["Courtyard cleaning","Limpieza de patio","Nettoyage de la cour"],
"limpeza de pisos industriais":["Industrial floor cleaning","Limpieza de pisos industriales","Nettoyage des sols industriels"],
"remocao de oleo e residuos":["Oil and residue removal","Remoción de aceite y residuos","Élimination de l'huile et des résidus"],
"limpeza de vestiarios":["Locker room cleaning","Limpieza de vestuarios","Nettoyage des vestiaires"],
"higienizacao do refeitorio":["Cafeteria sanitization","Higienización del comedor","Désinfection du réfectoire"],
"coleta seletiva":["Recycling collection","Recolección selectiva","Tri sélectif"],
"limpeza de escritorios":["Office cleaning","Limpieza de oficinas","Nettoyage des bureaux"],
"controle de acesso":["Access control","Control de acceso","Contrôle d'accès"],
"registro de visitantes":["Visitor registration","Registro de visitantes","Enregistrement des visiteurs"],
"ronda perimetral":["Perimeter patrol","Ronda perimetral","Ronde périmétrique"],
"ronda interna":["Internal patrol","Ronda interna","Ronde intérieure"],
"registro de ocorrencias":["Incident reporting","Registro de novedades","Consignation des incidents"],
"abertura e fechamento":["Opening and closing","Apertura y cierre","Ouverture et fermeture"],
"monitoramento cftv":["CCTV monitoring","Monitoreo CCTV","Surveillance vidéo (CCTV)"],
"limpeza de areas comuns":["Common area cleaning","Limpieza de áreas comunes","Nettoyage des parties communes"],
"limpeza de elevadores":["Elevator cleaning","Limpieza de ascensores","Nettoyage des ascenseurs"],
"limpeza de garagem":["Garage cleaning","Limpieza del garaje","Nettoyage du parking"],
"jardinagem":["Gardening","Jardinería","Jardinage"],
"manutencao predial":["Building maintenance","Mantenimiento del edificio","Maintenance du bâtiment"],
"limpeza de vidros e fachada":["Glass and façade cleaning","Limpieza de vidrios y fachada","Nettoyage des vitres et de la façade"],
"limpeza do salao":["Sales floor cleaning","Limpieza del salón","Nettoyage de la surface de vente"],
"limpeza de vitrines":["Showcase cleaning","Limpieza de vitrinas","Nettoyage des vitrines"],
"limpeza de provadores":["Fitting room cleaning","Limpieza de probadores","Nettoyage des cabines d'essayage"],
"limpeza dos caixas":["Checkout cleaning","Limpieza de las cajas","Nettoyage des caisses"],
"higienizacao de carrinhos e cestas":["Cart and basket sanitization","Higienización de carros y cestas","Désinfection des chariots et paniers"],
"limpeza de banheiros de clientes":["Customer restroom cleaning","Limpieza de baños de clientes","Nettoyage des toilettes clients"],
"limpeza de vitrines e vidros":["Showcase and glass cleaning","Limpieza de vitrinas y vidrios","Nettoyage des vitrines et vitres"],
"limpeza de expositores":["Display case cleaning","Limpieza de exhibidores","Nettoyage des présentoirs"],
"limpeza de balcoes":["Counter cleaning","Limpieza de mostradores","Nettoyage des comptoirs"],
"limpeza da sala vip":["VIP room cleaning","Limpieza de la sala VIP","Nettoyage du salon VIP"],
"limpeza de area restrita (acompanhada)":["Restricted area cleaning (accompanied)","Limpieza de área restringida (acompañada)","Nettoyage de zone réservée (accompagné)"],
"higienizacao de terminais de autoatendimento":["Self-service terminal sanitization","Higienización de terminales de autoservicio","Désinfection des bornes libre-service"],
"limpeza de guiches":["Service window cleaning","Limpieza de ventanillas","Nettoyage des guichets"],
"limpeza de porta giratoria":["Revolving door cleaning","Limpieza de puerta giratoria","Nettoyage de la porte tambour"],
"limpeza da area de espera":["Waiting area cleaning","Limpieza del área de espera","Nettoyage de la salle d'attente"],
"limpeza de tesouraria (acompanhada)":["Treasury room cleaning (accompanied)","Limpieza de tesorería (acompañada)","Nettoyage de la trésorerie (accompagné)"],
"limpeza de banheiros":["Bathroom cleaning","Limpieza de baños","Nettoyage des salles de bain"],
"limpeza de superficies":["Surface cleaning","Limpieza de superficies","Nettoyage des surfaces"],
"limpeza de superficies de contato":["High-touch surface cleaning","Limpieza de superficies de contacto","Nettoyage des surfaces de contact"]
};
var FREQ_T={"Diária":["Daily","Diaria","Quotidienne"],"Semanal":["Weekly","Semanal","Hebdomadaire"],"Quinzenal":["Every 15 days","Quincenal","Bimensuelle"],"Mensal":["Monthly","Mensual","Mensuelle"],"Eventual":["As needed","Eventual","Ponctuelle"]};
var CF_UI={en:["How to do it","activity","activities"],es:["Cómo hacerlo","actividad","actividades"],fr:["Comment faire","activité","activités"]};
var LIX={en:0,es:1,fr:2};
function cfIdx(nome){var n=normT(nome);for(var i=0;i<COMO_FAZER.length;i++)if(COMO_FAZER[i][0].test(n))return i;return -1}
function nomeAtivT(g,l){if(!l||l=="pt-BR"||LIX[l]==null)return g.nome;if(g.nomeT&&g.nomeT[l])return g.nomeT[l];var t=ATIV_T[normT(g.nome)];return t?t[LIX[l]]:g.nome}
function ativNomeTemT(g,l){return !!((g.nomeT&&g.nomeT[l])||ATIV_T[normT(g.nome)])}
/* retorna {txt, oficial:bool} */
function obsAtivT(g,l){
  if(!l||l=="pt-BR"||LIX[l]==null)return {txt:g.obs||"",oficial:true};
  if(g.obsT&&g.obsT[l])return {txt:g.obsT[l],oficial:true};
  if(g.obsAuto&&g.obs){var i=cfIdx(g.nome);if(i>=0&&CF_T[l]&&CF_T[l][i])return {txt:CF_T[l][i],oficial:true}}
  return {txt:g.obs||"",oficial:!g.obs};
}
function freqT(f,l){var t=FREQ_T[f||"Diária"];return t&&LIX[l]!=null?t[LIX[l]]:(f||"Diária")}
function setPgTrad(id,campo,l,v){var g=PG.filter(function(x){return x.id==id})[0];if(!g||!arChk(g.l))return;g[campo]=g[campo]||{};v=String(v||"").trim();if(v)g[campo][l]=v;else delete g[campo][l];svPg()}
function cfTradEditor(g){
  var langs=[["en","🇺🇸"],["es","🇪🇸"],["fr","🇫🇷"]];
  var st=langs.map(function(x){var o=obsAtivT(g,x[0]);return x[1]+(o.oficial&&(o.txt||!g.obs)&&(ativNomeTemT(g,x[0])||!g.nome)?'✔':'✖')}).join(" ");
  return '<details style="margin-top:4px"><summary class="s" style="cursor:pointer">🌐 Traduções '+st+'</summary>'
    +langs.map(function(x){var l=x[0],o=obsAtivT(g,l),padrao=!(g.obsT&&g.obsT[l])&&g.obsAuto&&o.oficial&&o.txt;
      return '<div style="border-left:3px solid var(--line);padding-left:8px;margin:6px 0"><div class="s"><b>'+x[1]+' '+l.toUpperCase()+'</b>'+(padrao?' · tradução padrão incluída':'')+'</div>'
        +'<input placeholder="Nome da atividade ('+l+')" value="'+esc((g.nomeT||{})[l]||"")+'" onchange="setPgTrad('+g.id+',\'nomeT\',\''+l+'\',this.value)" style="font-size:14px">'
        +'<textarea rows="2" placeholder="'+esc(padrao?o.txt.slice(0,90)+'…':'Como fazer ('+l+')')+'" onchange="setPgTrad('+g.id+',\'obsT\',\''+l+'\',this.value)" style="width:100%;padding:8px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font:inherit;font-size:13px">'+esc((g.obsT||{})[l]||"")+'</textarea>'
        +(g.obs?'<a class="s" href="'+gTradUrl(g.nome+"\n\n"+g.obs,l)+'" target="_blank" rel="noopener" style="color:var(--brand)">🌐 Traduzir com o Google</a>':'')+'</div>'}).join("")
    +'<div class="s">Atividades com “como fazer” padrão já têm tradução revisada. Se você editar o texto em português, cadastre a tradução aqui.</div></details>';
}

/* ===================================================================
   NAVEGAÇÃO
   =================================================================== */
var NAV_KEYS=["perfil","ag","c","g","hOP","hAD","CLI","imp","eq","ct","ar","pti","plano","p","t","l","k","i","q","es","rq","st","r","ac","fc","dc","cont"];
var NAV={};NAV_KEYS.forEach(function(k){NAV[k]=k});
function nomeNav(k){
  var m={perfil:"menu_perfil",c:"menu_rotina",g:"menu_giro",hOP:"menu_hOP",hAD:"menu_hAD",CLI:"menu_cli",imp:"menu_imp",eq:"menu_eq",ct:"menu_ct",ar:"menu_ar",ag:"menu_ag",pti:"menu_pti",plano:"menu_plano",p:"menu_painel",t:"menu_trein",l:"menu_aval",k:"menu_cfg",i:"menu_ind",q:"menu_qr",es:"menu_estoque",rq:"menu_req",st:"menu_set",r:"menu_rel",cont:"menu_cont"};
  var nm=k=="ac"?"Acessos e clientes":k=="fc"?"Fábrica de Contratos":k=="dc"?"Documentos e vídeos":m[k]?T(m[k]):k;
  if(k=="ag"&&session){var n=agPendentes().length;if(n)nm+=" ("+n+")"}
  return nm;
}
function buildNav(){
  var temHosp=meusContratos().some(function(c){return ctTemTipo(c,"hosp")})||!meusContratos().length;
  var meus=["perfil"].concat(acessosDoUsuario().filter(function(k){return k!=="perfil"&&k!=="cont"&&k!=="dc"&&NAV[k]&&(k!="g"||temHosp)}));
  meus.push("dc");meus.push("cont");
  var btn=function(k){return '<button class="'+(k==tab?'on':'')+'" onclick="go(\''+k+'\')"><span class="ico">'+NN_ICONE(k)+'</span>'+esc(nomeNav(k))+'</button>'};
  var usados={},html="";
  NAV_GRUPOS.forEach(function(g){var ks=g[1].filter(function(k){return meus.indexOf(k)>=0});ks.forEach(function(k){usados[k]=1});
    if(ks.length)html+=(g[0]?'<div class="navg">'+esc(g[0])+'</div>':'')+ks.map(btn).join("")});
  var resto=meus.filter(function(k){return !usados[k]});if(resto.length)html+='<div class="navg">Outros</div>'+resto.map(btn).join("");
  document.getElementById("sbNav").innerHTML=html;
  updateUserInfo();
}
function go(t){
  stopCam();stopNFC();
  var u=session?userById(session.userId):null;
  if(u&&u.pinTrocar&&t!=="perfil"){alert("Defina seu novo PIN primeiro.");t="perfil"}
  else if(u&&!u.perfilCompleto&&t!=="perfil"){alert("Complete seu perfil primeiro.");t="perfil"}
  var permitidos=acessosDoUsuario();
  var especiais=["perfil","cont","dc"];
  if(especiais.indexOf(t)<0 && permitidos.indexOf(t)<0){
    alert("Sem acesso a esta área.");return;
  }
  tab=t;if(t!="dc"){DC.sel=null;DC.ed=null}closeMenu();buildNav();draw();window.scrollTo(0,0);
}
function opts(sel){return areasPermitidas().map(function(i){return '<option value="'+i+'"'+(i==sel?' selected':'')+'>'+esc(LOC[i].n)+'</option>'}).join("")}
function setLoc(v){loc=+v;if(!LOC[loc])loc=-1;sv("rad_loc",loc);draw()}

function draw(){
  if(!session)return;
  sfCtx=null;
  var u=userById(session.userId);
  if(!u){logout(true);return}

  if((!u.perfilCompleto||u.pinTrocar)&&tab!=="perfil"){tab="perfil"}
  ajustaLoc();

  var banner="";
  if(perfilReal){
    banner+='<div class="card noprint" style="border-left:4px solid var(--run);cursor:pointer" onclick="simularPerfil(\'\')">'
      +'<b style="color:#8a6400">👁️ Vendo como '+esc(perfil)+'</b>'
      +'<div class="s">Toque para voltar a '+esc(perfilReal)+'.</div></div>';
  }
  if(u.perfilCompleto){
    if(tab!="ag")banner+=agBanner();
    banner+=alertaVencidos();
    banner+=alertaPtisPendentes();
    if(podeEstoque()&&acessosDoUsuario().indexOf("rq")>=0){
      var pendReq=reqPend();
      if(pendReq.length){
        banner+='<div class="card noprint" style="border-left:4px solid var(--run);background:rgba(224,161,0,.06);cursor:pointer" onclick="go(\'rq\')">'
          +'<div style="display:flex;align-items:center;gap:10px">'
            +'<span style="font-size:24px">🛒</span>'
            +'<span style="flex:1"><b style="color:#8a6400">'+pendReq.length+' requisição(ões) pendente(s)</b></span>'
          +'</div></div>';
      }
    }
    if(podeEstMov()&&acessosDoUsuario().indexOf("es")>=0){
      var mids=meusCtIds();var crit=EST.filter(function(it){return estCritico(it)&&(verTudo()||mids.indexOf(it.ct)>=0)});
      if(crit.length){
        banner+='<div class="card noprint" style="border-left:4px solid var(--warn);background:rgba(198,40,40,.06);cursor:pointer" onclick="go(\'es\')">'
          +'<div style="display:flex;align-items:center;gap:10px">'
            +'<span style="font-size:24px">📦</span>'
            +'<span style="flex:1"><b style="color:var(--warn)">'+crit.length+' item(ns) abaixo do mínimo</b></span>'
          +'</div></div>';
      }
    }
  }

  var v="";
  try{
    if(tab=="perfil")v=vMeuPerfil();
    else if(tab=="c")v=(ehColabCargo()&&!open?bannerInterno():"")+vC();
    else if(tab=="g")v=vG();
    else if(tab=="hOP")v=vOP();
    else if(tab=="hAD")v=vAD();
    else if(tab=="CLI")v=vCLI();
    else if(tab=="imp")v=vImp();
    else if(tab=="eq")v=vEquipe();
    else if(tab=="ct")v=vContratos();
    else if(tab=="pti")v=vPTI();
    else if(tab=="plano")v=vPlanoGeral();
    else if(tab=="l")v=vL();
    else if(tab=="r")v=vR();
    else if(tab=="k")v=vK();
    else if(tab=="t")v=vT();
    else if(tab=="i")v=vI();
    else if(tab=="p")v=vP();
    else if(tab=="q")v=vQ();
    else if(tab=="st")v=vSetores();
    else if(tab=="es")v=vEstoque();
    else if(tab=="rq")v=vReq();
    else if(tab=="cont")v=vContatos();
    else if(tab=="ar")v=vAreas();
    else if(tab=="ag")v=vAgenda();
    else if(tab=="ac")v=usersCard();
    else if(tab=="fc")v=vFabrica();
    else if(tab=="dc")v=vDocs();
    else v=vMeuPerfil();
  }catch(err){
    v='<div class="card" style="border-left:4px solid var(--warn)"><b style="color:var(--warn)">Erro ao carregar</b><div class="s">'+esc(err.message||String(err))+'</div></div>';
    console.error(err);
  }

  document.getElementById("app").innerHTML=banner+v;
  tickSf();hideCfg();
  if(tab=="q")setTimeout(qrs,50);
  if(tab=="eq"&&tsel)setTimeout(geraQRColab,100);
  if(tab=="pti")setTimeout(function(){
    if(document.getElementById("novo_pti_colab"))novoPtiContratos();
  },50);
}

/* ===================================================================
   BOOT
   =================================================================== */
/* Abrir o app pelo link de um QR (…?l=0&e=3) inicia/finaliza a atividade após o login */
var PEND_QR=parseQR(location.search)?location.search:null;
function startApp(){
  document.removeEventListener("keydown",pinKeyHandler);
  login.step="";
  document.getElementById("login").classList.add("hide");
  document.getElementById("shell").classList.remove("hide");
  var u=session?userById(session.userId):null;
  if(u&&(!u.perfilCompleto||u.pinTrocar)){tab="perfil"}
  else{
    var acessos=acessosDoUsuario().filter(function(k){return k!="perfil"});
    tab=acessos.length?acessos[0]:"perfil";
  }
  buildNav();
  try{aplicarTema()}catch(e){}
  draw();
  setTimeout(agVerificar,800);
  if(PEND_QR&&u&&u.perfilCompleto&&!u.pinTrocar){
    var q=PEND_QR;PEND_QR=null;
    try{history.replaceState(null,"",location.pathname)}catch(e){}
    if(acessosDoUsuario().indexOf("c")>=0){tab="c";buildNav();lido(q)}
  }
}

(function init(){
  document.documentElement.lang=IDIOMA_UI;
  renderLangLogin();
  document.getElementById("loginTitle").textContent=T("app_title");
  document.getElementById("loginSub").textContent=T("app_sub");
  var bp=document.getElementById("btnAlterPin");if(bp)bp.textContent=T("sb_changePin");
  var bl=document.getElementById("btnLogout");if(bl)bl.textContent=T("sb_logout");
  try{aplicarTema()}catch(e){console.error(e)}
  /* Mantém contratos alinhados aos departamentos do efetivo (inclui bases importadas antes desta versão) */
  try{
    ensureModelChecklists();
    /* migração: áreas que estavam marcadas em contrato.locais passam a pertencer ao contrato */
    var migr=false;
    CONTRATOS.forEach(function(c){if(!c.tipo){c.tipo="out";migr=true}if(!c.tipos||!c.tipos.length){c.tipos=[c.tipo];migr=true}(c.locais||[]).forEach(function(i){if(LOC[i]&&LOC[i].ct==null){LOC[i].ct=c.id;migr=true}});if(c.locais){delete c.locais;migr=true}});
    if(migr){svCt();svLoc()}
    if(COL.some(function(c){return chaveDepto(c)})){syncContratosDepto();syncSupervisores()}
    prepararAreas();
    autoClassificar();
    if(aplicarComoFazer())sv("rad_pg",PG);
  }catch(e){console.error(e)}
  if(ptIniciar()){}
  else if(session&&session.userId&&userById(session.userId)){
    aplicaSession(session);
    startApp();
  }else{
    session=null;try{localStorage.removeItem("rad_session")}catch(e){}
    document.getElementById("login").classList.remove("hide");
    document.getElementById("shell").classList.add("hide");
    renderLogin();
  }
})();
</script>
</body>
</html>
