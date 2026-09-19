/* MediLink — shared core used by both the Staff (index.html) and Administrator
   (admin.html) pages: the encrypted-directory reader, the searchable list, the
   Add-to-Home-Screen button, and small helpers for the approval workflow.
   The encrypted directory itself lives in medilink-data.js as window.ENC. */
(function(){
  "use strict";
  var MDL = {};
  var ROWS = [], TYPE = 'all';

  function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c];});}
  MDL.esc = esc;

  // A small, correct CSV reader: honours "quoted, fields" that contain commas,
  // doubled "" quotes and stray newlines, so a value like "6701,6716" stays one field.
  function parseCSV(text){
    var rows=[], row=[], field='', inQ=false, i=0, n=text.length;
    function endField(){row.push(field);field='';}
    function endRow(){endField();rows.push(row);row=[];}
    while(i<n){
      var c=text[i];
      if(inQ){
        if(c==='"'){ if(text[i+1]==='"'){field+='"';i+=2;continue;} inQ=false;i++;continue; }
        field+=c;i++;continue;
      }
      if(c==='"'){inQ=true;i++;continue;}
      if(c===','){endField();i++;continue;}
      if(c==='\r'){i++;continue;}
      if(c==='\n'){endRow();i++;continue;}
      field+=c;i++;
    }
    if(field.length||row.length)endRow();
    return rows;
  }
  MDL.parseCSV = parseCSV;

  function classForType(t){
    var s=String(t||'').toLowerCase();
    if(s.indexOf('pager')>-1)return 'pager';
    if(s.indexOf('phone')>-1)return 'phone';
    return 'ext';
  }
  function tagLabel(t){
    var s=String(t||'').toLowerCase();
    if(s.indexOf('pager')>-1)return 'Pager';
    if(s.indexOf('phone')>-1)return 'Phone';
    return 'Ext';
  }
  // Numbers can be a single value, a comma-list ("6701,6716") or shorthand
  // ("4030/1/9"). Comma-lists become separate call chips; anything else stays a
  // single chip that dials its first real number.
  function numbersHTML(value){
    var v=String(value||'').trim();
    if(!v)return '<span class="num muted">—</span>';
    var parts = v.indexOf(',')>-1 ? v.split(',') : [v];
    return parts.map(function(p){
      p=p.trim(); if(!p)return '';
      var m=p.match(/\d{3,}/);
      return m ? '<a class="num" href="tel:'+esc(m[0])+'">'+esc(p)+'</a>'
               : '<span class="num muted">'+esc(p)+'</span>';
    }).join('');
  }

  function render(){
    var qEl=document.getElementById('q');
    var q=((qEl&&qEl.value)||'').trim().toLowerCase();
    var terms=q.split(/\s+/).filter(Boolean);
    var out=ROWS.filter(function(r){
      if(TYPE!=='all' && r.type!==TYPE)return false;
      if(terms.length){
        var hay=(r.name+' '+r.value).toLowerCase();
        for(var k=0;k<terms.length;k++){ if(hay.indexOf(terms[k])<0)return false; }
      }
      return true;
    });
    var countEl=document.getElementById('count');
    if(countEl)countEl.textContent=
      out.length+' of '+ROWS.length+' entr'+(ROWS.length===1?'y':'ies');
    var list=document.getElementById('list');
    if(!list)return;
    if(!out.length){ list.innerHTML='<div class="note">No entry matches that search.</div>'; return; }
    var html=new Array(out.length);
    for(var j=0;j<out.length;j++){
      var r=out[j], cls=classForType(r.type);
      html[j]='<div class="row"><div class="body"><div class="name">'+esc(r.name)+
        '</div><div class="nums">'+numbersHTML(r.value)+'</div></div>'+
        '<span class="tag '+cls+'">'+tagLabel(r.type)+'</span></div>';
    }
    list.innerHTML=html.join('');
  }
  MDL.render = render;

  function setType(t){
    TYPE=t;
    var chips=document.querySelectorAll('#filters .chip');
    for(var i=0;i<chips.length;i++)chips[i].classList.toggle('on',chips[i].getAttribute('data-type')===t);
    render();
  }
  MDL.setType = setType;

  function boot(rows){
    var seen={}, clean=[];
    for(var i=1;i<rows.length;i++){ // row 0 is the header
      var r=rows[i]; if(!r)continue;
      var type=(r[0]||'').trim(), name=(r[1]||'').trim(), value=(r[2]||'').trim();
      if(!name && !value)continue;
      var key=type+'|'+name+'|'+value;
      if(seen[key])continue; seen[key]=1;
      clean.push({type:type,name:name||'(unnamed)',value:value});
    }
    clean.sort(function(a,b){ return a.name.toLowerCase()<b.name.toLowerCase()?-1:a.name.toLowerCase()>b.name.toLowerCase()?1:0; });
    ROWS=clean;
    render();
  }
  MDL.boot = boot;

  // ── Decryption ──────────────────────────────────────────────────────────────
  function b64buf(b){var s=atob(b),n=s.length,u=new Uint8Array(n);for(var i=0;i<n;i++)u[i]=s.charCodeAt(i);return u.buffer;}
  async function decryptDirectory(code){
    var ENC=window.ENC;
    var subtle=window.crypto&&window.crypto.subtle;
    if(!subtle)throw new Error('unsupported');
    if(!ENC)throw new Error('no data');
    var km=await subtle.importKey('raw',new TextEncoder().encode(code),'PBKDF2',false,['deriveKey']);
    var key=await subtle.deriveKey({name:'PBKDF2',salt:b64buf(ENC.salt),iterations:200000,hash:'SHA-256'},
      km,{name:'AES-GCM',length:256},false,['decrypt']);
    var plain=await subtle.decrypt({name:'AES-GCM',iv:b64buf(ENC.iv)},key,b64buf(ENC.data));
    return new TextDecoder().decode(plain);
  }
  MDL.decryptDirectory = decryptDirectory;

  // Decrypt with `code`, and on success load the directory. Returns true/false.
  MDL.unlock = async function(code){
    try{
      var text=await decryptDirectory(code);
      boot(parseCSV(text));
      return true;
    }catch(e){ return false; }
  };

  // ── Add-to-Home-Screen ───────────────────────────────────────────────────────
  // Android/desktop Chrome fire beforeinstallprompt, so the button installs in
  // one tap. iOS never fires it, so the button shows the Share → Add steps. When
  // the page is already running as an installed app, the button stays hidden.
  var deferredPrompt=null;
  function isStandalone(){
    try{ return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone===true; }
    catch(e){ return false; }
  }
  MDL.isStandalone = isStandalone;
  function isIOS(){ return /iphone|ipad|ipod/i.test(navigator.userAgent) ||
    (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1); }
  MDL.isIOS = isIOS;
  function platform(){
    if(isIOS())return 'ios';
    if(/android/i.test(navigator.userAgent))return 'android';
    return 'other';
  }
  MDL.platform = platform;

  function openSheet(){ var s=document.getElementById('sheet'); if(s)s.classList.add('show'); }
  function closeSheet(){ var s=document.getElementById('sheet'); if(s)s.classList.remove('show'); }
  MDL.openSheet = openSheet;
  window.closeSheet = closeSheet;

  function wireInstall(){
    var installBtn=document.getElementById('install-btn');
    function showInstall(){ if(installBtn && !isStandalone()) installBtn.hidden=false; }
    window.addEventListener('beforeinstallprompt',function(e){ e.preventDefault(); deferredPrompt=e; showInstall(); });
    window.addEventListener('appinstalled',function(){ if(installBtn)installBtn.hidden=true; deferredPrompt=null;
      window.dispatchEvent(new Event('mdl-installed')); });
    if(installBtn){
      installBtn.addEventListener('click',function(){
        if(deferredPrompt){ var ev=deferredPrompt; deferredPrompt=null; ev.prompt(); if(ev.userChoice&&ev.userChoice.catch)ev.userChoice.catch(function(){}); return; }
        openSheet();
      });
    }
    var sheet=document.getElementById('sheet');
    if(sheet)sheet.addEventListener('click',function(e){ if(e.target===this)closeSheet(); });
    document.body.setAttribute('data-plat', platform());
    if(!isStandalone()) showInstall();
  }
  // Trigger the real Chrome/Android install prompt from anywhere (e.g. the gate).
  MDL.promptInstall = function(){
    if(deferredPrompt){ var ev=deferredPrompt; deferredPrompt=null; ev.prompt(); if(ev.userChoice&&ev.userChoice.catch)ev.userChoice.catch(function(){}); return true; }
    openSheet(); return false;
  };
  MDL.hasInstallPrompt = function(){ return !!deferredPrompt; };

  // No service worker: the page always loads straight from the network. Any copy
  // registered by an earlier version is removed here, and its caches cleared, so
  // a stale worker can never keep an old or broken page on screen.
  function cleanupServiceWorker(){
    if('serviceWorker' in navigator){
      try{
        navigator.serviceWorker.getRegistrations().then(function(regs){
          regs.forEach(function(r){ r.unregister(); });
        }).catch(function(){});
      }catch(e){}
      try{
        if(window.caches && caches.keys){
          caches.keys().then(function(keys){ keys.forEach(function(k){ if(/medilink/.test(k)) caches.delete(k); }); }).catch(function(){});
        }
      }catch(e){}
    }
  }

  // Wire the shared directory UI (search box, filter chips, install button, SW
  // cleanup). Safe to call once the DOM is ready; missing elements are ignored.
  MDL.wireDirectoryUI = function(){
    var q=document.getElementById('q');
    if(q)q.addEventListener('input',render);
    var filters=document.getElementById('filters');
    if(filters)filters.addEventListener('click',function(e){
      var chip=e.target.closest('.chip'); if(chip)setType(chip.getAttribute('data-type'));
    });
    wireInstall();
    cleanupServiceWorker();
  };

  // ── Approval workflow helpers ─────────────────────────────────────────────────
  // Requests and approvals are kept in localStorage, which is shared between
  // index.html and admin.html on the same site — so on one device a staff
  // request shows up in the admin tab automatically. Across two devices the
  // staff request travels as a short code (packReq/unpackReq) the admin pastes.
  var LS = {
    pass:'mdh-pass',        // decryption code, remembered on this device
    requests:'mdh-requests',// admin-side queue of staff requests
    me:'mdh-me',            // this device's own staff request
    device:'mdh-device'     // a stable id for this device
  };
  MDL.LS = LS;

  MDL.getPass = function(){ try{ return localStorage.getItem(LS.pass)||''; }catch(e){ return ''; } };
  MDL.setPass = function(c){ try{ localStorage.setItem(LS.pass,c); }catch(e){} };
  MDL.clearPass = function(){ try{ localStorage.removeItem(LS.pass); }catch(e){} };

  MDL.deviceId = function(){
    var id='';
    try{ id=localStorage.getItem(LS.device)||''; }catch(e){}
    if(!id){ id='d'+Date.now().toString(36)+Math.random().toString(36).slice(2,8);
      try{ localStorage.setItem(LS.device,id); }catch(e){} }
    return id;
  };

  MDL.loadRequests = function(){
    try{ var a=JSON.parse(localStorage.getItem(LS.requests)||'[]'); return Array.isArray(a)?a:[]; }
    catch(e){ return []; }
  };
  MDL.saveRequests = function(a){ try{ localStorage.setItem(LS.requests,JSON.stringify(a)); }catch(e){} };

  MDL.getMe = function(){
    try{ var m=JSON.parse(localStorage.getItem(LS.me)||'null'); return m&&typeof m==='object'?m:null; }
    catch(e){ return null; }
  };
  MDL.setMe = function(m){ try{ localStorage.setItem(LS.me,JSON.stringify(m)); }catch(e){} };
  MDL.clearMe = function(){ try{ localStorage.removeItem(LS.me); }catch(e){} };

  // Add or merge a request into the admin queue (dedup by id).
  MDL.upsertRequest = function(req){
    var all=MDL.loadRequests(), i;
    for(i=0;i<all.length;i++){ if(all[i].id===req.id){
      // keep an existing decision; otherwise refresh the details
      if(all[i].status==='pending'){ all[i]=req; } else { all[i].name=req.name; all[i].surname=req.surname; all[i].ward=req.ward; }
      MDL.saveRequests(all); return all[i];
    } }
    all.push(req); MDL.saveRequests(all); return req;
  };
  MDL.findRequest = function(id){
    var all=MDL.loadRequests();
    for(var i=0;i<all.length;i++){ if(all[i].id===id)return all[i]; }
    return null;
  };

  // Base64-URL of a JSON payload, unicode-safe — used for the request code and
  // for the shareable access link, so awkward characters survive copy/paste.
  function b64uEncode(str){
    var b64=btoa(unescape(encodeURIComponent(str)));
    return b64.replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');
  }
  function b64uDecode(s){
    s=String(s||'').replace(/-/g,'+').replace(/_/g,'/');
    while(s.length%4)s+='=';
    return decodeURIComponent(escape(atob(s)));
  }
  MDL.b64uEncode=b64uEncode; MDL.b64uDecode=b64uDecode;

  MDL.packReq = function(req){
    return b64uEncode(JSON.stringify({v:1,id:req.id,n:req.name,s:req.surname,w:req.ward,t:req.ts}));
  };
  MDL.unpackReq = function(code){
    var o=JSON.parse(b64uDecode(String(code||'').trim()));
    if(!o||o.v!==1||!o.id)throw new Error('bad code');
    return {id:o.id,name:o.n||'',surname:o.s||'',ward:o.w||'',ts:o.t||Date.now(),status:'pending'};
  };

  // Build the shareable access link for an approved staff member. The code rides
  // in the URL fragment (#k=…) so opening the link unlocks the directory.
  MDL.buildAccessLink = function(code){
    var base=location.origin + location.pathname.replace(/[^/]*$/,'') + 'index.html';
    return base + '#k=' + encodeURIComponent(code);
  };

  MDL.fmtWhen = function(ts){
    try{
      var d=new Date(ts);
      return d.toLocaleString(undefined,{day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'});
    }catch(e){ return ''; }
  };

  MDL.copyText = function(text){
    if(navigator.clipboard && navigator.clipboard.writeText){
      return navigator.clipboard.writeText(text).then(function(){return true;}).catch(function(){return false;});
    }
    try{
      var ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0';
      document.body.appendChild(ta); ta.select(); var ok=document.execCommand('copy'); document.body.removeChild(ta);
      return Promise.resolve(ok);
    }catch(e){ return Promise.resolve(false); }
  };

  window.MDL = MDL;
})();
