/* MediLink — Ordering tab.
   The Stoma Care Clinic's weekly order (due every Saturday) is made of two forms:
     Part 1  M.M.M.U. IRO Top-Up System List, Disposables   (forms/stoma-care-mmml-part1.xlsx)
     Part 2  Cleaning Consumables Order List                 (forms/stoma-care-mmml-part2.xlsx)
   Quantities typed here are written straight into copies of those original Excel files
   (layout, borders and print settings untouched), then emailed as two attachments in one
   email, or printed. Everything stays on this device — nothing is uploaded anywhere. */
(function(){
  'use strict';

  var ORDER_TO = 'disposablessupplies.mdh@gov.mt';
  var ORDER_CC = 'stomacare.mdh@gov.mt';
  var STORE_KEY = 'medilink-order-v1';
  var XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // [sheet row, stock code, description, top-up quantity]
  var P1_ITEMS=[[14,"P2019","ALCOHOL WIPES",2],[15,"EQ02050","BASIC DRESSING PACK STERILE",18],[16,"F011","DRESSING STERILE - HYDROCOLLOID 10CM X 10CM",5],[17,"F012","DRESSING STERILE - HYDROCOLLOID 20CM X 20CM",5],[18,"F027","DRESSING STERILE - HYDROCOLLOID THIN FILM 5CM X 20CM",5],[19,"EQ07023","GLOVES - HYPOALLERGENIC LATEX EXAM SIZE MEDIUM",200],[20,"ICU014","GLOVES - HYPOALLERGENIC LATEX EXAM SIZE LARGE",200],[21,"EQ02033","LOW ADHERENT DRESSING STERILE - 5CM X 5CM",5],[22,"EQ12036","LANCET - BLOOD",50],[23,"EQ13096","GLUCOSE IN BLOOD MONITORING STRIPS",200],[24,"EQ17019","REMOVAL OF SUTURE PACK STERILE",1],[25,"EQ02037","SELF ADHESIVE FABRIC ROLL - 10CM X 10M",1],[26,"EQ02038","SELF ADHESIVE FABRIC ROLL - 15CM X 10M",1],[27,"EQ02013","SWAB DISP NON WOVEN STERILE . 10 X 10",225],[28,"EQ03004","TAPE - HYPOALLERGIC SURGICAL",1],[29,"ICU025","SANITIZING WIPES (CLINELL SURFACE WIPES)",1],[30,"EQ13534","SURGICAL HAIR CLIPPER BLADE",10],[31,"P2020","ALCOHOL HAND RUB",1],[32,"P3021","SKIN CARE LOTION HIGH QUALITY NON - GREASY MOISTURIZING PERFUME FREE 500ML",1],[33,"EQ16162","ADHESIEVE REMOVER SPRAY",5],[34,"MASK005","MASK FACE",50]];
  // [sheet row, stock code, description, unit]
  var P2_ITEMS=[[8,"10FOIL-005","ALUMINIUM FOIL","PACK"],[9,"15APRON015","APRONS DISPOSABLE APRONS FOR DISPENSERERS","Number"],[10,"11BAGS-070","BAGS HOT WATER SOLUBLE","BAGS"],[11,"11BAGS-040","BAGS PLASTIC CLEAR W/RED HD 165 X 230 X 25 (F/BLOOD SAMPLES)","BAGS"],[12,"11BAGS-080","BAGS PLASTIC CLEAR 19\" X 21\" W/MDH PRINT",""],[13,"11BAGS-062","BAGS REFUSE TRANSPARENT 32\" X 40\" FOR DOMESTIC WASTE ONLY","BAGS"],[14,"11BAGS-050","BAGS REFUSE YELLOW 32\" X 40\" FOR BIOHAZARD","BAGS"],[15,"11BAGS-065","BAGS VOMITING 123mm X 80mm X 237mm","BAGS"],[16,"01BINS002","BINS","NUMBER"],[17,"CC-BLADE-001","BLADES FOR SAFETY RAZORS",""],[18,"01BOWLS007","BOWLS DISPOSABLE","NUMBER"],[19,"15BRACE010","BRACELETS IDENTIFICATION ADULT","Number"],[20,"15BRACE020","BRACELETS IDENTIFICATION ADULT ALLERGY RED","Number"],[21,"15BRACE005","BRACELETS IDENTIFICATION PAEDIATRIC","Number"],[22,"15BRACE015","BRACELETS IDENTIFICATION PAEDIATRIC ALLERGY RED","Number"],[23,"03BROOM015","BROOM FIBRE","Number"],[24,"03BRUSH015","BRUSHES SCRUBBING","Number"],[25,"03BRUSH005","BRUSHES FOR BOTTLES","Number"],[26,"15CAPS-015","CAPS FOR FEMALE NURSES ITEM A FOR LONG HAIR","Number"],[27,"15CAPS-005","CAPS FOR FEMALE NURSES ITEM B FOR SHORT HAIR","Number"],[28,"15CAPS-010","CAPS FOR MALE NURSES","Number"],[29,"03CLOTH005","CLOTH KITCHEN","Number"],[30,"03DETER020","DETERGENT FLOOR X 5 LTRS","LTRS."],[31,"03DETER025","DISH WASHING LIQUID X 5 LITRES","LTRS."],[32,"01ENDM005","ENDS FOR COTTON MOPS 400gr","Number"],[33,"15MASKS005","FACE MASKS DISPOSABLE C/W FILTER","Number"],[34,"15FOAMS005","FOAM SHAVING",""],[35,"","FRENCH MOPS COMPLETE (60 WIDE)","Number"],[36,"","FRENCH MOPS COMPLETE (40 WIDE)","Number"],[37,"03POLIS","FURNITURE POLISH","Number"],[38,"03CLEAN020","GLASS CLEANING LIQUID X 1000ML","Number"],[39,"03GLOVE005","GLOVES RUBBER HOUSEHOLD","LARGE"],[40,"15GOWNS005","GOWNS PLASTIC DISPOSABLE","Number"],[41,"03DETER075","HAND & BODY WASH X 3.78 LTRS (F/OUTPATIENTS ONLY)","LTRS."],[42,"03DETER030","HAND & BODY WASH X 5 LTRS","LTRS."],[43,"03DETER032","HAND & BODY WASH X 500ML C/W DISPENSER PUMP (F/WARDS ONLY)","BOTTLES"],[44,"03DETER105","HAND WASH SERAMAN SENSITIVE 1000ml",""],[45,"01HANDL005","HANDLES ALUMINUIM FOR MOPS","Number"],[46,"","HYGEINE LIQUID X 5LTRS",""],[47,"15INCON005","INCONTINENCE BED PADS 40cm x 60cm","Number"],[48,"15TOWEL005","INTERFOLD HAND TOWELS","SHEETS"],[49,"01JUGSE010","JUGS ELECTRIC","NUMBER"],[50,"01KNIVE030","KNIVES TABLE","Number"],[51,"03MOPS010","KENTUCKY MOPS COTTON 400gr","Number"],[52,"01MATCH005","MATCHES","Number"],[53,"03MOPS020","MOP FLAT 60cm",""],[54,"03MOPS015","MOP FLAT 40cm",""],[55,"15NAPPI010","NAPPIES DISPOSABLE ADULT LARGE","Number"],[56,"15NAPPI015","NAPPIES PAEDIATRIC 12 - 25 KILOS","NUMBER"],[57,"15NAPPI020","NAPPIES PAEDIATRIC 2 - 5 KILOS","NUMBER"],[58,"10CTIES010","NYLON CABLE TIES (F/TRANSPARENT & YELLOW REFUSE BAGS ONLY)","Number"],[59,"CC-SHOE--001","OVERSHOES (ITU & THEATERS)",""],[60,"03PADS 005","PADS SCOURING",""],[61,"06PADLO005","PADLOCK BRASS 1\"",""],[62,"06PADLO010","PADLOCK BRASS 1.1/2\"",""],[63,"01PANS005","PANS DUST PLASTIC",""],[64,"12FOAM-010","PILLOWS SOFT SILICONISED FIBRE 50 X 70CM","Number"],[65,"03DETER070","PINE DISINFECTANT X 5 LITRES","LTRS."],[66,"01PLATE030","PLATES DISPOSABLE","NUMBER"],[67,"03POWDE010","POWDER DETERGENT FOR AUTOMATIC WASHING MACHINE",""],[68,"03POWDE020","POWDER SCOURING VIM","Number"],[69,"15SHEET005","SHEET COUCH ROLLS","Roll"],[70,"03SHAMP005","SHAMPOO ADULT",""],[71,"03SHAMP010","SHAMPOO BABY",""],[72,"03SKINS005","SKINS CHAMOIS LEATHER","Number"],[73,"03SOAP015","SOAP FOR BABY",""],[74,"03SOAP010","SOAP TOILET","Number"],[75,"03SOAP005","SOAP LAUNDRY","Number"],[76,"03SPONG010","SPONGES WIRE",""],[77,"03SPONG015","SPONGES BATH","Number"],[78,"03SPONG005","SPONGES SCOURING","Number"],[79,"08SPRAY005","SPRAY AIRFRESHENER 300ML","TINS"],[80,"08SPRAY010","SPRAY INSECTICIDE","TINS"],[81,"03SQUEE005","SQUEEGEES FLOOR 13\"LONG","Number"],[82,"03SQUEE010","SQUEEGEES FLOOR 22\"LONG","Number"],[83,"01STRAWS005","STRAWS DRINGING",""],[84,"03STEEL005","STEEL WOOL SMALL","PACK OF 10"],[85,"03STICK005","STICKS FOR BROOMS/MOPS","Number"],[86,"01STRET005","STRETCH AND SEAL","PACKS"],[87,"03SWABS005","SWABS FLOOR","Number"],[88,"15TAPE015","TAPE COTTON 1/2\"",""],[89,"03DETER005","THICK CLEANING BLEACH X 5 LITRES","LTRS."],[90,"01THERMOS","THERMOS FLASKS","Number"],[91,"03CLEAN015","TOILET CLEANER LIQUID ACID X 500ML","BOTTLES"],[92,"01TORCH005","TORCHES HAND 2 CELL TYPE",""],[93,"CC-TWINE-001","TWINE","ROLL"],[94,"15HOLDU005","UNIVERSAL ATTACHMENT FOR URINE & SECRE.",""],[95,"15APRON015","WHITE DISPOSABLE APRONS FOR DISPENSERS","Number"],[96,"01WRING005","WRINGERS PLASTIC 4100P","Number"],[97,"15APRON020","YELLOW DISPOSABLE APRONS FOR DISPENSERS","Number"],[98,"15APRON015","WHITE DISPOSABLE APRONS FOR DISPENSERS","NUMBER"],[99,"15TOWEL005","INTERFOLD HAND TOWELS","NUMBER"]];
  // Rows the clinic keeps visible on its own copy of Part 2 (its usual items).
  var P2_USUAL = {42:1, 98:1, 99:1};

  var DEFAULTS = { officer:'Jason Fenech', team:'Tuesday Team 1', remarks:'TEAM 1 - TUESDAY' };

  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g,function(c){
    return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]; }); }
  function pad(n){ return (n<10?'0':'')+n; }
  function iso(d){ return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate()); }
  function fromIso(s){ var p=String(s||'').split('-'); return p.length===3 ? new Date(+p[0],+p[1]-1,+p[2]) : null; }
  function dmy(d){ return pad(d.getDate())+'/'+pad(d.getMonth()+1)+'/'+d.getFullYear(); }
  function today(){ var d=new Date(); return new Date(d.getFullYear(),d.getMonth(),d.getDate()); }
  // This Saturday if today is Saturday, otherwise the coming one.
  function nextSaturday(from){ var d=new Date((from||today()).getTime()); d.setDate(d.getDate()+((6-d.getDay()+7)%7)); return d; }
  var DAYS=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  // ── Saved draft (this device only) ─────────────────────────────────────────
  var S;
  function load(){
    var s=null;
    try{ s=JSON.parse(localStorage.getItem(STORE_KEY)||'null'); }catch(e){}
    s=s||{};
    S={ date:s.date||'', officer:s.officer!=null?s.officer:DEFAULTS.officer,
        team:s.team!=null?s.team:DEFAULTS.team, remarks:s.remarks!=null?s.remarks:DEFAULTS.remarks,
        p1:s.p1||{}, p2:s.p2||{}, sent:s.sent||null };
    // A past order date rolls forward to the next Saturday.
    var d=fromIso(S.date);
    if(!d || d<today()) S.date=iso(nextSaturday());
  }
  function save(){ try{ localStorage.setItem(STORE_KEY, JSON.stringify(S)); }catch(e){} }

  function qtyOf(map,row){ var v=map[row]; return v==null?'':String(v).trim(); }
  function countFilled(map){ var n=0; for(var k in map){ if(qtyOf(map,k)!=='') n++; } return n; }
  function orderDate(){ return fromIso(S.date)||nextSaturday(); }

  // ── Rendering the two forms ────────────────────────────────────────────────
  var p2Mode='usual', p2Query='';
  function qtyInput(part,row,val,label){
    return '<input class="qty" type="text" inputmode="numeric" autocomplete="off" data-part="'+part+'" data-row="'+row+
      '" value="'+esc(val)+'" placeholder="0" aria-label="'+esc(label)+'"/>';
  }
  function renderP1(){
    var html=P1_ITEMS.map(function(it){
      var v=qtyOf(S.p1,it[0]);
      return '<div class="orow'+(v?' has':'')+'"><div class="body"><div class="oname">'+esc(it[2])+'</div>'+
        '<div class="ometa"><span>'+esc(it[1])+'</span><span>Top-up '+esc(it[3])+'</span></div></div>'+
        qtyInput(1,it[0],v,'Required: '+it[2])+'</div>';
    });
    $('p1-list').innerHTML=html.join('');
  }
  function renderP2(){
    var terms=p2Query.toLowerCase().split(/\s+/).filter(Boolean);
    var rows=P2_ITEMS.filter(function(it){
      if(terms.length){
        var hay=(it[1]+' '+it[2]).toLowerCase();
        for(var i=0;i<terms.length;i++){ if(hay.indexOf(terms[i])<0) return false; }
        return true;
      }
      return p2Mode==='all' || P2_USUAL[it[0]] || qtyOf(S.p2,it[0])!=='';
    });
    var html=rows.map(function(it){
      var v=qtyOf(S.p2,it[0]);
      return '<div class="orow'+(v?' has':'')+'"><div class="body"><div class="oname">'+esc(it[2])+'</div>'+
        '<div class="ometa"><span>'+esc(it[1]||'—')+'</span>'+(it[3]?'<span>'+esc(it[3])+'</span>':'')+'</div></div>'+
        qtyInput(2,it[0],v,'Demand: '+it[2])+'</div>';
    });
    $('p2-list').innerHTML=html.length?html.join(''):'<div class="note">No item matches that search.</div>';
    $('p2-hint').textContent = terms.length ? rows.length+' match'+(rows.length===1?'':'es')
      : (p2Mode==='usual' ? 'Usual items and anything ordered · switch to All items to add more' : P2_ITEMS.length+' items');
  }
  function renderCounts(){
    var a=countFilled(S.p1), b=countFilled(S.p2);
    $('p1-count').textContent=a+' item'+(a===1?'':'s');
    $('p2-count').textContent=b+' item'+(b===1?'':'s');
    $('p1-count').classList.toggle('on',a>0);
    $('p2-count').classList.toggle('on',b>0);
  }

  // ── "Due every Saturday" status ────────────────────────────────────────────
  function isDue(){
    var t=today();
    return t.getDay()===6 && !(S.sent && S.sent.date===iso(t));
  }
  function renderDue(){
    var t=today(), sat=nextSaturday(), b=$('due-banner'), msg, cls;
    var sentThisWeek = S.sent && S.sent.date===iso(sat);
    if(sentThisWeek){
      var at=new Date(S.sent.at);
      cls='ok'; msg='✓ Order for <b>Sat '+dmy(sat)+'</b> prepared for sending on '+DAYS[at.getDay()]+' '+dmy(at)+' at '+pad(at.getHours())+':'+pad(at.getMinutes())+'.'+
        ' <button type="button" class="linkbtn" id="new-order">Start next week’s order</button>';
    } else if(t.getDay()===6){
      cls='today'; msg='⏰ <b>Order due today</b> (Saturday '+dmy(t)+'). Fill in the quantities, then tap <b>Send email</b>.';
    } else {
      var days=Math.round((sat-t)/864e5);
      cls=''; msg='📅 This order is sent <b>every Saturday</b>. Next due <b>Sat '+dmy(sat)+'</b> ('+(days===1?'tomorrow':'in '+days+' days')+').';
    }
    b.className='due '+cls; b.innerHTML=msg;
    var nb=$('new-order'); if(nb) nb.addEventListener('click',newOrder);
    var dot=$('tab-order-dot'); if(dot) dot.hidden=!isDue();
  }
  function newOrder(){
    if(!confirm('Clear all quantities and start next week’s order?')) return;
    S.p1={}; S.p2={}; S.sent=null; S.date=iso(nextSaturday(new Date(today().getTime()+864e5)));
    save(); renderAll();
  }

  function renderAll(){
    $('o-date').value=S.date; $('o-officer').value=S.officer; $('o-team').value=S.team; $('o-remarks').value=S.remarks;
    renderP1(); renderP2(); renderCounts(); renderDue();
  }

  // ── Filling the original Excel files ───────────────────────────────────────
  function xmlEsc(s){ return String(s).replace(/[&<>"]/g,function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
  function colNum(ref){ var m=ref.match(/^[A-Z]+/)[0], n=0; for(var i=0;i<m.length;i++) n=n*26+(m.charCodeAt(i)-64); return n; }
  // Set one cell, keeping its style. Numbers stay numbers; text is written inline.
  function setCell(xml, ref, value){
    var isNum = typeof value==='number' || (/^\d+(\.\d+)?$/.test(String(value)));
    var re=new RegExp('<c r="'+ref+'"([^>]*?)(?:/>|>[\\s\\S]*?</c>)');
    var m=xml.match(re), attrs='';
    if(m){ var sm=m[1].match(/\ss="\d+"/); attrs=sm?sm[0]:''; }
    var cell = String(value)==='' ? '<c r="'+ref+'"'+attrs+'/>'
      : isNum ? '<c r="'+ref+'"'+attrs+'><v>'+Number(value)+'</v></c>'
      : '<c r="'+ref+'"'+attrs+' t="inlineStr"><is><t xml:space="preserve">'+xmlEsc(value)+'</t></is></c>';
    if(m) return xml.replace(re, function(){ return cell; });
    // Cell not present: insert it into its row in column order.
    var row=ref.match(/\d+$/)[0], rre=new RegExp('(<row r="'+row+'"[^>]*>)([\\s\\S]*?)(</row>)');
    var rm=xml.match(rre); if(!rm) return xml;
    var cells=rm[2], col=colNum(ref), pos=cells.length, cre=/<c r="([A-Z]+)\d+"/g, c;
    while((c=cre.exec(cells))){ if(colNum(c[1])>col){ pos=c.index; break; } }
    return xml.replace(rre, function(_,a,b,z){ return a+b.slice(0,pos)+cell+b.slice(pos)+z; });
  }
  function setRowHidden(xml,row,hidden){
    var re=new RegExp('<row r="'+row+'"[^>]*>');
    return xml.replace(re,function(tag){
      tag=tag.replace(/\shidden="1"/,'');
      return hidden ? tag.replace(/>$/,' hidden="1">') : tag;
    });
  }
  var templates={};
  function template(name){
    if(templates[name]) return Promise.resolve(templates[name]);
    return fetch('forms/'+name,{cache:'no-cache'}).then(function(r){
      if(!r.ok) throw new Error('Could not load '+name+' (HTTP '+r.status+')');
      return r.arrayBuffer();
    }).then(function(buf){ templates[name]=buf; return buf; });
  }
  function fillWorkbook(name, edit){
    if(!window.JSZip) return Promise.reject(new Error('The Excel helper did not load. Check the connection and reload.'));
    return template(name).then(function(buf){ return JSZip.loadAsync(buf); }).then(function(zip){
      var path='xl/worksheets/sheet1.xml';
      return zip.file(path).async('string').then(function(xml){
        zip.file(path, edit(xml));
        return zip.generateAsync({type:'blob', mimeType:XLSX_TYPE, compression:'DEFLATE'});
      });
    });
  }
  function buildPart1(){
    var d=orderDate();
    return fillWorkbook('stoma-care-mmml-part1.xlsx', function(x){
      x=setCell(x,'B3',S.team||' ');
      x=setCell(x,'A9','DATE:    '+dmy(d));
      x=setCell(x,'B38','Prepared By: '+(S.officer||''));
      P1_ITEMS.forEach(function(it){ x=setCell(x,'D'+it[0],qtyOf(S.p1,it[0])); });
      return x;
    });
  }
  function buildPart2(){
    var d=orderDate();
    return fillWorkbook('stoma-care-mmml-part2.xlsx', function(x){
      x=setCell(x,'D3',dmy(d));
      x=setCell(x,'B4',S.officer||'');
      x=setCell(x,'B5',S.remarks||'');
      // Like the clinic's own copy: show the usual items plus anything ordered; hide the rest.
      x=x.replace(/<sheetPr filterMode="1"\/>/,'').replace(/<autoFilter[\s\S]*?<\/autoFilter>/,'');
      P2_ITEMS.forEach(function(it){
        var v=qtyOf(S.p2,it[0]);
        x=setCell(x,'E'+it[0],v);
        x=setRowHidden(x,it[0],!(v!=='' || P2_USUAL[it[0]]));
      });
      return x;
    });
  }
  function fileNames(){
    var d=orderDate(), tag=d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());
    return ['Stoma Care MMML Part 1 - '+tag+'.xlsx','Stoma Care MMML Part 2 - '+tag+'.xlsx'];
  }
  function subject(){ return 'Stoma Care Clinic (STO-01) - Weekly Order - Saturday '+dmy(orderDate()); }
  function bodyText(){
    var a=countFilled(S.p1), b=countFilled(S.p2);
    return 'Good day,\n\nPlease find attached the Stoma Care Clinic (STO-01) weekly order for Saturday '+dmy(orderDate())+':\n\n'+
      '  1. Stoma Care MMML Part 1 - M.M.M.U. Top-Up System List, Disposables ('+a+' item'+(a===1?'':'s')+')\n'+
      '  2. Stoma Care MMML Part 2 - Cleaning Consumables Order List ('+b+' item'+(b===1?'':'s')+')\n\n'+
      'Kind regards,\n'+(S.officer||'')+'\nStoma Care Clinic, Out Patients Level -1\nExt. 4431/2';
  }

  // ── Sending ────────────────────────────────────────────────────────────────
  var built=null; // {files:[File,File], blobs:[Blob,Blob], names:[..]}
  function canShareFiles(files){
    try{ return !!(navigator.canShare && navigator.share && navigator.canShare({files:files})); }catch(e){ return false; }
  }
  function isPhone(){ return /android|iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform==='MacIntel' && navigator.maxTouchPoints>1); }
  function download(blob,name){
    var url=URL.createObjectURL(blob), a=document.createElement('a');
    a.href=url; a.download=name; a.rel='noopener'; document.body.appendChild(a); a.click();
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(url); },4000);
  }
  function markSent(){ var now=new Date(); S.sent={date:iso(orderDate()), at:now.getTime()}; save(); renderDue(); }
  function copyText(text,btn){
    var label=btn?btn.textContent:'';
    function done(ok){ if(btn){ btn.textContent= ok?'✓ Copied':'Copy failed'; setTimeout(function(){ btn.textContent=label; },1800);} }
    if(navigator.clipboard && navigator.clipboard.writeText){ return navigator.clipboard.writeText(text).then(function(){done(true);},function(){done(false);}); }
    try{ var ta=document.createElement('textarea'); ta.value=text; ta.style.position='fixed'; ta.style.opacity='0'; document.body.appendChild(ta); ta.select(); var ok=document.execCommand('copy'); document.body.removeChild(ta); done(ok); }catch(e){ done(false); }
  }

  function openSend(){
    var a=countFilled(S.p1), b=countFilled(S.p2);
    if(!a && !b && !confirm('No quantities have been entered on either form. Prepare the email anyway?')) return;
    var sheet=$('send-sheet');
    $('send-status').textContent='Preparing the two Excel attachments…';
    $('send-status').className='send-status';
    $('send-options').hidden=true;
    $('send-summary').innerHTML='Order for <b>Saturday '+esc(dmy(orderDate()))+'</b> · Part 1: <b>'+a+'</b> item'+(a===1?'':'s')+' · Part 2: <b>'+b+'</b> item'+(b===1?'':'s');
    sheet.classList.add('show');
    built=null;
    Promise.all([buildPart1(),buildPart2()]).then(function(blobs){
      var names=fileNames(), files=null;
      try{ files=blobs.map(function(bl,i){ return new File([bl],names[i],{type:XLSX_TYPE}); }); }catch(e){}
      built={blobs:blobs,names:names,files:files};
      $('send-files').innerHTML=names.map(function(n){ return '<li>📎 '+esc(n)+'</li>'; }).join('');
      var share = files && canShareFiles(files);
      $('opt-share').hidden=!share;
      // Phones get the share sheet first; computers get the Outlook draft first.
      var list=$('send-options');
      if(share && isPhone()) list.insertBefore($('opt-share'),list.firstChild);
      else list.insertBefore($('opt-eml'),list.firstChild);
      $('send-status').textContent='';
      $('send-options').hidden=false;
    }).catch(function(err){
      $('send-status').textContent='Could not prepare the attachments: '+(err&&err.message||err);
      $('send-status').className='send-status err';
    });
  }
  function closeSend(){ $('send-sheet').classList.remove('show'); }

  // 1) Phone: share sheet with both files attached (Gmail / Outlook / Mail).
  function sendShare(){
    if(!built||!built.files) return;
    // Put the recipients on the clipboard; share sheets can't pre-fill To/CC.
    copyText(ORDER_TO);
    navigator.share({files:built.files, title:subject(), text:bodyText()})
      .then(function(){ markSent(); flash('Shared. Check the To ('+ORDER_TO+') and CC ('+ORDER_CC+') fields, then send.'); })
      .catch(function(e){ if(e && e.name!=='AbortError') flash('Sharing failed: '+(e.message||e), true); });
  }
  // 2) Computer: an Outlook-ready draft (.eml) with To, CC, subject, text and both attachments.
  function b64(buf){
    var bytes=new Uint8Array(buf), bin='', CH=0x8000;
    for(var i=0;i<bytes.length;i+=CH) bin+=String.fromCharCode.apply(null, bytes.subarray(i,i+CH));
    return btoa(bin).replace(/.{76}/g,'$&\r\n');
  }
  function b64text(s){ return btoa(unescape(encodeURIComponent(s))).replace(/.{76}/g,'$&\r\n'); }
  function sendEml(){
    if(!built) return;
    Promise.all(built.blobs.map(function(b){ return b.arrayBuffer ? b.arrayBuffer() : new Response(b).arrayBuffer(); })).then(function(bufs){
      var B='----=_MediLink_'+Date.now().toString(36), CRLF='\r\n';
      var out=['To: '+ORDER_TO,'Cc: '+ORDER_CC,'Subject: '+subject(),'X-Unsent: 1','MIME-Version: 1.0',
        'Content-Type: multipart/mixed; boundary="'+B+'"','','This is a multi-part message in MIME format.',
        '--'+B,'Content-Type: text/plain; charset="utf-8"','Content-Transfer-Encoding: base64','',b64text(bodyText().replace(/\n/g,'\r\n'))];
      bufs.forEach(function(buf,i){
        var n=built.names[i];
        out.push('--'+B,'Content-Type: '+XLSX_TYPE+'; name="'+n+'"','Content-Disposition: attachment; filename="'+n+'"',
          'Content-Transfer-Encoding: base64','',b64(buf));
      });
      out.push('--'+B+'--','');
      var d=orderDate();
      download(new Blob([out.join(CRLF)],{type:'message/rfc822'}),'Stoma Care Order - '+d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate())+'.eml');
      markSent();
      flash('Email file downloaded — open it: Outlook shows a ready email with both attachments. Click Send.');
    });
  }
  // 3) Any device: download both files and open the email app with To/CC/subject filled in.
  function sendMailto(){
    if(!built) return;
    download(built.blobs[0],built.names[0]);
    setTimeout(function(){ download(built.blobs[1],built.names[1]); },600);
    var url='mailto:'+ORDER_TO+'?cc='+encodeURIComponent(ORDER_CC)+'&subject='+encodeURIComponent(subject())+
      '&body='+encodeURIComponent(bodyText()+'\n\n[Attach the two downloaded files: '+built.names.join(' and ')+']');
    setTimeout(function(){ location.href=url; },1400);
    markSent();
    flash('Both files downloaded. Attach them to the email that opens, then send.');
  }
  function flash(msg,err){
    var el=$('send-status'); el.textContent=msg; el.className='send-status '+(err?'err':'ok');
  }

  // ── Weekly Saturday reminder (calendar) ────────────────────────────────────
  function addReminder(){
    var first=nextSaturday(), d=first.getFullYear()+pad(first.getMonth()+1)+pad(first.getDate());
    var stamp=new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+/,'');
    var link=location.href.split('#')[0]+'#ordering';
    var ics=['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//MediLink//Stoma Care Order//EN','CALSCALE:GREGORIAN','BEGIN:VEVENT',
      'UID:medilink-stoma-order-'+d+'@medilink','DTSTAMP:'+stamp,'DTSTART:'+d+'T080000','DTEND:'+d+'T081500',
      'RRULE:FREQ=WEEKLY;BYDAY=SA','SUMMARY:Send Stoma Care weekly order',
      'DESCRIPTION:Open MediLink > Ordering\\, fill in Part 1 and Part 2\\, then tap Send email.\\nTo: '+ORDER_TO+'\\nCC: '+ORDER_CC+'\\n'+link,
      'URL:'+link,'BEGIN:VALARM','ACTION:DISPLAY','DESCRIPTION:Send Stoma Care weekly order','TRIGGER:PT0M','END:VALARM',
      'END:VEVENT','END:VCALENDAR',''].join('\r\n');
    download(new Blob([ics],{type:'text/calendar'}),'Stoma Care order reminder.ics');
  }

  // ── Printing ───────────────────────────────────────────────────────────────
  function printForms(){
    var d=orderDate();
    var p1rows=P1_ITEMS.map(function(it){
      return '<tr><td>'+esc(it[1])+'</td><td class="l">'+esc(it[2])+'</td><td>'+esc(it[3])+'</td><td class="q">'+esc(qtyOf(S.p1,it[0]))+'</td><td></td><td>'+esc(it[1])+'</td></tr>';
    }).join('');
    var p2items=P2_ITEMS.filter(function(it){ return qtyOf(S.p2,it[0])!=='' || P2_USUAL[it[0]]; });
    var p2rows=p2items.map(function(it){
      return '<tr><td>'+esc(it[1])+'</td><td class="l">'+esc(it[2])+'</td><td>'+esc(it[3])+'</td><td></td><td class="q">'+esc(qtyOf(S.p2,it[0]))+'</td><td></td></tr>';
    }).join('');
    $('print-area').innerHTML=
      '<section class="pform">'+
        '<table class="phead"><tr><td class="box">Your Ref:</td><td class="mid"><div>'+esc(S.team)+'</div><div>OUT PATIENTS LEVEL -1</div></td><td class="box">Our Ref.</td></tr></table>'+
        '<h2>M.M.M.U. - IRO Top-Up System List, Disposables</h2>'+
        '<h3>STO-01 &nbsp; Stoma Care Clinic &nbsp; Tel 25454431/2</h3>'+
        '<div class="pdate">DATE: &nbsp; '+esc(dmy(d))+'</div>'+
        '<table class="pt"><thead><tr><th>EQ Number</th><th>Item Description</th><th>Top Up Quantity</th><th>Required</th><th>Supplied</th><th>EQ Number</th></tr></thead><tbody>'+p1rows+'</tbody></table>'+
        '<div class="psign"><div><b>For Office Use Only:</b></div><div>Prepared By: '+esc(S.officer)+'</div><div>Issued By:</div><div>Recd. By:</div><div>Authorising Officer Dept:</div><div>Signature:</div></div>'+
      '</section>'+
      '<section class="pform">'+
        '<div class="pcap">MDH General Stores - Cleaning and Consumables Order List</div>'+
        '<h2>CLEANING CONSUMABLES ORDER LIST</h2>'+
        '<table class="pinfo"><tr><th>Ward / Section:</th><td>STOMA CARE CLINIC</td><th>Date:</th><td>'+esc(dmy(d))+'</td></tr>'+
        '<tr><th>Name and Surname Officer I/c:</th><td>'+esc(S.officer)+'</td><th>Ext. Nos:</th><td>4431/2</td></tr>'+
        '<tr><th>Remarks:</th><td>'+esc(S.remarks)+'</td><th>Cost Centre No:</th><td>STO-01</td></tr></table>'+
        '<table class="pt"><thead><tr><th>SLH General Stores Stock Codes</th><th>Stock Item Description</th><th>Unit</th><th>Quota</th><th>Demand</th><th>Supply</th></tr></thead><tbody>'+p2rows+'</tbody></table>'+
      '</section>';
    window.print();
  }

  // ── Wiring ─────────────────────────────────────────────────────────────────
  function onQty(e){
    var t=e.target; if(!t.classList || !t.classList.contains('qty')) return;
    var map = t.getAttribute('data-part')==='1' ? S.p1 : S.p2, row=t.getAttribute('data-row');
    var v=t.value.trim(); if(v==='') delete map[row]; else map[row]=v;
    t.closest('.orow').classList.toggle('has', v!=='');
    save(); renderCounts();
  }
  function init(){
    load();
    $('p1-list').addEventListener('input',onQty);
    $('p2-list').addEventListener('input',onQty);
    $('o-date').addEventListener('change',function(){ if(this.value){ S.date=this.value; save(); } else this.value=S.date; });
    $('o-officer').addEventListener('input',function(){ S.officer=this.value; save(); });
    $('o-team').addEventListener('input',function(){ S.team=this.value; save(); });
    $('o-remarks').addEventListener('input',function(){ S.remarks=this.value; save(); });
    $('p1-fill').addEventListener('click',function(){
      P1_ITEMS.forEach(function(it){ if(qtyOf(S.p1,it[0])==='') S.p1[it[0]]=String(it[3]); });
      save(); renderP1(); renderCounts();
    });
    $('p1-clear').addEventListener('click',function(){ if(countFilled(S.p1) && !confirm('Clear all Part 1 quantities?')) return; S.p1={}; save(); renderP1(); renderCounts(); });
    $('p2-clear').addEventListener('click',function(){ if(countFilled(S.p2) && !confirm('Clear all Part 2 quantities?')) return; S.p2={}; save(); renderP2(); renderCounts(); });
    $('p2-modes').addEventListener('click',function(e){
      var c=e.target.closest('[data-mode]'); if(!c) return;
      p2Mode=c.getAttribute('data-mode');
      var all=this.querySelectorAll('[data-mode]'); for(var i=0;i<all.length;i++) all[i].classList.toggle('on',all[i]===c);
      renderP2();
    });
    $('p2-q').addEventListener('input',function(){ p2Query=this.value.trim(); renderP2(); });
    $('btn-send').addEventListener('click',openSend);
    $('btn-print').addEventListener('click',printForms);
    $('btn-remind').addEventListener('click',addReminder);
    $('opt-share').addEventListener('click',sendShare);
    $('opt-eml').addEventListener('click',sendEml);
    $('opt-mailto').addEventListener('click',sendMailto);
    $('send-close').addEventListener('click',closeSend);
    $('send-sheet').addEventListener('click',function(e){ if(e.target===this) closeSend(); });
    $('copy-to').addEventListener('click',function(){ copyText(ORDER_TO,this); });
    $('copy-cc').addEventListener('click',function(){ copyText(ORDER_CC,this); });
    renderAll();
    // Warm the template cache so preparing the email is instant.
    template('stoma-care-mmml-part1.xlsx').catch(function(){}); template('stoma-care-mmml-part2.xlsx').catch(function(){});
  }
  window.MediLinkOrdering={ refresh:function(){ if(S){ load(); renderAll(); } }, isDue:function(){ return S?isDue():false; } };
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init); else init();
})();
