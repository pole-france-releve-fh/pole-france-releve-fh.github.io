/* Pôle France Relève · service worker : notifications et mise en cache de la page */
var API='https://script.google.com/macros/s/AKfycbxm3C8aW7kjv6PQCQ151w3swN17ugD58ADi8hiAgwNe16uHQDzLRbR15HKcy8GILsOJ/exec',V='pfr-v34';
self.addEventListener('install',function(e){self.skipWaiting();e.waitUntil(caches.open(V).then(function(c){return c.addAll(['./','index.html','manifest.json','icon-192.png','badge.png']);}).catch(function(){}));});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(l){return Promise.all(l.filter(function(k){return k!==V&&k!=='pfr';}).map(function(k){return caches.delete(k);}));}).then(function(){return self.clients.claim();}));});
self.addEventListener('fetch',function(e){var u=new URL(e.request.url);if(e.request.method!=='GET'||u.origin!==location.origin)return;
  e.respondWith(fetch(e.request).then(function(r){var c=r.clone();caches.open(V).then(function(x){x.put(e.request,c);});return r;}).catch(function(){return caches.match(e.request).then(function(r){return r||caches.match('index.html');});}));});
function montrer(t,o){o.icon='icon-192.png';o.badge='badge.png';return self.registration.showNotification(t,o);}
self.addEventListener('push',function(e){
  /* iOS retire l'abonnement si une notification n'est pas affichée tout de suite : on montre d'abord un message d'attente,
     remplacé ensuite (même étiquette) par le vrai contenu lu sur le serveur */
  var ATT='pfr-msg';
  e.waitUntil(montrer('Pôle France Relève',{body:'Nouveau message…',tag:ATT,data:{lien:''}}).then(function(){return caches.open('pfr');}).then(function(c){return c.match('cle');}).then(function(r){return r?r.text():'';}).then(function(k){
    return Promise.race([fetch(API,{method:'POST',body:JSON.stringify({action:'api',fn:'mesNotifs',args:[k]})}).then(function(r){return r.json();}),new Promise(function(_,rej){setTimeout(function(){rej(new Error('lent'));},20000);})]);
  }).then(function(j){var l=(j&&j.ok)||[];
    /* chaque appareil (téléphone, ordinateur) retient ce qu'il a déjà affiché : un message n'est plus « consommé » par le premier appareil */
    return caches.open('pfr').then(function(c){return c.match('vus').then(function(r){return r?r.json():[];}).catch(function(){return [];}).then(function(vus){
      var neufs=l.filter(function(n){return vus.indexOf(n.id)<0;}).slice(0,4);
      if(!neufs.length)neufs=l.slice(0,1);
      vus=vus.concat(neufs.map(function(n){return n.id;})).slice(-200);
      return c.put('vus',new Response(JSON.stringify(vus))).then(function(){
        if(!neufs.length)return;
        /* pastille rouge sur l'icône : nombre de messages non vus depuis la dernière ouverture */
        var nb=neufs.filter(function(n){return vus.slice(0,-neufs.length).indexOf(n.id)<0;}).length;
        var pastille=c.match('pastille').then(function(r){return r?r.text():'0';}).then(function(t){var v=(Number(t)||0)+nb;return c.put('pastille',new Response(String(v))).then(function(){if(self.navigator&&self.navigator.setAppBadge)return self.navigator.setAppBadge(v).catch(function(){});});});
        return Promise.all(neufs.map(function(n,i){return montrer(n.titre,{body:n.corps,tag:i===0?ATT:n.id,data:{lien:n.lien||'',id:n.id}});}).concat([pastille]));});});});
  }).catch(function(){}));
});
self.addEventListener('notificationclick',function(e){e.notification.close();var url=self.registration.scope+(e.notification.data&&e.notification.data.lien||'');
  var lien=e.notification.data&&e.notification.data.lien||'',nid=e.notification.data&&e.notification.data.id||'';if(nid)url=self.registration.scope+'?lu='+encodeURIComponent(nid)+lien;
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(function(l){if(l.length){var c=l[0];c.postMessage({type:'ouvrir',lien:lien,id:nid});return c.focus?c.focus():c;}return self.clients.openWindow(url);}));});
