/* Pôle France Relève · service worker : notifications et mise en cache de la page */
var API='https://script.google.com/macros/s/AKfycbxm3C8aW7kjv6PQCQ151w3swN17ugD58ADi8hiAgwNe16uHQDzLRbR15HKcy8GILsOJ/exec',V='pfr-v22';
self.addEventListener('install',function(e){self.skipWaiting();e.waitUntil(caches.open(V).then(function(c){return c.addAll(['./','index.html','manifest.json','icon-192.png','badge.png']);}).catch(function(){}));});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(l){return Promise.all(l.filter(function(k){return k!==V&&k!=='pfr';}).map(function(k){return caches.delete(k);}));}).then(function(){return self.clients.claim();}));});
self.addEventListener('fetch',function(e){var u=new URL(e.request.url);if(e.request.method!=='GET'||u.origin!==location.origin)return;
  e.respondWith(fetch(e.request).then(function(r){var c=r.clone();caches.open(V).then(function(x){x.put(e.request,c);});return r;}).catch(function(){return caches.match(e.request).then(function(r){return r||caches.match('index.html');});}));});
function montrer(t,o){o.icon='icon-192.png';o.badge='badge.png';return self.registration.showNotification(t,o);}
self.addEventListener('push',function(e){
  e.waitUntil(caches.open('pfr').then(function(c){return c.match('cle');}).then(function(r){return r?r.text():'';}).then(function(k){
    return fetch(API,{method:'POST',body:JSON.stringify({action:'api',fn:'mesNotifs',args:[k]})}).then(function(r){return r.json();});
  }).then(function(j){var l=(j&&j.ok)||[];if(!l.length)return montrer('Pôle France Relève',{body:'Nouveau message',data:{lien:''}});
    return Promise.all(l.map(function(n){return montrer(n.titre,{body:n.corps,tag:n.id,data:{lien:n.lien||''}});}));
  }).catch(function(){return montrer('Pôle France Relève',{body:'Nouveau message',data:{lien:''}});}));
});
self.addEventListener('notificationclick',function(e){e.notification.close();var url=self.registration.scope+(e.notification.data&&e.notification.data.lien||'');
  e.waitUntil(self.clients.matchAll({type:'window',includeUncontrolled:true}).then(function(l){for(var i=0;i<l.length;i++){if('navigate' in l[i])return l[i].navigate(url).then(function(c){return c&&c.focus();});}return self.clients.openWindow(url);}));});
