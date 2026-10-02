(function(){
  var css = `
  nav.topnav{position:sticky;top:0;z-index:50;display:flex;justify-content:center;gap:14px;flex-wrap:wrap;background:rgba(0,0,0,0.75);padding:10px 16px;border-bottom:1px solid rgba(255,255,255,0.15);backdrop-filter:blur(4px);}
  nav.topnav .nav-item{width:46px;height:46px;border-radius:50%;overflow:hidden;border:2px solid rgba(255,255,255,0.35);display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,0.08);transition:border-color .15s,transform .15s;flex-shrink:0;}
  nav.topnav .nav-item:hover{border-color:#ffeb3b;transform:scale(1.12);}
  nav.topnav .nav-item.active{border-color:#d32f2f;border-width:3px;}
  nav.topnav .nav-item img{width:100%;height:100%;object-fit:cover;display:block;}
  @media (max-width:520px){nav.topnav{gap:9px;padding:8px 10px;}nav.topnav .nav-item{width:38px;height:38px;}}
  `;
  var style = document.createElement('style');
  style.textContent = css;
  document.head.appendChild(style);

  var pages = [
    {href:"analyse.html", img:"https://i.postimg.cc/DZVM1DhN/urne.png", title:"Élections"},
    {href:"extreme-droite.html", img:"https://i.postimg.cc/6qd4xJ6k/maxresdefault.jpg", title:"Extrême droite", style:"object-position:top;"},
    {href:"immigration.html", img:"https://i.postimg.cc/4dCSpMGb/karl-marx.png", title:"Immigration"},
    {href:"securite.html", img:"https://i.postimg.cc/43qQRQ9C/ministere-interieur.png", title:"Sécurité"},
    {href:"economie.html", img:"https://i.postimg.cc/QtLycPDp/INSEE.png", title:"Économie"},
    {href:"emploi.html", img:"https://i.postimg.cc/9M3nyS2N/pole-emploi.png", title:"Emploi"}
  ];

  var current = location.pathname.split('/').pop() || 'index.html';

  var html = '<nav class="topnav"><a href="index.html" class="nav-item" title="Accueil"><img src="https://i.postimg.cc/tT0MFfbS/pcf43.png" alt="Accueil PCF 43"></a>';
  pages.forEach(function(p){
    var activeClass = (p.href === current) ? ' active' : '';
    var styleAttr = p.style ? ' style="'+p.style+'"' : '';
    html += '<a href="'+p.href+'" class="nav-item'+activeClass+'" title="'+p.title+'"><img src="'+p.img+'" alt="'+p.title+'"'+styleAttr+'></a>';
  });
  html += '</nav>';

  document.addEventListener('DOMContentLoaded', function(){
    document.body.insertAdjacentHTML('afterbegin', html);
  });
})();