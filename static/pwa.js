(() => {
  let deferredPrompt = null;
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;

  function createInstallButton(){
    if(isStandalone || document.getElementById('pwaInstallButton')) return;
    const btn=document.createElement('button');
    btn.id='pwaInstallButton';
    btn.className='pwa-install-button';
    btn.innerHTML='📲 تثبيت Shalfa Ops';
    btn.hidden=true;
    document.body.appendChild(btn);
    btn.addEventListener('click',async()=>{
      if(deferredPrompt){
        deferredPrompt.prompt();
        await deferredPrompt.userChoice;
        deferredPrompt=null;
        btn.hidden=true;
      }else if(/iPhone|iPad|iPod/i.test(navigator.userAgent)){
        alert('على iPhone: افتح زر المشاركة في Safari ثم اختر "إضافة إلى الشاشة الرئيسية".');
      }
    });
  }

  window.addEventListener('beforeinstallprompt',e=>{
    e.preventDefault();
    deferredPrompt=e;
    createInstallButton();
    const btn=document.getElementById('pwaInstallButton');
    if(btn) btn.hidden=false;
  });

  window.addEventListener('appinstalled',()=>{
    const btn=document.getElementById('pwaInstallButton');
    if(btn) btn.remove();
  });

  document.addEventListener('DOMContentLoaded',()=>{
    createInstallButton();

    // iOS helper
    if(!isStandalone && /iPhone|iPad|iPod/i.test(navigator.userAgent)){
      const btn=document.getElementById('pwaInstallButton');
      if(btn){
        btn.hidden=false;
        btn.innerHTML='📲 إضافة للشاشة الرئيسية';
      }
    }

    // Mobile sidebar toggle
    const sidebar=document.querySelector('.app-sidebar');
    if(sidebar && !document.getElementById('mobileMenuButton')){
      const menu=document.createElement('button');
      menu.id='mobileMenuButton';
      menu.className='mobile-menu-button';
      menu.setAttribute('aria-label','فتح القائمة');
      menu.innerHTML='☰';
      document.body.appendChild(menu);
      menu.addEventListener('click',()=>sidebar.classList.toggle('mobile-open'));
      document.addEventListener('click',e=>{
        if(window.innerWidth<=900 && sidebar.classList.contains('mobile-open') &&
           !sidebar.contains(e.target) && e.target!==menu){
          sidebar.classList.remove('mobile-open');
        }
      });
    }
  });

  if('serviceWorker' in navigator){
    window.addEventListener('load',()=>{
      navigator.serviceWorker.register('/sw.js',{scope:'/'}).then(reg=>{
        reg.addEventListener('updatefound',()=>{
          const worker=reg.installing;
          worker.addEventListener('statechange',()=>{
            if(worker.state==='installed' && navigator.serviceWorker.controller){
              const bar=document.createElement('div');
              bar.className='pwa-update-bar';
              bar.innerHTML='<span>يتوفر تحديث جديد للنظام.</span><button>تحديث الآن</button>';
              document.body.appendChild(bar);
              bar.querySelector('button').onclick=()=>{
                worker.postMessage({type:'SKIP_WAITING'});
              };
            }
          });
        });
      }).catch(err=>console.warn('Service worker registration failed',err));

      navigator.serviceWorker.addEventListener('controllerchange',()=>location.reload());
    });
  }
})();