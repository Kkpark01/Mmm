document.addEventListener('DOMContentLoaded',()=>{
 const root=document.querySelector('.mcm3007u-screen');if(!root)return;
 const rows=kind=>[...document.querySelectorAll('#'+kind+'Table tbody tr')];
 const key='mcm3007u.selection';let saved={};try{saved=JSON.parse(sessionStorage.getItem(key)||'{}');}catch{}
 if(saved.token!==root.dataset.searchToken)saved={token:root.dataset.searchToken};
 const remember=()=>{try{sessionStorage.setItem(key,JSON.stringify(saved));}catch{}};
 const mark=(kind,row)=>rows(kind).forEach(r=>r.classList.toggle('selected',r===row));
 const show=(kind,predicate)=>{const list=rows(kind);list.forEach(r=>r.hidden=!predicate(r));const n=list.filter(r=>!r.hidden).length;document.querySelector('#count-'+kind).textContent=n+'件';document.querySelector('#'+kind+'Scroll .empty-grid').hidden=n!==0;return list.filter(r=>!r.hidden);};
 const chooseUm=row=>{mark('um',row);saved.um=row?.dataset.index;show('tk',r=>!!row&&r.dataset.quote===row.dataset.quote);remember();};
 const chooseMaster=row=>{mark('master',row);saved.master=row?.dataset.index;const available=show('um',r=>!!row&&r.dataset.plant===row.dataset.plant&&r.dataset.nonyusaki===row.dataset.nonyusaki);chooseUm(available.find(r=>r.dataset.index===saved.um)||available[0]);remember();};
 for(const kind of ['master','um','tk'])for(const row of rows(kind)){
  const choose=()=>kind==='master'?chooseMaster(row):kind==='um'?chooseUm(row):mark(kind,row);
  row.addEventListener('click',choose);row.addEventListener('keydown',e=>{if(e.target!==row)return;if(['Enter',' '].includes(e.key)){e.preventDefault();choose();}else if(['ArrowUp','ArrowDown'].includes(e.key)){e.preventDefault();const a=rows(kind).filter(r=>!r.hidden),i=a.indexOf(row)+(e.key==='ArrowDown'?1:-1);if(a[i]){a[i].focus();a[i].click();}}});
 }
 show('master',()=>true);chooseMaster(rows('master').find(r=>r.dataset.index===saved.master)||rows('master')[0]);
 document.querySelectorAll('.discard-form').forEach(f=>f.addEventListener('submit',e=>{if(!window.confirm('選択した契約を破棄します。よろしいですか？'))e.preventDefault();}));
});
