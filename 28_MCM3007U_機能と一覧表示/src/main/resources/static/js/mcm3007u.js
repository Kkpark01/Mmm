document.addEventListener('DOMContentLoaded', () => {
 const root=document.querySelector('.mcm3007u-screen');if(!root)return;
 const kinds=['master','um','tk'];
 const tables=Object.fromEntries(kinds.map(k=>[k,document.getElementById(k+'Table')]));
 const lists=Object.fromEntries(kinds.map(k=>[k,[...tables[k].querySelectorAll('tbody tr')]]));
 const key='mcm3007u.selection';let saved={};
 try{saved=JSON.parse(sessionStorage.getItem(key)||'{}')||{};}catch{}
 if(saved.token!==root.dataset.searchToken)saved={token:root.dataset.searchToken};
 const remember=()=>{try{sessionStorage.setItem(key,JSON.stringify(saved));}catch{}};
 const mark=(kind,row,cell)=>{
  lists[kind].forEach(r=>{r.classList.toggle('selected',r===row&&!cell);r.setAttribute('aria-selected',String(r===row));});
  tables[kind].querySelectorAll('.cell-active').forEach(c=>c.classList.remove('cell-active'));
  cell?.classList.add('cell-active');saved[kind]=row?.dataset.index;
 };
 const show=(kind,predicate)=>{
  const visible=[];for(const row of lists[kind]){row.hidden=!predicate(row);if(!row.hidden){visible.push(row);row.querySelector('.row-number').textContent=visible.length;}}
  document.getElementById('count-'+kind).textContent=visible.length+'件';
  document.querySelector('#'+kind+'Scroll .empty-grid').hidden=visible.length!==0;return visible;
 };
 const chooseTk=(row,cell)=>{mark('tk',row,cell);remember();};
 const chooseUm=(row,cell)=>{
  mark('um',row,cell);const available=show('tk',r=>!!row&&r.dataset.quote===row.dataset.quote);
  chooseTk(available.find(r=>r.dataset.index===saved.tk)||available[0]);remember();
 };
 const chooseMaster=(row,cell)=>{
  mark('master',row,cell);const available=show('um',r=>!!row&&r.dataset.plant===row.dataset.plant&&r.dataset.nonyusaki===row.dataset.nonyusaki);
  chooseUm(available.find(r=>r.dataset.index===saved.um)||available[0]);remember();
 };
 for(const kind of kinds)for(const row of lists[kind]){
  const choose=cell=>kind==='master'?chooseMaster(row,cell):kind==='um'?chooseUm(row,cell):chooseTk(row,cell);
  row.addEventListener('click',e=>{const cell=e.target.closest('td');choose(cell&&!cell.classList.contains('row-head')?cell:null);});
  row.addEventListener('keydown',e=>{
   if(e.target!==row)return;
   if(['Enter',' '].includes(e.key)){e.preventDefault();choose();}
   else if(['ArrowUp','ArrowDown'].includes(e.key)){
    e.preventDefault();const visible=lists[kind].filter(r=>!r.hidden),next=visible[visible.indexOf(row)+(e.key==='ArrowDown'?1:-1)];
    if(next){next.focus();next.click();}
   }
  });
 }
 show('master',()=>true);chooseMaster(lists.master.find(r=>r.dataset.index===saved.master)||lists.master[0]);
 for(const kind of kinds){
  const scroll=document.getElementById(kind+'Scroll'),position=saved[kind+'Scroll'];
  if(Array.isArray(position)){scroll.scrollLeft=position[0];scroll.scrollTop=position[1];}
  scroll.addEventListener('scroll',()=>{saved[kind+'Scroll']=[scroll.scrollLeft,scroll.scrollTop];remember();},{passive:true});
 }
 document.querySelectorAll('.discard-form').forEach(f=>f.addEventListener('submit',e=>{
  if(!window.confirm('契約を破棄します。よろしいですか？'))e.preventDefault();
 }));
});
