document.addEventListener('DOMContentLoaded', () => {
 const root=document.querySelector('.mcm3007u-screen');if(!root)return;
 const kinds=['master','um','tk'];
 const tables=Object.fromEntries(kinds.map(k=>[k,document.getElementById(k+'Table')]));
 const lists=Object.fromEntries(kinds.map(k=>[k,[...tables[k].querySelectorAll('tbody tr')]]));
 const key='mcm3007u.selection';let saved={},widths={};
 try{saved=JSON.parse(sessionStorage.getItem(key)||'{}')||{};widths=JSON.parse(sessionStorage.getItem('mcm3007u.widths')||'{}')||{};}catch{}
 // 操作トークンは破棄時に更新するが、選択・スクロールは同じ検索結果の間保持する。
 if(saved.token!==root.dataset.selectionToken)saved={token:root.dataset.selectionToken};
 const remember=()=>{try{sessionStorage.setItem(key,JSON.stringify(saved));}catch{}};
 const mark=(kind,row,cell)=>{
  lists[kind].forEach(r=>{r.classList.toggle('selected',r===row&&!cell);r.setAttribute('aria-selected',String(r===row));});
  tables[kind].querySelectorAll('.cell-active').forEach(c=>c.classList.remove('cell-active'));
  cell?.classList.add('cell-active');saved[kind]=row?.dataset.key;
 };
 const show=(kind,predicate)=>{
  const visible=[];for(const row of lists[kind]){row.hidden=!predicate(row);if(!row.hidden){visible.push(row);row.querySelector('.row-number').textContent=visible.length;}}
  document.querySelector('#'+kind+'Scroll .empty-grid').hidden=visible.length!==0;return visible;
 };
 const chooseTk=(row,cell)=>{mark('tk',row,cell);remember();};
 const chooseUm=(row,cell)=>{
  mark('um',row,cell);const available=show('tk',r=>!!row&&r.dataset.quote===row.dataset.quote);
  chooseTk(available.find(r=>r.dataset.key===saved.tk)||available[0]);remember();
 };
 const chooseMaster=(row,cell)=>{
  mark('master',row,cell);const available=show('um',r=>!!row&&r.dataset.plant===row.dataset.plant&&r.dataset.nonyusaki===row.dataset.nonyusaki);
  chooseUm(available.find(r=>r.dataset.key===saved.um)||available[0]);remember();
 };
 for(const kind of kinds)for(const row of lists[kind]){
  const choose=cell=>kind==='master'?chooseMaster(row,cell):kind==='um'?chooseUm(row,cell):chooseTk(row,cell);
  row.addEventListener('click',e=>{
   const cell=e.target.closest('td');choose(cell&&!cell.classList.contains('row-head')?cell:null);
   if(!e.target.closest('a,button,input'))row.focus({preventScroll:true});
  });
  row.addEventListener('keydown',e=>{
   if(['ArrowUp','ArrowDown'].includes(e.key)){
    e.preventDefault();const visible=lists[kind].filter(r=>!r.hidden),next=visible[visible.indexOf(row)+(e.key==='ArrowDown'?1:-1)];
    if(next){next.focus({preventScroll:true});next.scrollIntoView({block:'nearest',inline:'nearest'});next.click();}
   } else if(e.target===row&&['Enter',' '].includes(e.key)){e.preventDefault();choose();}
  });
 }
 show('master',()=>true);chooseMaster(lists.master.find(r=>r.dataset.key===saved.master)||lists.master[0]);
 for(const kind of kinds){
  const scroll=document.getElementById(kind+'Scroll'),position=saved[kind+'Scroll'];
  if(Array.isArray(position)){scroll.scrollLeft=position[0];scroll.scrollTop=position[1];}
  scroll.addEventListener('scroll',()=>{saved[kind+'Scroll']=[scroll.scrollLeft,scroll.scrollTop];remember();},{passive:true});
  const table=tables[kind],cols=[...table.querySelectorAll('col')],heads=[...table.querySelectorAll('thead th')];
  const colWidths=cols.map((col,i)=>Math.max(i===0?24:40,Number(widths[kind]?.[i])||parseFloat(col.style.width)||80));
  const apply=()=>{
   cols.forEach((col,i)=>col.style.width=colWidths[i]+'px');table.style.width=colWidths.reduce((a,b)=>a+b,0)+'px';
   let left=0,fixed=0;const max={master:5,um:3,tk:3}[kind];
   while(fixed<max&&(fixed===0||left+colWidths[fixed]<=scroll.clientWidth-80)){left+=colWidths[fixed++];}
   for(const row of table.rows){let offset=0;[...row.cells].forEach((cell,i)=>{cell.classList.toggle('fixed-col',i<fixed);cell.style.left=i<fixed?offset+'px':'';offset+=colWidths[i];});}
  };
  heads.forEach((head,i)=>{
   if(i===0)return;
   const handle=document.createElement('span');handle.className='col-resizer';handle.title='ドラッグして列幅を変更';handle.setAttribute('aria-hidden','true');head.append(handle);
   handle.addEventListener('pointerdown',e=>{
    if(e.button!==0)return;e.preventDefault();e.stopPropagation();
    const x=e.clientX,initial=colWidths[i];handle.setPointerCapture(e.pointerId);root.classList.add('resizing');
    const move=ev=>{colWidths[i]=Math.max(40,Math.round(initial+ev.clientX-x));apply();};
    const end=()=>{handle.removeEventListener('pointermove',move);handle.removeEventListener('pointerup',end);handle.removeEventListener('pointercancel',end);handle.removeEventListener('lostpointercapture',end);root.classList.remove('resizing');widths[kind]=colWidths;try{sessionStorage.setItem('mcm3007u.widths',JSON.stringify(widths));}catch{}};
    handle.addEventListener('pointermove',move);handle.addEventListener('pointerup',end);handle.addEventListener('pointercancel',end);handle.addEventListener('lostpointercapture',end);
   });
  });
  apply();new ResizeObserver(apply).observe(scroll);
 }
 document.querySelectorAll('.discard-form').forEach(f=>{
  let confirming=false,approved=false,submitting=false;
  window.addEventListener('pageshow',e=>{if(e.persisted){confirming=false;approved=false;submitting=false;}});
  f.addEventListener('submit',async e=>{
   if(submitting){e.preventDefault();return;}
   if(approved){submitting=true;window.McmCommon?.showProgress();return;}
   e.preventDefault();const submitter=e.submitter;
   if(confirming||!submitter)return;
   confirming=true;
   try {
    if(await McmNotifications.confirm('契約を破棄します。よろしいですか？')){approved=true;f.requestSubmit(submitter);}
   } finally {confirming=false;}
  });
 });
});
