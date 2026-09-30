(() => {
  'use strict';
  document.addEventListener('DOMContentLoaded', () => {
    const screen=document.querySelector('.mcm1011u-screen'); if(!screen)return;
    const grid=screen.querySelector('.grid-wrap'), rows=Array.from(screen.querySelectorAll('#mainTbody tr[data-row-key]'));
    function setupColumns(table, wrap, main) {
      const cols=Array.from(table.querySelectorAll('colgroup col')), heads=Array.from(table.querySelectorAll('thead th'));
      const widths=cols.map(c=>parseFloat(c.style.width));
      function fixed(){
        // VBはプラント名まで固定。小画面では契約NOまで残して右側の列への到達を確保する。
        const count=main?(widths.slice(0,12).reduce((a,b)=>a+b,0)<wrap.clientWidth-80?12:4):0;
        table.querySelectorAll('tr').forEach(row=>{let left=0;Array.from(row.cells).forEach((cell,i)=>{
          const pinned=i<count&&cell.colSpan===1;cell.classList.toggle('col-fixed',pinned);cell.style.left=pinned?left+'px':'';left+=widths[i]||0;
        });});
      }
      function update(){cols.forEach((c,i)=>c.style.width=widths[i]+'px');table.style.width=widths.reduce((a,b)=>a+b,0)+'px';fixed();}
      update();new ResizeObserver(fixed).observe(wrap);
      heads.forEach((head,index)=>{
        if(index===0||(main&&index===1))return;
        const handle=document.createElement('span');handle.className='col-resizer';handle.title='ドラッグして列幅を変更';head.append(handle);
        handle.addEventListener('pointerdown',event=>{
          event.preventDefault();event.stopPropagation();handle.setPointerCapture(event.pointerId);
          const startX=event.clientX,startWidth=widths[index];document.body.classList.add('col-resizing');
          const move=e=>{widths[index]=Math.max(40,startWidth+e.clientX-startX);update();};
          const end=()=>{handle.removeEventListener('pointermove',move);handle.removeEventListener('pointerup',end);handle.removeEventListener('pointercancel',end);document.body.classList.remove('col-resizing');};
          handle.addEventListener('pointermove',move);handle.addEventListener('pointerup',end);handle.addEventListener('pointercancel',end);
        });
      });
    }
    setupColumns(screen.querySelector('#mainGrid'),grid,true);
    screen.querySelectorAll('.tenpu-table').forEach(table=>{
      setupColumns(table,screen.querySelector('#tenpuContent'),false);
      const files=Array.from(table.tBodies[0].rows);
      function selectFile(row){files.forEach(r=>{const active=r===row;r.classList.toggle('selected',active);r.classList.toggle('row-selected',active);r.tabIndex=active?0:-1;});}
      files.forEach(row=>{
        row.addEventListener('click',()=>selectFile(row));
        row.addEventListener('keydown',e=>{
          if(e.target.closest('a')||!['ArrowUp','ArrowDown',' '].includes(e.key))return;
          e.preventDefault();const next=files[Math.max(0,Math.min(files.length-1,files.indexOf(row)+(e.key==='ArrowUp'?-1:e.key==='ArrowDown'?1:0)))];
          selectFile(next);next.focus({preventScroll:true});next.scrollIntoView({block:'nearest'});
        });
      });
    });
    const key='mcm1011u.gridState';
    const storage={get(){try{return JSON.parse(sessionStorage.getItem(key)||'{}')||{};}catch{return {};}},set(v){try{sessionStorage.setItem(key,JSON.stringify(v));}catch{}},clear(){try{sessionStorage.removeItem(key);}catch{}}};
    function save(){const row=rows.find(r=>r.classList.contains('row-selected'));storage.set({id:row?.dataset.rowKey,top:grid.scrollTop,left:grid.scrollLeft});}
    function select(row){
      rows.forEach(r=>{const active=r===row;r.classList.toggle('row-selected',active);r.classList.toggle('selected',active);r.setAttribute('aria-selected',String(active));r.tabIndex=active?0:-1;});
      screen.querySelectorAll('.tenpu-item').forEach(item=>item.style.display=item.id==='tenpu-'+row.dataset.index?'':'none');
      save();
    }
    if(screen.dataset.resetGrid==='true')storage.clear();
    const state=storage.get();
    if(rows.length)select(rows.find(r=>r.dataset.rowKey===state.id)||rows[0]);
    grid.scrollTop=Number(state.top)||0;grid.scrollLeft=Number(state.left)||0;
    grid.addEventListener('scroll',save,{passive:true});
    rows.forEach(row=>{
      row.addEventListener('click',()=>select(row));
      row.addEventListener('keydown',event=>{
        if(event.target.matches('input,a,button'))return;
        if(!['ArrowUp','ArrowDown',' '].includes(event.key))return;
        event.preventDefault();const delta=event.key==='ArrowUp'?-1:event.key==='ArrowDown'?1:0;
        const next=rows[Math.max(0,Math.min(rows.length-1,rows.indexOf(row)+delta))];select(next);next.focus({preventScroll:true});next.scrollIntoView({block:'nearest',inline:'nearest'});
      });
    });
    screen.querySelector('.search-area').addEventListener('submit',()=>storage.clear());
    const all=screen.querySelector('#chkAll');
    const boxes=()=>Array.from(screen.querySelectorAll('[name=checkedIds]:not(:disabled)'));
    function sync(){const c=boxes(),n=c.filter(b=>b.checked).length;all.checked=c.length>0&&n===c.length;all.indeterminate=n>0&&n<c.length;}
    all.addEventListener('change',()=>{boxes().forEach(c=>c.checked=all.checked);sync();});
    boxes().forEach(c=>c.addEventListener('change',sync));
    let pending=false;
    for(const id of ['approveButton','rejectButton']){
      const button=screen.querySelector('#'+id);
      button.addEventListener('click',async()=>{
        if(pending||button.disabled)return;
        let selected=boxes().filter(c=>c.checked);
        if(id==='approveButton'){
          selected.filter(c=>['3','4'].includes(c.dataset.jotai)).forEach(c=>c.checked=false);
          selected=boxes().filter(c=>c.checked);sync();
        }
        if(!selected.length){await McmNotifications.error(id==='approveButton'?'審査・承認可能な申請が選定されていません。':'差戻し可能な申請が選択されていません。');return;}
        pending=true;
        try{
          if(!await McmNotifications.confirm(id==='approveButton'?'審査・承認処理を実施します。よろしいですか？':'差し戻し処理を実施します。よろしいですか？')){pending=false;return;}
          save();const form=screen.querySelector('#actionForm');form.action=button.dataset.action;
          screen.querySelectorAll('.action-bar button').forEach(b=>b.disabled=true);
          form.submit();
        }catch(error){pending=false;throw error;}
      });
    }
  });
})();
