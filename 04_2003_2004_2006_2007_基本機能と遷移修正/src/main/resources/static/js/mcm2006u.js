'use strict';
(()=>{const form=document.getElementById('saveForm');if(!form)return;
 // テキスト入力のEnterで先頭の「行追加」が暗黙実行されることを防ぐ。
 form.addEventListener('keydown',event=>{if(event.key==='Enter'&&event.target instanceof HTMLInputElement&&['text','number'].includes(event.target.type))event.preventDefault();});
 document.querySelectorAll('.flag').forEach(box=>box.addEventListener('change',()=>{box.previousElementSibling.value=box.checked?'1':'0';}));
 const rows=id=>[...document.querySelectorAll('#'+id+'Table tbody tr')];
 function select(id,row){rows(id).forEach(r=>r.classList.toggle('selected',r===row));if(id==='brand')filter('kosei',row?.dataset.id);if(id==='kosei')filter('detail',row?.dataset.id);}
 function filter(id,parent){const list=rows(id);list.forEach(r=>r.hidden=!parent||r.dataset.parent!==parent);const visible=list.filter(r=>!r.hidden);select(id,visible.find(r=>r.classList.contains('selected'))||visible[0]);document.querySelector('#'+id+'Scroll .empty-grid').hidden=visible.length>0;}
 ['seiban','brand','kosei','detail','inspection'].forEach(id=>{const list=rows(id);list.forEach(row=>{row.addEventListener('click',()=>select(id,row));row.addEventListener('focusin',()=>select(id,row));row.addEventListener('keydown',e=>{if(e.target!==row||!['ArrowUp','ArrowDown'].includes(e.key))return;e.preventDefault();const visible=rows(id).filter(r=>!r.hidden),i=visible.indexOf(row);visible[i+(e.key==='ArrowDown'?1:-1)]?.focus();});});select(id,list.find(r=>!r.hidden));if(['seiban','inspection'].includes(id))document.querySelector('#'+id+'Scroll .empty-grid').hidden=list.length>0;});
 document.querySelectorAll('a[data-tab-index]').forEach(link=>link.addEventListener('click',event=>{if(form.dataset.readonly==='true')return;event.preventDefault();const button=document.createElement('button');button.type='submit';button.hidden=true;button.name='tabIndex';button.value=link.dataset.tabIndex;button.formAction=form.dataset.tabUrl;form.appendChild(button);form.requestSubmit(button);}));
 document.querySelectorAll('.remove-row').forEach(button=>button.addEventListener('click',()=>{const id=button.dataset.table,row=rows(id).find(r=>r.classList.contains('selected'));if(!row)return;const submit=document.createElement('button');submit.type='submit';submit.hidden=true;submit.formAction=form.dataset.rowsUrl+'?table='+id+'&operation=remove&rowIndex='+row.dataset.rowIndex;form.appendChild(submit);form.requestSubmit(submit);/* 非同期の日付検査が終わるまで送信ボタンをフォームに残す。 */}));

 const start=form.elements.namedItem('kaisiDt'),end=form.elements.namedItem('syuryoDt'),notice=document.getElementById('periodNotice');
 let rememberedStart=start?.value,rememberedEnd=end?.value,pending=null,replay=false,finishing=false;
 const message=(text,error=false)=>{notice.textContent=text;notice.hidden=!text;notice.classList.toggle('error',error);};
 const dateChanged=()=>start&&end&&(start.value!==rememberedStart||end.value!==rememberedEnd);
 async function synchronize(){
  if(pending){if(!await pending)return false;return synchronize();}
  if(!dateChanged())return true;
  const sentStart=start.value,sentEnd=end.value;
  pending=(async()=>{try{
   const response=await fetch(form.dataset.periodUrl,{method:'POST',body:new URLSearchParams(new FormData(form)),credentials:'same-origin',headers:{'Accept':'application/json'}});
   if(!response.ok)throw Error('HTTP '+response.status);const result=await response.json();
   if(result.workflowToken)form.elements.namedItem('workflowToken').value=result.workflowToken;
   const index=Number(form.elements.namedItem('selectedIndex').value),period=result.periods[index];
   if(period){if(start.value===sentStart)start.value=period.start;if(end.value===sentEnd)end.value=period.end;rememberedStart=period.start;rememberedEnd=period.end;}
   document.querySelectorAll('a[data-tab-index]').forEach(link=>{const t=result.periods[Number(link.dataset.tabIndex)];if(t)link.querySelector('span').textContent=t.label;});
   if(result.nextRenewal)document.getElementById('nextRenewal').value=result.nextRenewal;
   message(result.error||'前後の期間を確認し、日付を連動しました。登録で確定します。',!!result.error);
   return !result.error;
  }catch(error){message('期間を確認できませんでした。入力を保持しています。もう一度操作してください。',true);return false;}
  finally{pending=null;}})();return pending;
 }
 [start,end].filter(Boolean).forEach(field=>field.addEventListener('blur',()=>{if(!field.readOnly)synchronize();}));
 form.addEventListener('submit',async event=>{
  if(replay){replay=false;return;}
  const button=event.submitter,isUnlock=button?.id==='releaseLock';
  if(finishing){event.preventDefault();return;}
  if(!isUnlock&&!pending&&!dateChanged())return;
  event.preventDefault();finishing=true;
  try{
   if(isUnlock){const text='未登録の入力を破棄して契約を読み直し、管理者用の編集制限を解除します。よろしいですか？';if(!await (window.McmNotifications?window.McmNotifications.confirm(text):Promise.resolve(window.confirm(text))))return;if(pending)await pending;}
   else if(!await synchronize())return;
   replay=true;form.requestSubmit(button||undefined);
  }finally{finishing=false;}
 });
})();
