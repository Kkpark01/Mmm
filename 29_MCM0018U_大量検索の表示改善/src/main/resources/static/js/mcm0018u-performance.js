/* 全結果を保持し、編集候補だけを共有する。検索条件・登録データ・行インデックスは変えない。 */
(() => {
 'use strict';
 let expanded=null, submitting=false;
 const sources={seizomakerId:'searchSeizomakerId',kikibunruiId:'searchKikibunruiId'};
 const key=v=>String(v??'').replace(/\.0+$/,'');
 function compact(select) {
  if(!select)return;
  const value=select.value,label=select.selectedOptions[0]?.textContent||'';
  select.replaceChildren(new Option('',''));
  if(value!=='')select.add(new Option(label,value,true,true));
  select.value=value;
  if(expanded===select)expanded=null;
 }
 function expand(select) {
  if(!select||select.disabled||!select.matches('select[data-shared-options]')||expanded===select)return;
  if(expanded)compact(expanded);
  const source=document.getElementById(sources[select.dataset.sharedOptions]);if(!source)return;
  const value=select.value,label=select.selectedOptions[0]?.textContent||'';
  const fragment=document.createDocumentFragment();let matched=value==='';
  for(const option of source.options){const v=key(option.value);const selected=v===key(value);matched=matched||selected;fragment.append(new Option(option.textContent,v,selected,selected));}
  if(!matched)fragment.append(new Option(label,value,true,true));
  select.replaceChildren(fragment);select.value=key(value);expanded=select;
 }
 // captureで既存の2クリック編集・新規行生成より前に候補を用意する。
 document.addEventListener('focusin',e=>expand(e.target),true);
 document.addEventListener('pointerdown',e=>expand(e.target.closest?.('select[data-shared-options]')),true);
 document.addEventListener('focusout',e=>{if(e.target===expanded)compact(expanded);},true);
 const overlay=()=>document.getElementById('mcm0018Loading');
 function hide(){if(overlay())overlay().hidden=true;submitting=false;}
 window.Mcm0018Performance={
  submitSearch(form){
   if(submitting)return;submitting=true;
   const el=overlay();if(el){el.querySelector('span').textContent='検索しています。結果の表示が完了するまでお待ちください。';el.hidden=false;}
   // ネットワーク待ちに入る前に処理中表示を描画する。
   requestAnimationFrame(()=>requestAnimationFrame(()=>{try{HTMLFormElement.prototype.submit.call(form);}catch(error){hide();throw error;}}));
  }
 };
 document.addEventListener('DOMContentLoaded',()=>requestAnimationFrame(()=>requestAnimationFrame(hide)));
 window.addEventListener('pageshow',hide);
})();
