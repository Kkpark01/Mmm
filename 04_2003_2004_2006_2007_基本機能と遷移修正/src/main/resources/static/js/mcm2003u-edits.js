/* VB Mcm2003uTabControl の Leave 計算。金額は BigInt の 1/1000000 円で保持する。 */
(() => {
  const form = document.getElementById('mainForm');
  if (!form) return;
  const fields = [...form.querySelectorAll('[data-brand-calc]')];
  const units = value => { const text=String(value||'0'), negative=text.startsWith('-'), parts=text.replace(/^-/, '').split('.'); return (negative?-1n:1n)*(BigInt(parts[0]||0)*1000000n+BigInt((parts[1]||'').padEnd(6,'0'))); };
  const decimal = value => { const sign=value<0n?'-':''; value=value<0n?-value:value;return sign+(value/1000000n).toString()+(value%1000000n?'.'+(value%1000000n).toString().padStart(6,'0').replace(/0+$/,''):''); };
  const get = key => units(document.getElementById('cost-'+key).value);
  const set = (key,value) => { document.getElementById('cost-'+key).value=decimal(value); };
  const roundedQuarter = value => { const sign=value<0n?-1n:1n;value=value<0n?-value:value;let result=value/4000000n; const rest=value%4000000n;if(rest>2000000n || rest===2000000n && result%2n===1n)result++;return sign*result*1000000n; };
  const limits=[5000000n,10000000n,15000000n,20000000n,25000000n,30000000n,35000000n,40000000n,45000000n,50000000n,55000000n,60000000n,65000000n,70000000n,80000000n,90000000n,100000000n];
  const rates=[150000n,300000n,450000n,700000n,1000000n,1200000n,1400000n,1500000n,1600000n,1700000n,1750000n,1800000n,1850000n,1900000n,2000000n,2100000n,2150000n];
  for(const field of fields)field.dataset.previous=field.value;
  function apply(field) {
    if(field.disabled || field.readOnly || field.value===field.dataset.previous)return;
    const key=field.dataset.brandCalc,value=field.value;
    field.setCustomValidity('');
    if(!field.checkValidity())return;
    if(key==='keiyakujikantai' && value!=='' && Number(value)>24){field.setCustomValidity('契約時間帯は0～24で入力してください。');return;}
    const event=document.createElement('input');event.type='hidden';event.name='brandCalculation';event.value=key+'='+value;form.appendChild(event);
    if(key==='keiyakujikantai') {
      const cost=get('systemSekkeiKin')+get('kihonSekkeiKin')+get('programSakuseiKin');
      const index=limits.findIndex(limit=>cost<limit*1000000n);
      let system=value===''?0n:index<0?cost:rates[index]*1000000n;
      if(value==='24')system=system*12n/10n;
      set('systemSupportKin',system);set('dtsSupportKin',system/4n);set('softHosyuKin',system+system/4n+get('daifukuGijutsuKin'));
    } else if(key==='hoseiSystemSupportKin') {
      const dts=roundedQuarter(get(key));set('hoseiDtsSupportKin',dts);set('hoseiSoftHosyuKin',get(key)+dts+get('hoseiDaifukuGijutsuKin'));
    } else {
      const difference=get('hoseiSoftHosyuKin')-get('hoseiDaifukuGijutsuKin');
      if(difference>0n){const system=(difference*4n/5000000n)*1000000n;set('hoseiSystemSupportKin',system);set('hoseiDtsSupportKin',difference-system);}
    }
    for(const input of fields)input.dataset.previous=input.value;
  }
  for(const field of fields){field.addEventListener('input',()=>field.setCustomValidity(''));field.addEventListener('change',()=>apply(field));}
  form.addEventListener('submit',event=>{for(const field of fields)apply(field);if(!form.reportValidity())event.preventDefault();});
})();
