'use strict';
// VB同様に、親の解除は子へ、子の選択は親へ連動する。
function family(cb, kind) {
  const all = Array.from(document.querySelectorAll('.chk-'+kind));
  return all.filter(other => other.dataset.brand === cb.dataset.brand &&
      other.dataset.kosei === cb.dataset.kosei &&
      (kind === 'kosei' || !cb.dataset.meisai || other.dataset.meisai === cb.dataset.meisai));
}
function syncSelection(cb) {
  if (cb.classList.contains('chk-kosei')) {
    family(cb,'meisai').concat(family(cb,'kotai')).forEach(x => { x.checked=cb.checked; });
  } else if (cb.classList.contains('chk-meisai')) {
    family(cb,'kotai').forEach(x => { x.checked=cb.checked; });
    if(cb.checked)family(cb,'kosei').forEach(x => { x.checked=true; });
  } else if(cb.checked) {
    family(cb,'meisai').concat(family(cb,'kosei')).forEach(x => { x.checked=true; });
  }
}
function toggleAll(kind,checked) {
  document.querySelectorAll('.chk-'+kind).forEach(cb => {cb.checked=checked;syncSelection(cb);});
}
document.querySelectorAll('.chk-kosei,.chk-meisai,.chk-kotai').forEach(cb => cb.addEventListener('change',()=>syncSelection(cb)));
