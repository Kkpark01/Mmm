'use strict';
(() => {
    const rows = table => Array.from(table.tBodies[0].rows).filter(r => !r.classList.contains('empty-row'));
    const visibleRows = table => rows(table).filter(r => !r.hidden);
    const grids = new Map();
    // DB数値のscale差（1 / 1.0）をDOM側でも同一IDとして扱う。
    const id = (row, name) => (row?.dataset[name] || '').replace(/\.0+$/, '');
    function select(table, row, focus = false) {
        rows(table).forEach(r => { r.classList.toggle('selected', r === row); r.tabIndex = r === row ? 0 : -1; r.setAttribute('aria-selected', String(r === row)); });
        if (row && focus) { row.focus({preventScroll:true}); row.scrollIntoView({block:'nearest',inline:'nearest'}); }
        table.dispatchEvent(new CustomEvent('gridselect', {detail:row}));
    }
    function filter(table, match) {
        rows(table).forEach(row => { row.hidden = !match(row); });
        const available = visibleRows(table);
        table.querySelector('.empty-row').hidden = available.length !== 0;
        select(table, available[0] || null);
    }
    function setup(table) {
        grids.set(table.id, table);
        const cols = Array.from(table.querySelectorAll('colgroup col'));
        function size() { table.style.width = cols.reduce((n,c) => n + parseFloat(c.style.width),0) + 'px'; }
        size();
        table.querySelectorAll('thead th').forEach((th,i) => {
            if (i === 0) return;
            const handle = document.createElement('span'); handle.className='col-resizer'; handle.setAttribute('aria-hidden','true'); th.append(handle);
            handle.addEventListener('pointerdown', e => {
                e.preventDefault(); e.stopPropagation();
                const x=e.clientX, width=parseFloat(cols[i].style.width); handle.setPointerCapture(e.pointerId); document.body.classList.add('col-resizing');
                const move = ev => { cols[i].style.width=Math.max(40,width+ev.clientX-x)+'px'; size(); };
                const stop = () => { document.body.classList.remove('col-resizing'); handle.removeEventListener('pointermove',move); handle.removeEventListener('pointerup',stop); handle.removeEventListener('pointercancel',stop); };
                handle.addEventListener('pointermove',move); handle.addEventListener('pointerup',stop); handle.addEventListener('pointercancel',stop);
            });
        });
        rows(table).forEach(row => {
            row.tabIndex=-1;
            row.addEventListener('click', e => select(table,row,e.target.tagName !== 'INPUT'));
            row.addEventListener('keydown', e => {
                if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
                    e.preventDefault(); const all=visibleRows(table), index=all.indexOf(row);
                    select(table,all[Math.max(0,Math.min(all.length-1,index+(e.key==='ArrowDown'?1:-1)))],true);
                } else if (e.code==='Space' && e.target.tagName!=='INPUT') {
                    e.preventDefault(); select(table,row); const box=row.querySelector('input[type=checkbox][name]');
                    if (box && !box.disabled) { box.checked=!box.checked; box.dispatchEvent(new Event('change',{bubbles:true})); }
                }
            });
        });
        table.querySelector('.empty-row').hidden=rows(table).length!==0;
    }
    document.querySelectorAll('.data-grid').forEach(setup);
    if (grids.has('mitsumoriGrid')) {
        const master=grids.get('mitsumoriGrid'), periods=grids.get('kikanGrid'), prices=grids.get('tankaGrid');
        periods.addEventListener('gridselect', e => filter(prices,r => e.detail && id(r,'kikanId')===id(e.detail,'kikanId')));
        master.addEventListener('gridselect', e => filter(periods,r => e.detail && id(r,'keiyakujikanId')===id(e.detail,'keiyakujikanId')));
        select(master, rows(master).find(r=>r.querySelector('input:checked')) || rows(master)[0] || null);
        const date=document.querySelector('#kaisiDt:not([readonly]), #syuryoDt:not([readonly])'); if(date)date.focus();
    } else if (grids.has('koseiGrid')) {
        const master=grids.get('koseiGrid'), details=grids.get('meisaiGrid'), items=grids.get('kotaiGrid');
        // VB: 明細・個体はいずれも構成の子。明細行選択で個体を更に絞らない。
        master.addEventListener('gridselect', e => {
            filter(details,r=>e.detail && id(r,'kikikoseiId')===id(e.detail,'kikikoseiId'));
            filter(items,r=>e.detail && id(r,'kikikoseiId')===id(e.detail,'kikikoseiId'));
        });
        const box = row => row.querySelector('input[type=checkbox][name]');
        // 未選定行もクリックできるようにしてVBのMSG_0081Eを通知。権限なしのdisabledは維持する。
        let selectionErrorShowing = false;
        const inEstimate = b => b.dataset.eligible === 'true';
        const showSelectionError = async () => {
            if (selectionErrorShowing) return;
            selectionErrorShowing = true;
            try { await McmNotifications.error('前画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。'); }
            finally { selectionErrorShowing = false; }
        };
        [master, details, items].forEach(table => table.querySelectorAll('input[type=checkbox][name]').forEach(b => {
            const reject = e => {
                if (b.disabled || inEstimate(b)) return;
                e.preventDefault(); e.stopImmediatePropagation(); b.checked = false;
                void showSelectionError();
            };
            b.addEventListener('click', reject);
            b.addEventListener('change', reject); // 行のSpace操作・プログラム経由も同じ制御
        }));
        const eligibleSet = (row,on) => { const b=box(row); if(b&&!b.disabled&&inEstimate(b))b.checked=on; };
        const sameParent=(a,b)=>id(a,'kikikoseiId')===id(b,'kikikoseiId');
        const sameDetail=(a,b)=>sameParent(a,b)&&id(a,'kikimeisaiId')===id(b,'kikimeisaiId');
        const syncParent = row => rows(master).filter(r=>sameParent(r,row)).forEach(r=>eligibleSet(r,rows(details).some(d=>sameParent(d,row)&&box(d).checked)));
        master.addEventListener('change', e => { const row=e.target.closest('tr'); if(!row)return; rows(details).filter(r=>sameParent(r,row)).forEach(r=>eligibleSet(r,e.target.checked)); rows(items).filter(r=>sameParent(r,row)).forEach(r=>eligibleSet(r,e.target.checked)); });
        details.addEventListener('change', e => { const row=e.target.closest('tr'); if(!row)return; rows(items).filter(r=>sameDetail(r,row)).forEach(r=>eligibleSet(r,e.target.checked)); syncParent(row); });
        items.addEventListener('change', e => { const row=e.target.closest('tr'); if(!row)return; rows(details).filter(r=>sameDetail(r,row)).forEach(r=>eligibleSet(r,rows(items).some(t=>sameDetail(t,row)&&box(t).checked))); syncParent(row); });
        select(master,rows(master)[0] || null);
    }
    document.querySelectorAll('form.screen-form').forEach(form => form.addEventListener('submit', () => {
        form.querySelectorAll('input[data-unchecked]').forEach(e=>e.remove());
        // nameは元のインデックスのまま保持する。非表示の子表のチェックも送信する。
        form.querySelectorAll('input[type=checkbox][name]').forEach(box => {
            if (!box.checked || box.disabled) { const hidden=document.createElement('input'); hidden.type='hidden'; hidden.name=box.name; hidden.value='0'; hidden.dataset.unchecked='true'; form.append(hidden); }
        });
    }));
})();
