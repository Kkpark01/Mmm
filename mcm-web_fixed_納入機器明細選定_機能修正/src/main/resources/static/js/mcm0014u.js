/* 全操作で同じ編集値を送り、検証エラー・通信エラーでも入力を保持する。 */
(() => {
    'use strict';
    const $ = id => document.getElementById(id), form = $('editForm');
    const base = form.action.replace(/\/settei$/, '');
    let busy = false, leaving = false, edited = form.dataset.changed === 'true';
    const isEdit = el => el.matches('[name="setNm"],[name="editHyojijun"],[name="editSuryoNm"],[name="editNounyuKbn"],[name="editBiko"]');
    function resolveMaker() {
        const matches = Array.from($('makerOptions').options).filter(o => o.value === $('seizomakerText').value);
        // 候補を一意に選べるときはID一致。直接入力・同名候補はメーカー名前方一致。
        $('seizomakerId').value = matches.length === 1 ? matches[0].dataset.id : '';
    }
    $('seizomakerText').addEventListener('input', resolveMaker);
    $('seizomakerText').addEventListener('change', resolveMaker);
    form.addEventListener('input', e => {if (isEdit(e.target)) edited = true;});
    form.addEventListener('change', e => {if (isEdit(e.target)) edited = true;});
    function dialog(text, confirm = false, title = '入力エラー') {
        window.McmCommon?.hideProgress(true);
        const modal = $('messageDialog'), previous = document.activeElement;
        $('dialogTitle').textContent = confirm ? '確認' : title;
        $('dialogMessage').textContent = text;
        const info = /必ず入力|桁以下|数値で|行が選択されていません/.test(text);
        $('dialogIcon').textContent = confirm ? '?' : info ? 'i' : '!';
        $('dialogIcon').style.background = confirm ? '#2764a2' : info ? '#1a6fd4' : '#b07800';
        $('dialogCancel').hidden = !confirm;
        return new Promise(resolve => {
            function done(value) {
                modal.close(); modal.removeEventListener('cancel', escape);
                $('dialogOk').onclick = $('dialogCancel').onclick = $('dialogClose').onclick = null;
                previous?.focus(); resolve(value);
            }
            function escape(event) {event.preventDefault(); done(false);}
            $('dialogOk').onclick = () => done(true);
            $('dialogCancel').onclick = $('dialogClose').onclick = () => done(false);
            modal.addEventListener('cancel', escape); modal.showModal(); $('dialogOk').focus();
        });
    }
    function clearSearch() {
        for (const tableId of ['atsukaikikiTable','partnerTable']) {
            const body = $(tableId).tBodies[0]; body.replaceChildren();
            const cell = body.insertRow().insertCell(); cell.colSpan = tableId === 'atsukaikikiTable' ? 6 : 3;
            cell.className = 'no-data'; cell.textContent = tableId === 'atsukaikikiTable' ? '検索結果はありません' : '取扱機器を選択してください';
        }
    }
    async function request(path, values) {
        values.set('editToken', $('editToken').value);
        const abort = new AbortController(), timer = setTimeout(() => abort.abort(), 30000);
        try {
            const response = await fetch(base + '/' + path, {method:'POST', body:values, signal:abort.signal});
            if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error('response');
            const result = await response.json();
            if (result.editToken) $('editToken').value = result.editToken;
            if (result.clearSearch) clearSearch();
            if (result.errors?.length) {await dialog(result.errors[0], false, result.title); return null;}
            return result;
        } finally {clearTimeout(timer);}
    }
    async function run(path, extra = {}) {
        if (busy) return;
        busy = true;
        resolveMaker();
        const values = new URLSearchParams(new FormData(form));
        Object.entries(extra).forEach(([k,v]) => values.set(k, v));
        const controls = Array.from(form.elements).filter(el => !el.disabled);
        controls.forEach(el => {el.disabled = true;});
        try {
            if (path === 'discard' && edited && !await dialog('データが変更されています。破棄されますがよろしいですか？', true)) return;
            if (path === 'deleteMeisaiRow') {
                if (!await request('validateDelete',values)) return;
                if (!await dialog('行を削除します。よろしいですか？',true)) return;
            }
            const result = await request(path,values);
            if (result?.redirect) {leaving = true; location.assign(result.redirect);}
        } catch (e) {
            await dialog(e.name === 'AbortError'
                ? '通信が30秒以内に完了しませんでした。入力内容を保持しています。接続を確認して再度操作してください。'
                : '処理を完了できませんでした。入力内容を保持しています。もう一度操作してください。');
        } finally {controls.forEach(el => {el.disabled = false;});busy = false;}
    }
    form.addEventListener('submit', e => {e.preventDefault(); run('settei');});
    $('btnSearch14').addEventListener('click', () => run('search'));
    document.querySelector('.search-area').addEventListener('keydown', e => {
        if (e.key === 'Enter') {e.preventDefault();run('search');}
    });
    $('btnClose14').addEventListener('click', () => run('discard'));
    document.querySelectorAll('.link-btn').forEach(btn => btn.addEventListener('click', () => run('selectAtsukaikiki',{atsukaikikiId:btn.dataset.equipmentId})));
    $('deleteSelectedRow').addEventListener('click', () => {
        const row = document.querySelector('.detail-row.selected-row');
        if (!row) {dialog('行が選択されていません。');return;}
        run('deleteMeisaiRow',{rowIndex:row.dataset.rowIndex});
    });
    const equipmentRows = Array.from(document.querySelectorAll('.equipment-row'));
    const partnerRows = Array.from(document.querySelectorAll('.partner-row'));
    const idKey = value => String(value || '').trim().replace(/\.0+$/, '').replace(/^0+(?=\d)/, '');
    function selectEquipment(row) {
        equipmentRows.forEach(r => r.classList.toggle('selected-row', r === row));
        let count = 0;
        partnerRows.forEach(r => {r.hidden = !row || idKey(r.dataset.equipmentId) !== idKey(row.dataset.equipmentId);if(!r.hidden)count++;});
        const empty = $('partnerEmpty');
        if (empty) {empty.hidden = count > 0;empty.cells[0].textContent = row ? '取引先情報はありません' : '取扱機器を選択してください';}
        document.querySelector('.partner-container').scrollTop = 0;
    }
    equipmentRows.forEach(row => {
        row.addEventListener('click', () => {if(!busy)selectEquipment(row);});
        row.addEventListener('keydown', e => {
            if (busy || e.target !== row) return;
            if (e.key === 'Enter' || e.key === ' ') {e.preventDefault();selectEquipment(row);}
            if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
                e.preventDefault();const next = equipmentRows[equipmentRows.indexOf(row)+(e.key === 'ArrowDown'?1:-1)];
                if (next) {next.focus();selectEquipment(next);}
            }
        });
    });
    const selected = $('atsukaikikiTable').dataset.selectedEquipment;
    selectEquipment(equipmentRows.find(r=>idKey(r.dataset.equipmentId)===idKey(selected)) || equipmentRows[0] || null);
    const detailRows = Array.from(document.querySelectorAll('.detail-row'));
    function selectDetail(row) {
        if(busy)return;
        detailRows.forEach(r => {r.classList.toggle('selected-row',r===row);r.querySelector('input[type=radio]').checked=r===row;});
        $('deleteSelectedRow').disabled=!row;
    }
    detailRows.forEach(row=>{
        row.addEventListener('click',()=>selectDetail(row));row.addEventListener('focusin',()=>selectDetail(row));
    });
    selectDetail(detailRows[0]||null);
    window.addEventListener('beforeunload',e=>{if(edited&&!leaving){e.preventDefault();e.returnValue='';}});
})();
