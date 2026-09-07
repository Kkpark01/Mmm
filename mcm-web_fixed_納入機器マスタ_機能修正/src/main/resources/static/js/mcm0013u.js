/* MCM0013U: 入力は画面／セッションに保持し、登録時だけDBに反映する。 */
document.addEventListener('DOMContentLoaded', function () {
    'use strict';
    const $ = id => document.getElementById(id);
    const base = (document.querySelector('meta[name="contextPath"]')?.content || '').replace(/\/$/, '');
    const form = $('updateForm');
    let busy = false, leaving = false, edited = false;
    const original = new URLSearchParams(new FormData(form)).toString();
    const msg = {
        save: '登録を実施します。よろしいですか？', remove: '行を削除します。よろしいですか？',
        discard: 'データが変更されています。破棄されますがよろしいですか？',
        required: '必須項目を入力して下さい。', selected: '行が選択されていません。',
        min: 'セット数は1以上の値を入力して下さい。', max: 'セット数は4桁以下で入力してください。'
    };
    function dialog(text, confirm = false, title) {
        if (window.McmCommon) window.McmCommon.hideProgress(true);
        const overlay = $('miniDialogOverlay'), ok = $('miniDialogOk'), cancel = $('miniDialogCancel'), close = $('miniDialogCloseX');
        const previous = document.activeElement;
        $('miniDialogTitle').textContent = title || (confirm ? '確認' : text.includes('既に使用されている為') ? '削除エラー' : '入力エラー');
        $('miniDialogMsg').textContent = text;
        const icon = document.querySelector('.mini-icon');
        const information = title === '完了' || /必ず入力|桁以下|数値で|書式で|値が変更されていません|行が選択されていません/.test(text);
        icon.textContent = confirm ? '?' : information ? 'i' : '!';
        icon.style.cssText = 'display:inline-flex;align-items:center;justify-content:center;flex:0 0 32px;height:32px;border-radius:50%;font:bold 24px Arial;color:white;background:' + (confirm ? '#2764a2' : information ? '#1a6fd4' : '#b07800');
        cancel.hidden = !confirm; overlay.style.display = 'flex';
        return new Promise(resolve => {
            function done(value) {
                overlay.style.display = 'none'; ok.onclick = cancel.onclick = close.onclick = null;
                overlay.removeEventListener('keydown', keyboard); previous?.focus(); resolve(value);
            }
            function keyboard(e) {
                if (e.key === 'Escape') {e.preventDefault(); done(false);}
                if (e.key === 'Tab') {
                    const controls = confirm ? [close, ok, cancel] : [close, ok];
                    const at = controls.indexOf(document.activeElement);
                    e.preventDefault(); controls[(at + (e.shiftKey ? controls.length - 1 : 1)) % controls.length].focus();
                }
            }
            ok.onclick = () => done(true); cancel.onclick = close.onclick = () => done(false);
            overlay.addEventListener('keydown', keyboard); ok.focus();
        });
    }
    function data(extra = {}) {
        const values = new FormData(form);
        const csrf = document.querySelector('meta[name="_csrf"]')?.content;
        if (csrf) values.set(document.querySelector('meta[name="_csrf_parameter"]')?.content || '_csrf', csrf);
        Object.entries(extra).forEach(([key,value]) => values.set(key,value == null ? '' : value));
        return values;
    }
    function send(path, values) {
        leaving = true;
        const post = document.createElement('form'); post.method = 'post'; post.action = base + '/mcm0013u/' + path;
        for (const [name,value] of values) {const input = document.createElement('input');input.type='hidden';input.name=name;input.value=value;post.appendChild(input);}
        document.body.appendChild(post); post.submit();
    }
    function numberError() {
        for (const input of document.querySelectorAll('.set-nm-input,[name$=".hyojijun"]')) {
            if (!input.value.trim()) continue;
            const label = input.matches('.set-nm-input') ? 'セット数' : 'No';
            if (!/^-?\d+$/.test(input.value)) return label + 'は数値で入力してください。';
            if (input.matches('.set-nm-input') && Number(input.value) < 1) return msg.min;
            if (input.matches('.set-nm-input') && Number(input.value) > 9999) return msg.max;
            if (!input.matches('.set-nm-input') && input.value.replace(/^-/, '').length > 6) return 'Noは6桁以下で入力してください。';
        }
        return '';
    }
    function required(row) {
        return !row.querySelector('.kosei-nk-input').value.trim() || !row.querySelector('.set-nm-input').value || !row.querySelector('.tani-select').value;
    }
    async function request(path, values) {
        const abort = new AbortController();
        const timer = setTimeout(() => abort.abort(), 30000);
        try {
            const response = await fetch(base + '/mcm0013u/' + path, {method:'POST',body:values,signal:abort.signal});
            if (!response.ok || !response.headers.get('content-type')?.includes('application/json')) throw new Error('response');
            return await response.json();
        } finally {clearTimeout(timer);}
    }
    // 通信中の二重操作や入力変更を防ぎ、送信した値と画面を一致させる。
    async function run(task) {
        if (busy) return;
        const error = numberError(); if (error) {await dialog(error);return;}
        busy = true;
        const locked = Array.from(document.querySelectorAll('.section input,.section select,.section button,#btnUpdate')).filter(el=>!el.disabled);
        try {await task(() => locked.forEach(el=>el.disabled=true));}
        catch (e) {await dialog(e.name === 'AbortError'
            ? '通信が30秒以内に完了しませんでした。入力内容を保持しています。接続を確認して再度操作してください。'
            : '処理を完了できませんでした。入力内容を保持しています。もう一度操作してください。');}
        finally {locked.forEach(el=>el.disabled=false);busy=false;}
    }
    async function confirmAction(action, target, path, extra) {
        await run(async lock => {
            const values = data(extra); const check = data({action,targetId:target}); lock();
            const result = await request('validate',check);
            if (result.errors?.length) {await dialog(result.errors[0],false,result.title);return;}
            if (await dialog(action === 'update' ? msg.save : msg.remove,true)) send(path,values);
        });
    }
    form.addEventListener('submit', e => {e.preventDefault();confirmAction('update',null,'update',{});});
    $('btnAddKosei').addEventListener('click', () => run(async lock => {
        if (Array.from(document.querySelectorAll('#koseiTable tr[data-kosei-id]')).some(required)) {await dialog(msg.required);return;}
        const values=data();lock();send('addKoseiRow',values);
    }));
    $('btnDeleteKosei').addEventListener('click', () => {
        const id=document.querySelector('#koseiTable .selected-row')?.dataset.koseiId;
        if(!id){dialog(msg.selected);return;}confirmAction('deleteKosei',id,'deleteKoseiRow',{kikikoseiId:id});
    });
    let selectedKotaiId=null;
    document.querySelectorAll('.kotai-row').forEach(row => row.addEventListener('click', () => {
        if(busy)return;
        document.querySelectorAll('.selected-kotai-row').forEach(r=>r.classList.remove('selected-kotai-row'));
        row.classList.add('selected-kotai-row');selectedKotaiId=row.dataset.kotaiId;
    }));
    $('btnDeleteKotai').addEventListener('click', () => {
        if(!selectedKotaiId){dialog(msg.selected);return;}
        confirmAction('deleteKotai',selectedKotaiId,'deleteKotaiRow',{kotaikanriId:selectedKotaiId});
    });
    document.querySelectorAll('#koseiTable tr[data-kosei-id]').forEach(row => {
        row.addEventListener('click', e => {
            if(e.target.closest('a,input,select,button,textarea') || row.classList.contains('selected-row'))return;
            run(async lock=>{const values=data({kikikoseiId:row.dataset.koseiId});lock();send('selectKoseiRow',values);});
        });
        row.querySelector('.select-kosei-link').addEventListener('click',e=>{
            e.preventDefault();run(async lock=>{if(required(row)){await dialog(msg.required);return;}const values=data({kikikoseiId:row.dataset.koseiId});lock();send('goToMcm0014u',values);});
        });
    });
    document.querySelectorAll('.controller-flg-checkbox').forEach(cb=>cb.addEventListener('change',()=>{
        cb.closest('td').querySelector('.controller-flg-hidden').value=cb.checked?'1':'0';edited=true;
    }));
    document.querySelectorAll('.kosei-edit,.kotai-edit').forEach(el=>el.addEventListener('input',()=>edited=true));
    document.querySelectorAll('.set-nm-input').forEach(input=>input.addEventListener('change',()=>run(async lock=>{
        if(!input.value){await dialog(msg.required);return;}
        const values=data({kikikoseiId:input.closest('tr').dataset.koseiId,setNm:input.value});lock();
        const result=await request('syncKotaiForSetNum',values);
        if(result.errors?.length){await dialog(result.errors[0],false,result.title);return;}
        leaving=true;location.reload();
    })));
    function dirty(){return $('mcmHasChanges')?.value==='true'||edited||new URLSearchParams(new FormData(form)).toString()!==original;}
    for(const [id,destination] of [['btnBackLink','back'],['btnCloseLink','close']]) $(id).addEventListener('click',async e=>{
        e.preventDefault();if(busy)return;
        if(!dirty()||await dialog(msg.discard,true))send('discard',data({destination}));
    });
    window.addEventListener('beforeunload',e=>{if(!leaving&&dirty()){e.preventDefault();e.returnValue='';}});
    const success=$('successMsg')?.textContent.trim();
    const error=$('errorMsg')?.querySelector('li')?.textContent.trim();
    if(error)dialog(error);else if(success)dialog(success,false,'完了');
});
