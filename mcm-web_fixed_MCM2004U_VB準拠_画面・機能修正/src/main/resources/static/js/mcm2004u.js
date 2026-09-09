/* MCM2004U: VB row selection/confirmation, Web MCM dialog. */
(() => {
    'use strict';
    const $ = id => document.getElementById(id), rows = [...document.querySelectorAll('.result-row')];
    let busy = false, anchor = 0;
    function lock(value) { busy = value; $('searchButton').disabled = value; if ($('deleteButton')) $('deleteButton').disabled = value; }
    function dialog(message, confirm = false, title = '入力エラー') {
        window.McmCommon?.hideProgress(true);
        const modal = $('messageDialog'), previous = document.activeElement;
        $('dialogTitle').textContent = confirm ? '確認' : title; $('dialogMessage').textContent = message;
        $('dialogIcon').textContent = confirm ? '?' : '!'; $('dialogIcon').style.background = confirm ? '#2764a2' : '#b07800';
        $('dialogCancel').hidden = !confirm;
        return new Promise(resolve => {
            function done(ok) { modal.close(); modal.removeEventListener('cancel', cancel); previous?.focus(); resolve(ok); }
            function cancel(e) { e.preventDefault(); done(false); }
            $('dialogOk').onclick = () => done(true); $('dialogCancel').onclick = $('dialogClose').onclick = () => done(false);
            modal.addEventListener('cancel', cancel); modal.showModal(); $('dialogOk').focus();
        });
    }
    function select(row, event = {}) {
        const index = rows.indexOf(row);
        if (event.shiftKey) {
            rows.forEach((r,i) => r.classList.toggle('selected', i >= Math.min(anchor,index) && i <= Math.max(anchor,index)));
        } else if (event.ctrlKey || event.metaKey) { row.classList.toggle('selected'); anchor = index; }
        else { rows.forEach(r => r.classList.toggle('selected',r===row)); anchor = index; }
        rows.forEach(r => r.setAttribute('aria-selected',r.classList.contains('selected')));
    }
    rows.forEach(row => {
        row.addEventListener('click',e => {if(!busy)select(row,e);});
        row.addEventListener('keydown',e => {
            if(busy || e.target !== row)return;
            if(e.key===' '){e.preventDefault();select(row,e);}
            if(e.key==='ArrowUp'||e.key==='ArrowDown'){
                e.preventDefault(); const next=rows[rows.indexOf(row)+(e.key==='ArrowUp'?-1:1)];
                if(next){select(next,e);next.focus();}
            }
        });
    });
    async function act(operation,row) {
        if(busy)return; lock(true);
        const selected = operation==='delete-row' ? rows.filter(r=>r.classList.contains('selected')) : [row];
        if(!selected.length){await dialog('行が選択されていません。');lock(false);return;}
        const message = operation==='delete-row'?'行を削除します。よろしいですか？':operation==='discard-mitsumori'?'見積を破棄します。よろしいですか？':'契約を破棄します。よろしいですか？';
        if(!await dialog(message,true)){lock(false);return;}
        const form=$('actionForm');form.querySelectorAll('[data-action-value]').forEach(e=>e.remove());
        const add=(name,value)=>{const input=document.createElement('input');input.type='hidden';input.name=name;input.value=value;input.dataset.actionValue='true';form.append(input);};
        [...new Set(selected.map(r=>r.dataset.estimate))].forEach(id=>add('umKihonMitsumoriId',id));
        if(operation!=='delete-row')add('ukKeiyakuId',row.dataset.contract||'');
        form.action=form.action.replace(/\/[^/]+$/, '/'+operation);form.submit();
    }
    document.querySelectorAll('[data-operation]').forEach(button=>button.addEventListener('click',e=>{e.stopPropagation();const row=button.closest('tr');if(!busy)select(row);act(button.dataset.operation,row);}));
    $('deleteButton')?.addEventListener('click',()=>act('delete-row'));
    $('searchForm').addEventListener('submit',e=>{if(busy)e.preventDefault();else lock(true);});
    document.querySelectorAll('a').forEach(link=>link.addEventListener('click',e=>{if(busy)e.preventDefault();}));
    window.addEventListener('pageshow',()=>lock(false));
    const errors=$('serverErrors');
    if(errors){lock(true);dialog([...errors.children].map(e=>e.textContent).join('\n'),false,errors.dataset.title).finally(()=>lock(false));}
})();
