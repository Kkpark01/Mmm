(() => {
    'use strict';
    const form = document.querySelector('.mcm2007u-step1 #selectionForm');
    if (!form) return;
    const rows = [...form.querySelectorAll('#estimateTable tbody tr')];
    const current = form.querySelector('#selectedEstimateIndex');
    const panels = [...form.querySelectorAll('[data-estimate-panel]')];
    const brandsFor = index => form.querySelectorAll(`#brandTable tbody[data-estimate-panel="${index}"] input[name="brandSelection"]`);

    function show(index) {
        current.value = String(index);
        rows.forEach((row, i) => {
            row.classList.toggle('selected', i === index);
            row.setAttribute('aria-current', String(i === index));
        });
        panels.forEach(panel => { panel.hidden = Number(panel.dataset.estimatePanel) !== index; });
        for (const [table, empty] of [['brandTable', 'brandEmpty'], ['previewTable', 'previewEmpty']]) {
            form.querySelector('#' + empty).hidden = !!form.querySelector(`#${table} tbody:not([hidden]) tr`);
        }
    }
    rows.forEach((row, index) => {
        const check = row.querySelector('.estimate-check');
        row.addEventListener('click', () => show(index));
        row.addEventListener('keydown', event => {
            if (event.target !== row) return;
            if (event.key === 'Enter' || event.key === ' ') {
                event.preventDefault(); show(index);
            } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                event.preventDefault();
                const next = Math.max(0, Math.min(rows.length - 1, index + (event.key === 'ArrowDown' ? 1 : -1)));
                show(next); rows[next].focus();
            }
        });
        // VBの見積チェックは配下の全ブランドに連動。表示行の切替だけでは選定状態を変えない。
        check.addEventListener('change', () => {
            brandsFor(index).forEach(brand => { brand.checked = check.checked; });
            show(index);
        });
    });
    let index = Number(current.value);
    if (!Number.isInteger(index) || index < 0 || index >= rows.length) {
        index = rows.findIndex(row => row.querySelector('.estimate-check').checked);
        if (index < 0 && rows.length) index = 0;
    }
    show(index);
    // 日付欄のEnterで意図せずStep2へ進まない。
    form.addEventListener('keydown', event => {
        if (event.key === 'Enter' && event.target instanceof HTMLInputElement && event.target.type === 'text') event.preventDefault();
    });
})();
