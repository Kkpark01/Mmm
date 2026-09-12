(() => {
    'use strict';
    const form = document.querySelector('.mcm2007u-step2 #equipmentForm');
    if (!form) return;
    const rows = kind => [...form.querySelectorAll(`#${kind}Table tbody tr`)];
    const checks = kind => [...form.querySelectorAll(`.chk-${kind}`)];
    const key = value => String(value ?? '').replace(/\.0+$/, '');
    const same = (a, b, prop) => key(a.dataset[prop]) === key(b.dataset[prop]);
    const family = (source, kind) => checks(kind).filter(other => same(source, other, 'brand') && same(source, other, 'kosei') && (kind === 'kosei' || !source.dataset.meisai || same(source, other, 'meisai')));
    const quoted = check => check.dataset.quoted === 'true';
    const set = (check, value) => { check.checked = value && quoted(check); };
    const brandInput = form.querySelector('#selectedBrandIndex');
    const koseiInput = form.querySelector('#selectedKoseiIndex');
    const remembered = new Map();

    function refreshChecks() {
        ['kosei', 'meisai', 'kotai'].forEach(kind => {
            checks(kind).forEach(check => {
                check.disabled = !quoted(check) && !check.checked;
                check.closest('tr').classList.toggle('unquoted', !quoted(check));
            });
            const eligible = checks(kind).filter(check => !check.closest('tr').hidden && (quoted(check) || check.checked));
            const head = form.querySelector(`.check-visible[data-kind="${kind}"]`);
            head.disabled = eligible.length === 0;
            head.checked = eligible.length > 0 && eligible.every(check => check.checked);
            head.indeterminate = eligible.some(check => check.checked) && !head.checked;
        });
    }
    function syncSelection(check) {
        if (check.checked && !quoted(check)) check.checked = false;
        if (check.classList.contains('chk-kosei')) {
            family(check, 'meisai').forEach(child => {
                set(child, check.checked);
                family(child, 'kotai').forEach(individual => set(individual, child.checked));
            });
        } else {
            if (check.classList.contains('chk-meisai')) {
                family(check, 'kotai').forEach(child => set(child, check.checked));
            } else {
                family(check, 'meisai').forEach(parent => set(parent, family(parent, 'kotai').some(child => child.checked)));
            }
            family(check, 'kosei').forEach(parent => set(parent, family(parent, 'meisai').some(child => child.checked)));
        }
        refreshChecks();
    }
    function showKosei(index) {
        const current = rows('kosei').find(row => Number(row.dataset.index) === index && !row.hidden);
        koseiInput.value = current ? String(index) : '-1';
        if (current) remembered.set(brandInput.value, index);
        rows('kosei').forEach(row => row.classList.toggle('selected', row === current));
        ['meisai', 'kotai'].forEach(kind => {
            rows(kind).forEach(row => { row.hidden = !current || !same(current, row, 'brand') || !same(current, row, 'kosei'); });
            form.querySelector(`#${kind}Empty`).hidden = rows(kind).some(row => !row.hidden);
        });
        refreshChecks();
    }
    function showBrand(index, preferredKosei) {
        const current = rows('brand').find(row => Number(row.dataset.index) === index);
        brandInput.value = current ? String(index) : '-1';
        rows('brand').forEach(row => { row.classList.toggle('selected', row === current); row.setAttribute('aria-current', String(row === current)); });
        rows('kosei').forEach(row => { row.hidden = !current || !same(current, row, 'brand'); });
        const available = rows('kosei').filter(row => !row.hidden);
        form.querySelector('#brandEmpty').hidden = !!current;
        form.querySelector('#koseiEmpty').hidden = available.length > 0;
        const desired = preferredKosei ?? remembered.get(brandInput.value);
        const selected = available.find(row => Number(row.dataset.index) === desired) || available.find(row => row.querySelector('input').checked) || available.find(row => quoted(row.querySelector('input'))) || available[0];
        showKosei(selected ? Number(selected.dataset.index) : -1);
    }
    ['brand', 'kosei'].forEach(kind => rows(kind).forEach(row => {
        const show = target => kind === 'brand' ? showBrand(Number(target.dataset.index)) : showKosei(Number(target.dataset.index));
        row.addEventListener('click', () => show(row));
        row.addEventListener('keydown', event => {
            if (event.target !== row) return;
            if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); show(row); }
            if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                event.preventDefault(); const visible = rows(kind).filter(r => !r.hidden);
                const next = visible[Math.max(0, Math.min(visible.length - 1, visible.indexOf(row) + (event.key === 'ArrowDown' ? 1 : -1)))];
                if (next) { show(next); next.focus(); }
            }
        });
    }));
    ['kosei', 'meisai', 'kotai'].forEach(kind => checks(kind).forEach(check => check.addEventListener('change', () => syncSelection(check))));
    form.querySelectorAll('.check-visible').forEach(head => head.addEventListener('change', () => {
        const value = head.checked;
        checks(head.dataset.kind).filter(check => !check.closest('tr').hidden && !check.disabled).forEach(check => { set(check, value); syncSelection(check); });
        refreshChecks();
    }));
    const brands = rows('brand'); let initial = Number(brandInput.value);
    if (!brands.some(row => Number(row.dataset.index) === initial)) {
        const checked = checks('kosei').find(check => check.checked);
        initial = Number((brands.find(row => checked && same(row, checked, 'brand')) || brands[0])?.dataset.index ?? -1);
    }
    showBrand(initial, Number(koseiInput.value));
})();
