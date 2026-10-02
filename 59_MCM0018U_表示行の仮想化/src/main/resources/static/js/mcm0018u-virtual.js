/* MCM0018U専用。検索結果はデータとして保持し、表示範囲だけ入力要素を作る。 */
(function () {
    'use strict';
    window.createMcm0018VirtualGrid = function (tbody, data, selectedIndex) {
        if (!tbody || !Array.isArray(data) || data.length <= 200) return null;
        const table = tbody.closest('table'), scroller = table.closest('.grid-scroll-parent');
        const prototype = tbody.querySelector('tr').cloneNode(true);
        const height = 23, overscan = 8, mounted = new Map(), records = new Map();
        const fields = Array.from(prototype.querySelectorAll('[name]'), el => ({
            key: el.name.split('.').pop(), checkbox: el.type === 'checkbox', money: el.classList.contains('money-input')
        }));
        const labels = {};
        for (const [key, id] of Object.entries({seizomakerId: 'searchSeizomakerId', kikibunruiId: 'searchKikibunruiId'})) {
            labels[key] = new Map(Array.from(document.getElementById(id).options, o => [o.value.replace(/\.0+$/, ''), o.textContent]));
        }
        let visible = [], positions = new Map(), frame = null, nextIndex = data.length;
        const text = v => v == null ? '' : String(v);
        function clean(v) {
            const s = text(v);
            if (/^[\s\u3000]*(?:<<|<|«|\()?\s*(?:NULL|ヌル)\s*(?:>>|>|»|\))?\s*$/i.test(s)) return '';
            return s.replace(/(?:<<|«|＜＜|‹‹)\s*NULL\s*(?:>>|»|＞＞|››)/gi, '');
        }
        function placeholder(row) { return row.atsukaikikiId == null && (!row.rowStatus || row.rowStatus === 'unchanged'); }
        data.forEach((row, index) => {
            const values = {};
            Object.entries(row).forEach(([key, value]) => { values[key] = clean(value); });
            records.set(index, {index, data: values, placeholder: placeholder(row), removed: false});
        });
        // JSONを別の大きな配列として二重保持しない。
        window.__atsukaikikiRows = null;
        data = null;
        function rebuildPositions() {
            visible = Array.from(records.values()).filter(r => !r.removed && r.data.rowStatus !== 'deleted');
            positions = new Map(visible.map((r, i) => [r.index, i]));
            tbody.style.height = (visible.length * height) + 'px';
            table.setAttribute('aria-rowcount', visible.length + 1);
        }
        function createRow(record, reusable) {
            const tr = reusable || prototype.cloneNode(true), d = record.data;
            tr.classList.remove('row-active', 'selected-row', 'new-row');
            tr.removeAttribute('style');
            tr.dataset.parentIndex = record.index;
            tr.dataset.parentId = text(d.atsukaikikiId);
            if (record.placeholder) { tr.dataset.isNew = 'true'; tr.classList.add('new-row'); }
            else tr.removeAttribute('data-is-new');
            tr.setAttribute('aria-rowindex', (positions.get(record.index) ?? record.index) + 2);
            tr.querySelectorAll('[name]').forEach(el => {
                const key = el.name.split('.').pop(), value = text(d[key]);
                el.name = 'atsukaikikiRows[' + record.index + '].' + key;
                el.classList.remove('field-error');
                el.removeAttribute('data-auto-no');
                if (el.type === 'checkbox') el.checked = value === '1';
                else if (el.tagName === 'SELECT') {
                    const normalized = value.replace(/\.0+$/, '');
                    el.replaceChildren(new Option('', ''));
                    if (normalized) el.add(new Option(labels[key]?.get(normalized) || text(d[key.replace(/Id$/, 'Nk')]), normalized, true, true));
                    el.value = normalized;
                } else {
                    el.value = value;
                    if (el.matches('input.edit-input')) el.readOnly = true;
                    if (el.classList.contains('money-input') && /^\d+$/.test(value)) el.value = '￥' + value.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
                }
            });
            tr.querySelectorAll('.cell-selected, .cell-error').forEach(c => c.classList.remove('cell-selected', 'cell-error'));
            const dates = tr.querySelectorAll('.col-dt.readonly'), users = tr.querySelectorAll('.col-user.readonly');
            ['createdDt', 'lastupdateDt'].forEach((key, i) => { if (dates[i]) dates[i].textContent = text(d[key]); });
            ['createdBy', 'lastupdateBy'].forEach((key, i) => { if (users[i]) users[i].textContent = text(d[key]); });
            return tr;
        }
        function recordFor(tr) { return records.get(Number(tr.dataset.parentIndex)); }
        function save(tr) {
            const record = recordFor(tr); if (!record) return;
            tr.querySelectorAll('[name]').forEach(el => {
                const key = el.name.split('.').pop();
                record.data[key] = el.type === 'checkbox' ? (el.checked ? '1' : '0') : el.value;
            });
            record.placeholder = tr.dataset.isNew === 'true';
        }
        function sync() { mounted.forEach(save); }
        function render() {
            if (frame !== null) { cancelAnimationFrame(frame); frame = null; }
            const start = Math.max(0, Math.floor(scroller.scrollTop / height) - overscan);
            const end = Math.min(visible.length, Math.ceil((scroller.scrollTop + scroller.clientHeight) / height) + overscan);
            const needed = new Set(visible.slice(start, end).map(r => r.index));
            // 編集中・選択中の1行は残す。フォーカス・必須検査・親子の対応を失わない。
            const active = window.Mcm0018Performance.activeRow(tbody);
            const focused = document.activeElement?.closest('tr');
            for (const row of [active, focused]) {
                if (row && row.parentElement === tbody && positions.has(Number(row.dataset.parentIndex))) needed.add(Number(row.dataset.parentIndex));
            }
            for (const [index, tr] of mounted) {
                if (!needed.has(index)) { save(tr); tr.remove(); mounted.delete(index); }
            }
            const ordered = Array.from(needed).sort((a, b) => positions.get(a) - positions.get(b));
            let anchor = tbody.firstElementChild;
            for (const index of ordered) {
                let tr = mounted.get(index);
                if (!tr) { tr = createRow(records.get(index)); mounted.set(index, tr); }
                tr.style.top = (positions.get(index) * height) + 'px';
                tr.setAttribute('aria-rowindex', positions.get(index) + 2);
                if (tr !== anchor) tbody.insertBefore(tr, anchor);
                else anchor = anchor.nextElementSibling;
            }
        }
        function schedule() { if (frame === null) frame = requestAnimationFrame(render); }
        function ensure(index) {
            const record = records.get(Number(index));
            if (!record || !positions.has(record.index)) return null;
            const top = positions.get(record.index) * height;
            const viewport = scroller.clientHeight - table.tHead.offsetHeight;
            if (top < scroller.scrollTop) scroller.scrollTop = top;
            else if (top + height > scroller.scrollTop + viewport) scroller.scrollTop = top + height - viewport;
            render();
            return mounted.get(record.index);
        }
        function previous(tr) {
            const pos = positions.get(Number(tr.dataset.parentIndex));
            if (pos == null || pos === 0) return null;
            const record = visible[pos - 1];
            return mounted.get(record.index) || createRow(record);
        }
        function revealInput(input) {
            const row = input?.closest('tr'); if (!row) return input;
            const live = ensure(row.dataset.parentIndex);
            return live ? Array.from(live.querySelectorAll('[name]')).find(el => el.name === input.name) : input;
        }
        tbody.replaceChildren();
        tbody.classList.add('virtual-rows');
        window.Mcm0018Performance.resetGrid(tbody);
        rebuildPositions(); render();
        const selected = records.get(Number(selectedIndex));
        if (selected && !selected.placeholder && positions.has(selected.index)) {
            const tr = mounted.get(selected.index) || createRow(selected);
            if (!mounted.has(selected.index)) { mounted.set(selected.index, tr); tbody.appendChild(tr); tr.style.top = positions.get(selected.index) * height + 'px'; }
            window.Mcm0018Performance.markActive(tr);
        }
        scroller.addEventListener('scroll', schedule, {passive: true});
        new ResizeObserver(schedule).observe(scroller);
        return {
            render, sync, previous, revealInput,
            hasChanges() {
                sync();
                return Array.from(records.values()).some(r => !r.removed && (['modified', 'deleted'].includes(r.data.rowStatus) || (!r.placeholder && r.data.rowStatus === 'added')));
            },
            forEachRow(callback) {
                sync();
                // 全件検査では画面外用の1行だけ再利用し、数千行の入力要素を再作成しない。
                let scratch = null;
                for (const record of records.values()) {
                    if (record.removed) continue;
                    let row = mounted.get(record.index);
                    if (!row) { scratch = createRow(record, scratch); row = scratch; }
                    if (callback(row) === false) break;
                }
            },
            payload() {
                sync();
                return Array.from(records.values()).filter(r => !r.removed && !r.placeholder).map(r => {
                    const row = {};
                    fields.forEach(field => {
                        let value = text(r.data[field.key]);
                        if (field.checkbox) value = value === '1' ? '1' : '0';
                        if (field.money) value = value.replace(/[^\d-]/g, '');
                        row[field.key] = value === '' || value === '-' ? null : value;
                    });
                    return row;
                });
            },
            neighbor(tr, direction) {
                const pos = positions.get(Number(tr.dataset.parentIndex)), next = visible[pos + direction];
                return next ? ensure(next.index) : null;
            },
            byId(id) {
                const record = visible.find(r => text(r.data.atsukaikikiId) === text(id));
                return record ? ensure(record.index) : null;
            },
            append(tr, clone) {
                save(tr);
                clone.dataset.parentIndex = nextIndex;
                clone.dataset.parentId = '';
                clone.querySelectorAll('[name]').forEach(el => { el.name = el.name.replace(/atsukaikikiRows\[\d+\]/, 'atsukaikikiRows[' + nextIndex + ']'); });
                records.set(nextIndex, {index: nextIndex, data: {}, placeholder: true, removed: false});
                mounted.set(nextIndex, clone); save(clone); nextIndex++;
                tbody.appendChild(clone); rebuildPositions(); render();
            },
            remove(tr, added) {
                save(tr);
                const record = recordFor(tr); if (!record) return;
                if (added) record.removed = true;
                else record.data.rowStatus = 'deleted';
                window.Mcm0018Performance.markSelected(tbody, null);
                mounted.delete(record.index); tr.remove(); rebuildPositions(); render();
            }
        };
    };
})();
