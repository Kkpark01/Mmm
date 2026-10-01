(function () {
    var changeStatusField = document.getElementById('changeStatusField');
    var discardConfirmedFld = document.getElementById('discardConfirmed');
    var searchForm = document.getElementById('searchForm');
    var updateBtn = document.getElementById('btnUpdate');
    var updateForm = document.getElementById('updateForm');
    var atsukaikikiTbody = document.getElementById('atsukaikikiTbody');
    var koseiTbody = document.getElementById('koseiTbody');
    // 候補の原本は、画面行の増減やエラーに影響されないよう保持する。
    var koseiRowTemplate = koseiTbody && koseiTbody.querySelector('tr[data-is-new="true"]');
    koseiRowTemplate = koseiRowTemplate ? koseiRowTemplate.cloneNode(true) : null;
    var selectedAtsukaikikiIdField = document.getElementById('selectedAtsukaikikiIdField');
    var btnDeleteAtsukaikiki = document.getElementById('btnDeleteAtsukaikiki');
    var btnDeleteKosei = document.getElementById('btnDeleteKosei');
    var tempIdCounter = 0;
    window.__activeNewParentTempId = null;
    window.__pendingKoseiChanges = {};
    window.__currentParentKey = null; // "id:X" or "tmp:X"

    /*
     * ===== 金額表示（現行MCM準拠：「￥」付きカンマ区切り表示） =====
     * 【修正】従来 controllerKin は type="number" のため「￥」等のプレフィックス文字を
     *   表示できなかった。type="text" に変更し、非編集時は「￥1,000」形式で表示、
     *   フォーカス時は数値のみに戻す（MCM0015U money-input と同一パターン）。
     *   送信時は moneyDigits() で「￥」とカンマを除去し、サーバー側 BigDecimal への
     *   デシリアライズエラーを防ぐ。
     */
    function moneyDigits(s) { return (s == null ? "" : String(s)).replace(/[^\d-]/g, ""); }
    function isNumericMoneyValue(v) {
        var s = String(v == null ? "" : v).trim().replace(/[,￥]/g, "");
        return s === "" || /^[+-]?\d+$/.test(s);
    }
    function formatYen(s) {
        var d = moneyDigits(s);
        if (d === "" || d === "-") return "";
        var neg = d.charAt(0) === "-"; if (neg) d = d.slice(1);
        d = d.replace(/^0+(?=\d)/, "");
        return "￥" + (neg ? "-" : "") + d.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    }
    function applyYenAll() {
        document.querySelectorAll("#updateForm .money-input").forEach(function (inp) {
            if (inp.value !== "" && isNumericMoneyValue(inp.value)) inp.value = formatYen(inp.value);
        });
    }
    document.addEventListener("focusin", function (e) {
        if (e.target.classList && e.target.classList.contains("money-input") && isNumericMoneyValue(e.target.value)) {
            e.target.value = moneyDigits(e.target.value);
        }
    });
    document.addEventListener("focusout", function (e) {
        if (e.target.classList && e.target.classList.contains("money-input") && isNumericMoneyValue(e.target.value)) {
            e.target.value = formatYen(e.target.value);
        }
        /*
         * ★セル移動時（フォーカスアウト）にも桁数・数値チェックを実施する。
         *   従来は登録ボタン押下時のみチェックしており、UPS交換周期のように
         *   セル移動時点でのエラー表示が仕様上求められる項目に対応できていなかった。
         *   checkCellLength() は登録前チェックと同一ロジックを再利用する。
         */
        if (e.target.classList && e.target.classList.contains("edit-input") && e.target.name) {
            var tr = e.target.closest("tr");
            var st = tr && tr.querySelector(".row-status-input");
            if (st && String(st.value || "") === "deleted") return;
            var msg = checkCellLength(e.target);
            if (!msg && tr) {
                // 桁数・数値チェックに問題なければ、行単位の相関チェック（コントローラ⇔金額、
                // 契約可能期間⇔起算日区分）もセル移動時に行う。対象項目のセルでのみ判定する。
                if (/\.(controllerFlg|controllerKin)$/.test(e.target.name)) {
                    msg = checkAtsukaikikiRowConsistency(tr);
                } else if (/\.(keiyakukanokikan|kisanbiKbn)$/.test(e.target.name)) {
                    msg = checkKoseiRowConsistency(tr);
                }
            }
            if (msg) {
                e.target.classList.add("field-error");
                customAlert(msg, 'input');
            } else {
                e.target.classList.remove("field-error");
            }
        }
    });
    document.addEventListener("DOMContentLoaded", applyYenAll);
    if (document.readyState !== "loading") applyYenAll();

    // ========== Dialog helpers ==========
    var ICONS = {
        warning: '<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48"><path fill="#f9a825" d="M24 4 L44 42 L4 42 Z" stroke="#7a5a00" stroke-width="1.5" stroke-linejoin="round"/><rect x="22" y="18" width="4" height="14" fill="#fff"/><rect x="22" y="34" width="4" height="4" fill="#fff"/></svg>',
        info: '<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48"><circle cx="24" cy="24" r="20" fill="#2196f3"/><text x="24" y="35" font-family="Georgia,serif" font-size="34" font-weight="bold" font-style="italic" text-anchor="middle" fill="#fff">i</text></svg>',
        question: '<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48"><circle cx="24" cy="24" r="20" fill="#1976d2"/><text x="24" y="34" font-family="sans-serif" font-size="30" font-weight="bold" text-anchor="middle" fill="#fff">?</text></svg>',
        success: '<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 48 48"><circle cx="24" cy="24" r="20" fill="#2196f3"/><text x="24" y="35" font-family="Georgia,serif" font-size="34" font-weight="bold" font-style="italic" text-anchor="middle" fill="#fff">i</text></svg>'
    };
    var DIALOG_CONFIG = {
        search:  { title: '検索エラー', icon: 'warning' },
        input:   { title: '入力エラー', icon: 'info' },
        warning: { title: '入力エラー', icon: 'warning' },
        confirm: { title: '確認',       icon: 'question' },
        success: { title: '完了',       icon: 'success' },
        info:    { title: '情報',       icon: 'info' }
    };
    function showCustomDialog(type, message, onOk, onCancel) {
        var overlay = document.getElementById('customDialogOverlay');
        if (!overlay) {
            if (type === 'confirm') { if (confirm(message)) { if(onOk) onOk(); } else { if(onCancel) onCancel(); } }
            else { alert(message); if(onOk) onOk(); }
            return;
        }
        var titleEl = document.getElementById('customDialogTitle');
        var iconEl = document.getElementById('customDialogIcon');
        var msgEl = document.getElementById('customDialogMessage');
        var okBtn = document.getElementById('customDialogOkBtn');
        var cancelBtn = document.getElementById('customDialogCancelBtn');
        var closeX = document.getElementById('customDialogCloseX');
        var config = DIALOG_CONFIG[type] || DIALOG_CONFIG.info;
        titleEl.textContent = config.title;
        iconEl.innerHTML = ICONS[config.icon];
        msgEl.textContent = message;
        cancelBtn.style.display = (type === 'confirm') ? 'inline-block' : 'none';
        overlay.style.display = 'flex';
        var handled = false;
        function cleanup() {
            overlay.style.display = 'none';
            okBtn.onclick = null; cancelBtn.onclick = null; closeX.onclick = null;
        }
        okBtn.onclick = function () { if (handled) return; handled = true; cleanup(); if (onOk) onOk(); };
        cancelBtn.onclick = function () { if (handled) return; handled = true; cleanup(); if (onCancel) onCancel(); };
        closeX.onclick = function () { if (handled) return; handled = true; cleanup(); if (onCancel) onCancel(); };
        setTimeout(function () { okBtn.focus(); }, 100);
    }
    function customConfirm(message, onOk, onCancel) { showCustomDialog('confirm', message, onOk, onCancel); }
    function customAlert(message, type, onOk) { showCustomDialog(type || 'info', message, onOk); }

    // ========== Server messages on load ==========
    (function showServerMessagesOnLoad() {
        var errorsContainer = document.getElementById('serverErrors');
        var successEl = document.getElementById('serverSuccessMessage');
        var lines = [];
        var errorType = 'input';
        if (errorsContainer) {
            errorType = errorsContainer.getAttribute('data-error-type') || 'input';
            var items = errorsContainer.querySelectorAll('.server-error-item');
            for (var i = 0; i < items.length; i++) {
                var text = items[i].textContent.trim();
                if (text) lines.push(text);
            }
        }
        if (lines.length > 0) {
            setTimeout(function () { customAlert(lines.join('\n'), errorType); }, 100);
        } else if (successEl) {
            var successMsg = successEl.textContent.trim();
            var msgType = successEl.getAttribute('data-message-type') || 'success';
            if (successMsg) setTimeout(function () { customAlert(successMsg, msgType); }, 100);
        }
    })();

    function generateTempId() { return Date.now() * 1000 + (++tempIdCounter); }

    // ========== Part 3.2.2: 変更検知強化（既存行編集時 rowStatus="modified" 自動セット） ==========
    function bindChangeDetection() {
        function markChanged() { if (changeStatusField) changeStatusField.value = 'true'; }

        // 既存行が編集されたら rowStatus="modified" を自動セット
        function markRowModified(e) {
            var tr = e.target.closest('tr');
            if (!tr) return;
            // 空の placeholder（新規行）はスキップ
            if (tr.getAttribute('data-is-new') === 'true') return;
            var rowStatusEl = tr.querySelector('.row-status-input');
            if (!rowStatusEl) return;
            // "unchanged" のときのみ "modified" にする（added/deleted は上書きしない）
            var currentStatus = rowStatusEl.value;
            if (currentStatus === 'unchanged' || currentStatus === '' || currentStatus === null) {
                rowStatusEl.value = 'modified';
            }
        }

        [atsukaikikiTbody, koseiTbody].forEach(function (tb) {
            if (!tb) return;
            tb.addEventListener('input', function (e) {
                /*
                 * ★No欄は数値（半角0-9）以外を入力自体できないようにする。
                 *   キーボード入力・貼り付け・ドラッグ&ドロップ・IME確定後のいずれも
                 *   input イベントで捕捉されるため、ここで半角数字以外を即時除去する。
                 */
                if (e.target.name && /\.hyojijun$/.test(e.target.name)) {
                    var filtered = e.target.value.replace(/[^0-9]/g, "");
                    if (filtered !== e.target.value) e.target.value = filtered;
                }
                markChanged();
                markRowModified(e);
            });
            tb.addEventListener('change', function (e) {
                markChanged();
                markRowModified(e);
            });
        });
    }

    if (searchForm) {
        searchForm.addEventListener('submit', function (e) {
            var changed = changeStatusField && changeStatusField.value === 'true';
            if (changed) {
                /* 未保存変更は確認後に検索する。VB準拠で検索中の待機画面は表示しない。 */
                e.preventDefault();
                customConfirm('データが変更されています。破棄されますがよろしいですか？', function () {
                    if (discardConfirmedFld) discardConfirmedFld.value = 'true';
                    window.Mcm0018Performance.submitSearch(searchForm);
                });
            } else {
                /* 変更なし時は待機画面を表示せず検索する。 */
                if (discardConfirmedFld) discardConfirmedFld.value = 'false';
                e.preventDefault();
                window.Mcm0018Performance.submitSearch(searchForm);
            }
        });
    }

    // ========== Utils ==========
    /* index.html の共有候補処理と同じ正規化（BigDecimal の "123.0" 等を "123" に揃える） */
    function normalizeOptionKey(v) { return String(v == null ? '' : v).replace(/\.0+$/, ''); }
    /* 取引先 ID→名称。原本select(#mcm0018TorihikisakiSource)から初回だけ作って使い回す。 */
    var torihikisakiLabels = null;
    function getTorihikisakiLabels() {
        if (torihikisakiLabels) return torihikisakiLabels;
        torihikisakiLabels = {};
        var src = document.getElementById('mcm0018TorihikisakiSource');
        if (src) {
            for (var i = 0; i < src.options.length; i++) {
                var op = src.options[i];
                if (op.value !== '') torihikisakiLabels[normalizeOptionKey(op.value)] = op.textContent;
            }
        }
        return torihikisakiLabels;
    }
    /*
     * 【性能改善】子グリッド再描画の後処理。
     *   ・保持している ▶/行選択の状態を捨てる（行を作り直したため）
     *   ・再描画で発生した MutationRecord を破棄し、行ごとの遅延スイープを起こさない
     *   ・ヌル表記の除去は子グリッドに1回だけ。option は除去済みの原本由来のため走査しない
     */
    var nullSweepObserver = null;
    function finishKoseiRender() {
        window.Mcm0018Performance.resetGrid(koseiTbody);
        if (nullSweepObserver) nullSweepObserver.takeRecords();
        if (koseiTbody) sweepNullLiterals(koseiTbody, true);
    }
    function escapeHtml(s) {
        if (s === null || s === undefined) return '';
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    function escapeAttr(v) {
        if (v === null || v === undefined) return '';
        return String(v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }
    /* [修正] 備考などに紛れ込む "<<NULL>>" 等のヌル表記を空文字にする */
    function nzText(v) {
        if (v === null || v === undefined) return '';
        var t = String(v).trim();
        return (/^(<<NULL>>|<NULL>|\(NULL\)|NULL)$/i.test(t)) ? '' : v;
    }
    function getRowStatus(tr) {
        var el = tr.querySelector('.row-status-input');
        return el ? el.value : '';
    }

    // ========== Capture / Restore kosei rows (Case A) ==========
    function captureCurrentKoseiRows(parentKey) {
        if (!koseiTbody || !parentKey) return;
        var rows = koseiTbody.querySelectorAll('tr');
        var captured = [];
        rows.forEach(function (tr) {
            if (tr.getAttribute('data-is-new') === 'true') return; // skip placeholder
            /*
             * 【修正】従来は rowStatus が added/modified/deleted の行のみをキャプチャし、
             *   unchanged（未編集の既存行）を除外していた。しかし復元側の
             *   renderKoseiFromPending() はキャプチャした配列だけで子グリッドを丸ごと
             *   再構築するため、unchanged行を除外すると親を切り替えて戻った際に
             *   「新規追加行だけ表示され、既存の未変更データが消える」不具合が発生していた。
             *   表示されている全行（unchanged含む）をキャプチャし、完全なスナップショットを
             *   保持することで、復元時にデータが欠落しないようにする。
             */
            var data = {};
            // Collect all inputs and selects
            tr.querySelectorAll('[name]').forEach(function (el) {
                var m = el.getAttribute('name').match(/koseiRows\[\d+\]\.(.+)/);
                if (!m) return;
                var field = m[1];
                var v;
                if (el.type === 'checkbox') { v = el.checked ? el.value : ''; }
                else { v = el.value; }
                data[field] = v;
            });
            // 表示用フィールド（既存行のtorihikisakiNk等）
            var tCell = tr.querySelector('td.col-torihikisaki');
            if (tCell && !tCell.querySelector('select')) {
                data.torihikisakiNk = tCell.textContent.trim();
            }
            var dtCells = tr.querySelectorAll('td.col-dt.readonly');
            var userCells = tr.querySelectorAll('td.col-user.readonly');
            if (dtCells.length >= 2) {
                data.createdDt = dtCells[0].textContent.trim();
                data.lastupdateDt = dtCells[1].textContent.trim();
            }
            if (userCells.length >= 2) {
                data.createdBy = userCells[0].textContent.trim();
                data.lastupdateBy = userCells[1].textContent.trim();
            }
            captured.push(data);
        });
        if (captured.length > 0) {
            window.__pendingKoseiChanges[parentKey] = captured;
        } else {
            delete window.__pendingKoseiChanges[parentKey];
        }
    }

    // ========== Parent row click ==========
    if (atsukaikikiTbody) {
        /*
         * 製造メーカー・機器分類(select)セルは、ブラウザ仕様と2クリック編集（mousedownでの
         * preventDefault／ネイティブ一覧の展開）により click が発火しないことがあり、
         * 取引先グリッドが切り替わらなかった。select セルは mousedown 時点で行切替を行い、
         * 後続の click では二重に処理しない。
         */
        var lastActivatedParentTr = null;  // 直近で取引先を表示した親行
        var selectHandledTr = null;        // mousedown で処理済みの行（直後の click を無視するため）
        atsukaikikiTbody.addEventListener('mousedown', function (e) {
            if (e.button !== 0) return;
            var sel = e.target.closest && e.target.closest('select');
            if (!sel) return;
            var tr = sel.closest('tr');
            if (!tr || tr.parentNode !== atsukaikikiTbody) return;
            selectHandledTr = tr;
            if (tr === lastActivatedParentTr) return; // 同じ行内の select は再描画不要
            activateParentRow(tr, false);
        });
        atsukaikikiTbody.addEventListener('click', function (e) {
            var tr = e.target.closest('tr');
            if (!tr || tr.parentNode !== atsukaikikiTbody) return;
            var handled = selectHandledTr === tr && e.target.closest && e.target.closest('select');
            selectHandledTr = null;
            if (handled) return;
            activateParentRow(tr, !!(e.target.closest && e.target.closest('td.col-select')));
        });

        window.selectMcm0018ParentRow = activateParentRow;
        function activateParentRow(tr, onMarkerCell) {
            if (tr.style.display === 'none') return;
            lastActivatedParentTr = tr;
            // 子グリッドに今表示している親のキー（取引先の退避先）
            var previousKey = window.__currentParentKey;

            // Step 2: Update selection
            /* MCM0015U 準拠: 行選択グレーは最左▶列(td.col-select)を押した時だけ。
               通常のセルクリックでは行グレーにせず、セル選択グレーのみとする。
               【性能改善】全行(6000行規模)の検索をやめ、直前の1行だけクラスを外す。 */
            window.Mcm0018Performance.markSelected(atsukaikikiTbody, onMarkerCell ? tr : null);
            /* ▶マーカー（行アクティブ）は常に移動 */
            window.Mcm0018Performance.markActive(tr);
            var parentId = tr.getAttribute('data-parent-id') || '';
            var isNewRow = tr.getAttribute('data-is-new') === 'true';
            var rowStatus = getRowStatus(tr);
            window.__selectedAtsukaikikiId = parentId || null;
            if (selectedAtsukaikikiIdField) selectedAtsukaikikiIdField.value = parentId;
            if (isNewRow) {
                window.__activeNewParentTempId = null;
                window.__currentParentKey = null;
            } else if (rowStatus === 'added') {
                var tempIdInput = tr.querySelector('input[name$=".tempId"]');
                window.__activeNewParentTempId = tempIdInput ? tempIdInput.value : null;
                window.__selectedAtsukaikikiId = null;
                if (selectedAtsukaikikiIdField) selectedAtsukaikikiIdField.value = '';
                window.__currentParentKey = window.__activeNewParentTempId ? ('tmp:' + window.__activeNewParentTempId) : null;
            } else {
                window.__activeNewParentTempId = null;
                window.__currentParentKey = parentId ? ('id:' + parentId) : null;
            }
            updateKoseiNewRowParentId(parentId);

            // Step 1+3: 退避と子グリッド描画は scheduleKoseiRender() でまとめて行う
            var newKey = window.__currentParentKey;
            if (newKey === previousKey && !koseiRenderPending) return; // 同じ親のセル移動では子表を再生成しない
            scheduleKoseiRender(previousKey, { newKey: newKey, parentId: parentId, isNew: isNewRow || rowStatus === 'added' });
        }
    }

    /*
     * 子グリッドの再描画は次の描画フレームにまとめ、最後に選んだ親の分だけ行う。
     *   ↑↓キー連打・長押しで親行を移動した場合、途中の親ごとに取引先を退避・再構築していたのを
     *   1回にまとめる。60msの固定待機をなくし、親の選択と子一覧を同じ描画で反映する。
     *   captureKey は「現在子グリッドに表示されている親」のキーで、最初の予約時の値を保持する。
     *   子グリッド・登録・行削除・変更判定の前には flushKoseiRender() で必ず確定させる。
     */
    var koseiRenderFrame = null;
    var koseiRenderPending = null; // { captureKey, target }
    function scheduleKoseiRender(previousKey, target) {
        if (!koseiRenderPending) koseiRenderPending = { captureKey: previousKey, target: target };
        else koseiRenderPending.target = target;
        if (koseiRenderFrame === null) koseiRenderFrame = requestAnimationFrame(flushKoseiRender);
    }
    function flushKoseiRender() {
        if (koseiRenderFrame !== null) { cancelAnimationFrame(koseiRenderFrame); koseiRenderFrame = null; }
        var p = koseiRenderPending;
        koseiRenderPending = null;
        if (!p) return;
        var t = p.target;
        // A→B→A のように表示中の親へ戻った場合は作り直さない（親キーが無い新規行同士は従来どおり再描画）
        if (t.newKey && t.newKey === p.captureKey) return;
        // Step 1: Capture current kosei before switching
        if (p.captureKey) captureCurrentKoseiRows(p.captureKey);
        // Step 3: Render kosei - prefer pending if available
        if (t.newKey && window.__pendingKoseiChanges[t.newKey]) {
            renderKoseiFromPending(window.__pendingKoseiChanges[t.newKey]);
        } else if (t.isNew) {
            renderKoseiFromPending([]);
        } else {
            renderKosei((window.__koseiRowsMap || {})[t.parentId] || []);
        }
    }
    window.flushMcm0018KoseiRender = flushKoseiRender;

    function updateKoseiNewRowParentId(parentId) {
        if (!koseiTbody) return;
        var newRow = koseiTbody.querySelector('tr[data-is-new="true"]');
        if (!newRow) return;
        var parentIdInput = newRow.querySelector('input[name$=".atsukaikikiId"]');
        var tempParentIdInput = newRow.querySelector('input[name$=".tempParentId"]');
        if (window.__activeNewParentTempId) {
            if (parentIdInput) parentIdInput.value = '';
            if (tempParentIdInput) tempParentIdInput.value = window.__activeNewParentTempId;
        } else {
            if (parentIdInput) parentIdInput.value = parentId || '';
            if (tempParentIdInput) tempParentIdInput.value = '';
        }
    }

    // ========== Part 3.2.2: 統一されたkosei行ビルダー（常に編集可能セル + No.編集可） ==========
    function buildKoseiRow(k, idx, templateForOptions) {
        var tr = document.createElement('tr');

        // kisanbi のoptionはplaceholderから借用（件数が少ないため従来どおり）
        var kisanbiOptionsHtml = '';
        if (templateForOptions) {
            var kSel = templateForOptions.querySelector('select[name$=".kisanbiKbn"]');
            if (kSel) kisanbiOptionsHtml = kSel.innerHTML;
        }
        /*
         * 【性能改善】取引先は全候補を行ごとに複製せず、空＋選択中の1件だけを持たせる。
         *   候補はプルダウンを開く時に index.html の共有候補処理（data-shared-options）が
         *   #mcm0018TorihikisakiSource から展開する。原本に無い値は従来どおり未選択になる。
         */
        var torihikisakiOptionsHtml = '<option value=""></option>';
        var tKey = normalizeOptionKey(k.torihikisakiId);
        if (tKey !== '') {
            var tLabel = getTorihikisakiLabels()[tKey];
            if (tLabel !== undefined) {
                torihikisakiOptionsHtml += '<option value="' + escapeAttr(tKey) + '" selected>' + escapeHtml(tLabel) + '</option>';
            }
        }

        // No.以外のhidden inputs
        var hiddenHtml =
            '<input type="hidden" name="koseiRows[' + idx + '].rowStatus" class="row-status-input" value="' + escapeAttr(k.rowStatus || 'unchanged') + '">' +
            '<input type="hidden" name="koseiRows[' + idx + '].tempId" value="' + escapeAttr(k.tempId) + '">' +
            '<input type="hidden" name="koseiRows[' + idx + '].tempParentId" value="' + escapeAttr(k.tempParentId) + '">' +
            '<input type="hidden" name="koseiRows[' + idx + '].atsukaikikikoseiId" value="' + escapeAttr(k.atsukaikikikoseiId) + '">' +
            '<input type="hidden" name="koseiRows[' + idx + '].atsukaikikiId" value="' + escapeAttr(k.atsukaikikiId) + '">' +
            '<input type="hidden" name="koseiRows[' + idx + '].originalMakerhosyuDt" value="' + escapeAttr(k.originalMakerhosyuDt || k.makerhosyuDt) + '">';

        var hosyuChecked = (k.hosyukeiyakutaisyoFlg == 1 || k.hosyukeiyakutaisyoFlg === '1') ? 'checked' : '';

        // Part 3.2.2: No.列を編集可能な number input に変更
        // .hyojijun-input と .hyojijun-display 両方の役割を1つの visible input が担う
        tr.innerHTML =
            '<td class="col-select"><span class="row-marker" aria-hidden="true"></span></td>' +
            '<td class="col-no">' + hiddenHtml +
                '<input type="number" name="koseiRows[' + idx + '].hyojijun" class="edit-input hyojijun-input text-right" value="' + escapeAttr(k.hyojijun) + '" style="width: 40px;">' +
            '</td>' +
            '<td class="col-torihikisaki"><select name="koseiRows[' + idx + '].torihikisakiId" class="edit-input" data-shared-options="torihikisakiId">' + torihikisakiOptionsHtml + '</select></td>' +
            '<td class="col-maker-dt"><input type="text" name="koseiRows[' + idx + '].makerhosyuDt" class="edit-input" value="' + escapeAttr(k.makerhosyuDt) + '" style="width:85px;"></td>' +
            '<td class="col-keiyaku"><input type="text" name="koseiRows[' + idx + '].keiyakukanokikan" class="edit-input" value="' + escapeAttr(k.keiyakukanokikan) + '"></td>' +
            '<td class="col-kisanbi"><select name="koseiRows[' + idx + '].kisanbiKbn" class="edit-input">' + kisanbiOptionsHtml + '</select></td>' +
            '<td class="col-hosyu-check"><input type="checkbox" name="koseiRows[' + idx + '].hosyukeiyakutaisyoFlg" value="1" ' + hosyuChecked + '></td>' +
            '<td class="col-biko-child"><input type="text" name="koseiRows[' + idx + '].biko" class="edit-input" value="' + escapeAttr(nzText(k.biko)) + '"></td>' +
            '<td class="col-dt readonly">' + escapeHtml(k.createdDt || '') + '</td>' +
            '<td class="col-user readonly">' + escapeHtml(k.createdBy || '') + '</td>' +
            '<td class="col-dt readonly">' + escapeHtml(k.lastupdateDt || '') + '</td>' +
            '<td class="col-user readonly">' + escapeHtml(k.lastupdateBy || '') + '</td>';

        // innerHTML後に選択値を設定（取引先は上で selected 済み）
        var kSel2 = tr.querySelector('select[name$=".kisanbiKbn"]');
        if (kSel2 && k.kisanbiKbn) kSel2.value = k.kisanbiKbn;

        /* MCM0015U 準拠: 生成行も既定 readonly（2クリック編集の対象にする） */
        tr.querySelectorAll('input.edit-input').forEach(function (el) { el.setAttribute('readonly', 'readonly'); });

        /*
         * ★削除保留（rowStatus="deleted"）の行は非表示のまま復元する。
         *   captureCurrentKoseiRows() は display:none の削除保留行もスナップショットに
         *   含めているが、この buildKoseiRow() で行を再構築する際に display 状態を
         *   引き継いでいなかったため、親行を切り替えて戻る（セル移動）と削除保留に
         *   した取引先行が再表示されてしまう不具合があった。
         */
        if (k.rowStatus === 'deleted') {
            tr.style.display = 'none';
        }

        return tr;
    }

    // ========== Render kosei from server data (initial display) ==========
    // Part 3.2.2: buildKoseiRow() を使って編集可能セルを生成
    function renderKosei(list) {
        if (!koseiTbody) return;
        var newRow = koseiTbody.querySelector('tr[data-is-new="true"]');
        var newRowClone = (newRow || koseiRowTemplate) ? (newRow || koseiRowTemplate).cloneNode(true) : null;
        if (newRowClone) resetNewRowClone(newRowClone);
        koseiTbody.innerHTML = '';
        var idx = 0;
        (list || []).forEach(function (k) {
            var tr = buildKoseiRow(k, idx, newRowClone);
            koseiTbody.appendChild(tr);
            idx++;
        });
        if (newRowClone) {
            reindexNewRow(newRowClone, idx, 'koseiRows');
            koseiTbody.appendChild(newRowClone);
        }
        finishKoseiRender();
    }

    function resetNewRowClone(clone) {
        /* 展開中の取引先候補を複製していても、新規行は空＋選択1件の形に戻す */
        window.Mcm0018Performance.compactRow(clone);
        clone.classList.remove('selected-row', 'row-active');
        clone.querySelectorAll('.cell-selected').forEach(function (cell) { cell.classList.remove('cell-selected'); });
        clone.querySelectorAll('input, select').forEach(function (el) {
            if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
            else if (el.tagName === 'SELECT') el.selectedIndex = 0;
            else el.value = '';
        });
        var hDisp = clone.querySelector('.hyojijun-display');
        if (hDisp) hDisp.textContent = '';
        var pid = clone.querySelector('input[name$=".atsukaikikiId"]');
        var tpid = clone.querySelector('input[name$=".tempParentId"]');
        if (window.__activeNewParentTempId) {
            if (pid) pid.value = '';
            if (tpid) tpid.value = window.__activeNewParentTempId;
        } else if (window.__selectedAtsukaikikiId) {
            if (pid) pid.value = window.__selectedAtsukaikikiId;
            if (tpid) tpid.value = '';
        } else {
            if (pid) pid.value = '';
            if (tpid) tpid.value = '';
        }
    }

    // ========== Render kosei from pending changes (preserves user edits) ==========
    // Part 3.2.2: buildKoseiRow() を統一使用（旧 buildKoseiRowFromPending は使用しない）
    function renderKoseiFromPending(pendingRows) {
        if (!koseiTbody) return;
        var newRow = koseiTbody.querySelector('tr[data-is-new="true"]');
        var newRowClone = (newRow || koseiRowTemplate) ? (newRow || koseiRowTemplate).cloneNode(true) : null;
        if (newRowClone) resetNewRowClone(newRowClone);
        koseiTbody.innerHTML = '';
        var idx = 0;
        (pendingRows || []).forEach(function (k) {
            var tr = buildKoseiRow(k, idx, newRowClone);
            koseiTbody.appendChild(tr);
            idx++;
        });
        if (newRowClone) {
            reindexNewRow(newRowClone, idx, 'koseiRows');
            koseiTbody.appendChild(newRowClone);
        }
        finishKoseiRender();
    }

    // ========== 旧 buildKoseiRowFromPending（後方互換用に残置。renderKoseiFromPendingからは使用しない）==========
    function buildKoseiRowFromPending(k, idx, templateForOptions) {
        // Part 3.2.2 以降は buildKoseiRow() を推奨
        return buildKoseiRow(k, idx, templateForOptions);
    }

    // ========== New row activation ==========
    function bindNewRowActivation(tbody, prefix) {
        if (!tbody) return;
        function handle(e) {
            var tr = e.target.closest('tr[data-is-new="true"]');
            if (!tr || !tbody.contains(tr)) return;
            /* MCM0015U 準拠: No 列だけの入力では新規行を確定させない */
            var td = e.target.closest ? e.target.closest('td') : null;
            if (td && td.classList.contains('col-no')) return;
            if (isRowContentEmptyTr(tr)) return;
            activateNewRow(tr, tbody, prefix);
        }
        tbody.addEventListener('input', handle);
        tbody.addEventListener('change', handle);
    }

    function activateNewRow(tr, tbody, prefix) {
        var rowStatusInput = tr.querySelector('.row-status-input');
        if (rowStatusInput) rowStatusInput.value = 'added';
        if (prefix === 'atsukaikikiRows') {
            var tempIdInput = tr.querySelector('input[name$=".tempId"]');
            var tempId = tempIdInput ? tempIdInput.value : '';
            if (!tempId) {
                tempId = String(generateTempId());
                if (tempIdInput) tempIdInput.value = tempId;
            }
            window.__activeNewParentTempId = tempId;
            window.__selectedAtsukaikikiId = null;
            window.__currentParentKey = 'tmp:' + tempId;
            if (selectedAtsukaikikiIdField) selectedAtsukaikikiIdField.value = '';
        } else if (prefix === 'koseiRows') {
            var parentIdInput = tr.querySelector('input[name$=".atsukaikikiId"]');
            var tempParentIdInput = tr.querySelector('input[name$=".tempParentId"]');
            if (window.__activeNewParentTempId) {
                if (parentIdInput) parentIdInput.value = '';
                if (tempParentIdInput) tempParentIdInput.value = window.__activeNewParentTempId;
            } else if (window.__selectedAtsukaikikiId) {
                if (parentIdInput) parentIdInput.value = window.__selectedAtsukaikikiId;
                if (tempParentIdInput) tempParentIdInput.value = '';
            }
        }
        /* MCM0015U 準拠: 採番は「直前行の No + 1」 */
        var newH = 1;
        var prevR = tr.previousElementSibling;
        while (prevR) {
            if (prevR.style.display !== 'none') {
                var phi = prevR.querySelector('.hyojijun-input');
                var pv = null;
                if (phi && phi.value) pv = parseInt(phi.value, 10);
                else {
                    var phd = prevR.querySelector('.hyojijun-display');
                    if (phd && phd.textContent) pv = parseInt(phd.textContent, 10);
                }
                if (pv !== null && !isNaN(pv)) { newH = pv + 1; break; }
            }
            prevR = prevR.previousElementSibling;
        }
        var hInput = tr.querySelector('.hyojijun-input');
        /* MCM0015U 準拠: ユーザーが入力済み（自動採番でない）なら上書きしない */
        if (hInput && (hInput.value === '' || hInput.getAttribute('data-auto-no') === '1')) { hInput.value = newH; hInput.removeAttribute('data-auto-no'); }
        else if (hInput && hInput.value !== '') { newH = parseInt(hInput.value, 10) || newH; }
        var anyName = null;
        var nameEls = tr.querySelectorAll('[name]');
        for (var i = 0; i < nameEls.length; i++) {
            if (nameEls[i].name.indexOf(prefix + '[') === 0) { anyName = nameEls[i].name; break; }
        }
        if (!anyName) return;
        var m = anyName.match(new RegExp(prefix + '\\[(\\d+)\\]'));
        if (!m) return;
        var currentIdx = parseInt(m[1], 10);
        var clone = tr.cloneNode(true);
        /* 展開中の候補(数百件)を複製しないよう、次の新規行は空＋選択1件の形にする */
        window.Mcm0018Performance.compactRow(clone);
        reindexNewRow(clone, currentIdx + 1, prefix);
        clone.querySelectorAll('input, select').forEach(function (el) {
            if (el.type === 'checkbox' || el.type === 'radio') el.checked = false;
            else if (el.tagName === 'SELECT') el.selectedIndex = 0;
            else el.value = '';
        });
        var cloneHDisp = clone.querySelector('.hyojijun-display');
        if (cloneHDisp) cloneHDisp.textContent = '';
        clone.querySelectorAll('input.edit-input').forEach(function (el) { el.setAttribute('readonly', 'readonly'); });
        var cloneHInput = clone.querySelector('.hyojijun-input');
        if (cloneHInput) { cloneHInput.value = ''; cloneHInput.removeAttribute('data-auto-no'); }
        clone.setAttribute('data-is-new', 'true');
        clone.classList.add('new-row');
        clone.classList.remove('selected-row');
        /* 選択状態は1行だけで保持するため、複製元の ▶/セル選択を引き継がない */
        clone.classList.remove('row-active');
        clone.querySelectorAll('.cell-selected').forEach(function (cell) { cell.classList.remove('cell-selected'); });
        if (prefix === 'koseiRows') {
            var cParentIdInput = clone.querySelector('input[name$=".atsukaikikiId"]');
            var cTempParentIdInput = clone.querySelector('input[name$=".tempParentId"]');
            if (window.__activeNewParentTempId) {
                if (cParentIdInput) cParentIdInput.value = '';
                if (cTempParentIdInput) cTempParentIdInput.value = window.__activeNewParentTempId;
            } else if (window.__selectedAtsukaikikiId) {
                if (cParentIdInput) cParentIdInput.value = window.__selectedAtsukaikikiId;
                if (cTempParentIdInput) cTempParentIdInput.value = '';
            }
        }
        tr.removeAttribute('data-is-new');
        tr.classList.remove('new-row');
        tbody.appendChild(clone);
    }

    // ========== MCM0015U 準拠: 新規行ヘルパ ==========
    /* No 列を除いた入力がすべて空か */
    function isRowContentEmptyTr(tr) {
        if (!tr) return true;
        var empty = true;
        tr.querySelectorAll('td:not(.col-no) input.edit-input, td:not(.col-no) select.edit-input').forEach(function (el) {
            if (el.value && String(el.value).trim() !== '') empty = false;
        });
        tr.querySelectorAll('input[type="checkbox"]').forEach(function (el) { if (el.checked) empty = false; });
        return empty;
    }
    /* 新規行を選択したら 直前行の No + 1 を自動採番 */
    function assignNoForRow(tr) {
        if (!tr || tr.getAttribute('data-is-new') !== 'true') return;
        var noInput = tr.querySelector('.hyojijun-input');
        if (!noInput || noInput.value !== '') return;
        var prev = tr.previousElementSibling;
        while (prev) {
            if (prev.style.display !== 'none') {
                var pno = prev.querySelector('.hyojijun-input');
                if (pno && pno.value !== '') { var v = parseInt(pno.value, 10); if (!isNaN(v)) { noInput.value = v + 1; break; } }
            }
            prev = prev.previousElementSibling;
        }
        if (noInput.value === '') noInput.value = 1;
        noInput.setAttribute('data-auto-no', '1');
    }
    /* No 以外が空のままの新規行から 自動採番した No を消す */
    function clearAutoNoIfUntouched(tr) {
        if (!tr || !document.body.contains(tr)) return;
        if (tr.getAttribute('data-is-new') !== 'true') return;
        if (!isRowContentEmptyTr(tr)) return;
        var noInput = tr.querySelector('.hyojijun-input');
        if (noInput && noInput.getAttribute('data-auto-no') === '1') { noInput.value = ''; noInput.removeAttribute('data-auto-no'); }
    }
    /* index.html 側の 2クリック編集から参照するため公開 */
    window.assignNoForRow = assignNoForRow;
    window.clearAutoNoIfUntouched = clearAutoNoIfUntouched;

    /* ===== MCM0015U 準拠: 必須チェック（一部入力あり＆必須未入力ならエラー） ===== */
    /* 親: No / 製造メーカー / 取扱機器名 / 機器分類   子: No / 取引先 */
    function getRequiredDefs(tr) {
        if (!tr) return null;
        var tb = tr.parentNode;
        if (tb === atsukaikikiTbody) {
            return [
                { el: tr.querySelector('.hyojijun-input'), label: 'No' },
                { el: tr.querySelector('td.col-seizomaker select'), label: '製造メーカー' },
                { el: tr.querySelector('td.col-atsukaikiki-nk input'), label: '取扱機器名' },
                { el: tr.querySelector('td.col-kikibunrui select'), label: '機器分類' }
            ];
        }
        if (tb === koseiTbody) {
            return [
                { el: tr.querySelector('.hyojijun-input'), label: 'No' },
                { el: tr.querySelector('td.col-torihikisaki select'), label: '取引先' }
            ];
        }
        return null;
    }
    function isElEmpty(el) { return !el || el.value === null || String(el.value).trim() === ''; }
    /* 必須チェック。mark=true なら未入力欄に .field-error を付ける */
    function validateRowRequired(tr, mark) {
        if (!tr || !document.body.contains(tr)) return true;
        var defs = getRequiredDefs(tr);
        if (!defs) return true;
        defs.forEach(function (d) {
            if (!d.el) return;
            d.el.classList.remove('field-error');
            var td0 = d.el.closest ? d.el.closest('td') : null;
            if (td0) td0.classList.remove('cell-error');
        });
        /* 行がまったく未入力（触っていない新規行）ならチェックしない */
        if (getRowStatus(tr) === "deleted") return true;
        if (tr.getAttribute("data-is-new") === "true" && isRowContentEmptyTr(tr)) return true;
        var missing = defs.filter(function (d) { return isElEmpty(d.el); });
        if (missing.length === 0) return true;
        if (mark) { missing.forEach(function (d) {
            if (!d.el) return;
            d.el.classList.add('field-error');
            /* select は appearance:none を使えない(▼が消える)ため、セル側に赤背景を付ける */
            if (d.el.tagName === 'SELECT') { var td1 = d.el.closest ? d.el.closest('td') : null; if (td1) td1.classList.add('cell-error'); }
        }); }
        return false;
    }
    function getMissingLabelsForRow(tr) {
        var defs = getRequiredDefs(tr);
        if (!defs || isRowContentEmptyTr(tr)) return [];
        var labels = [];
        defs.forEach(function (d) { if (isElEmpty(d.el)) labels.push(d.label); });
        return labels;
    }
    function buildRequiredMsg(labels) {
        if (!labels || labels.length === 0) return '未入力の必須項目があります。';
        return '未入力の必須項目があります。（' + labels.join('／') + '）';
    }

    /*
     * A027: 退避データ（__pendingKoseiChanges）の1行が「実際の変更」かどうかを判定する。
     *   captureCurrentKoseiRows() は親行を切り替えるたびに未変更行（unchanged）も含めて
     *   退避するため、件数だけで判定すると行を選び直しただけで「変更あり」になっていた。
     *   added 行は、画面行の判定（isRowContentEmptyTr）と同じく No 以外に入力がある場合のみ変更扱い。
     */
    function isPendingRowChanged(r) {
        if (!r) return false;
        var s = r.rowStatus || 'unchanged';
        if (s === 'modified' || s === 'deleted') return true;
        if (s === 'added') {
            var fields = ['torihikisakiId', 'makerhosyuDt', 'keiyakukanokikan', 'kisanbiKbn', 'biko'];
            for (var i = 0; i < fields.length; i++) {
                if (String(r[fields[i]] == null ? '' : r[fields[i]]).trim() !== '') return true;
            }
            return r.hosyukeiyakutaisyoFlg === '1';
        }
        return false;
    }

    /* ===== MCM0015U 準拠: 変更有無の判定（「閉じる」の確認ダイアログ用） ===== */
    function hasUnsavedChanges() {
        flushKoseiRender(); // 予約中の子グリッド切り替えを確定してから判定する
        var changed = false;
        [atsukaikikiTbody, koseiTbody].forEach(function (tb) {
            if (!tb || changed) return;
            tb.querySelectorAll('tr').forEach(function (tr) {
                if (changed) return;
                var st = tr.querySelector('.row-status-input');
                var status = st ? st.value : 'unchanged';
                if (status === 'modified' || status === 'deleted') { changed = true; return; }
                if ((status === 'added' || tr.getAttribute('data-is-new') === 'true') && !isRowContentEmptyTr(tr)) { changed = true; return; }
            });
        });
        /* 他の親行に紐づく保留中の子データ変更も「変更あり」とみなす
           A027: 件数ではなく、実際に変更された行が含まれる場合のみ「変更あり」とする */
        if (!changed && window.__pendingKoseiChanges) {
            for (var k in window.__pendingKoseiChanges) {
                if (!Object.prototype.hasOwnProperty.call(window.__pendingKoseiChanges, k)) continue;
                var v = window.__pendingKoseiChanges[k] || [];
                for (var j = 0; j < v.length; j++) {
                    if (isPendingRowChanged(v[j])) { changed = true; break; }
                }
                if (changed) break;
            }
        }
        return changed;
    }

    /* index.html 側のインラインスクリプトから参照するため公開 */
    window.validateRowRequired = validateRowRequired;
    window.getMissingLabelsForRow = getMissingLabelsForRow;
    window.buildRequiredMsg = buildRequiredMsg;
    window.hasUnsavedChanges = hasUnsavedChanges;
    window.customAlertMcm = customAlert;
    window.customConfirmMcm = customConfirm;   /* 確認ダイアログ（？アイコン付き）を index.html から使う */

    function reindexNewRow(row, newIdx, prefix) {
        row.querySelectorAll('[name]').forEach(function (el) {
            var oldName = el.getAttribute('name');
            var newName = oldName.replace(new RegExp(prefix + '\\[\\d+\\]'), prefix + '[' + newIdx + ']');
            el.setAttribute('name', newName);
        });
    }

    /* ===== 桁数・文字数チェック（VB CPValidate 相当）===== */
    var LENGTH_RULES = [
        /* hyojijun(No) は checkCellLength() 内で専用チェックするため、ここには含めない */
        { pattern: /\.atsukaikikiNk$/,      maxLen: 60,   label: "取扱機器名" },
        { pattern: /\.katashiki$/,          maxLen: 60,   label: "型式" },
        { pattern: /\.keiyakukanokikan$/,   maxLen: 20,   label: "契約可能期間" },
        { pattern: /\.biko$/,               maxLen: 4000, label: "備考" }
    ];
    /*
     * ★行単位の相関チェック（コントローラフラグ⇔金額、契約可能期間⇔起算日区分）。
     *   登録時チェックとセル移動時チェックの両方から共通で呼び出す。
     *   サーバー側 validate() と同一仕様。戻り値: エラーメッセージ文字列 or null。
     */
    function checkAtsukaikikiRowConsistency(tr) {
        if (!tr || isRowContentEmptyTr(tr)) return null;
        var flgInp = tr.querySelector('input[name$=".controllerFlg"]');
        var kinInp = tr.querySelector('input[name$=".controllerKin"]');
        if (flgInp && kinInp) {
            var hasKin = kinInp.value.trim() !== "";
            if (flgInp.checked && !hasKin) return "コントローラのチェックがされていますが、金額が入力されていません。";
            if (!flgInp.checked && hasKin) return "金額が入力されていますが、コントローラのチェックがされていません。";
        }
        return null;
    }
    function checkKoseiRowConsistency(tr) {
        if (!tr || isRowContentEmptyTr(tr)) return null;
        var kikanInp = tr.querySelector('input[name$=".keiyakukanokikan"]');
        var kisanbiSel = tr.querySelector('select[name$=".kisanbiKbn"]');
        if (kikanInp && kisanbiSel) {
            var hasKikan = kikanInp.value.trim() !== "";
            var hasKisanbi = kisanbiSel.value.trim() !== "";
            if (hasKikan && !hasKisanbi) return "契約可能期間を入力した場合は、起算日区分も入力してください。";
            if (!hasKikan && hasKisanbi) return "起算日区分を入力した場合は、契約可能期間も入力してください。";
        }
        return null;
    }

    function checkCellLength(inp) {
        if (!inp || inp.value === "") return null;
        var name = inp.name || "";
        /*
         * 【修正】金額（controllerKin）は type="text"（money-input）のため、ブラウザの
         *   type="number" のような自動フィルタが働かず、"e"や英字・日本語等の
         *   数値以外の文字がそのまま入力できてしまっていた。送信前の moneyDigits() は
         *   非数字を単純除去するだけで "12e3" → "123" のような不正な値化けを防げず、
         *   誤った値がチェックを通過して登録されてしまう不具合があった。
         *   桁数チェックと同じタイミング（登録ボタン押下時）で数値形式チェックを行う。
         *   DB列は NUMBER(10) で整数・マイナス不可（実運用上、金額は負値を取らない）。
         */
        if (/\.controllerKin$/.test(name)) {
            var raw = inp.value.trim().replace(/[,￥]/g, "");
            if (!/^\d+$/.test(raw)) {
                return "金額は数値で入力してください。";
            }
            if (raw.replace(/^0+(?=\d)/, "").length > 10) {
                return "金額は10桁以下で入力してください。";
            }
            return null;
        }
        /*
         * 【修正】UPS交換周期（upskokanshuki）は従来 type="number" だったため、
         *   "e"（指数表記の一部として許容される）等の無効な値を入力すると value が
         *   空文字化してしまい、数値以外入力チェック自体が素通りしていた。
         *   type="text" に変更（テンプレート側）した上で、桁数チェックの対象
         *   （LENGTH_RULES）にも含まれていなかったこの項目を明示的にチェックする。
         */
        if (/\.upskokanshuki$/.test(name)) {
            var upsRaw = inp.value.trim();
            if (!/^\d+$/.test(upsRaw)) {
                return "UPS交換周期は数値で入力してください。";
            }
            if (upsRaw.replace(/^0+(?=\d)/, "").length > 5) {
                return "UPS交換周期は5桁以下で入力してください。";
            }
            return null;
        }
        /*
         * ★No（hyojijun）専用の数値チェック（MCM0015U準拠、半角数字のみ許容）。
         *   全角数字・英字・記号・小数点・指数表記(e/E)・正負符号は不可とする。
         *   従来は decimal:true のみで numeric 判定が無く、parseFloat("e") が NaN に
         *   なることでチェックをすり抜けていた（数値以外を数値チェックが検出できなかった）。
         */
        if (/\.hyojijun$/.test(name)) {
            if (!/^\d+$/.test(inp.value.trim())) {
                return "Noは数値で入力してください。";
            }
            if (inp.value.trim().replace(/^0+(?=\d)/, "").length > 6) {
                return "Noは6桁以下で入力してください。";
            }
            return null;
        }
        for (var i = 0; i < LENGTH_RULES.length; i++) {
            var rule = LENGTH_RULES[i];
            if (!rule.pattern.test(name)) continue;
            if (rule.decimal) {
                var num = parseFloat(inp.value);
                if (!isNaN(num) && Math.abs(Math.trunc(num)) >= Math.pow(10, rule.maxLen))
                    return rule.label + "は" + rule.maxLen + "桁以下で入力してください。";
            } else if (inp.value.length > rule.maxLen) {
                return rule.label + "は" + rule.maxLen + "文字以内で入力してください。";
            }
            return null;
        }
        return null;
    }

    /*
     * 【性能改善（MCM0015U準拠）】行データをtr内のname付き要素から汎用的に収集し、
     *   JSONオブジェクトとして構築する。従来は FormData（application/x-www-form-urlencoded）
     *   による通常フォーム送信（ページ全体再読み込み）で、行数の多いフォームでは
     *   Spring MVC側のバインディングが極端に遅くなる問題があった。
     */
    function buildRowsPayloadGeneric(tbodyId, prefix) {
        var tb = document.getElementById(tbodyId);
        var rows = [];
        if (!tb) return rows;
        tb.querySelectorAll("tr").forEach(function (tr) {
            if (tr.getAttribute("data-is-new") === "true" && isRowContentEmptyTr(tr)) return;
            if (tr.getAttribute("data-is-new") === "true" && isRowContentEmptyTr(tr)) return;
            var row = {};
            var hasField = false;
            tr.querySelectorAll("[name^='" + prefix + "[']").forEach(function (el) {
                if (el.disabled) return;
                var m = el.name.match(/^\w+\[\d+\]\.(.+)$/);
                if (!m) return;
                var field = m[1];
                if (el.type === "checkbox") {
                    row[field] = el.checked ? "1" : "0";
                } else if (el.classList.contains("money-input")) {
                    // 【修正】「￥」付き表示（money-input）は送信前に数値のみへ戻す
                    //   （BigDecimal等の数値項目でJSONデシリアライズエラーを防ぐ）。
                    var digits = moneyDigits(el.value);
                    row[field] = (digits === "" || digits === "-") ? null : digits;
                } else {
                    // 空文字はnullに変換する（BigDecimal等の数値項目でJSONデシリアライズエラーを防ぐ）
                    row[field] = (el.value === "") ? null : el.value;
                }
                hasField = true;
            });
            if (hasField) rows.push(row);
        });
        return rows;
    }
    /*
     * フォーム直下（行以外）のトップレベルhiddenフィールドを収集する。
     * 空文字はnullに変換する（BigDecimal等の数値項目が空文字のままだと
     * サーバー側でJSONデシリアライズエラーになるため）。
     */
    function buildTopLevelPayload(form) {
        var payload = {};
        Array.prototype.forEach.call(form.elements, function (el) {
            if (!el.name || el.name.indexOf("[") !== -1) return; // 行フィールドは除外
            payload[el.name] = (el.value === "") ? null : el.value;
        });
        return payload;
    }

    // ========== Update button click - MERGE all pending changes ==========
    if (updateBtn && updateForm) {
        updateBtn.addEventListener('click', function () {
            flushKoseiRender(); // 予約中の子グリッド切り替えを確定してからチェック・送信する
            // 必須エラー時は確認ダイアログも送信も行わず、DOMを保持する。
            var invalid = null;
            [atsukaikikiTbody, koseiTbody].forEach(function (tbody) {
                if (!tbody) return;
                tbody.querySelectorAll('tr').forEach(function (tr) {
                    if (!validateRowRequired(tr, true) && !invalid) invalid = tr;
                });
            });
            if (invalid) {
                customAlert(buildRequiredMsg(getMissingLabelsForRow(invalid)), 'input');
                return;
            }
            var pendingError = false;
            Object.keys(window.__pendingKoseiChanges).forEach(function (key) {
                if (key === window.__currentParentKey) return;
                (window.__pendingKoseiChanges[key] || []).forEach(function (row) {
                    if (row.rowStatus !== 'deleted' && (!String(row.hyojijun || '').trim() || !String(row.torihikisakiId || '').trim())) pendingError = true;
                });
            });
            if (pendingError) {
                customAlert('別の取扱機器の取引先に未入力の必須項目があります。（No／取引先）', 'input');
                return;
            }
            /* ★桁数・文字数チェック（VB CPValidate 相当） */
            var lenErr = null, lenErrInp = null;
            document.querySelectorAll("#updateForm input.edit-input, #updateForm select.edit-input").forEach(function (inp) {
                if (lenErr || !inp.name) return;
                var tr = inp.closest("tr");
                var st = tr && tr.querySelector(".row-status-input");
                if (st && String(st.value || "") === "deleted") return;
                var msg = checkCellLength(inp);
                if (msg) { inp.classList.add("field-error"); lenErr = msg; lenErrInp = inp; }
            });
            if (lenErr) {
                customAlert(lenErr, 'input');
                if (lenErrInp) lenErrInp.focus();
                return;
            }
            /*
             * ★コントローラフラグと金額の入力整合性チェック（サーバー側 validate() と同一仕様）。
             *   ON+金額未入力 / OFF+金額入力あり のいずれもエラーとする。クライアント側でも
             *   同じ判定を行うことで、サーバーとの往復を待たずに即座にエラーを表示できる。
             *   最終判定はサーバー側 validate() で行われるため、こちらはUX向上のための
             *   先行チェックであり、サーバー側のロジックは変更しない。
             */
            var ctrlErr = null, ctrlErrInp = null;
            document.querySelectorAll("#atsukaikikiTbody tr").forEach(function (tr) {
                if (ctrlErr) return;
                var st = tr.querySelector(".row-status-input");
                if (st && String(st.value || "") === "deleted") return;
                var msg2 = checkAtsukaikikiRowConsistency(tr);
                if (msg2) { ctrlErr = msg2; ctrlErrInp = tr.querySelector('input[name$=".controllerKin"]'); }
            });
            if (ctrlErr) {
                customAlert(ctrlErr, 'input');
                if (ctrlErrInp) ctrlErrInp.focus();
                return;
            }
            /*
             * ★契約可能期間⇔起算日区分の入力整合性チェック（サーバー側 validate() と同一仕様）。
             *   従来はサーバー側のみでチェックされ、クライアント側の先行チェックが無かった。
             */
            var kisanbiErr = null, kisanbiErrInp = null;
            document.querySelectorAll("#koseiTbody tr").forEach(function (tr) {
                if (kisanbiErr) return;
                var st = tr.querySelector(".row-status-input");
                if (st && String(st.value || "") === "deleted") return;
                var msg3 = checkKoseiRowConsistency(tr);
                if (msg3) { kisanbiErr = msg3; kisanbiErrInp = tr.querySelector('input[name$=".keiyakukanokikan"]'); }
            });
            if (kisanbiErr) {
                customAlert(kisanbiErr, 'input');
                if (kisanbiErrInp) kisanbiErrInp.focus();
                return;
            }
            /*
             * ★変更有無チェックを登録確認ダイアログより先に実施する。
             *   従来は「登録を実施しますか」を先に表示し、OK押下後にサーバー側で
             *   変更なし判定（DataNotChangedException）をエラーとして返す流れだったため、
             *   変更が無い場合でも不要な確認操作が発生していた。
             *   MCM0017U等と同様、クライアント側で先にチェックし、変更が無ければ
             *   確認ダイアログを出さずに「値が変更されていません。」を表示する。
             */
            if (!hasUnsavedChanges()) {
                customAlert('値が変更されていません。', 'input');
                return;
            }
            customConfirm('登録を実施します。よろしいですか？', function () {
                // Step 1: Capture current parent's kosei to pending
                if (window.__currentParentKey) {
                    captureCurrentKoseiRows(window.__currentParentKey);
                }
                // 画面の空行・候補・他の親の表示を変更せず、JSONのみを組み立てる。
                var payload = buildTopLevelPayload(updateForm);
                payload.atsukaikikiRows = buildRowsPayloadGeneric("atsukaikikiTbody", "atsukaikikiRows");
                payload.koseiRows = buildRowsPayloadGeneric("koseiTbody", "koseiRows");
                Object.keys(window.__pendingKoseiChanges).forEach(function (key) {
                    if (key === window.__currentParentKey) return;
                    (window.__pendingKoseiChanges[key] || []).forEach(function (row) {
                        var copy = {};
                        Object.keys(row).forEach(function (field) { copy[field] = row[field] === '' ? null : row[field]; });
                        payload.koseiRows.push(copy);
                    });
                });
                var headers = { "X-Requested-With": "XMLHttpRequest", "Content-Type": "application/json" };
                var mh = document.querySelector('meta[name="_csrf_header"]');
                var mt = document.querySelector('meta[name="_csrf"]');
                if (mh && mt && mh.content && mt.content) headers[mh.content] = mt.content;

                if (window.McmCommon) window.McmCommon.showProgress();
                fetch(updateForm.getAttribute("action"), {
                    method: "POST", headers: headers, body: JSON.stringify(payload)
                }).then(function (res) {
                    if (!res.ok) {
                        return res.json().catch(function () { return null; }).then(function (data) {
                            customAlert((data && data.message) ? data.message : "登録に失敗しました。時間をおいて再度お試しください。",
                                (data && data.errorType) || "warning");
                        });
                    }
                    return res.json().catch(function () { return {}; }).then(function (data) {
                        customAlert((data && data.message) ? data.message : "登録処理が完了しました。", "success", function () {
                            saveViewStateBeforeReload();
                            window.location.reload();
                        });
                    });
                }).catch(function () {
                    customAlert("通信エラーが発生しました。時間をおいて再度お試しください。", "warning");
                }).finally(function () {
                    if (window.McmCommon) window.McmCommon.hideProgress();
                });
            });
        });
    }

    /*
     * ===== 登録後の表示位置維持（スクロール位置・選択行の保存/復元） =====
     * 【修正】登録後は window.location.reload() で画面全体を再読込しているため、
     *   常に一覧の先頭へスクロール位置が戻り、選択行も失われていた。
     *   sessionStorage にスクロール位置と選択中の行IDを保存しておき、
     *   再読込後の初期化時に復元することで、登録前の表示位置・選択行を維持する。
     *   （画面全体の再読込自体は変更しない。DOM部分更新への変更は既存の
     *    サーバー応答仕様・テンプレート構造の変更を伴い影響範囲が大きいため見送った。）
     */
    var VIEW_STATE_KEY = 'mcm0018u_viewState';
    function saveViewStateBeforeReload() {
        try {
            var parentScroller = atsukaikikiTbody && atsukaikikiTbody.closest('.grid-scroll-parent');
            var childScroller = koseiTbody && koseiTbody.closest('.grid-scroll-child');
            /*
             * A027: 保存の基準を ▶マーカー行（row-active）に変更する。
             *   selected-row は ▶列クリック時のみ付与されるため、通常のセルクリックで
             *   選択した取扱機器が保存されず、登録後に先頭行へ戻っていた。
             */
            var activeRow = atsukaikikiTbody && atsukaikikiTbody.querySelector('tr.row-active');
            var selectedKoseiRow = koseiTbody && koseiTbody.querySelector('tr.selected-row');
            var state = {
                parentScrollTop: parentScroller ? parentScroller.scrollTop : 0,
                childScrollTop: childScroller ? childScroller.scrollTop : 0,
                selectedAtsukaikikiId: activeRow ? (activeRow.querySelector('input[name$=".atsukaikikiId"]') || {}).value : '',
                parentRowSelected: !!(activeRow && activeRow.classList.contains('selected-row')),
                selectedAtsukaikikikoseiId: selectedKoseiRow ? (selectedKoseiRow.querySelector('input[name$=".atsukaikikikoseiId"]') || {}).value : ''
            };
            sessionStorage.setItem(VIEW_STATE_KEY, JSON.stringify(state));
        } catch (e) { /* sessionStorage不可時は無視（表示位置維持の失敗は致命的ではない） */ }
    }
    function restoreViewStateAfterReload() {
        var raw;
        try { raw = sessionStorage.getItem(VIEW_STATE_KEY); } catch (e) { return; }
        if (!raw) return;
        try { sessionStorage.removeItem(VIEW_STATE_KEY); } catch (e) { /* ignore */ }
        var state;
        try { state = JSON.parse(raw); } catch (e) { return; }
        if (!state) return;

        // 選択行を復元（IDが一致する行を選択状態にする）
        if (state.selectedAtsukaikikiId && atsukaikikiTbody) {
            var input = atsukaikikiTbody.querySelector(
                'input[name$=".atsukaikikiId"][value="' + state.selectedAtsukaikikiId + '"]');
            var tr = input && input.closest('tr');
            if (tr) {
                /*
                 * A027: CSSクラスの付け替えだけでは、取引先グリッド・__currentParentKey・
                 *   selectedAtsukaikikiIdField が先頭行のまま残り、▶と取引先の表示が食い違っていた。
                 *   親行クリック処理（Parent row click）をそのまま実行して、すべて同期させる。
                 */
                tr.click();
                flushKoseiRender(); // 直後に子行の選択を復元するため即時に描画する
                if (state.parentRowSelected) window.Mcm0018Performance.markSelected(atsukaikikiTbody, tr);
            }
        }
        if (state.selectedAtsukaikikikoseiId && koseiTbody) {
            var kInput = koseiTbody.querySelector(
                'input[name$=".atsukaikikikoseiId"][value="' + state.selectedAtsukaikikikoseiId + '"]');
            var kTr = kInput && kInput.closest('tr');
            if (kTr) {
                window.Mcm0018Performance.markSelected(koseiTbody, kTr);
            }
        }

        // スクロール位置を復元（DOM構築後に確実に効くよう次フレームで実行）
        requestAnimationFrame(function () {
            var parentScroller = atsukaikikiTbody && atsukaikikiTbody.closest('.grid-scroll-parent');
            var childScroller = koseiTbody && koseiTbody.closest('.grid-scroll-child');
            if (parentScroller && typeof state.parentScrollTop === 'number') parentScroller.scrollTop = state.parentScrollTop;
            if (childScroller && typeof state.childScrollTop === 'number') childScroller.scrollTop = state.childScrollTop;
        });
    }
    restoreViewStateAfterReload();

    // ========== Delete button ==========
    /*
     * ★取引先（構成）グリッドで既に削除保留（rowStatus="deleted"）になっている行の
     *   atsukaikikikoseiId 一覧を返す。取扱機器（親）行の削除チェック時に、
     *   この保留分をDB残存件数から除外するために使う。
     *   koseiTbody は現在選択中の親の構成行を表示しているため、ここに含まれる
     *   削除保留行が「これから削除しようとしている親」に対応する保留分となる。
     */
    function collectPendingDeletedKoseiIds() {
        var ids = [];
        if (!koseiTbody) return ids;
        koseiTbody.querySelectorAll('tr').forEach(function (tr) {
            var rs = tr.querySelector('.row-status-input');
            if (!rs || rs.value !== 'deleted') return;
            var idInput = tr.querySelector('input[name$=".atsukaikikikoseiId"]');
            if (idInput && idInput.value) ids.push(idInput.value);
        });
        return ids;
    }

    /**
     * @param {string} idFieldSuffix 参照チェック対象IDのinput name末尾（例: ".atsukaikikiId"）。
     *   nullの場合は参照チェックを行わない。
     * @param {string} checkUrl 参照チェックAPIのURL（GET, クエリパラメータ paramName=値）。
     * @param {string} paramName チェックAPIに渡すクエリパラメータ名。
     * @param {function} [extraParamsFn] 追加のクエリパラメータ文字列を返す関数（省略可）。
     */
    function setupDeleteButton(btn, tbody, confirmMsg, idFieldSuffix, checkUrl, paramName, extraParamsFn) {
        if (!btn || !tbody) return;
        btn.addEventListener('click', function () {
            flushKoseiRender(); // 親削除の保留子ID収集・子行選択は表示確定後に行う
            var selected = tbody.querySelector('tr.selected-row');
            if (!selected) { customAlert('行が選択されていません。', 'input'); return; }
            if (selected.getAttribute('data-is-new') === 'true') { customAlert('末尾の新規行は削除できません。', 'input'); return; }

            function doConfirmAndDelete() {
                customConfirm(confirmMsg, function () {
                    var rs = selected.querySelector('.row-status-input');
                    if (rs && rs.value === 'added') {
                        selected.remove();
                    } else if (rs) {
                        rs.value = 'deleted';
                        selected.style.display = 'none';
                    }
                    if (changeStatusField) changeStatusField.value = 'true';
                });
            }

            /*
             * ★行削除ボタン押下時点で参照整合性チェックを実施する。
             *   従来はチェックが登録ボタン押下時（updateAll/validate）にしか行われず、
             *   参照ありのデータでも「削除保留（rowStatus=deleted）」にはできてしまい、
             *   登録ボタンを押すまでエラーに気付けなかった。既存の行（IDを持つ行）に対しては
             *   削除ボタン押下時にサーバーへ参照チェックを問い合わせ、参照があれば
             *   その場でエラーを表示し、確認ダイアログ自体を出さない。
             */
            var idInput = idFieldSuffix ? selected.querySelector('input[name$="' + idFieldSuffix + '"]') : null;
            var idVal = idInput ? idInput.value : '';
            if (checkUrl && idVal) {
                var url = checkUrl + '?' + paramName + '=' + encodeURIComponent(idVal);
                if (extraParamsFn) {
                    var extra = extraParamsFn();
                    if (extra) url += '&' + extra;
                }
                if (window.McmCommon) window.McmCommon.showProgress();
                fetch(url, { method: 'GET' })
                    .then(function (res) { return res.json(); })
                    .then(function (data) {
                        if (data && data.hasReference) {
                            customAlert(data.message || '取扱機器情報として、既に使用されている為、削除する事が出来ません。', 'warning');
                        } else {
                            doConfirmAndDelete();
                        }
                    })
                    .catch(function () {
                        customAlert('通信エラーが発生しました。時間をおいて再度お試しください。', 'warning');
                    })
                    .finally(function () {
                        if (window.McmCommon) window.McmCommon.hideProgress();
                    });
            } else {
                doConfirmAndDelete();
            }
        });
    }

    // ========== Initialization ==========
    // Determine initial current parent key
    (function initCurrentParentKey() {
        if (window.__selectedAtsukaikikiId) {
            window.__currentParentKey = 'id:' + window.__selectedAtsukaikikiId;
        }
    })();

    bindChangeDetection();
    bindNewRowActivation(atsukaikikiTbody, 'atsukaikikiRows');
    bindNewRowActivation(koseiTbody, 'koseiRows');
    /*
     * ★コンテキストパス対応：本アプリは application.yml で server.servlet.context-path=/mcm
     *   を設定しているため、素の絶対パス（/mcm0018u/...）ではコンテキストパスが付かず
     *   404になる。updateForm の action 属性（Thymeleaf の @{...} でコンテキストパスが
     *   自動付与されたURL）を基準にして、確実に正しいベースパスを組み立てる。
     */
    function resolveContextPath() {
        if (updateForm) {
            var action = updateForm.getAttribute('action') || '';
            var idx = action.indexOf('/mcm0018u/update');
            if (idx >= 0) return action.substring(0, idx);
        }
        return '';
    }
    var ctxPath = resolveContextPath();
    setupDeleteButton(btnDeleteAtsukaikiki, atsukaikikiTbody, '行を削除します。よろしいですか？',
        '.atsukaikikiId', ctxPath + '/mcm0018u/checkDeleteAtsukaikiki', 'atsukaikikiId',
        function () {
            /* ★取引先グリッドで削除保留中の行は、親削除の参照カウントから除外する。 */
            var ids = collectPendingDeletedKoseiIds();
            return ids.map(function (id) { return 'excludeKoseiIds=' + encodeURIComponent(id); }).join('&');
        });
    setupDeleteButton(btnDeleteKosei, koseiTbody, '行を削除します。よろしいですか？',
        '.atsukaikikikoseiId', ctxPath + '/mcm0018u/checkDeleteKosei', 'atsukaikikikoseiId');
    /* ===== MCM0015U 準拠: 行選択グレーは最左▶列クリック時のみ ===== */
    function bindRowSelectByMarker(tbody) {
        if (!tbody) return;
        tbody.addEventListener('click', function (e) {
            var cell = e.target.closest ? e.target.closest('td.col-select') : null;
            if (!cell) return;
            var prevSel = window.Mcm0018Performance.activeRow(tbody);
            var trTarget = cell.closest('tr');
            /* 新規行を編集した状態で他の行の▶を押した場合は必須チェックエラー（MCM0015U 準拠） */
            if (prevSel && trTarget && prevSel !== trTarget && typeof validateRowRequired === 'function' &&
                !validateRowRequired(prevSel, true)) {
                e.preventDefault();
                e.stopPropagation();
                customAlert(buildRequiredMsg(getMissingLabelsForRow(prevSel)), 'input');
                var fmk = prevSel.querySelector('.field-error');
                if (fmk) { try { fmk.removeAttribute('readonly'); fmk.focus({ preventScroll: true }); } catch (ev) { } }
                return;
            }
            var tr = cell.closest('tr');
            if (!tr || tr.parentNode !== tbody) return;
            /* 【性能改善】全行検索をやめ、保持している直前の1行だけを外す */
            window.Mcm0018Performance.markSelected(tbody, tr);
            window.Mcm0018Performance.markActive(tr);
        });
    }
    bindRowSelectByMarker(atsukaikikiTbody);
    bindRowSelectByMarker(koseiTbody);

    /* ===== MCM0015U 準拠: 初期表示時は全入力欄を readonly（2クリックで編集開始） =====
     * 【性能改善】index.html の setupTwoClickEdit() の lockAll() が同じ処理を初期化時に1回行うため、
     *   ここでの全入力(数万件)への重複設定はやめた。JS生成行は buildKoseiRow() で設定済み。 */

    /* 子グリッドを触る前に、予約中の再描画を確定させる（capture で他の処理より先に実行） */
    (function bindKoseiRenderFlush() {
        var koseiGridEl = document.getElementById('koseiGrid');
        if (!koseiGridEl) return;
        ['pointerdown', 'mousedown', 'focusin', 'keydown'].forEach(function (ev) {
            koseiGridEl.addEventListener(ev, flushKoseiRender, true);
        });
    })();

    /* ===== 入力されたら .field-error を解除する（MCM0015U 準拠） ===== */
    [atsukaikikiTbody, koseiTbody].forEach(function (tb) {
        if (!tb) return;
        ['input', 'change'].forEach(function (evt) {
            tb.addEventListener(evt, function (e) {
                var el = e.target;
                if (el && el.classList && el.classList.contains('field-error') &&
                    el.value !== null && String(el.value).trim() !== '') {
                    el.classList.remove('field-error');
                    var tdc = el.closest ? el.closest('td') : null;
                    if (tdc) tdc.classList.remove('cell-error');
                }
            });
        });
    });
    /* ===== [修正] "<<NULL>>" 等のヌル表記を画面上から一掃する（備考に限らず全セル対象） ===== */
    /*   DB値・Thymeleaf描画・JS動的生成のどの経路で入っても、描画後にまとめて空欄化する。 */
    /* skipOptions=true: 候補(option)は除去済みの原本由来のため走査しない（子グリッド再描画用） */
    function sweepNullLiterals(root, skipOptions) {
        root = root || document.body;
        if (!root || !root.querySelectorAll) return;
        var NULLPAT = /^[\s\u3000]*(?:<<|\u00AB|\uFF1C\uFF1C|\u2039\u2039|\uFF62|\u300E|\u300C|\()?[\s\u3000]*(?:NULL|\u30CC\u30EB)[\s\u3000]*(?:>>|\u00BB|\uFF1E\uFF1E|\u203A\u203A|\uFF63|\u300F|\u300D|\))?[\s\u3000]*$/i;
        /* 1) 入力欄（text/number/date/hidden/textarea）の値と value 属性 */
        root.querySelectorAll("input, textarea").forEach(function (el) {
            if (el.type === "checkbox" || el.type === "radio") return;
            if (NULLPAT.test(el.value || "")) { el.value = ""; }
            var dv = el.getAttribute("value");
            if (dv !== null && NULLPAT.test(dv)) { el.setAttribute("value", ""); }
        });
        /* 2) プルダウンの選択肢テキスト */
        if (!skipOptions) root.querySelectorAll("option").forEach(function (op) {
            if (NULLPAT.test(op.textContent || "")) { op.textContent = ""; }
        });
        /* 3) 子要素を持たない末端要素のテキスト（td/th/span/label 等） */
        root.querySelectorAll("td, th, span, label, div, p, a").forEach(function (el) {
            if (el.children.length === 0 && NULLPAT.test(el.textContent || "")) { el.textContent = ""; }
        });
        /* 4) 文章の一部に混ざった "<<NULL>>" "«NULL»" 等の表記も除去 */
        var wk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
        var tn, hits = [];
        while ((tn = wk.nextNode())) {
            if (skipOptions && tn.parentNode && tn.parentNode.nodeName === "OPTION") continue;
            if (/(?:<<|\u00AB|\uFF1C\uFF1C|\u2039\u2039)[\s\u3000]*NULL[\s\u3000]*(?:>>|\u00BB|\uFF1E\uFF1E|\u203A\u203A)/i.test(tn.nodeValue || "")) { hits.push(tn); }
        }
        hits.forEach(function (n) {
            n.nodeValue = n.nodeValue.replace(/(?:<<|\u00AB|\uFF1C\uFF1C|\u2039\u2039)[\s\u3000]*NULL[\s\u3000]*(?:>>|\u00BB|\uFF1E\uFF1E|\u203A\u203A)/gi, "");
        });
    }
    function sweepAllGrids() { sweepNullLiterals(document.body); }
    /*
     * 【性能改善】従来は描画タイミングのズレに備えて、DOMContentLoaded / load / 0ms / 300ms / 1000ms
     *   と計6回、画面全体（6000行規模・数十万ノード）を走査していた。表示直後の約1秒間に
     *   クリックしても反応が遅れる原因になっていた。
     *   グリッドはサーバー側で描画済みで、このスクリプトは本文末尾（グリッドの後）で実行されるため、
     *   初期DOMは1回の走査で足りる。後から追加・再描画される行は下の MutationObserver と
     *   finishKoseiRender() が対象にする。
     */
    sweepAllGrids();
    /*
     * 【性能改善】動的な行追加・再描画への追従は必要だが、従来は
     *   document.documentElement 全体（childList/subtree/characterData）を監視し、
     *   変更があるたびに sweepNullLiterals(document.body) でドキュメント全体を
     *   毎回再走査していたため、行編集や再描画のたびに全DOM走査が発生していた。
     *
     * 改善: 監視対象を「実際に動的更新されるグリッド本体（atsukaikikiTbody/koseiTbody）」
     *   に限定し、かつ MutationRecord から実際に追加されたノードのみを対象にスイープする。
     *   これにより、動的更新の都度ドキュメント全体を再走査する必要がなくなる。
     */
    if (window.MutationObserver) {
        var __swTimer = null;
        var __pendingTargets = [];
        function scheduleSweep(target) {
            if (target && __pendingTargets.indexOf(target) === -1) __pendingTargets.push(target);
            if (__swTimer) return;
            __swTimer = setTimeout(function () {
                __swTimer = null;
                var targets = __pendingTargets;
                __pendingTargets = [];
                targets.forEach(function (t) { if (t && t.isConnected !== false) sweepNullLiterals(t); });
            }, 50);
        }
        var watchTargets = [
            document.getElementById("atsukaikikiTbody"),
            document.getElementById("koseiTbody")
        ].filter(Boolean);
        var observer = new MutationObserver(function (mutations) {
            mutations.forEach(function (m) {
                /*
                 * 【性能改善】共有候補の展開・縮小（select 直下の option 入れ替え）は対象外。
                 *   option は除去済みの原本から作るため不要で、展開のたびに数百件の option を
                 *   個別にスイープ予約すると、それ自体がセル選択時のもたつきになる。
                 */
                if (m.type === "childList" && m.target && m.target.nodeName === "SELECT") return;
                // 追加ノードがある場合はそのノードだけをスイープ対象にする
                if (m.addedNodes && m.addedNodes.length > 0) {
                    m.addedNodes.forEach(function (n) {
                        if (n.nodeType === 1) scheduleSweep(n);
                    });
                }
                // テキスト変更（characterData）の場合はその親要素をスイープ対象にする
                if (m.type === "characterData" && m.target && m.target.parentElement) {
                    scheduleSweep(m.target.parentElement);
                }
            });
        });
        watchTargets.forEach(function (t) {
            observer.observe(t, { childList: true, subtree: true, characterData: true });
        });
        nullSweepObserver = observer;
    }
})();

/*
 * ===== 検索条件「製造メーカー」入力可能コンボボックス =====
 * 表示用 input に文字を入力すると、候補一覧（部分一致・大文字小文字無視）を直下に表示する。
 * 確定値は非表示の select(#searchSeizomakerId) に反映し、送信値・グリッド編集候補の原本は従来のまま。
 * 候補に一致しない入力は、確定時に直前の選択値へ戻す（ID検索のため曖昧な値は送らない）。
 */
(function () {
    var input = document.getElementById('searchSeizomakerText');
    var list = document.getElementById('searchSeizomakerList');
    var source = document.getElementById('searchSeizomakerId');
    var form = document.getElementById('searchForm');
    if (!input || !list || !source) return;

    // select の option から候補 li を生成（先頭の空 option も「未選択」として残す）
    var items = Array.prototype.map.call(source.options, function (opt, i) {
        var li = document.createElement('li');
        li.id = 'searchSeizomakerOpt' + i;
        li.setAttribute('role', 'option');
        li.dataset.value = opt.value;
        li.textContent = opt.textContent;
        if (opt.value === '') li.innerHTML = '&nbsp;';
        list.appendChild(li);
        return li;
    });
    function labelOf(value) {
        var opt = Array.prototype.find.call(source.options, function (o) { return o.value === value; });
        return opt ? opt.textContent : '';
    }
    function syncFromSource() { input.value = labelOf(source.value); }

    function visibleItems() { return items.filter(function (li) { return !li.hidden; }); }
    function setActive(li) {
        items.forEach(function (o) { o.classList.remove('active'); o.removeAttribute('aria-selected'); });
        if (li) {
            li.classList.add('active'); li.setAttribute('aria-selected', 'true');
            input.setAttribute('aria-activedescendant', li.id);
            li.scrollIntoView({ block: 'nearest' });
        } else {
            input.removeAttribute('aria-activedescendant');
        }
    }
    function open(filterText) {
        var f = (filterText || '').trim().toLowerCase();
        var any = false;
        items.forEach(function (li) {
            // 絞り込み中は空の「未選択」候補を出さない
            var show = f === '' ? true : (li.dataset.value !== '' && li.textContent.toLowerCase().indexOf(f) >= 0);
            li.hidden = !show; if (show) any = true;
        });
        list.hidden = !any;
        input.setAttribute('aria-expanded', String(any));
        // 絞り込み時は先頭候補、未入力時は現在値をアクティブにする
        var vis = visibleItems();
        setActive(f === '' ? items.find(function (li) { return li.dataset.value === source.value; }) || null : (vis[0] || null));
    }
    function close() {
        list.hidden = true;
        input.setAttribute('aria-expanded', 'false');
        setActive(null);
    }
    function choose(li) {
        if (source.value !== li.dataset.value) {
            source.value = li.dataset.value;
            source.dispatchEvent(new Event('change', { bubbles: true }));
        }
        syncFromSource();
        close();
    }
    // 入力文字列を候補に確定させる（完全一致 → 絞り込み結果が1件 → 空欄はクリア → それ以外は元に戻す）
    function commit() {
        var text = input.value.trim();
        if (text === '') { choose(items[0]); return; }
        // 現在値の表示名のままなら変更しない（同名メーカーがあっても選択IDを維持）
        if (text === labelOf(source.value).trim()) { syncFromSource(); close(); return; }
        var lower = text.toLowerCase();
        var exact = items.filter(function (li) { return li.dataset.value !== '' && li.textContent.trim().toLowerCase() === lower; });
        if (exact.length >= 1) { choose(exact[0]); return; }
        var partial = items.filter(function (li) { return li.dataset.value !== '' && li.textContent.toLowerCase().indexOf(lower) >= 0; });
        if (partial.length === 1) { choose(partial[0]); return; }
        syncFromSource();
        close();
    }

    input.addEventListener('focus', function () { input.select(); open(''); });
    input.addEventListener('click', function () { if (list.hidden) open(''); });
    input.addEventListener('input', function () { open(input.value); });
    input.addEventListener('blur', commit);
    input.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') { syncFromSource(); close(); return; }
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
            e.preventDefault();
            if (list.hidden) { open(''); return; }
            var vis = visibleItems(); if (!vis.length) return;
            var idx = vis.findIndex(function (li) { return li.classList.contains('active'); });
            idx = e.key === 'ArrowDown' ? Math.min(idx + 1, vis.length - 1) : Math.max(idx - 1, 0);
            setActive(vis[idx]);
            return;
        }
        if (e.key === 'Enter') {
            // 一覧表示中の Enter は候補確定のみ（検索は実行しない）
            if (!list.hidden) {
                e.preventDefault();
                var active = items.find(function (li) { return li.classList.contains('active') && !li.hidden; });
                if (active) choose(active); else commit();
            } else {
                commit(); // 一覧非表示時はそのまま検索（送信前に確定）
            }
        }
    });
    // mousedown で確定（blur より先に処理し、入力欄のフォーカスを維持）
    list.addEventListener('mousedown', function (e) {
        var li = e.target.closest('li');
        e.preventDefault();
        if (li && !li.hidden) choose(li);
    });
    document.addEventListener('mousedown', function (e) {
        if (!e.target.closest || !e.target.closest('.combo-wrap')) close();
    });
    // 送信前の保険（キャプチャで既存の submit ハンドラより先に確定させる）
    if (form) form.addEventListener('submit', function () { if (document.activeElement === input) commit(); }, true);
    // ブラウザの戻る等でフォーム値が復元された場合も表示を合わせる
    window.addEventListener('pageshow', syncFromSource);

    syncFromSource();
})();
