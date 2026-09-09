/* Shared MCM message dialogs. PDF metadata is inserted by build-runtime.cjs. */
(function (window, document) {
    'use strict';
    if (window.McmNotifications) return;

    const definitions = [{"id":"MSG_0001","title":"検索エラー","icon":"Exclamation","text":"検索条件は、1項目以上選択して下さい。"},{"id":"MSG_0002","title":"検索エラー","icon":"Exclamation","text":"検索結果が1件も存在しません。"},{"id":"MSG_0003","title":"検索エラー","icon":"Exclamation","text":"検索結果から遷移したが、対応するデータが存在しない場合のエラー"},{"id":"MSG_0004","title":"削除エラー","icon":"Exclamation","text":"削除チェックがされているが、新規登録で遷移してきた場合のエラー"},{"id":"MSG_0005","title":"削除エラー","icon":"Exclamation","text":"値が変更されていません。"},{"id":"MSG_0006","title":"削除エラー","icon":"Exclamation","text":"既にプラント情報が登録されている為、削除する事が出来ません。"},{"id":"MSG_0007","title":"確認","icon":"Question","text":"登録を実施します。よろしいですか？"},{"id":"MSG_0008","title":"完了","icon":"Information","text":"登録を完了しました。"},{"id":"MSG_0009","title":"削除エラー","icon":"Exclamation","text":"{0}既に使用されている為、削除する事が出来ません。"},{"id":"MSG_0010","title":"入力エラー","icon":"Exclamation","text":"個体管理ありが有効な場合は、親機器分類コードを入力して下さい。"},{"id":"MSG_0011","title":"入力エラー","icon":"Exclamation","text":"ダイフクの情報は削除できません。"},{"id":"MSG_0012","title":"入力エラー","icon":"Exclamation","text":"保守時間は1～24までの数値を入力してください。"},{"id":"MSG_0013","title":"入力エラー","icon":"Exclamation","text":"開始時間、終了時間を入力してください。"},{"id":"MSG_0014","title":"入力エラー","icon":"Exclamation","text":"例外開始時間、例外終了時間を入力してください。"},{"id":"MSG_0015","title":"入力エラー","icon":"Exclamation","text":"例外曜日を入力してください。"},{"id":"MSG_0016","title":"入力エラー","icon":"Exclamation","text":"”9999/12/31”以降の期間は設定できません"},{"id":"MSG_0017","title":"入力エラー","icon":"Exclamation","text":"例外終了時間が例外開始時間よりも前の時間で入力されています"},{"id":"MSG_0018","title":"入力エラー","icon":"Exclamation","text":"ダイフク保守契約条件を選択してください。"},{"id":"MSG_0019","title":"入力エラー","icon":"Exclamation","text":"選択したダイフク保守契約条件が、保守曜日・時間を越えています。"},{"id":"MSG_0020","title":"入力エラー","icon":"Exclamation","text":"個体管理が必要な機器情報が一致していません。"},{"id":"MSG_0021","title":"入力エラー","icon":"Exclamation","text":"構成パターンのセット数と機器個体の数が一致していません。"},{"id":"MSG_0022","title":"削除エラー","icon":"Exclamation","text":"{0}既に登録されている為、削除する事が出来ません。"},{"id":"MSG_0023","title":"入力エラー","icon":"Exclamation","text":"見積依頼する機器を選定してください。"},{"id":"MSG_0024","title":"入力エラー","icon":"Exclamation","text":"保守契約時間帯を選択してください。"},{"id":"MSG_0025","title":"排他エラー","icon":"Exclamation","text":"店舗への見積依頼 又は 取引先との契約として使用されている為、破棄する事が出来ません。"},{"id":"MSG_0026","title":"排他エラー","icon":"Exclamation","text":"取引先への支払が発生している為、破棄する事が出来ません。"},{"id":"MSG_0027","title":"入力エラー","icon":"Exclamation","text":"開始日は終了日よりも前の日付で入力して下さい。"},{"id":"MSG_0028","title":"入力エラー","icon":"Exclamation","text":"数量は1以上の値を入力して下さい。"},{"id":"MSG_0029","title":"入力エラー","icon":"Exclamation","text":"出精値引きのある見積を選択されています。期間の設定を確認してください。"},{"id":"MSG_0030","title":"入力エラー","icon":"Exclamation","text":"支払が発生済み期間の機器情報を変更することは出来ません。"},{"id":"MSG_0031","title":"入力エラー","icon":"Exclamation","text":"コントローラのチェックがされていますが、金額が入力されていません。"},{"id":"MSG_0032","title":"入力エラー","icon":"Exclamation","text":"金額が入力されていますが、コントローラのチェックがされていません。"},{"id":"MSG_0033","title":"入力エラー","icon":"Exclamation","text":"重複したログインIDが存在します。"},{"id":"MSG_0034","title":"入力エラー","icon":"Exclamation","text":"個体管理が必要な機器を必ず一つ入力してください。"},{"id":"MSG_0035","title":"入力エラー","icon":"Exclamation","text":"契約可能期間を入力した場合は、起算日区分も入力してください。"},{"id":"MSG_0036","title":"入力エラー","icon":"Exclamation","text":"起算日区分を入力した場合は、契約可能期間も入力してください。"},{"id":"MSG_0037","title":"入力エラー","icon":"Exclamation","text":"１件もチェックされていません。"},{"id":"MSG_0038","title":"入力エラー","icon":"Exclamation","text":"選択した見積内に重複した機器が存在します。"},{"id":"MSG_0039","title":"入力エラー","icon":"Exclamation","text":"上位者IDに対応する上位者のログインIDが存在しません。"},{"id":"MSG_0040","title":"入力エラー","icon":"Exclamation","text":"タブを追加するには期間の終了日の入力が必要です。"},{"id":"MSG_0041","title":"入力エラー","icon":"Exclamation","text":"期間は重複しないように入力してください。"},{"id":"MSG_0042","title":"入力エラー","icon":"Exclamation","text":"空いている日付が存在しないように期間を指定してください。"},{"id":"MSG_0043","title":"入力エラー","icon":"Exclamation","text":"保守時間が24の場合、開始時間、終了時間、例外曜日、例外開始時間、例外終了時間は入力しないでください。"},{"id":"MSG_0044","title":"入力エラー","icon":"Exclamation","text":"構成を成すセットの中で、ブランドが異なっています。"},{"id":"MSG_0045","title":"入力エラー","icon":"Exclamation","text":"{0}を先に入力してください。"},{"id":"MSG_0046","title":"入力エラー","icon":"Exclamation","text":"指定されたデータのプラントIDが見つかりません"},{"id":"MSG_0047","title":"入力エラー","icon":"Exclamation","text":"付替え元が設定されていません。"},{"id":"MSG_0048","title":"入力エラー","icon":"Exclamation","text":"付替え先が設定されていません。"},{"id":"MSG_0049","title":"入力エラー","icon":"Exclamation","text":"ブランド情報が設定されていません。"},{"id":"MSG_0050","title":"入力エラー","icon":"Exclamation","text":"セット数は1以上の値を入力して下さい。"},{"id":"MSG_0051","title":"入力エラー","icon":"Exclamation","text":"機器構成パターン情報のセット数と機器個体情報の数量が一致しません。"},{"id":"MSG_0052","title":"入力エラー","icon":"Exclamation","text":"機器明細情報に値が存在しない為、更新できません。"},{"id":"MSG_0053","title":"入力エラー","icon":"Exclamation","text":"機器個体管理情報に値が存在しない為、更新できません。"},{"id":"MSG_0054","title":"入力エラー","icon":"Exclamation","text":"必須項目を入力して下さい。"},{"id":"MSG_0055","title":"入力エラー","icon":"Exclamation","text":"セット数と一致しない為、削除できません。"},{"id":"MSG_0056","title":"入力エラー","icon":"Exclamation","text":"個体管理情報が重複しています。"},{"id":"MSG_0057","title":"入力エラー","icon":"Exclamation","text":"取引先見積情報の行がチェックされていません。"},{"id":"MSG_0058","title":"入力エラー","icon":"Exclamation","text":"開始日付の書式を指定して下さい。(YYYY/MM/DD)"},{"id":"MSG_0059","title":"入力エラー","icon":"Exclamation","text":"終了日付の書式を指定して下さい。(YYYY/MM/DD)"},{"id":"MSG_0060","title":"入力エラー","icon":"Exclamation","text":"同じ分類となる機器は複数登録できません。"},{"id":"MSG_0061","title":"入力エラー","icon":"Exclamation","text":"既に同じ型式が指定されています。"},{"id":"MSG_0062","title":"確認","icon":"Question","text":"見積を依頼します。よろしいですか？"},{"id":"MSG_0063","title":"確認","icon":"Question","text":"見積を複製します。よろしいですか？"},{"id":"MSG_0064","title":"確認","icon":"Question","text":"見積を破棄します。よろしいですか？"},{"id":"MSG_0065","title":"確認","icon":"Question","text":"契約を作成します。よろしいですか？"},{"id":"MSG_0066","title":"確認","icon":"Question","text":"契約を破棄します。よろしいですか？"},{"id":"MSG_0067","title":"入力エラー","icon":"Exclamation","text":"納入先情報が存在しません。"},{"id":"MSG_0068","title":"入力エラー","icon":"Exclamation","text":"個体管理ありを無効にすることはできません。"},{"id":"MSG_0069","title":"入力エラー","icon":"Exclamation","text":"状態が{0}のため、更新できません。"},{"id":"MSG_0070","title":"完了","icon":"Information","text":"登録を完了しました。"},{"id":"MSG_0071","title":"入力エラー","icon":"Exclamation","text":"見積資料を添付してください。"},{"id":"MSG_0072","title":"完了","icon":"Information","text":"ファイルの追加を完了しました。"},{"id":"MSG_0073","title":"完了","icon":"Information","text":"ファイルの削除を完了しました。"},{"id":"MSG_0074","title":"入力エラー","icon":"Exclamation","text":"店舗見積を作成する機器を選定してください。"},{"id":"MSG_0075","title":"入力エラー","icon":"Exclamation","text":"店舗見積・ユーザ契約作成検索画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。"},{"id":"MSG_0076","title":"入力エラー","icon":"Exclamation","text":"{0}から{1}の間の日付を設定してください。"},{"id":"MSG_0077","title":"情報","icon":"Information","text":"{0}タブの終了日は{1}に変更されます。"},{"id":"MSG_0078","title":"入力エラー","icon":"Exclamation","text":"店舗見積を選択してください。"},{"id":"MSG_0079","title":"入力エラー","icon":"Exclamation","text":"選択された見積には、出精値引きが含まれている為、見積の期間と契約の期間が一致する必要があります。\n期間を確認してください。"},{"id":"MSG_0080","title":"入力エラー","icon":"Exclamation","text":"選択された見積には、契約の期間に対応する見積情報が存在しません。\n見積内容を確認、または期間の設定を変更してください。"},{"id":"MSG_0081","title":"入力エラー","icon":"Exclamation","text":"前画面で、選択した取引先見積に含まれていない為、選択する事は出来ません。"},{"id":"MSG_0082","title":"入力エラー","icon":"Exclamation","text":"選択された見積には、契約の期間に対応する見積情報が存在しません。\n見積内容を確認、または期間の設定を変更してください。"},{"id":"MSG_0083","title":"入力エラー","icon":"Exclamation","text":"チェック不足のエラー"},{"id":"MSG_0084","title":"入力エラー","icon":"Exclamation","text":"ブランド名称が未入力です。"},{"id":"MSG_0085","title":"入力エラー","icon":"Exclamation","text":"取引先名称が未入力です。"},{"id":"MSG_0086","title":"入力エラー","icon":"Exclamation","text":"付替え先ブランドが選択されていないデータがあります。"},{"id":"MSG_0087","title":"情報","icon":"Information","text":"{0}タブの開始日は{1}に変更されます。"},{"id":"MSG_0088","title":"入力エラー","icon":"Information","text":"審査・承認可能な申請が選定されていません。"},{"id":"MSG_0089","title":"確認","icon":"Question","text":"添付ファイルを削除します。よろしいですか？"},{"id":"MSG_0090","title":"完了","icon":"Information","text":"審査・承認を完了しました。"},{"id":"MSG_0091","title":"入力エラー","icon":"Exclamation","text":"機器構成情報のコントローラが有効ではありません。"},{"id":"MSG_0092","title":"入力エラー","icon":"Exclamation","text":"機器構成情報のコントローラを無効にして下さい。"},{"id":"MSG_0093","title":"入力エラー","icon":"Exclamation","text":"機器明細情報と一致していないデータが個体管理情報に存在します。"},{"id":"MSG_0094","title":"確認","icon":"Question","text":"審査・承認処理を実施します。よろしいですか？"},{"id":"MSG_0095","title":"確認","icon":"Question","text":"差し戻し処理を実施します。よろしいですか？"},{"id":"MSG_0096","title":"完了","icon":"Information","text":"差戻しを完了しました。"},{"id":"MSG_0097","title":"入力エラー","icon":"Exclamation","text":"サポートIDは必ず入力してください。"},{"id":"MSG_0098","title":"確認","icon":"Question","text":"付替えを完了しました。付替え元、付替え先をクリアしてよろしいですか？"},{"id":"MSG_0099","title":"確認","icon":"Question","text":"付替えしてもよろしいですか？"},{"id":"MSG_0100","title":"入力エラー","icon":"Exclamation","text":"ブランド”{0}”に紐付く個体の中で機器構成”{1}”を形成できない設定があるため、付替えすることができません。\n機器構成”{1}”の内容を確認してください。"},{"id":"MSG_0101","title":"確認","icon":"Question","text":"機器個体の契約期限、契約満了日を設定します。よろしいですか？"},{"id":"MSG_0102","title":"確認","icon":"Question","text":"新規登録の為、機器個体の契約期限、契約満了日を設定します。\n一次納入日、二次納入日、初回契約日の設定はされていますか？\n登録を実施します。よろしいですか？"},{"id":"MSG_0103","title":"入力エラー","icon":"Exclamation","text":"点検回数と月のチェックの数が一致していません。"},{"id":"MSG_0104","title":"入力エラー","icon":"Exclamation","text":"支払回数と月のチェックの数が一致していません。"},{"id":"MSG_0105","title":"確認","icon":"Question","text":"期間の開始月に支払チェックがされてません。処理を続行してもよろしいですか？"},{"id":"MSG_0106","title":"入力エラー","icon":"Exclamation","text":"承認済みのファイルの為、削除できません。"},{"id":"MSG_0107","title":"入力エラー","icon":"Exclamation","text":"承認済みな契約手続依頼書が存在する為、再発行できません。\n期間タブを作成し直して下さい。"},{"id":"MSG_0108","title":"入力エラー","icon":"Exclamation","text":"再申請はできません。期間タブを作成し直して下さい。"},{"id":"MSG_0109","title":"入力エラー","icon":"Exclamation","text":"該当ファイルが存在しません。"},{"id":"MSG_0110","title":"情報","icon":"Exclamation","text":"ブランド構成が設定されていない為、機器構成の設定画面には遷移することは出来ません。"},{"id":"MSG_0111","title":"入力エラー","icon":"Exclamation","text":"ユーザ契約されているため、マスタ情報は変更できません。\nユーザー契約情報を変更して下さい。"},{"id":"MSG_0112","title":"入力エラー","icon":"Exclamation","text":"契約依頼資料を添付してください。"},{"id":"MSG_0113","title":"入力エラー","icon":"Exclamation","text":"品名：”{0}”\n型式：”{1}”\nに該当する親機器が見つかりません。"},{"id":"MSG_0114","title":"入力エラー","icon":"Exclamation","text":"ブランド構成情報を削除して下さい。"},{"id":"MSG_0115","title":"入力エラー","icon":"Exclamation","text":"{0}タブの終了日が変更不可の為、開始日を変更する事はできません。"},{"id":"MSG_0116","title":"入力エラー","icon":"Exclamation","text":"{0}タブの終了日は必須項目です。"},{"id":"MSG_0117","title":"確認","icon":"Question","text":"開始日に過去日を設定すると、登録後に開始日が編集出来なくなります。"},{"id":"MSG_0118","title":"確認","icon":"Question","text":"未登録タブのため、変更を行うと現在のタブの情報が変更内容に差し変わります。"},{"id":"MSG_0119","title":"入力エラー","icon":"Exclamation","text":"出精値引きがある場合は、期間の変更は出来ません。"},{"id":"MSG_0120","title":"確認","icon":"Question","text":"申請を実施します。よろしいですか？"},{"id":"MSG_0121","title":"完了","icon":"Information","text":"申請を完了しました。"},{"id":"MSG_0122","title":"入力エラー","icon":"Exclamation","text":"契約金額変動時期は、1件以上選択してください。"},{"id":"MSG_0123","title":"入力エラー","icon":"Exclamation","text":"期間は1年以内で設定してください。"},{"id":"MSG_0124","title":"入力エラー","icon":"Exclamation","text":"付替え対象が検索されていません。"},{"id":"MSG_0125","title":"入力エラー","icon":"Exclamation","text":"プラント情報にデータが存在しません。"},{"id":"MSG_0126","title":"入力エラー","icon":"Exclamation","text":"差戻し可能な申請が選択されていません。"},{"id":"MSG_0127","title":"入力エラー","icon":"Exclamation","text":"付替え元と付替え先は異なるデータを指定して下さい。"},{"id":"MSG_0128","title":"入力エラー","icon":"Exclamation","text":"DTSマスタ連携ログを検索する場合は、ログ出力日を入力して下さい。"},{"id":"MSG_0129","title":"入力エラー","icon":"Exclamation","text":"親機器分類コードが個体管理ありの機器分類コードに存在しません。"},{"id":"MSG_0130","title":"入力エラー","icon":"Exclamation","text":"付替え元が１件もチェックされていません。"},{"id":"MSG_0131","title":"入力エラー","icon":"Exclamation","text":"取引先見積情報が存在しません。"},{"id":"MSG_0132","title":"確認","icon":"Question","text":"付替えにより、プラントが複数となる契約が存在します。\n処理を続行してよろしいですか？"},{"id":"MSG_0133","title":"入力エラー","icon":"Exclamation","text":"機器情報が存在しない為、付け替え元に設定することは出来ません。"},{"id":"MSG_0134","title":"入力エラー","icon":"Exclamation","text":"既に機器明細情報で使用されているため、変更できません。"},{"id":"MSG_0135","title":"入力エラー","icon":"Exclamation","text":"既に取扱機器構成情報で使用されているため、変更できません。"},{"id":"MSG_0136","title":"入力エラー","icon":"Exclamation","text":"既に取引先契約点検情報で使用されているため、変更できません。"}];
    const escapeRegex = text => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const patterns = definitions.map(def => ({
        ...def,
        pattern: new RegExp('^' + def.text.split(/(\{\d+\})/).map(part =>
            /^\{\d+\}$/.test(part) ? '[\\s\\S]*?' : escapeRegex(part)).join('') + '$')
    }));
    const queue = [];
    let active = null;
    let dialog, title, message, icon, ok, cancel;

    function metadata(text, kind) {
        const definition = patterns.find(def => def.pattern.test(text));
        return {
            title: definition ? definition.title : kind === 'error' ? 'エラー' : kind === 'confirm' ? '確認' : '情報',
            icon: definition ? definition.icon : kind === 'error' ? 'Exclamation' : kind === 'confirm' ? 'Question' : 'Information'
        };
    }

    function element(tag, className, text) {
        const el = document.createElement(tag);
        el.className = className;
        if (text !== undefined) el.textContent = text;
        return el;
    }

    function ensureDialog() {
        if (dialog) return;
        dialog = element('dialog', 'mcm-notification');
        dialog.id = 'mcmNotificationDialog';
        dialog.setAttribute('role', 'alertdialog');
        dialog.setAttribute('aria-modal', 'true');
        dialog.setAttribute('aria-labelledby', 'mcmNotificationTitle');
        dialog.setAttribute('aria-describedby', 'mcmNotificationMessage');
        const head = element('div', 'mcm-notification-head');
        title = element('h2', 'mcm-notification-title');
        title.id = 'mcmNotificationTitle';
        const close = element('button', 'mcm-notification-close', '×');
        close.type = 'button';
        close.setAttribute('aria-label', '閉じる');
        close.addEventListener('click', () => finish(false));
        head.append(title, close);
        const body = element('div', 'mcm-notification-body');
        icon = element('span', 'mcm-notification-icon');
        icon.setAttribute('aria-hidden', 'true');
        message = element('p', 'mcm-notification-message');
        message.id = 'mcmNotificationMessage';
        body.append(icon, message);
        const foot = element('div', 'mcm-notification-foot');
        ok = element('button', 'mcm-notification-button', 'OK');
        cancel = element('button', 'mcm-notification-button', 'キャンセル');
        ok.type = cancel.type = 'button';
        ok.addEventListener('click', () => finish(true));
        cancel.addEventListener('click', () => finish(false));
        foot.append(ok, cancel);
        dialog.append(head, body, foot);
        // Native modal dialogs block pointer/focus access to the page underneath.
        dialog.addEventListener('cancel', event => { event.preventDefault(); finish(false); });
        dialog.addEventListener('keydown', event => {
            // Do not let grid shortcuts or a surrounding screen's Enter handler run.
            event.stopPropagation();
            if (event.key === 'Enter' && event.target === dialog) {
                event.preventDefault();
                finish(true);
            }
        });
        document.body.appendChild(dialog);
    }

    function next() {
        if (active || !queue.length) return;
        ensureDialog();
        active = queue.shift();
        window.McmCommon?.hideProgress(true);
        title.textContent = active.meta.title;
        message.textContent = active.text;
        icon.dataset.icon = active.meta.icon;
        icon.textContent = active.meta.icon === 'Exclamation' ? '!' : active.meta.icon === 'Question' ? '?' : 'i';
        cancel.hidden = active.kind !== 'confirm';
        dialog.showModal();
        ok.focus({preventScroll: true});
    }

    function finish(accepted) {
        if (!active) return;
        const current = active;
        active = null;
        dialog.close();
        // Close/X/Escape dismiss an OK notice; they cancel a confirmation.
        current.resolve(current.kind === 'confirm' ? accepted : true);
        if (current.focus && current.focus.isConnected && !current.focus.disabled) {
            current.focus.focus({preventScroll: true});
        }
        queueMicrotask(next);
    }

    function show(text, kind) {
        text = text == null ? '' : String(text);
        if (!text.trim()) return Promise.resolve(false);
        return new Promise(resolve => {
            queue.push({text, kind, meta: metadata(text, kind), resolve, focus: document.activeElement});
            next();
        });
    }

    function collect(root) {
        const unique = new Set();
        const groups = new Map();
        root.querySelectorAll('[data-mcm-notifications]:not([data-mcm-consumed])').forEach(container => {
            container.dataset.mcmConsumed = 'true';
            container.querySelectorAll('[data-mcm-notice]').forEach(node => {
                const text = node.textContent;
                if (!text.trim()) return;
                const kind = node.dataset.mcmNotice;
                if (unique.has(kind + '\0' + text)) return;
                unique.add(kind + '\0' + text);
                const meta = metadata(text, kind);
                const key = kind + '\0' + meta.title + '\0' + meta.icon;
                if (!groups.has(key)) groups.set(key, {kind, meta, texts: []});
                groups.get(key).texts.push(text);
            });
        });
        for (const group of groups.values()) {
            const text = group.texts.join('\n');
            queue.push({text, kind: group.kind, meta: group.meta, resolve: () => {}, focus: document.activeElement});
        }
        next();
    }

    window.McmNotifications = Object.freeze({
        info: text => show(text, 'info'),
        error: text => show(text, 'error'),
        confirm: text => show(text, 'confirm'),
        collect: () => collect(document)
    });
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => collect(document), {once: true});
    } else {
        collect(document);
    }
})(window, document);
