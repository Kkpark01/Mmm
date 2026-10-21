/**
 * 【共通JS】MCM全画面共通ユーティリティ
 *
 * 各画面(mcm00xxu/index.html等)がコピペで個別実装していた処理のうち、
 * 全画面横断で使うべきものをここに集約する。
 *
 * 提供機能:
 *   1. プログレスインジケーター（サーバー処理中の画面ブロック）
 *   2. フォーム送信/fetch送信への組み込みヘルパー
 *
 * 使い方（各画面側）:
 *   <link rel="stylesheet" th:href="@{/css/common.css}"/>
 *   <script th:src="@{/js/common.js}"></script>
 *   ...
 *   McmCommon.withProgress(function () {
 *       form.submit();  // または fetch(...).then(...).finally(McmCommon.hideProgress)
 *   });
 *
 * 注意:
 *   通常のフォーム送信(form.submit())はページ遷移を伴うため、
 *   オーバーレイは新しいページが読み込まれるまで表示され続ける（＝隠す必要がない）。
 *   fetch送信の場合は、応答後に必ず McmCommon.hideProgress() を呼ぶこと。
 */
(function (window, document) {
    "use strict";

    var OVERLAY_ID = "mcmProgressOverlay";
    var overlayEl = null;
    var activeCount = 0; // 多重呼び出し（fetchの並列実行等）に対応するための参照カウント

    /** オーバーレイ要素を取得する。無ければ生成してbodyに追加する。 */
    function ensureOverlay() {
        if (overlayEl && document.body.contains(overlayEl)) return overlayEl;
        overlayEl = document.getElementById(OVERLAY_ID);
        if (!overlayEl) {
            overlayEl = document.createElement("div");
            overlayEl.id = OVERLAY_ID;
            overlayEl.className = "progress-overlay";
            overlayEl.setAttribute("aria-hidden", "true");
            var spinner = document.createElement("div");
            spinner.className = "progress-spinner";
            overlayEl.appendChild(spinner);
            document.body.appendChild(overlayEl);
        }
        return overlayEl;
    }

    /**
     * プログレスオーバーレイを表示し、画面全体の操作を無効化する。
     * 参照カウント方式のため、複数箇所から呼ばれても最後の hideProgress() までは表示され続ける。
     */
    function showProgress() {
        activeCount++;
        var el = ensureOverlay();
        el.classList.add("is-active");
    }

    /**
     * プログレスオーバーレイを隠す。
     * showProgress() の呼び出し回数分 hideProgress() が呼ばれるまでは隠れない。
     * @param {boolean} [force] true の場合、参照カウントを無視して強制的に隠す（エラー時の保険用）。
     */
    function hideProgress(force) {
        if (force) {
            activeCount = 0;
        } else if (activeCount > 0) {
            activeCount--;
        }
        if (activeCount <= 0) {
            activeCount = 0;
            var el = ensureOverlay();
            el.classList.remove("is-active");
        }
    }

    /**
     * 処理を実行しながらプログレスオーバーレイを表示する。
     * 処理が Promise を返す場合は、その完了/失敗時に自動で非表示にする。
     * 通常の form.submit()（ページ遷移）のように戻り値が無い処理では、
     * 呼び出し側で明示的に hideProgress() を呼ぶ必要はない（ページ遷移で消える）。
     *
     * @param {Function} action 実行する処理。Promiseを返してもよい。
     */
    function withProgress(action) {
        showProgress();
        var result;
        try {
            result = action();
        } catch (e) {
            hideProgress();
            throw e;
        }
        if (result && typeof result.then === "function") {
            result.then(
                function (v) { hideProgress(); return v; },
                function (e) { hideProgress(); throw e; }
            );
        }
        return result;
    }

    // =========================================================================
    // 全画面への自動適用（オートフック）
    // 【目的】各画面のJSを個別に書き換えずに済むよう、通常のフォーム送信(submit)と
    //   fetch通信を自動的に検知してプログレスオーバーレイを表示する。
    //   画面側は <script th:src="@{/js/common.js}"></script> を読み込むだけでよい。
    // =========================================================================

    /**
     * 通常の <form> submit（ページ遷移を伴う送信）を捕捉してオーバーレイを表示する。
     * capture:true でバブリング前に検知するため、各画面が e.preventDefault() で
     * 独自にAjax化している場合でも submit イベント自体は発火するため影響はない。
     * ページ遷移が発生する送信ではオーバーレイは新しいページ読込まで表示されたままでよい。
     */
    document.addEventListener("submit", function (e) {
        var form = e.target;
        if (!form || form.tagName !== "FORM") return;
        // data-mcm-no-progress を明示的に付与したフォームは対象外にする（検索フォーム等の高速処理向け）
        if (form.hasAttribute("data-mcm-no-progress")) return;
        // 【不具合修正】別タブ／別ウィンドウへの送信（formtarget / target 指定）は自画面が遷移しないため、
        //   オーバーレイを表示すると消えずに操作不能になる（例: MCM1004Uの「照会」→MCM1007U）。
        if (isOtherWindowSubmit(e)) return;
        showProgress();
    }, true);

    /**
     * submit イベントの送信先が自画面以外（別タブ／別ウィンドウ／別フレーム）かを判定する。
     * 押下ボタンの formtarget を優先し、無ければ form の target を見る。
     */
    function isOtherWindowSubmit(e) {
        var form = e.target;
        var submitter = e.submitter;
        var target = submitter && submitter.hasAttribute && submitter.hasAttribute("formtarget")
            ? submitter.getAttribute("formtarget")
            : form.getAttribute("target");
        target = (target || "").toLowerCase();
        return target !== "" && target !== "_self";
    }

    /**
     * window.fetch をラップし、通信中は自動的にプログレスオーバーレイを表示する。
     * 各画面の fetch() 呼び出しコードは変更不要（透過的に効く）。
     * リクエスト初期化オプションに {mcmNoProgress: true} を指定した場合は対象外にする。
     */
    if (window.fetch) {
        var originalFetch = window.fetch;
        window.fetch = function (input, init) {
            var skip = init && init.mcmNoProgress === true;
            if (skip) return originalFetch.apply(window, arguments);
            showProgress();
            var promise = originalFetch.apply(window, arguments);
            promise.then(function () { hideProgress(); }, function () { hideProgress(); });
            return promise;
        };
    }

    // =========================================================================
    // 離脱確認ダイアログ（beforeunload）の共通化
    //
    // 【目的】入力値を変更した状態でブラウザのタブ「×」やブラウザ終了を行った場合に、
    //   離脱確認ダイアログを表示して意図しないデータ破棄を防ぐ。
    //   「閉じる」ボタン押下時の独自モーダル（「データが変更されています。破棄されますが
    //   よろしいですか？」）は各画面に実装済みだが、タブの「×」はブラウザ側の操作であり
    //   JSから捕捉できないため、これまで何の確認も行われずに破棄されていた。
    //
    // 【制約】表示文言はブラウザ既定のもの（例: Chrome「このサイトを離れますか？」）となり、
    //   独自文言は指定できない（Chrome 51 / Firefox 44 以降は returnValue が無視される）。
    //   そのため各画面の独自モーダルと文言は一致しない。これは仕様上の限界である。
    //   また、そのページで一度もユーザー操作が行われていない場合はブラウザが
    //   ダイアログをスキップするため、入力していない状態では表示されない。
    //
    // 【「閉じる」ボタン等の既存挙動を変えないための設計】
    //   画面内の意図した遷移（リンククリック、フォーム送信、「閉じる」ボタン）でも
    //   beforeunload は発火するため、そのまま実装すると独自モーダルとブラウザ既定
    //   ダイアログが二重に表示され、既存の「閉じる」挙動が変わってしまう。
    //   これを防ぐため isLeaving フラグで意図した遷移を判別し、その場合は
    //   離脱確認を抑止する（＝既存の挙動は一切変わらない）。
    //
    //   isLeaving は「タイマーで自動解除」ではなく「次のユーザー入力操作で解除」する。
    //   独自モーダルでの確認が何秒かかっても二重表示が起きず、かつ
    //   モーダルをキャンセルして編集を再開した場合は再び離脱確認が有効になる。
    // =========================================================================

    /** ユーザーが入力値を変更したか（＝未保存データが存在するか） */
    var isDirty = false;
    /** 画面内の意図した遷移（リンク／submit／閉じる）が開始されたか */
    var isLeaving = false;

    /** dirty 判定の対象とするフォーム部品のタグ名 */
    var DIRTY_TARGET_TAGS = { INPUT: true, SELECT: true, TEXTAREA: true };
    /** dirty 判定の対象外とする input の type（ユーザーの入力値変更ではないもの） */
    var DIRTY_IGNORE_INPUT_TYPES = {
        hidden: true, submit: true, button: true, reset: true, image: true
    };

    /**
     * 離脱確認の対象外要素かどうかを判定する。
     * 検索条件のみを入力して離脱する場合などに離脱確認を出したくない画面は、
     * 対象の入力部品またはその祖先要素に data-mcm-ignore-dirty を付与する。
     *
     * @param {Element} el 判定対象の要素
     * @return {boolean} true の場合は dirty 判定の対象外
     */
    function isIgnoredForDirty(el) {
        if (!el || !DIRTY_TARGET_TAGS[el.tagName]) return true;
        if (el.tagName === "INPUT") {
            var type = (el.getAttribute("type") || "text").toLowerCase();
            if (DIRTY_IGNORE_INPUT_TYPES[type]) return true;
        }
        if (el.closest && el.closest("[data-mcm-ignore-dirty]")) return true;
        return false;
    }

    /** 未保存データが存在する状態にする（画面側から明示的に呼ぶことも可能）。 */
    function markDirty() {
        isDirty = true;
        isLeaving = false;
    }

    /** 未保存データなしの状態に戻す（登録完了後など、画面側から呼ぶ）。 */
    function markClean() {
        isDirty = false;
    }

    /**
     * 画面内の意図した遷移が開始されたことを記録し、離脱確認を抑止する。
     * 独自の確認モーダルを経てページ遷移する画面は、遷移直前にこれを呼ぶことで
     * ブラウザ既定ダイアログとの二重表示を防げる（closeOrNavigate は内部で呼んでいる）。
     */
    function markLeaving() {
        isLeaving = true;
    }

    /**
     * タブを閉じる。閉じられなかった場合は離脱確認を再度有効に戻す。
     *
     * 【必要な理由】target="_blank" で開いたタブは「スクリプトが開いたウィンドウ」
     *   ではないため、そのタブ内で一度でも画面遷移するとブラウザが window.close() を
     *   拒否する（Chrome/Edge は "Scripts may close only windows that were opened
     *   by them." を出して何も起きない）。
     *   このとき離脱抑止フラグを立てたままにすると、閉じられていないのに
     *   その後のタブの × で離脱確認が出なくなる。
     */
    function closeTab() {
        markLeaving();
        window.close();
        // close() 自体は離脱確認ダイアログを表示しないため、
        // 閉じられなかった場合に抑止を解除しても副作用はない。
        setTimeout(function () {
            if (!window.closed) isLeaving = false;
        }, 0);
    }

    /**
     * ユーザーによる入力値変更を検知して未保存状態にする。
     * isTrusted でユーザー操作由来のイベントに限定するため、画面描画時の
     * プログラム的な値設定や合成イベントでは未保存扱いにならない。
     */
    function onUserEdit(e) {
        if (!e.isTrusted) return;
        if (isIgnoredForDirty(e.target)) return;
        markDirty();
    }

    document.addEventListener("input", onUserEdit, true);
    document.addEventListener("change", onUserEdit, true);

    /**
     * 画面内のリンククリックによる遷移を意図した遷移として記録する。
     * target="_blank" のリンクは自画面が残るため離脱ではなく対象外とする。
     *
     * 【preventDefault() 済みを対象外とする理由】
     *   preventDefault() されている場合は、そのクリックで遷移は発生しない。
     *   独自の確認モーダルを表示して後から遷移する画面や、「閉じる」押下時に
     *   サーバーエラーを表示して画面に留まる場合が該当する。
     *   ここで記録してしまうと、画面に留まったまま isLeaving が立ち続け、
     *   その後タブの × を押しても離脱確認が出なくなる。
     *
     *   独自モーダルのOK後に遷移する画面は、遷移直前に
     *   McmCommon.closeOrNavigate(href) または McmCommon.markLeaving() を
     *   呼ぶことでブラウザ既定ダイアログとの二重表示を防ぐ。
     */
    document.addEventListener("click", function (e) {
        if (e.defaultPrevented) return;
        var link = e.target.closest ? e.target.closest("a[href]") : null;
        if (!link) return;
        var target = (link.getAttribute("target") || "").toLowerCase();
        if (target && target !== "_self") return;          // 別タブ／別フレームは自画面が残る
        if (link.hasAttribute("download")) return;         // ダウンロードは遷移しない
        var href = link.getAttribute("href");
        if (!href || href.charAt(0) === "#") return;       // ページ内アンカーは遷移しない
        if (/^javascript:/i.test(href)) return;
        markLeaving();
    }, false);

    /**
     * フォーム送信による遷移を意図した遷移として記録する。
     * 各画面が preventDefault() で独自にAjax化している場合は遷移しないため対象外。
     * requestSubmit() は submit イベントを発火するためここで捕捉できる。
     */
    document.addEventListener("submit", function (e) {
        if (e.defaultPrevented) return;
        // 別タブ／別ウィンドウへの送信は自画面が残るため、離脱扱いにしない（離脱確認を無効化しない）
        if (e.target && e.target.tagName === "FORM" && isOtherWindowSubmit(e)) return;
        markLeaving();
    }, false);

    /**
     * form.submit() をラップして意図した遷移として記録する。
     *
     * 【必要な理由】JSから直接呼ぶ form.submit() は submit イベントを発火しないため、
     *   上記の submit リスナーでは捕捉できない。確認ダイアログのOK後に
     *   form.submit() で登録する画面（mcm0021u/0025u/0026u/1003u/1011u/2002u/2008u/3001u 等）
     *   が多数あり、捕捉できないと登録操作のたびにブラウザ既定の離脱確認が
     *   重なって表示され、既存の挙動が変わってしまう。
     */
    if (window.HTMLFormElement && HTMLFormElement.prototype.submit) {
        var originalFormSubmit = HTMLFormElement.prototype.submit;
        HTMLFormElement.prototype.submit = function () {
            markLeaving();
            return originalFormSubmit.apply(this, arguments);
        };
    }

    /**
     * history.back() / forward() / go() をラップして意図した遷移として記録する。
     * 「閉じる」「戻る」を history.back() で実装している画面（mcm1006u/step1 等）向け。
     */
    if (window.history && window.History && History.prototype) {
        ["back", "forward", "go"].forEach(function (name) {
            var original = History.prototype[name];
            if (typeof original !== "function") return;
            History.prototype[name] = function () {
                markLeaving();
                return original.apply(this, arguments);
            };
        });
    }

    /**
     * 未保存データがあり、かつ意図した遷移でない場合（＝タブの×、ブラウザ終了、
     * アドレスバーからの遷移など）に離脱確認ダイアログを要求する。
     * 表示される文言はブラウザ既定のものとなる。
     */
    window.addEventListener("beforeunload", function (e) {
        if (isLeaving || !isDirty) return;
        e.preventDefault();
        e.returnValue = "";
    });

    // =========================================================================
    // 【不具合修正】メニュー画面（MCM0001U）から target="_blank" で開いた
    //   業務画面タブの「閉じる」ボタン押下時の処理。
    //
    // 【現象】メニューの各ボタンは target="_blank" で別タブとして開くが、
    //   各画面の「閉じる」ボタンは <a href="/menu"> 等の通常のGETリンクに
    //   すぎず window.close() を呼んでいなかった。そのため「閉じる」を押すと
    //   タブが閉じずにそのタブ内で /menu へページ遷移してしまい、結果的に
    //   別タブ側のセッション状態によってはログイン画面へ遷移する不具合が発生していた。
    //
    // 【対応】メニューから開かれたタブ（window.opener が存在する）かどうかを判定し、
    //   そうであれば通常のページ遷移の代わりに window.close() でタブ自体を閉じる、
    //   という判定ロジックを McmCommon.closeOrNavigate(href) として共通化する。
    //
    //   ・ログアウトボタン（/menu/logout へのリンク。メニュー画面自身の「閉じる」）は対象外とする。
    //   ・window.opener が無い場合（メニュー経由以外での直接アクセス等）は、
    //     従来通り href 先へのページ遷移をそのまま行う（既存仕様を変更しない）。
    // =========================================================================

    /**
     * メニューから開かれたタブであれば window.close() でタブを閉じ、
     * そうでなければ通常のページ遷移（location.href = href）を行う。
     *
     * 画面側で「未保存の変更がある場合のみ独自の確認ダイアログを出し、
     * OK後に window.location.href = href で遷移する」実装（非同期モーダル等）
     * がある場合は、その遷移直前でこの関数を呼ぶことで、遷移とタブクローズの
     * 判定を一箇所に統一できる。
     *
     * @param {string} href 遷移先URL（閉じるボタンの本来の遷移先）
     */
    function closeOrNavigate(href) {
        // 独自モーダルで確認済みの意図した離脱であるため、
        // ブラウザ既定の離脱確認ダイアログ（beforeunload）が二重に出ないよう抑止する。
        if (window.opener && !window.opener.closed) {
            closeTab();
        } else {
            markLeaving();
            window.location.href = href;
        }
    }

    /**
     * <a class="btn-close">閉じる</a> のクリックを検知し、
     * メニューから開かれたタブであれば window.close() する。
     * 画面固有の確認ダイアログ（同期的な window.confirm 等で
     * event.preventDefault() 済みの場合）はスキップする。
     *
     * 【登録タイミングについて】
     *   画面固有の確認ダイアログ（例: mcm0021u の window.confirm() による
     *   クリア確認）は同じ <a class="btn-close"> 要素に対する別の click リスナーとして
     *   実装されており、event.preventDefault() でキャンセル時のみ遷移を止める作りになっている。
     *   click イベントの複数リスナーは登録順に実行されるため、このリスナーは
     *   各画面が登録する個別リスナーより後に登録する必要がある。
     *   これにより、個別リスナーの window.confirm() 結果（同期的に defaultPrevented へ反映される）
     *   を確認した上で window.close() の実行可否を判断できる。
     *   具体的には DOMContentLoaded 直後の setTimeout(...,0) で登録する。
     *   各画面の個別リスナー登録は、インラインスクリプトの実行時または
     *   DOMContentLoaded リスナー内で行われ、いずれも DOMContentLoaded の
     *   ディスパッチ完了までに終わるため、setTimeout のタスクは必ずそれより後になる。
     *
     * 【不具合修正：タブが開ききる前の「閉じる」押下】
     *   従来は window の load イベント内で登録していたため、タブの読み込みが
     *   完了する前（load 発火前）に「閉じる」を押すとリスナーが未登録で、
     *   タブが閉じられずにそのタブ内で /menu へ遷移し、メニュー画面が
     *   表示されてしまっていた。
     *   そこで、通常リスナーの登録前に発生したクリックは capture フェーズの
     *   早期リスナーで同じ処理を行うようにする。この時点では各画面の個別リスナーは
     *   まだ登録されていない（＝確認ダイアログの対象となる入力操作も発生し得ない）ため、
     *   早期リスナーが先に処理しても個別の確認ダイアログを奪うことはない。
     *
     * 【対象外となる画面について】
     *   「閉じる」ボタン押下時に非同期の独自モーダルダイアログ（showConfirm等）で
     *   確認を取り、OK時に window.location.href で遷移する画面は、クリックした
     *   瞬間に必ず event.preventDefault() されるため、このクリックリスナーでは
     *   意図通りに動作しない（キャンセルされたのか、確認待ち中なのか区別できないため）。
     *   このような画面は、遷移箇所（window.location.href = href）を
     *   McmCommon.closeOrNavigate(href) の呼び出しに置き換えて対応すること。
     */
    function handleCloseClick(e) {
        var link = e.target.closest ? e.target.closest("a.btn-close") : null;
        if (!link) return;
        if (link.textContent.trim() !== "閉じる") return; // 「戻る」等は対象外
        if (!link.getAttribute("href")) return;
        if (/\/menu\/logout(?:[?#]|$)/.test(link.getAttribute("href"))) return; // ログアウトは対象外

        // メニューから target="_blank" で開かれたタブでなければ、通常のページ遷移に任せる
        if (!window.opener || window.opener.closed) return;

        // 画面固有の確認ダイアログ等で既にキャンセルされている場合は何もしない
        if (e.defaultPrevented) return;

        e.preventDefault();

        var closeUrl = new URL(link.href, window.location.href);

        // 編集画面と1003U/3004Uは、状態を破棄してからタブを閉じる。
        if (closeUrl.origin === window.location.origin &&
            /\/mcm(?:200[36]|100[346]|3004)u\/close\/?$/.test(closeUrl.pathname)) {
            var resetUrl = new URL(closeUrl.href);
            // 検索画面の後始末は204応答で完結。fetchのリダイレクトで/menuを実行しない。
            if (/\/mcm(?:1003|3004)u\/close\/?$/.test(resetUrl.pathname)) {
                resetUrl.pathname = resetUrl.pathname.replace(/\/close\/?$/, "/reset");
            }
            fetch(resetUrl.href, { credentials: "same-origin", cache: "no-store" })
                .then(function (response) {
                    if (!response.ok) {
                        throw new Error("MCM2003Uの閉じる処理に失敗しました");
                    }
                    // 閉じる／遷移のいずれの場合も意図した離脱であるため抑止する。
                    markLeaving();
                    window.close();
                    if (!window.closed) {
                        window.location.assign(closeUrl.href);
                    }
                })
                .catch(function () {
                    // 通信に失敗した場合も、タブだけを閉じずに閉じる処理へ移動する
                    markLeaving();
                    window.location.assign(closeUrl.href);
                });
            return;
        }

        // 他画面は従来どおり
        closeTab();
    }

    /** 通常リスナー（各画面の個別リスナーより後に実行される）を登録済みか */
    var closeClickListenerRegistered = false;

    /** 通常リスナーを登録する（多重登録防止） */
    function registerCloseClickListener() {
        if (closeClickListenerRegistered) return;
        closeClickListenerRegistered = true;
        document.addEventListener("click", handleCloseClick, false);
    }

    /**
     * 早期リスナー（capture フェーズ）。
     * 通常リスナーが登録される前（＝ページの読み込み途中）に押された「閉じる」だけを
     * 処理する。登録後は通常リスナーに委ねるため何もしない。
     */
    document.addEventListener("click", function (e) {
        if (closeClickListenerRegistered) return;
        handleCloseClick(e);
    }, true);

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", function () {
            setTimeout(registerCloseClickListener, 0);
        });
    } else {
        // common.js が DOMContentLoaded 後に読み込まれた場合（遅延読み込み等）
        setTimeout(registerCloseClickListener, 0);
    }

    window.McmCommon = {
        showProgress: showProgress,
        hideProgress: hideProgress,
        withProgress: withProgress,
        closeOrNavigate: closeOrNavigate,
        markDirty: markDirty,
        markClean: markClean,
        markLeaving: markLeaving,
        isDirty: function () { return isDirty; }
    };
})(window, document);
