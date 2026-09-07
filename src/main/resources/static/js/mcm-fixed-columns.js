/** #155: 固定列の位置を実際の列幅に追従させる。 */
(function () {
    "use strict";

    function initialize() {
        document.querySelectorAll('table[data-mcm-fixed-columns]').forEach(function (table) {
            var count = Number(table.dataset.mcmFixedColumns);
            var header = table.tHead && table.tHead.rows[0];
            if (!header || !Number.isInteger(count) || count < 1 || count > header.cells.length) return;
            var cells = Array.prototype.slice.call(header.cells, 0, count);

            function updateOffsets() {
                // 非表示の子グリッドは幅0になる。再表示時のResizeObserverで更新する。
                if (table.getBoundingClientRect().width === 0) return;
                var widths = cells.map(function (cell) { return cell.getBoundingClientRect().width; });
                var left = 0;
                widths.forEach(function (width, index) {
                    var name = '--mcm-fixed-left-' + (index + 1);
                    var value = left + 'px';
                    if (table.style.getPropertyValue(name) !== value) {
                        table.style.setProperty(name, value);
                    }
                    left += width;
                });
            }

            updateOffsets();
            // 列幅ドラッグ・初期幅調整・親子グリッドの表示切替を検知。
            // leftの更新は幅を変えないため、監視による再帰的なリサイズは発生しない。
            if (typeof ResizeObserver !== 'undefined') {
                var observer = new ResizeObserver(updateOffsets);
                observer.observe(table);
                cells.forEach(function (cell) { observer.observe(cell); });
            } else {
                // ResizeObserver非対応環境では既存の列幅変更処理を監視する。
                var colgroup = table.querySelector('colgroup');
                if (colgroup) {
                    new MutationObserver(updateOffsets).observe(colgroup, {
                        attributes: true, attributeFilter: ['style'], subtree: true
                    });
                }
            }
            window.addEventListener('resize', updateOffsets);
        });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initialize, { once: true });
    } else {
        initialize();
    }
})();
