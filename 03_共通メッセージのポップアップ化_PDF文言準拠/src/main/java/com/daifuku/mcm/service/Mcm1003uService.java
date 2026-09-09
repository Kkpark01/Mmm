package com.daifuku.mcm.service;

import com.daifuku.mcm.constants.Mcm1003uConstants;
import com.daifuku.mcm.dto.Mcm1002pDataViewDto;
import com.daifuku.mcm.dto.Mcm1003uRowDto;
import com.daifuku.mcm.form.Mcm1003uForm;
import com.daifuku.mcm.repository.Mcm1003uJdbcRepository;
import com.daifuku.mcm.repository.Mcm1003uViewRepository;
import com.daifuku.mcm.repository.McmTkKeiyakuRepository;
import com.daifuku.mcm.repository.McmTmKeiyakujikanRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.List;

/**
 * 【変換元】Mcm1003uScreen.vb - MCM1003U 取引先見積・契約一覧 サービスクラス
 * 【変換元行数】約31,967行
 * 【作成者】T.Kajimura (2008/02/14)
 * 【更新者】T.Iwasawa (2009/08/12), M.Ohsuka (2009/12/02, 2010/01/15)
 *
 * 元コード: WinForms の画面ロジック（イベントハンドラ）から
 *           ビジネスロジック部分を抽出して Service に分離。
 *
 * 主要メソッドの変換元:
 *   search()             ← SearchButton_Click()
 *   discardMitsumori()   ← UVADataGridView_CellContentClick() - MITSUMORI_HAKI case
 *   discardKeiyaku()     ← UVADataGridView_CellContentClick() - KEIYAKU_HAKI case
 *   findPlantIds()       ← UVADataGridView_CellContentClick() - MITSUMORI_COPY case
 *   countUmTanka()       ← RowDeleteButton_Click()
 *   buildExcelDataList() ← IraiIchiranSyutsuryokuButton_Click()
 */
@Service
public class Mcm1003uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm1003uService.class);

    private final Mcm1003uViewRepository viewRepository;
    private final Mcm1003uJdbcRepository jdbcRepository;
    private final McmTkKeiyakuRepository tkKeiyakuRepository;
    private final McmTmKeiyakujikanRepository tmKeiyakujikanRepository;

    /**
     * コンストラクタインジェクション。
     */
    public Mcm1003uService(Mcm1003uViewRepository viewRepository,
                           Mcm1003uJdbcRepository jdbcRepository,
                           McmTkKeiyakuRepository tkKeiyakuRepository,
                           McmTmKeiyakujikanRepository tmKeiyakujikanRepository) {
        this.viewRepository = viewRepository;
        this.jdbcRepository = jdbcRepository;
        this.tkKeiyakuRepository = tkKeiyakuRepository;
        this.tmKeiyakujikanRepository = tmKeiyakujikanRepository;
    }

    /**
     * 依頼担当者リストを取得する。
     * @return 依頼担当者リスト
     */
    public List<String> getIraitantosyaList() {
        return jdbcRepository.findDistinctIraitantosya();
    }

    // ===================================================================
    // 検索
    // ===================================================================

    /**
     * 検索条件に基づき一覧を取得する。
     *
     * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
     *   元コード:
     *     '' 検索条件用項目が全て空白の場合
     *     If IsNull(MitsumoriIraiNoTextBox.Text) And IsNull(...) Then
     *       DisplayMessage(CPMessageConstant.MSG_0001)
     *       Return
     *     End If
     *     '' 依頼日が日付型でなければエラー
     *     If IsNotNull(Iraibi1TextBox.Text) Then
     *       If Not IsDate(Iraibi1TextBox.Text) Then
     *         DisplayMessage(CPMessageConstant.FWM_0006, ...)
     *       End If
     *     End If
     *     '' 状態チェックボックスの値を結合して検索パラメータ化
     *     If Me.SagyochuCheckBox.Checked Then strJotai &= "作成中,"
     *     ...
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "UVA", ...)
     *
     * @param form 検索フォーム
     * @return 検索結果と検証エラーのリスト
     */
    public SearchResult search(Mcm1003uForm form) {
        List<String> errors = new ArrayList<>();

        // --- バリデーション: 全条件が空の場合 ---
        if (isAllEmpty(form)) {
            errors.add("検索条件は、1項目以上選択して下さい。");
            return new SearchResult(List.of(), errors);
        }

        // --- バリデーション: 日付形式チェック ---
        /*
         * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
         *   元コード:
         *     If Not IsDate(Iraibi1TextBox.Text) Then
         *       DisplayMessage(CPMessageConstant.FWM_0006, New String() {"依頼日", "yyyy/MM/dd"})
         *     End If
         */
        if (isNotEmpty(form.getIraibi1()) && !isValidDate(form.getIraibi1())) {
            errors.add("開始日付の書式を指定して下さい。(YYYY/MM/DD)");
        }
        if (isNotEmpty(form.getIraibi2()) && !isValidDate(form.getIraibi2())) {
            errors.add("終了日付の書式を指定して下さい。(YYYY/MM/DD)");
        }

        if (!errors.isEmpty()) {
            return new SearchResult(List.of(), errors);
        }

        // --- 検索実行 ---
        List<Mcm1003uRowDto> results = viewRepository.search(form);

        /*
         * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
         *   元コード:
         *     '' 検索結果が0件の場合
         *     If Me.Mcm1003uDataSet.UVA.Count = 0 Then
         *       DisplayMessage(CPMessageConstant.MSG_0002)
         *     End If
         */
        if (results.isEmpty()) {
            errors.add("検索結果が1件も存在しません。");
        }

        log.info("MCM1003U 検索完了: {}件", results.size());
        return new SearchResult(results, errors);
    }

    // ===================================================================
    // 見積破棄
    // ===================================================================

    /**
     * 見積を破棄する。
     *
     * 【変換元】Mcm1003uScreen.vb - UVADataGridView_CellContentClick()
     *   Mcm1003uConstant.MITSUMORI_HAKI case 内のロジック:
     *
     *   元コード:
     *     '' 契約が存在する場合はエラー
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "TK_KEIYAKU", tkKeiyakuId)
     *     If Me.Mcm1003uDataSet.TK_KEIYAKU.Count > 0 Then
     *       DisplayMessage(CPMessageConstant.MSG_0025)
     *       Return
     *     End If
     *
     *     '' 見積契約時間の状態を「破棄」に更新
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "TM_KEIYAKUJIKAN", tmKeiyakujikanId)
     *     dvTM_KEIYAKUJIKAN = New DataView(Mcm1003uDataSet.TM_KEIYAKUJIKAN)
     *     dvTM_KEIYAKUJIKAN.RowFilter = "TM_KEIYAKUJIKAN_ID = " & tmKeiyakujikanId
     *     dvTM_KEIYAKUJIKAN.Item(0).Row("JOTAI") = "破棄"
     *     dvTM_KEIYAKUJIKAN.Item(0).Row("LASTUPDATE_BY") = loginUserId
     *     dvTM_KEIYAKUJIKAN.Item(0).Row("LASTUPDATE_DT") = Now
     *     Me.DaoContainer.Update(Mcm1003uDataSet, "TM_KEIYAKUJIKAN")
     *
     * @param tmKeiyakujikanId 見積契約時間ID
     * @param tkKeiyakuId      契約ID（存在チェック用）
     * @param loginUser        ログインユーザーID
     * @throws IllegalStateException 契約が既に存在する場合
     */
    @Transactional
    public void discardMitsumori(BigDecimal tmKeiyakujikanId,
                                 BigDecimal tkKeiyakuId,
                                 String loginUser) {

        log.info("MCM1003U 見積破棄開始: tmKeiyakujikanId={}", tmKeiyakujikanId);

        // --- 契約存在チェック ---
        if (tkKeiyakuId != null && tkKeiyakuId.compareTo(BigDecimal.ZERO) > 0) {
            int count = jdbcRepository.countTkKeiyakuById(tkKeiyakuId);
            if (count > 0) {
                /*
                 * 【変換元】MSG_0025
                 *   元メッセージ: 契約が存在するため、見積の破棄はできません。
                 */
                throw new IllegalStateException("店舗への見積依頼 又は 取引先との契約として使用されている為、破棄する事が出来ません。");
            }
        }

        // --- 見積契約時間の状態を「破棄」に更新 ---
        LocalDateTime now = LocalDateTime.now();
        tmKeiyakujikanRepository.updateJotai(
                Mcm1003uConstants.HAKI,
                loginUser,
                now,
                tmKeiyakujikanId);

        log.info("MCM1003U 見積破棄完了: tmKeiyakujikanId={}", tmKeiyakujikanId);
    }

    // ===================================================================
    // 契約破棄
    // ===================================================================

    /**
     * 契約を破棄する。
     *
     * 【変換元】Mcm1003uScreen.vb - UVADataGridView_CellContentClick()
     *   Mcm1003uConstant.KEIYAKU_HAKI case 内のロジック:
     *
     *   元コード:
     *     '' 支払月チェック
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "TK_SIHARAIMEISAI", tkKeiyakuId)
     *     If Me.Mcm1003uDataSet.TK_SIHARAIMEISAI.Count > 0 Then
     *       For i = 0 To .TK_SIHARAIMEISAI.Count - 1
     *         minTsuki = .TK_SIHARAIMEISAI(i).TSUKI
     *         '' 支払月と現在年月を比較
     *         If CInt(minTsuki) <= CInt(Format(Now, "yyyyMM")) Then
     *           DisplayMessage(CPMessageConstant.MSG_0026)
     *           Return
     *         End If
     *       Next
     *     End If
     *
     *     '' 契約の状態を「破棄」に更新
     *     dvTK_KEIYAKU.Item(0).Row("JOTAI") = "破棄"
     *     Me.DaoContainer.Update(Mcm1003uDataSet, "TK_KEIYAKU")
     *
     *     '' 見積契約時間の状態を「見積」に戻す
     *     dvTM_KEIYAKUJIKAN.Item(0).Row("JOTAI") = "見積"
     *     Me.DaoContainer.Update(Mcm1003uDataSet, "TM_KEIYAKUJIKAN")
     *
     * @param tkKeiyakuId      契約ID
     * @param tmKeiyakujikanId 見積契約時間ID
     * @param loginUser        ログインユーザーID
     * @throws IllegalStateException 支払が既に処理済みの場合
     */
    @Transactional
    public void discardKeiyaku(BigDecimal tkKeiyakuId,
                               BigDecimal tmKeiyakujikanId,
                               String loginUser) {

        log.info("MCM1003U 契約破棄開始: tkKeiyakuId={}, tmKeiyakujikanId={}",
                tkKeiyakuId, tmKeiyakujikanId);

        // --- 支払月チェック ---
        /*
         * 【変換元】Mcm1003uScreen.vb
         *   元コード: minTsuki に対して CInt(minTsuki) <= CInt(Format(Now, "yyyyMM"))
         */
        List<String> tsukiList = jdbcRepository.findTsukiByKeiyakuId(tkKeiyakuId);
        if (tsukiList != null && !tsukiList.isEmpty()) {
            String currentYearMonth = YearMonth.now().format(DateTimeFormatter.ofPattern("yyyyMM"));
            int currentYm = Integer.parseInt(currentYearMonth);

            for (String tsuki : tsukiList) {
                if (tsuki != null && !tsuki.trim().isEmpty()) {
                    try {
                        int tsukiInt = Integer.parseInt(tsuki.trim());
                        if (tsukiInt <= currentYm) {
                            /*
                             * 【変換元】MSG_0026
                             *   元メッセージ: 支払処理済みの期間があるため、契約の破棄はできません。
                             */
                            throw new IllegalStateException(
                                    "取引先への支払が発生している為、破棄する事が出来ません。");
                        }
                    } catch (NumberFormatException e) {
                        log.warn("MCM1003U 支払月の数値変換エラー: tsuki={}", tsuki);
                    }
                }
            }
        }

        LocalDateTime now = LocalDateTime.now();

        // --- 契約の状態を「破棄」に更新 ---
        tkKeiyakuRepository.updateJotai(
                Mcm1003uConstants.HAKI,
                loginUser,
                now,
                tkKeiyakuId);

        // --- 見積契約時間の状態を「見積」に戻す ---
        /*
         * 【変換元】Mcm1003uScreen.vb
         *   元コード: dvTM_KEIYAKUJIKAN.Item(0).Row("JOTAI") = "見積"
         *   → 契約破棄時は見積の状態に戻す
         */
        tmKeiyakujikanRepository.updateJotai(
                Mcm1003uConstants.MITSUMORI,
                loginUser,
                now,
                tmKeiyakujikanId);

        log.info("MCM1003U 契約破棄完了: tkKeiyakuId={}", tkKeiyakuId);
    }

    // ===================================================================
    // プラントID検索（見積複製用）
    // ===================================================================

    /**
     * 契約時間IDに紐づくプラントIDリストを取得する。
     *
     * 【変換元】Mcm1003uScreen.vb - UVADataGridView_CellContentClick()
     *   Mcm1003uConstant.MITSUMORI_COPY case 内のロジック:
     *
     *   元コード:
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "PLANT_ID",
     *       currentRow.Cells(Mcm1003uConstant.TM_KEIYAKUJIKAN_ID).Value)
     *     Dim plantID() As Decimal = ...
     *     For i = 0 To .PLANT_ID.Count - 1
     *       plantID(i) = .PLANT_ID(i).PLANT_ID
     *     Next
     *
     * @param tmKeiyakujikanId 見積契約時間ID
     * @return プラントIDリスト
     */
    public List<BigDecimal> findPlantIds(BigDecimal tmKeiyakujikanId) {
        return jdbcRepository.findPlantIdsByKeiyakujikanId(tmKeiyakujikanId);
    }

    // ===================================================================
    // ユーザ見積単価件数（行削除チェック用）
    // ===================================================================

    /**
     * 依頼IDに紐づくユーザ見積単価の件数を取得する。
     *
     * 【変換元】Mcm1003uScreen.vb - RowDeleteButton_Click()
     *   元コード:
     *     Me.DaoContainer.Fill(Mcm1003uDataSet, "UM_TANKA",
     *       currentRow.Cells(Mcm1003uConstant.TM_IRAI_ID).Value)
     *     If Me.Mcm1003uDataSet.UM_TANKA.Count > 0 Then
     *       '' ユーザ見積単価が存在する → 削除不可
     *     End If
     *
     * @param tmIraiId 依頼ID
     * @return ユーザ見積単価件数
     */
    public int countUmTanka(BigDecimal tmIraiId) {
        return jdbcRepository.countUmTankaByIraiId(tmIraiId);
    }

    // ===================================================================
    // Excel出力用データ変換
    // ===================================================================

    /**
     * 検索結果を依頼一覧Excel出力用のDTOリストに変換する。
     *
     * 【変換元】Mcm1003uScreen.vb - IraiIchiranSyutsuryokuButton_Click()
     *   元コード:
     *     mitsumoriList = New List(Of Mcm1002pData_View)
     *     uvaTable = Me.Mcm1003uDataSet.UVA.Select(Nothing, "TM_IRAI_NO DESC")
     *     For Each row As DataRow In uvaTable
     *       Dim data As Mcm1002pData_View = New Mcm1002pData_View
     *       data.IRAI_NO    = row("TM_IRAI_NO")
     *       data.JIKANTAI   = row("KEIYAKUJIKANTAI")
     *       data.JOTAI      = row("JOTAI")
     *       data.KEIYAKU_NO = row("KEIYAKU_NO")
     *       data.TORIHIKISAKI_NK = row("TORIHIKISAKI_NK")
     *       data.NONYUSAKI_NK    = row("NONYUSAKI_NK")
     *       data.SUPPORT_ID      = row("SUPPORT_ID")
     *       data.PLANT_NK        = row("PLANT_NK")
     *       data.KAITO_DT        = Format(row("KAITOKIZITSU_DT"), "yyyy/MM/dd")
     *       data.IRAI_DT         = Format(row("MITSUMORI_DT"), "yyyy/MM/dd")
     *       data.IRAISYA         = row("IRAITANTOSYA")
     *       mitsumoriList.Add(data)
     *     Next
     *
     *     '' Deliveryに格納してMCM1002P帳票画面へ遷移
     *     delivery.DgvArray = mitsumoriList
     *     ForwardScreen(McmScreenIdConstant.MCM1002P, delivery)
     *
     * @param rows 検索結果リスト
     * @return MCM1002P用のDTOリスト
     */
    public List<Mcm1002pDataViewDto> buildExcelDataList(List<Mcm1003uRowDto> rows) {
        List<Mcm1002pDataViewDto> excelList = new ArrayList<>();
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");

        for (Mcm1003uRowDto row : rows) {
            Mcm1002pDataViewDto data = new Mcm1002pDataViewDto();

            data.setIraiNo(row.getTmIraiNo());

            // 契約時間帯（String → BigDecimal）
            if (row.getKeiyakujikantai() != null && !row.getKeiyakujikantai().isEmpty()) {
                try {
                    data.setJikantai(new BigDecimal(row.getKeiyakujikantai()));
                } catch (NumberFormatException e) {
                    log.warn("MCM1003U 契約時間帯の数値変換エラー: {}", row.getKeiyakujikantai());
                }
            }

            data.setJotai(row.getJotai());
            data.setKeiyakuNo(row.getKeiyakuNo());
            data.setTorihikisakiNk(row.getTorihikisakiNk());
            data.setNonyusakiNk(row.getNonyusakiNk());
            data.setSupportId(row.getSupportId());
            data.setPlantNk(row.getPlantNk());

            /*
             * 元コード: Format(row("KAITOKIZITSU_DT"), "yyyy/MM/dd")
             */
            data.setKaitoDt(row.getKaitokizitsuDt());

            /*
             * 元コード: Format(row("MITSUMORI_DT"), "yyyy/MM/dd")
             */
            if (row.getMitsumoriDt() != null) {
                data.setIraiDt(row.getMitsumoriDt().format(dtf));
            }

            data.setIraisya(row.getIraitantosya());

            excelList.add(data);
        }

        log.info("MCM1003U Excel出力データ変換完了: {}件", excelList.size());
        return excelList;
    }

    // ===================================================================
    // バリデーション: 全条件空チェック
    // ===================================================================

    /**
     * 検索フォームの全フィールドが空かどうかを判定する。
     *
     * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
     *   元コード:
     *     If IsNull(MitsumoriIraiNoTextBox.Text) And IsNull(KeiyakuNoTextBox.Text) And ...
     *       And SagyochuCheckBox.Checked = False And ...
     *       DisplayMessage(CPMessageConstant.MSG_0001)
     *       Return
     *     End If
     *
     * @param form 検索フォーム
     * @return 全フィールドが空の場合 true
     */
    private boolean isAllEmpty(Mcm1003uForm form) {
        return isEmpty(form.getMitsumoriIraiNo())
                && isEmpty(form.getKeiyakuNo())
                && isEmpty(form.getIraibi1())
                && isEmpty(form.getIraibi2())
                && isEmpty(form.getIraisya())
                && isEmpty(form.getTorihikisakiCd())
                && isEmpty(form.getTorihikisakimei())
                && isEmpty(form.getNonyusakiCd())
                && isEmpty(form.getNonyusakimei())
                && isEmpty(form.getSupportId())
                && isEmpty(form.getPlantmei())
                && !form.isSagyochu()
                && !form.isIrai()
                && !form.isMitsumori()
                && !form.isKeiyaku()
                && !form.isHaki()
                && !form.isKaiyaku()
                && !form.isShoninSakuseichu()
                && !form.isShoninSinsachu()
                && !form.isShoninShoninchu()
                && !form.isShoninShoninzumi()
                && !form.isShoninSashimodoshi();
    }

    // ===================================================================
    // バリデーション: 日付形式チェック
    // ===================================================================

    /**
     * 日付文字列が yyyy/MM/dd 形式かどうかを判定する。
     *
     * 【変換元】Mcm1003uScreen.vb - SearchButton_Click()
     *   元コード: If Not IsDate(Iraibi1TextBox.Text) Then ...
     *
     * @param dateStr 日付文字列
     * @return 有効な日付形式の場合 true
     */
    private boolean isValidDate(String dateStr) {
        try {
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy/MM/dd");
            java.time.LocalDate.parse(dateStr, formatter);
            return true;
        } catch (DateTimeParseException e) {
            return false;
        }
    }

    // ===================================================================
    // ヘルパーメソッド
    // ===================================================================

    private boolean isEmpty(String value) {
        return value == null || value.trim().isEmpty();
    }

    private boolean isNotEmpty(String value) {
        return !isEmpty(value);
    }

    // ===================================================================
    // 検索結果ラッパー
    // ===================================================================

    /**
     * 検索結果とエラーメッセージを格納する内部クラス。
     *
     * 【変換元】
     *   元コード: VBではDisplayMessage()でメッセージ表示後にReturnしていた。
     *   変換後:   エラーメッセージをリストで返し、Controller側でFlashAttributeに設定。
     */
    public static class SearchResult {
        private final List<Mcm1003uRowDto> rows;
        private final List<String> errors;

        public SearchResult(List<Mcm1003uRowDto> rows, List<String> errors) {
            this.rows = rows;
            this.errors = errors;
        }

        public List<Mcm1003uRowDto> getRows() {
            return rows;
        }

        public List<String> getErrors() {
            return errors;
        }

        public boolean hasErrors() {
            return errors != null && !errors.isEmpty();
        }
    }
}
