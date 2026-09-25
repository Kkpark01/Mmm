/**
 * 【変換元】Mcm1005uScreen.vb / Mcm1005uTabControl.vb
 * 【説明】取引先契約内容（MCM1005U）サービスクラス
 *
 * 元イベント対応:
 *   Mcm1005uScreen_Load / SearchUpdate / SearchInsert → load()
 *   UpdateButton_Click / UpdButton()                  → save()
 *   SinseiButton_Click                                → sinsei()
 */
package com.daifuku.mcm.service;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.*;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.exception.ExclusiveControlException;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.repository.Mcm1005uRepository;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

@Service
public class Mcm1005uService {
    @Autowired private Mcm1003pExcelService reportExcel;

    /** VB KeiyakuHakoButton_Click: enforce period ownership and approved attachment restriction. */
    @Transactional(readOnly = true)
    public void checkReport(BigDecimal contractId, BigDecimal periodId, boolean cancellation, String user) {
        if (!canUpdate(user)) throw new McmBusinessException("権限がないため実行できません。");
        if (repository.findKeiyakuById(contractId) == null || repository.findKikanByKeiyakuId(contractId).stream().noneMatch(p -> sameId(p.getTkKikanId(), periodId))) throw conflict();
        if (!cancellation && repository.findTenpuByKikanId(periodId).stream().anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("承認済みな契約手続依頼書が存在する為、再発行できません。期間タブを作成し直して下さい。");
    }

    @Transactional(readOnly = true)
    public byte[] contractReport(BigDecimal contractId, BigDecimal periodId, String user, String lastName) throws java.io.IOException {
        var contract = repository.findKeiyakuById(contractId);
        checkReport(contractId, periodId, contract != null && contract.getKaiyakuDt() != null, user);
        var periods = repository.findKikanByKeiyakuId(contractId);
        int index = 0;
        while (!sameId(periods.get(index).getTkKikanId(), periodId)) index++;
        var selected = periods.get(index);
        McmTkKikanEntity previous = null;
        boolean cancellation = contract.getKaiyakuDt() != null;
        if (cancellation) previous = selected;
        else for (int i = index - 1; i >= 0; i--) {
            if (periods.get(i).getYukoFlg() != null && periods.get(i).getYukoFlg().signum() == 0) { previous = periods.get(i); break; }
        }
        var data = new com.daifuku.mcm.dto.Mcm1003pKihonDto();
        data.setKubun(cancellation ? com.daifuku.mcm.constants.Mcm1003pConstants.KUBUN_DELETE : previous == null ? com.daifuku.mcm.constants.Mcm1003pConstants.KUBUN_INSERT : com.daifuku.mcm.constants.Mcm1003pConstants.KUBUN_UPDATE);
        data.setCreatedDt(reportDate(LocalDateTime.now()));
        var contact = repository.findReportContact(selected.getTorihikisakiId());
        data.setTorihikisakiNk(reportText(contact.get("TORIHIKISAKI_NK")));
        data.setTorisyutantosyaNk(reportText(contact.get("TORISYUTANTOSYA_NK")));
        data.setNonyusakiNk(selected.getNonyusakiNk());
        data.setYubinNo(repository.findReportPostalCode(periodId));
        data.setNonyusakijusyo1Nk(reportText(selected.getNonyusakijusyo1Nk()) + reportText(selected.getNonyusakijusyo2Nk()));
        data.setNonyutelNo(selected.getNonyutelNo()); data.setNonyubusyoNk(selected.getNonyubusyoNk());
        data.setNonyutantosyaNk(selected.getNonyutantosyaNk()); data.setSupportId(selected.getSupportId());
        if (!cancellation) {
            data.setSikiriKeiyaku(reportAmount(selected.getSikirigokeiKin()));
            data.setSiharaiKeiyaku(reportPayment(periodId)); data.setKeiyakukaisiDt(reportDate(contract.getKeiyakuDt()));
            String hours = reportText(selected.getKeiyakujikantai()); data.setKeiyakujikantai(hours.isBlank() ? "" : hours + "H");
            data.setTmMitsumoriNo(reportJoin(repository.findReportQuotationNumbers(periodId)));
            data.setTehaiseiban(reportJoin(repository.findReportOrderNumbers(periodId)));
        }
        if (previous != null) {
            data.setSikiriKaiyaku(reportAmount(previous.getSikirigokeiKin())); data.setSiharaiKaiyaku(reportPayment(previous.getTkKikanId()));
            data.setKaiyakuDt(reportDate(contract.getKaiyakuDt())); data.setKeiyakuNo(contract.getKeiyakuNo());
        }
        var nowRows = repository.findReportDetails(periodId);
        var lines = new ArrayList<com.daifuku.mcm.dto.Mcm1003pMeisaiDto>();
        var currentInspection = reportInspections(periodId);
        if (cancellation || previous == null) {
            for (var row : nowRows) addReportLine(lines, row, reportNumber(row.get("SURYO_NM")), "", currentInspection);
        } else {
            var oldRows = repository.findReportDetails(previous.getTkKikanId());
            var oldInspection = reportInspections(previous.getTkKikanId());
            var oldByKey = new java.util.LinkedHashMap<String,Map<String,Object>>();
            var newByKey = new java.util.LinkedHashMap<String,Map<String,Object>>();
            for (var row : oldRows) oldByKey.putIfAbsent(reportKey(row), row);
            for (var row : nowRows) newByKey.putIfAbsent(reportKey(row), row);
            var ordered = new java.util.LinkedHashMap<String,Map<String,Object>>(oldByKey); ordered.putAll(newByKey);
            var keys = new ArrayList<>(ordered.keySet());
            keys.sort(java.util.Comparator.comparing((String k) -> reportNumber(ordered.get(k).get("MAE_HYOJIJUN"))).thenComparing(k -> reportNumber(ordered.get(k).get("HYOJIJUN"))));
            for (String key : keys) {
                var n = newByKey.get(key); var o = oldByKey.get(key);
                if (o == null) addReportLine(lines, n, reportNumber(n.get("SURYO_NM")), "契約", currentInspection);
                else if (n == null) addReportLine(lines, o, reportNumber(o.get("SURYO_NM")), "解約", oldInspection);
                else {
                    BigDecimal nq = reportNumber(n.get("SURYO_NM")), oq = reportNumber(o.get("SURYO_NM"));
                    addReportLine(lines, n, nq.min(oq), "契約(継続）", currentInspection);
                    if (nq.compareTo(oq) > 0) addReportLine(lines, n, nq.subtract(oq), "契約", currentInspection);
                    if (nq.compareTo(oq) < 0) addReportLine(lines, o, oq.subtract(nq), "解約", oldInspection);
                }
            }
        }
        return reportExcel.generate(data, lines, lastName);
    }

    private String reportPayment(BigDecimal period) {
        var row = repository.findSiharaiByKikanId(period);
        return row == null || row.getKaisu() == null ? "" : "年" + row.getKaisu().stripTrailingZeros().toPlainString() + "回";
    }
    private Map<String,String[]> reportInspections(BigDecimal period) {
        var map = new HashMap<String,String[]>();
        for (var row : repository.findTenkenByKikanId(period)) {
            String months = repository.findTenkenMeisaiTsukiMonths(row.getTkTenkenId()).stream().distinct().sorted().map(m -> m + "月").collect(Collectors.joining(","));
            map.put(reportText(row.getTkKikokoseiId()) + ":" + reportText(row.getOyakikibunruiCd()), new String[]{reportNumber(row.getTenkenkaisu()).stripTrailingZeros().toPlainString() + "回/年", months});
        }
        return map;
    }
    private static void addReportLine(List<com.daifuku.mcm.dto.Mcm1003pMeisaiDto> list, Map<String,Object> row, BigDecimal quantity, String note, Map<String,String[]> inspections) {
        var line = new com.daifuku.mcm.dto.Mcm1003pMeisaiDto();
        line.setHyojijun(Integer.toString(list.size()+1)); line.setKikihinmeiNk(reportText(row.get("KIKIHINMEI_NK")));
        line.setKikikatashiki(reportText(row.get("KIKIKATASHIKI"))); line.setSuryoNm(quantity); line.setBiko(note);
        var inspection = inspections.get(reportText(row.get("TK_KIKIKOSEI_ID")) + ":" + reportText(row.get("OYAKIKIBUNRUI_CD")));
        if (inspection != null) { line.setTenkenkaisu(inspection[0]); line.setTenkenTsuki(inspection[1]); }
        list.add(line);
    }
    private static String reportKey(Map<String,Object> row) { return reportText(row.get("KIKIKOSEI_ID")) + ":" + reportText(row.get("KIKIMEISAI_ID")); }
    private static String reportText(Object value) { return value == null ? "" : value instanceof BigDecimal b ? b.stripTrailingZeros().toPlainString() : value.toString(); }
    private static BigDecimal reportNumber(Object value) { return value == null ? BigDecimal.ZERO : new BigDecimal(value.toString()); }
    private static String reportAmount(BigDecimal value) { return value == null ? "" : String.format(java.util.Locale.JAPAN, "%,.0f", value); }
    private static String reportDate(LocalDateTime value) { return value == null ? "" : value.format(java.time.format.DateTimeFormatter.ofPattern("yyyy/MM/dd")); }
    private static String reportJoin(List<String> values) { return values.stream().filter(java.util.Objects::nonNull).collect(Collectors.joining(",")); }

    @Autowired
    private Mcm1005uRepository repository;

    @Autowired
    private Mcm2004uService permissions;

    @Autowired
    private com.daifuku.mcm.common.ValidationService validation;

    /**
     * 【移植】元VB: Mcm1005uScreen.UpdateButton/SinseiButton/KeiyakuHakoButton.AuthorityIsThrough=False
     *   「取引先契約関連」（KENGENBUNRUI_ID=3、機能ID=MCM1005U）の実効権限が
     *   作成（"2"）の場合のみ登録・申請・破棄・機器選定起動を許可する。
     *   参照専用ユーザーは画面上のボタン非活性に加え、サーバー側でも必ず検証する。
     */
    public boolean canUpdate(String user) {
        return "2".equals(permissions.getAuthority(user, "MCM1005U"));
    }

    // ===================================================================
    // 初期表示
    // ===================================================================

    /** 契約ヘッダー取得 */
    @Transactional(readOnly = true)
    public McmTkKeiyakuEntity loadKeiyaku(BigDecimal tkKeiyakuId) {
        return repository.findKeiyakuById(tkKeiyakuId);
    }

    /** 期間一覧取得（開始日昇順） */
    @Transactional(readOnly = true)
    public List<McmTkKikanEntity> loadKikanList(BigDecimal tkKeiyakuId) {
        return repository.findKikanByKeiyakuId(tkKeiyakuId);
    }

    /** 期間に紐づく点検情報取得 */
    @Transactional(readOnly = true)
    public List<McmTkTenkenEntity> loadTenkenList(BigDecimal tkKikanId) {
        List<McmTkTenkenEntity> list = repository.findTenkenByKikanId(tkKikanId);
        // 月チェックを点検明細から再構築
        for (McmTkTenkenEntity tenken : list) {
            List<Integer> months = repository.findTenkenMeisaiTsukiMonths(tenken.getTkTenkenId());
            String[] tsuki = new String[12];
            for (int i = 0; i < 12; i++) tsuki[i] = "0";
            for (int m : months) {
                if (m >= 1 && m <= 12) tsuki[m - 1] = "1";
            }
            tenken.setTsuki(tsuki);
        }
        return list;
    }

    /** 期間に紐づく支払情報取得 */
    @Transactional(readOnly = true)
    public McmTkSiharaiEntity loadSiharai(BigDecimal tkKikanId) {
        McmTkSiharaiEntity siharai = repository.findSiharaiByKikanId(tkKikanId);
        if (siharai == null) return null;
        List<Integer> months = repository.findSiharaiMeisaiTsukiMonths(siharai.getTkSiharaiId());
        boolean[] tsuki = new boolean[12];
        for (int m : months) {
            if (m >= 1 && m <= 12) tsuki[m - 1] = true;
        }
        siharai.setShiharaiTsuki(tsuki);
        return siharai;
    }

    /** 期間に紐づく添付ファイル一覧取得 */
    @Transactional(readOnly = true)
    public List<McmTkTenpuEntity> loadTenpuList(BigDecimal tkKikanId) {
        return repository.findTenpuByKikanId(tkKikanId);
    }

    // ===================================================================
    // 新規契約ID採番
    // ===================================================================

    @Transactional(readOnly = true)
    public BigDecimal getNextKeiyakuId() {
        return repository.getMaxKeiyakuId().add(BigDecimal.ONE);
    }

    // ===================================================================
    // 登録処理
    // 元VB: UpdateButton_Click → UpdButton()
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void save(McmTkKeiyakuEntity keiyaku,
                     List<McmTkKikanEntity> kikanList,
                     List<McmTkTenkenEntity> tenkenList,
                     List<McmTkSiharaiEntity> siharaiList,
                     List<Boolean[]> siharaiTsukiList,
                     List<McmTkTenpuEntity> tenpuList,
                     String loginUser) {

        if (!canUpdate(loginUser)) throw new McmBusinessException("権限がないため実行できません。");
        // Serialize writes to this contract, then compare the version originally displayed.
        McmTkKeiyakuEntity current = repository.lockKeiyaku(keiyaku.getTkKeiyakuId());
        if (current == null) throw new InputCheckException("対象の契約が見つかりません。検索画面から開き直してください。");
        checkVersion(keiyaku.getLastupdateDt(), current.getLastupdateDt());
        if (!canEditDetails(current)) {
            current.setBiko(keiyaku.getBiko()); current.setKosinnaiyo(keiyaku.getKosinnaiyo());
            repository.updateKeiyaku(current, loginUser);
            return;
        }
        List<McmTkKikanEntity> periods = repository.lockKikanList(current.getTkKeiyakuId());
        if (kikanList.isEmpty() || periods.size() != kikanList.size()) throw conflict();
        for (McmTkKikanEntity input : kikanList) {
            McmTkKikanEntity period = periods.stream().filter(p -> sameId(p.getTkKikanId(), input.getTkKikanId()))
                    .findFirst().orElseThrow(Mcm1005uService::conflict);
            checkVersion(input.getLastupdateDt(), period.getLastupdateDt());
            if (input.getKaisiDt() == null || input.getSyuryoDt() == null)
                throw new InputCheckException("未入力の必須項目があります。（期間）");
            if (input.getSyuryoDt().isBefore(input.getKaisiDt()))
                throw new InputCheckException("終了日が開始日より前になっています。");
            period.setKaisiDt(input.getKaisiDt()); period.setSyuryoDt(input.getSyuryoDt());
            period.setBiko(input.getBiko()); period.setKeiyakujikantai(input.getKeiyakujikantai());
        }
        // Only fields editable in this screen may change. Preserve state, totals and parent IDs.
        current.setKeiyakuNo(keiyaku.getKeiyakuNo()); current.setShokaiKeiyakuDt(keiyaku.getShokaiKeiyakuDt());
        current.setKeiyakuDt(keiyaku.getKeiyakuDt()); current.setJidokosinFlg(keiyaku.getJidokosinFlg());
        current.setKaiyakuDt(keiyaku.getKaiyakuDt()); current.setKeiyakumanryoDt(keiyaku.getKeiyakumanryoDt());
        current.setEntyokeiyakumanryoDt(keiyaku.getEntyokeiyakumanryoDt()); current.setPackFlg(keiyaku.getPackFlg());
        current.setPackkeiyakunaiyo(keiyaku.getPackkeiyakunaiyo()); current.setBiko(keiyaku.getBiko());
        current.setKosinnaiyo(keiyaku.getKosinnaiyo());
        current.setJikaikosinDt(periods.get(periods.size() - 1).getSyuryoDt().plusDays(1));

        List<McmTkTenkenEntity> inspectionUpdates = new ArrayList<>();
        for (McmTkKikanEntity period : periods) {
            for (McmTkTenkenEntity stored : repository.lockTenkenList(period.getTkKikanId())) {
                McmTkTenkenEntity input = tenkenList.stream().filter(t -> sameId(t.getTkTenkenId(), stored.getTkTenkenId()))
                        .findFirst().orElseThrow(Mcm1005uService::conflict);
                checkVersion(input.getLastupdateDt(), stored.getLastupdateDt());
                validateCount(input.getTenkenkaisu(), java.util.Arrays.stream(input.getTsuki()).filter("1"::equals).count(), "点検");
                stored.setTenkenkaisu(input.getTenkenkaisu()); stored.setTenkenkanoyobi(input.getTenkenkanoyobi());
                stored.setYakantaioumu(input.getYakantaioumu()); stored.setBiko(input.getBiko()); stored.setTsuki(input.getTsuki());
                inspectionUpdates.add(stored);
            }
        }
        if (inspectionUpdates.size() != tenkenList.size()) throw conflict();
        for (int i = 0; i < siharaiList.size(); i++) {
            McmTkSiharaiEntity input = siharaiList.get(i);
            if (periods.stream().noneMatch(p -> sameId(p.getTkKikanId(), input.getTkKikanId()))) throw conflict();
            McmTkSiharaiEntity stored = repository.lockSiharai(input.getTkKikanId());
            if (stored == null || !sameId(stored.getTkSiharaiId(), input.getTkSiharaiId())) throw conflict();
            checkVersion(input.getLastupdateDt(), stored.getLastupdateDt());
            validateCount(input.getKaisu(), java.util.Arrays.stream(siharaiTsukiList.get(i)).filter(Boolean.TRUE::equals).count(), "支払");
        }
        // Validation of the entire form precedes writes. Any later failure rolls back all updates.
        repository.updateKeiyaku(current, loginUser);
        for (McmTkKikanEntity period : periods) repository.updateKikan(period, loginUser);
        for (McmTkTenkenEntity inspection : inspectionUpdates) {
            repository.updateTenken(inspection, loginUser);
            McmTkKikanEntity period = periods.stream().filter(p -> sameId(p.getTkKikanId(), inspection.getTkKikanId())).findFirst().orElseThrow();
            boolean[] months = new boolean[12];
            for (int i = 0; i < 12; i++) months[i] = "1".equals(inspection.getTsukiAt(i));
            repository.saveMonthChecks(true, inspection.getTkTenkenId(), period.getKaisiDt(), period.getSyuryoDt(), months, loginUser);
        }
        for (int i = 0; i < siharaiList.size(); i++) {
            McmTkSiharaiEntity payment = siharaiList.get(i);
            repository.updateSiharai(payment, loginUser);
            McmTkKikanEntity period = periods.stream().filter(p -> sameId(p.getTkKikanId(), payment.getTkKikanId())).findFirst().orElseThrow();
            boolean[] months = new boolean[12];
            for (int m = 0; m < 12; m++) months[m] = Boolean.TRUE.equals(siharaiTsukiList.get(i)[m]);
            repository.saveMonthChecks(false, payment.getTkSiharaiId(), period.getKaisiDt(), period.getSyuryoDt(), months, loginUser);
        }
        // Attachment names/paths/status are read-only; never trust hidden values as updates.
    }

    /** Both actions share one transaction: failed application cannot leave a partial save. */
    @Transactional(rollbackFor = Exception.class)
    public void saveAndSinsei(McmTkKeiyakuEntity contract, List<McmTkKikanEntity> periods,
            List<McmTkTenkenEntity> inspections, List<McmTkSiharaiEntity> payments,
            List<Boolean[]> months, List<McmTkTenpuEntity> attachments, BigDecimal selectedPeriod, String user) {
        // Check approval preconditions before the first write as well as inside sinsei.
        var current = repository.lockKeiyaku(contract.getTkKeiyakuId());
        if (!canUpdate(user) || !canEditDetails(current)) throw new McmBusinessException("現在の状態では申請できません。");
        var files = repository.findTenpuByKikanId(selectedPeriod);
        if (contract.getKaiyakuDt() == null && files.stream().anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("再申請はできません。期間タブを作成し直して下さい。");
        if (files.stream().noneMatch(a -> McmConstants.SHONINJOTAI_SAKUSEICHU_CD.equals(a.getShoninjotai()) || McmConstants.SHONINJOTAI_SASHIMODOSHI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("契約依頼資料を添付してください。");
        save(contract, periods, inspections, payments, months, attachments, user);
        sinsei(contract.getTkKeiyakuId(), selectedPeriod, user);
    }

    private void validateCount(BigDecimal count, long checked, String label) {
        String error = validation.validateField(label + "回数", count == null ? null : count.toPlainString(),
                true, com.daifuku.mcm.common.AppConstants.VALIDATE_TYPE_NUMERIC, 0, 0);
        if (error != null) throw new InputCheckException(error);
        if (count == null || count.signum() < 0 || count.stripTrailingZeros().scale() > 0 || count.compareTo(new BigDecimal("99")) > 0)
            throw new InputCheckException(label + "回数は0～99の整数で入力してください。");
        if (count.compareTo(BigDecimal.valueOf(checked)) < 0)
            throw new InputCheckException(label + "回数と月のチェックの数が一致していません。");
    }

    public boolean canEditDetails(McmTkKeiyakuEntity contract) {
        return contract != null && !McmConstants.JOTAI_HAKI.equals(contract.getJotai())
                && !McmConstants.JOTAI_KAIYAKU.equals(contract.getJotai())
                && !McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(contract.getShoninjotai())
                && !McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(contract.getShoninjotai());
    }

    private static boolean sameId(BigDecimal left, BigDecimal right) {
        return left != null && right != null && left.compareTo(right) == 0;
    }

    private static ExclusiveControlException conflict() {
        return new ExclusiveControlException("他のユーザがデータを変更した可能性があります。処理をやり直してください。");
    }

    private static void checkVersion(LocalDateTime expected, LocalDateTime actual) {
        if (!java.util.Objects.equals(expected, actual)) throw conflict();
    }

    @Transactional(readOnly = true)
    public List<Map<String, Object>> loadKikiList(BigDecimal periodId) {
        return repository.findKikiList(periodId);
    }

    // ===================================================================
    // 添付ファイル削除
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void deleteTenpu(BigDecimal tkTenpuId, BigDecimal contractId, String loginUser) {
        if (!canUpdate(loginUser)) throw new McmBusinessException("権限がないため実行できません。");
        McmTkKeiyakuEntity contract = repository.lockKeiyaku(contractId);
        if (contract == null) throw conflict();
        McmTkTenpuEntity attachment = repository.findKikanByKeiyakuId(contractId).stream()
                .flatMap(p -> repository.findTenpuByKikanId(p.getTkKikanId()).stream())
                .filter(a -> sameId(a.getTkTenpuId(), tkTenpuId)).findFirst().orElseThrow(Mcm1005uService::conflict);
        if (!McmConstants.SHONINJOTAI_SAKUSEICHU_CD.equals(attachment.getShoninjotai())
                && !McmConstants.SHONINJOTAI_SASHIMODOSHI_CD.equals(attachment.getShoninjotai()))
            throw new McmBusinessException("承認処理中または承認済みの添付ファイルは削除できません。");
        repository.deleteTenpu(tkTenpuId);
    }

    // ===================================================================
    // MCM1006U 連携: 変更ボタン用 Delivery 構築
    // 元VB: Mcm1005uTabControl.vb HenkoButton_Click — buildDelivery相当
    // ===================================================================

    /**
     * 変更ボタン押下時に MCM1006U へ渡す Delivery Form を構築する。
     * 期間に既存の機器レコードが存在する場合はその内容を全て checkFlg=1 でセットする。
     * seniMotoKbn=1(INSERT) の場合はkiki行なし、seniMotoKbn=2(TAB_HENKO) の場合は既存行をロード。
     */
    @Transactional(readOnly = true)
    public Mcm1006uForm buildMcm1006uDelivery(McmTkKikanEntity kikan, McmTkKeiyakuEntity keiyaku,
                                               int seniMotoKbn, String kaisiDt, String syuryoDt) {
        Mcm1006uForm form = new Mcm1006uForm();
        form.setSeniMotoKbn(seniMotoKbn);
        form.setNonyusakiId(kikan.getNonyusakiId());
        form.setNonyusakiCd(kikan.getNonyusakiCd());
        form.setNonyusakiNk(kikan.getNonyusakiNk());
        form.setPlantId(kikan.getPlantId());
        form.setSupportId(kikan.getSupportId());
        form.setPlantNk(kikan.getPlantNk());
        form.setTorihikisakiId(kikan.getTorihikisakiId());
        form.setTorihikisakiCd(kikan.getTorihikisakiCd());
        form.setTorihikisakiNk(kikan.getTorihikisakiNk());
        form.setKeiyakuNo(keiyaku.getKeiyakuNo());
        form.setTkKeiyakuId(keiyaku.getTkKeiyakuId());
        form.setTkKikanId(kikan.getTkKikanId());
        form.setKaisiDt(kaisiDt);
        form.setSyuryoDt(syuryoDt);

        if (seniMotoKbn == 2 && kikan.getTkKikanId() != null) {
            // 既存機器レコードを読み込み delivery にセット（restoreCheckState で利用される）
            BigDecimal tkKikanId = kikan.getTkKikanId();
            List<Mcm1006uForm.KoseiRowForm> koseiRows = new ArrayList<>();
            List<Mcm1006uForm.MeisaiRowForm> meisaiRows = new ArrayList<>();
            List<Mcm1006uForm.KotaiRowForm> kotaiRows = new ArrayList<>();
            List<Mcm1006uForm.TankaRowForm> tankaRows = new ArrayList<>();

            for (java.util.Map<String, Object> r : repository.findKoseiRawByKikanId(tkKikanId)) {
                Mcm1006uForm.KoseiRowForm row = new Mcm1006uForm.KoseiRowForm();
                row.setKikikoseiId(toBD(r.get("KIKIKOSEI_ID")));
                row.setKikikoseiNk(str(r.get("KIKIKOSEI_NK")));
                row.setSetNm(str(r.get("SET_NM")));
                row.setTehaiseiban(str(r.get("TEHAISEIBAN")));
                row.setHyojijun(toBD(r.get("HYOJIJUN")));
                row.setTkKikikoseiId(toBD(r.get("TK_KIKIKOSEI_ID")));
                row.setCheckFlg(1);
                row.setCheckFlgOld(1);
                koseiRows.add(row);
            }
            for (java.util.Map<String, Object> r : repository.findMeisaiRawByKikanId(tkKikanId)) {
                Mcm1006uForm.MeisaiRowForm row = new Mcm1006uForm.MeisaiRowForm();
                row.setKikikoseiId(toBD(r.get("KIKIKOSEI_ID")));
                row.setKikimeisaiId(toBD(r.get("KIKIMEISAI_ID")));
                row.setSeizomakerId(toBD(r.get("SEIZOMAKER_ID")));
                row.setSeizomakerNk(str(r.get("SEIZOMAKER_NK")));
                row.setKikihinmeiNk(str(r.get("KIKIHINMEI_NK")));
                row.setKikikatashiki(str(r.get("KIKIKATASHIKI")));
                row.setSuryoNm(toBD(r.get("SURYO_NM")));
                row.setTkKikikoseiId(toBD(r.get("TK_KIKIKOSEI_ID")));
                row.setTkKikimeisaiId(toBD(r.get("TK_KIKIMEISAI_ID")));
                row.setCheckFlg(1);
                row.setCheckFlgOld(1);
                meisaiRows.add(row);
            }
            for (java.util.Map<String, Object> r : repository.findKotaiRawByKikanId(tkKikanId)) {
                Mcm1006uForm.KotaiRowForm row = new Mcm1006uForm.KotaiRowForm();
                row.setKikikoseiId(toBD(r.get("KIKIKOSEI_ID")));
                row.setKikimeisaiId(toBD(r.get("KIKIMEISAI_ID")));
                row.setKotaikanriId(toBD(r.get("KOTAIKANRI_ID")));
                row.setKotaiNk(str(r.get("KOTAI_NK")));
                row.setSerialNo(str(r.get("SERIAL_NO")));
                row.setTkKikikoseiId(toBD(r.get("TK_KIKIKOSEI_ID")));
                row.setTkKikimeisaiId(toBD(r.get("TK_KIKIMEISAI_ID")));
                row.setTkKotaimeisaiId(toBD(r.get("TK_KOTAIMEISAI_ID")));
                row.setCheckFlg(1);
                row.setCheckFlgOld(1);
                kotaiRows.add(row);
            }
            for (java.util.Map<String, Object> r : repository.findTankaRawByKikanId(tkKikanId)) {
                Mcm1006uForm.TankaRowForm row = new Mcm1006uForm.TankaRowForm();
                row.setKikikoseiId(toBD(r.get("KIKIKOSEI_ID")));
                row.setKikimeisaiId(toBD(r.get("KIKIMEISAI_ID")));
                row.setTmKikimeisaiId(toBD(r.get("TM_KIKIMEISAI_ID")));
                row.setTmTankaId(toBD(r.get("TM_TANKA_ID")));
                row.setTkKikikoseiId(toBD(r.get("TK_KIKIKOSEI_ID")));
                row.setTkKikimeisaiId(toBD(r.get("TK_KIKIMEISAI_ID")));
                row.setTkTankaId(toBD(r.get("TK_TANKA_ID")));
                row.setCheckFlg(1);
                row.setCheckFlgOld(1);
                tankaRows.add(row);
            }
            form.setKoseiRows(koseiRows);
            form.setMeisaiRows(meisaiRows);
            form.setKotaiRows(kotaiRows);
            form.setTankaRows(tankaRows);

            // 既存の TM_KEIYAKUJIKAN_ID リストをセット（step1 見積選択の初期チェック用）
            List<BigDecimal> tmKjIds = repository.findTmKeiyakujikanIdsByKikanId(tkKikanId);
            form.setTmKeiyakujikanIds(tmKjIds);
        }
        return form;
    }

    private BigDecimal toBD(Object v) {
        if (v == null) return null;
        if (v instanceof BigDecimal) return (BigDecimal) v;
        return new BigDecimal(v.toString());
    }

    private String str(Object v) {
        return v == null ? null : v.toString();
    }

    // ===================================================================
    // MCM1006U 連携: 選定結果適用
    // 元VB: Mcm1005uTabControl.vb — applyMcm1006uResult 相当
    // ===================================================================

    /**
     * MCM1006U（取引先契約機器選定）の選定結果を DB に適用する。
     * 対象期間の機器レコードを全削除後、checkFlg=1 の行のみ INSERT する。
     * 元VB: MCM1005Uタブ内 TabHenkoButtonClick → UpdateTKKikiData
     */
    @Transactional(rollbackFor = Exception.class)
    public void applyMcm1006uResult(Mcm1006uForm result, String loginUser) {
        BigDecimal tkKikanId = result.getTkKikanId();
        if (tkKikanId == null) return;

        // 既存機器レコード全削除（CASCADE: KOTAIMEISAI→TANKA→TENKEN→KIKIMEISAI→KIKIKOSEI）
        repository.deleteKikisForKikanId(tkKikanId);

        // 採番用最大ID取得
        BigDecimal koseiMaxId   = repository.getMaxKikikoseiId();
        BigDecimal meisaiMaxId  = repository.getMaxKikimeisaiId();
        BigDecimal kotaiMaxId   = repository.getMaxKotaimeisaiId();
        BigDecimal tankaMaxId   = repository.getMaxTankaId();

        // kikikoseiId → 新 TK_KIKIKOSEI_ID マッピング
        Map<BigDecimal, BigDecimal> koseiIdMap   = new HashMap<>();
        // kikimeisaiId → 新 TK_KIKIMEISAI_ID マッピング
        Map<BigDecimal, BigDecimal> meisaiIdMap  = new HashMap<>();

        // KIKIKOSEI INSERT
        for (Mcm1006uForm.KoseiRowForm r : result.getKoseiRows()) {
            if (r.getCheckFlg() != 1) continue;
            koseiMaxId = koseiMaxId.add(BigDecimal.ONE);
            koseiIdMap.put(r.getKikikoseiId(), koseiMaxId);
            repository.insertKikikosei(koseiMaxId, tkKikanId,
                r.getKikikoseiId(), r.getKikikoseiNk(), r.getSetNm(),
                r.getTehaiseiban(), r.getHyojijun(), loginUser);
        }

        // KIKIMEISAI INSERT
        for (Mcm1006uForm.MeisaiRowForm r : result.getMeisaiRows()) {
            if (r.getCheckFlg() != 1) continue;
            BigDecimal newKoseiId = koseiIdMap.get(r.getKikikoseiId());
            if (newKoseiId == null) continue;
            meisaiMaxId = meisaiMaxId.add(BigDecimal.ONE);
            meisaiIdMap.put(r.getKikimeisaiId(), meisaiMaxId);
            repository.insertKikimeisai(meisaiMaxId, newKoseiId,
                r.getKikikoseiId(), r.getKikimeisaiId(),
                r.getSeizomakerId(), r.getSeizomakerNk(),
                r.getKikihinmeiNk(), r.getKikikatashiki(),
                r.getSuryoNm(), null, r.getHyojijun(), loginUser);
        }

        // KOTAIMEISAI INSERT
        for (Mcm1006uForm.KotaiRowForm r : result.getKotaiRows()) {
            if (r.getCheckFlg() != 1) continue;
            BigDecimal newKoseiId  = koseiIdMap.get(r.getKikikoseiId());
            BigDecimal newMeisaiId = meisaiIdMap.get(r.getKikimeisaiId());
            if (newKoseiId == null || newMeisaiId == null) continue;
            kotaiMaxId = kotaiMaxId.add(BigDecimal.ONE);
            repository.insertKotaimeisai(kotaiMaxId, newKoseiId, newMeisaiId,
                r.getKikikoseiId(), r.getKikimeisaiId(),
                r.getKotaikanriId(), r.getKotaiNk(), r.getSerialNo(),
                r.getItijinonyugDt(), r.getSetchibasyo(),
                r.getTekkyoDt(), r.getEnchokeiyakukigenDt(), loginUser);
        }

        // TANKA INSERT
        for (Mcm1006uForm.TankaRowForm r : result.getTankaRows()) {
            if (r.getCheckFlg() != 1) continue;
            BigDecimal newKoseiId  = koseiIdMap.get(r.getKikikoseiId());
            BigDecimal newMeisaiId = meisaiIdMap.get(r.getKikimeisaiId());
            if (newKoseiId == null || newMeisaiId == null) continue;
            tankaMaxId = tankaMaxId.add(BigDecimal.ONE);
            repository.insertTanka(tankaMaxId, newKoseiId, newMeisaiId,
                r.getKikikoseiId(), r.getKikimeisaiId(),
                r.getTmKikimeisaiId(), r.getTmTankaId(),
                r.getPackFlg(), r.getKeiyakuNo(), r.getKeiyakunaiyo(),
                r.getTorihosyujikanId(), r.getTenkenumu(),
                r.getHosyuhoho(), r.getServicekeitai(),
                r.getHyojunKin(), r.getSikiriKin(), r.getSuryoNm(),
                r.getKaisiDt(), r.getSyuryoDt(), loginUser);
        }
    }

    // ===================================================================
    // 申請処理
    // 元VB: SinseiButton_Click
    //   MCM_TK_KEIYAKU.SHONINJOTAI → 審査中
    //   MCM_TK_KIKAN.YUKO_FLG      → 審査中(-1)
    //   MCM_TK_TENPU.SHONINJOTAI   → 審査中(作成中/差戻しのみ)
    // ===================================================================

    @Transactional(rollbackFor = Exception.class)
    public void sinsei(BigDecimal tkKeiyakuId, BigDecimal tkKikanId, String loginUser) {
        if (!canUpdate(loginUser)) throw new McmBusinessException("権限がないため実行できません。");
        McmTkKeiyakuEntity contract = repository.lockKeiyaku(tkKeiyakuId);
        if (contract == null || repository.findKikanByKeiyakuId(tkKeiyakuId).stream().noneMatch(p -> sameId(p.getTkKikanId(), tkKikanId))) throw conflict();
        if (!canEditDetails(contract)) throw new McmBusinessException("現在の状態では申請できません。");
        List<McmTkTenpuEntity> attachments = repository.findTenpuByKikanId(tkKikanId);
        if (contract.getKaiyakuDt() == null && attachments.stream().anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("再申請はできません。期間タブを作成し直して下さい。");
        if (attachments.stream().noneMatch(a -> McmConstants.SHONINJOTAI_SAKUSEICHU_CD.equals(a.getShoninjotai()) || McmConstants.SHONINJOTAI_SASHIMODOSHI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("契約依頼資料を添付してください。");
        repository.updateKeiyakuShoninjotai(tkKeiyakuId,
                McmConstants.SHONINJOTAI_SHINSACHU_CD, loginUser);
        repository.updateKikanShoninjotai(tkKikanId,
                new BigDecimal(McmConstants.YUKI_FLG_SHONIN), loginUser);
        repository.updateTenpuShoninjotaiToShinsachu(tkKikanId, loginUser);
    }
}
