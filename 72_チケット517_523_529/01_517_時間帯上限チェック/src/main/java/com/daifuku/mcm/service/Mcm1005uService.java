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

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.McmConstants;
import com.daifuku.mcm.entity.McmTkKeiyakuEntity;
import com.daifuku.mcm.entity.McmTkKikanEntity;
import com.daifuku.mcm.entity.McmTkSiharaiEntity;
import com.daifuku.mcm.entity.McmTkTenkenEntity;
import com.daifuku.mcm.entity.McmTkTenpuEntity;
import com.daifuku.mcm.exception.DataNotChangedException;
import com.daifuku.mcm.exception.ExclusiveControlException;
import com.daifuku.mcm.exception.InputCheckException;
import com.daifuku.mcm.exception.McmBusinessException;
import com.daifuku.mcm.form.Mcm1006uForm;
import com.daifuku.mcm.repository.Mcm1005uRepository;

@Service
public class Mcm1005uService {
    // VBの各DataTable_Item.xmlのVALUE。保存用コードは変換せず表示用のコピーに設定する。
    private static final Map<String, String> INSPECTION_LABELS = Map.of("0", "点検なし", "1", "点検あり");
    private static final Map<String, String> MAINTENANCE_METHOD_LABELS = Map.of(
            "F", "ｵﾝｻｲﾄ", "S", "ｾﾝﾄﾞﾊﾞｯｸ", "C", "ｺﾝﾃｯｸ製品",
            "H", "持ち帰り", "I", "ｽﾎﾟｯﾄ", "T", "TEL対応");
    private static final Map<String, String> SERVICE_TYPE_LABELS = Map.of(
            "1", "1：維持保守", "2", "2：出張修理", "3", "3：持込修理",
            "4", "4：引取修理", "5", "5：ﾌﾟﾘﾝﾀ出張修理", "6", "6：ｿﾌﾄｻﾎﾟｰﾄｻｰﾋﾞｽ");

    /** VBの契約時間SQLと同じ書式。IDを時間数として表示しない。 */
    public Map<String, String> maintenanceTimeLabels() {
        Map<String, String> labels = new HashMap<>();
        for (var row : repository.findMaintenanceTimes()) {
            String category = displayCode(row.get("HOSYU_KBN"));
            String label = displayCode(row.get("HOSYUJIKAN_DT")) + "H"
                    + (category.isEmpty() ? "" : "(" + category + ")")
                    + " " + weekdayLabel(row.get("KAISIYOBI")) + "～" + weekdayLabel(row.get("SYURYOYOBI"))
                    + " " + displayCode(row.get("KAISIJIKAN_DT")) + "～" + displayCode(row.get("SYURYOJIKAN_DT"));
            if (row.get("REIGAIYOBI") != null) {
                label += " (" + weekdayLabel(row.get("REIGAIYOBI")) + " "
                        + displayCode(row.get("REIGAIKAISIJIKAN_DT")) + "～"
                        + displayCode(row.get("REIGAISYURYOJIKAN_DT")) + ")";
            }
            labels.put(displayCode(row.get("TORIHOSYUJIKAN_ID")), label);
        }
        return labels;
    }

    public List<Map<String, Object>> equipmentDisplay(List<Map<String, Object>> rows, Map<String, String> timeLabels) {
        List<Map<String, Object>> display = new ArrayList<>();
        for (var row : rows) {
            Map<String, Object> copy = new HashMap<>(row);
            copy.put("TORIHOSYUJIKAN_LABEL", equipmentLabel(timeLabels, row.get("TORIHOSYUJIKAN_ID")));
            copy.put("TENKENUMU_LABEL", equipmentLabel(INSPECTION_LABELS, row.get("TENKENUMU")));
            copy.put("HOSYUHOHO_LABEL", equipmentLabel(MAINTENANCE_METHOD_LABELS, row.get("HOSYUHOHO")));
            copy.put("SERVICEKEITAI_LABEL", equipmentLabel(SERVICE_TYPE_LABELS, row.get("SERVICEKEITAI")));
            display.add(copy);
        }
        return display;
    }

    private static String equipmentLabel(Map<String, String> labels, Object value) {
        String code = displayCode(value);
        // 未定義値は隠さず原値を残し、未設定は空欄にする。
        return labels.getOrDefault(code, code);
    }

    private static String displayCode(Object value) {
        if (value == null) return "";
        if (value instanceof Number) return new BigDecimal(value.toString()).stripTrailingZeros().toPlainString();
        return value.toString().trim();
    }

    private static String weekdayLabel(Object value) {
        return switch (displayCode(value)) {
            case "1" -> "日"; case "2" -> "月"; case "3" -> "火"; case "4" -> "水";
            case "5" -> "木"; case "6" -> "金"; case "7" -> "土"; default -> "";
        };
    }

    /** Server-owned, unsaved selection. Negative IDs exist only in this preview. */
    @Autowired
    private com.daifuku.mcm.common.TantoNameResolver tantoNames;

    /** 表示情報を別に作り、排他判定・登録に用いるEntityの原値は保持する。 */
    public Map<String, Map<String, String>> auditDisplay(McmTkKeiyakuEntity contract,
            List<McmTkTenkenEntity> inspections) {
        List<String> ids = new ArrayList<>();
        ids.add(contract.getCreatedBy()); ids.add(contract.getLastupdateBy());
        for (var row : inspections) { ids.add(row.getCreatedBy()); ids.add(row.getLastupdateBy()); }
        Map<String, String> names = tantoNames.resolveAll(ids);
        Map<String, Map<String, String>> result = new java.util.HashMap<>();
        result.put("contract", auditValues(contract.getCreatedDt(), contract.getCreatedBy(), contract.getLastupdateDt(), contract.getLastupdateBy(), names));
        for (var row : inspections)
            result.put("tenken_" + row.getTkTenkenId(), auditValues(row.getCreatedDt(), row.getCreatedBy(), row.getLastupdateDt(), row.getLastupdateBy(), names));
        return result;
    }

    private Map<String, String> auditValues(LocalDateTime created, String creator, LocalDateTime updated,
            String updater, Map<String, String> names) {
        String c = creator == null ? "" : creator.trim(), u = updater == null ? "" : updater.trim();
        return Map.of("createdDt", com.daifuku.mcm.common.DateUtils.formatDateTimeMinutes(created),
                "lastupdateDt", com.daifuku.mcm.common.DateUtils.formatDateTimeMinutes(updated),
                "createdBy", names.getOrDefault(c,c), "lastupdateBy", names.getOrDefault(u,u));
    }

    public Map<String,Map<String,String>> periodAuditDisplay(List<McmTkKikanEntity> periods) {
        List<String> ids = new ArrayList<>();
        for(var period:periods){ids.add(period.getCreatedBy());ids.add(period.getLastupdateBy());}
        Map<String,String> names=tantoNames.resolveAll(ids);
        Map<String,Map<String,String>> result=new HashMap<>();
        for(var period:periods) result.put(period.getTkKikanId().toString(), auditValues(period.getCreatedDt(),period.getCreatedBy(),period.getLastupdateDt(),period.getLastupdateBy(),names));
        return result;
    }

    public Map<String,String> maintenanceCompanies(BigDecimal periodId, SelectionDraft draft) {
        boolean quote = periodId.signum() < 0;
        BigDecimal source = quote && draft != null && !draft.selection().getTmKikanIds().isEmpty()
            ? draft.selection().getTmKikanIds().get(0) : periodId;
        Map<String,String> choices = new java.util.LinkedHashMap<>();
        if (source.signum() > 0) for (var row : repository.findMaintenanceCompanies(source, quote))
            choices.put(uiId(toBD(row.get("HOSHUGAISHA_ID"))), str(row.get("HOSHUGAISHA_NK")));
        return choices;
    }

    /** VB Settings.ShinseiKisanDt。配備先の設定が異なる場合は上書き可能。 */
    @org.springframework.beans.factory.annotation.Value("${mcm.contract.application-start-date:2009-05-01}")
    private String applicationStartDate = "2009-05-01";

    public int initialPeriodIndex(List<McmTkKikanEntity> periods) {
        int selected = Math.max(0, periods.size() - 1);
        var today = java.time.LocalDate.now();
        for (int i = 0; i < periods.size(); i++) {
            var p = periods.get(i);
            if (p.getYukoFlg() != null && p.getYukoFlg().signum() == 0 && p.getKaisiDt() != null && p.getSyuryoDt() != null
                    && !today.isBefore(p.getKaisiDt().toLocalDate()) && !today.isAfter(p.getSyuryoDt().toLocalDate().plusDays(1))) selected = i;
        }
        return selected;
    }

    public String periodLabel(McmTkKikanEntity period, List<McmTkTenpuEntity> files) {
        String label = reportDate(period.getKaisiDt());
        if (period.getTkKikanId().signum() < 0) return label;
        if (period.getYukoFlg() != null && period.getYukoFlg().compareTo(BigDecimal.ONE) == 0) return label + "*";
        if (period.getYukoFlg() != null && period.getYukoFlg().signum() < 0) return label + "(申請中)";
        if (period.getYukoFlg() != null && period.getYukoFlg().signum() == 0 && period.getKaisiDt() != null
                && !period.getKaisiDt().toLocalDate().isBefore(java.time.LocalDate.parse(applicationStartDate))
                && files.stream().noneMatch(f -> java.util.Set.of("1", "2", "3").contains(String.valueOf(f.getShoninjotai())))) return label + "(未)";
        return label;
    }

    /** MSG_0076E（CPMessage.xml）。変更摘要開始日が適用期間の範囲外のとき。 */
    private static final String MSG_0076_CHANGE_START_RANGE = "{0}から{1}の間の日付を設定してください。";
    private static final java.time.format.DateTimeFormatter MSG_DATE_FMT = java.time.format.DateTimeFormatter.ofPattern("yyyy/MM/dd");

    /** VB変更摘要日の範囲。未来のタブは一つ前の有効期間を基準にする。 */
    public void validateChangeStart(McmTkKikanEntity period, List<McmTkKikanEntity> periods, java.time.LocalDate start) {
        var lower = period;
        boolean future = period.getKaisiDt().toLocalDate().isAfter(java.time.LocalDate.now());
        if (future) {
            lower = null;
            for (var p : periods) {
                if (sameId(p.getTkKikanId(), period.getTkKikanId())) break;
                if (p.getYukoFlg() != null && p.getYukoFlg().signum() == 0) lower = p;
            }
        }
        if (lower != null && (start.isBefore(lower.getKaisiDt().toLocalDate()) || start.isAfter(lower.getSyuryoDt().toLocalDate().plusDays(1))))
            // MSG_0076E「{0}から{1}の間の日付を設定してください。」（{0}=適用期間の開始日、{1}=終了日の翌日。yyyy/MM/dd）
            throw new InputCheckException(MSG_0076_CHANGE_START_RANGE
                    .replace("{0}", lower.getKaisiDt().toLocalDate().format(MSG_DATE_FMT))
                    .replace("{1}", lower.getSyuryoDt().toLocalDate().plusDays(1).format(MSG_DATE_FMT)));
    }

    public record SelectionDraft(String token, String user, Mcm1006uForm selection,
            List<McmTkKikanEntity> periods, List<Map<String,Object>> data, BigDecimal periodId) { }
    public record SavedSelection(BigDecimal contractId, BigDecimal periodId,
            Map<String,String> periodIds, Map<String,String> inspectionIds) { }
    private static String uiId(BigDecimal value) { return value.stripTrailingZeros().toPlainString(); }

    private static void requireSameIds(List<BigDecimal> expected, List<BigDecimal> actual) {
        var left = new java.util.TreeSet<BigDecimal>(); var right = new java.util.TreeSet<BigDecimal>();
        if (expected.stream().anyMatch(java.util.Objects::isNull) || actual.stream().anyMatch(java.util.Objects::isNull)) throw conflict();
        left.addAll(expected); right.addAll(actual);
        if (left.size() != expected.size() || right.size() != actual.size() || !left.equals(right)) throw conflict();
    }

    /** Selection, form values and optional application are committed together, only on POST save. */
    @Transactional(rollbackFor = Exception.class)
    @SuppressWarnings("unchecked")
    public SavedSelection saveSelection(SelectionDraft draft, McmTkKeiyakuEntity submitted,
            List<McmTkKikanEntity> periodInputs, List<McmTkTenkenEntity> inspectionInputs,
            List<McmTkSiharaiEntity> paymentInputs, List<Boolean[]> months,
            BigDecimal selectedId, String operation, String user) {
        if (draft == null || !draft.user().equals(user) || !canUpdate(user)) throw conflict();
        var selection = copySelection(draft.selection());
        boolean fresh = selection.getTkKeiyakuId() == null;
        if (!sameId(submitted.getTkKeiyakuId(), fresh ? BigDecimal.ZERO : selection.getTkKeiyakuId())) throw conflict();
        requireSameIds(draft.periods().stream().map(McmTkKikanEntity::getTkKikanId).toList(), periodInputs.stream().map(McmTkKikanEntity::getTkKikanId).toList());
        List<McmTkTenkenEntity> expectedInspections = new ArrayList<>();
        List<McmTkSiharaiEntity> expectedPayments = new ArrayList<>();
        for (var data : draft.data()) {
            expectedInspections.addAll((List<McmTkTenkenEntity>) data.get("tenken"));
            if (data.get("siharai") instanceof McmTkSiharaiEntity payment) expectedPayments.add(payment);
        }
        requireSameIds(expectedInspections.stream().map(McmTkTenkenEntity::getTkTenkenId).toList(), inspectionInputs.stream().map(McmTkTenkenEntity::getTkTenkenId).toList());
        requireSameIds(expectedPayments.stream().map(McmTkSiharaiEntity::getTkSiharaiId).toList(), paymentInputs.stream().map(McmTkSiharaiEntity::getTkSiharaiId).toList());
        if (months.size() != paymentInputs.size() || draft.periods().stream().noneMatch(p -> sameId(p.getTkKikanId(), selectedId))) throw conflict();
        var contract = copyBean(submitted, McmTkKeiyakuEntity::new);
        var originallyEditable = new java.util.TreeSet<BigDecimal>();
        if (!fresh) {
            var originalContract = repository.lockKeiyaku(selection.getTkKeiyakuId());
            var originals = repository.lockKikanList(selection.getTkKeiyakuId());
            originals.stream().filter(p -> canEditPeriodDetails(originalContract,p,originals)).forEach(p -> originallyEditable.add(p.getTkKikanId()));
        }
        // 新規・機器選定後の保存でも、最初のINSERT/UPDATE前に入力を検査する。
        for (var input : periodInputs) {
            if (fresh || sameId(input.getTkKikanId(), draft.periodId()) || originallyEditable.contains(input.getTkKikanId()))
                validateContractHours(input.getKeiyakujikantai());
        }
        BigDecimal actualPeriodId;
        if (fresh) {
            // Fail before the first write if the parent masters disappeared while editing.
            repository.newContractPeriod(selection.getPlantId(), selection.getTorihikisakiId());
            contract.setTkKeiyakuId(repository.getMaxKeiyakuId().add(BigDecimal.ONE));
            contract.setJotai(McmConstants.JOTAI_KEIYAKU); contract.setShoninjotai(McmConstants.SHONINJOTAI_SAKUSEICHU_CD);
            repository.insertKeiyaku(contract, user);
            var period = copyBean(draft.periods().get(0), McmTkKikanEntity::new);
            actualPeriodId = repository.getMaxKikanId().add(BigDecimal.ONE);
            period.setTkKikanId(actualPeriodId); period.setTkKeiyakuId(contract.getTkKeiyakuId());
            repository.insertKikan(period, user);
            var payment = new McmTkSiharaiEntity(); payment.setTkSiharaiId(repository.getMaxSiharaiId().add(BigDecimal.ONE));
            payment.setTkKikanId(actualPeriodId); payment.setKaisu(BigDecimal.ZERO); repository.insertSiharai(payment, user);
            selection.getKoseiRows().forEach(r -> r.setCheckFlgOld(0)); selection.getMeisaiRows().forEach(r -> r.setCheckFlgOld(0));
            selection.getKotaiRows().forEach(r -> r.setCheckFlgOld(0)); selection.getTankaRows().forEach(r -> r.setCheckFlgOld(0));
            selection.setTankaDelRows(new ArrayList<>());
            persistSelectionRows(selection, actualPeriodId, user);
        } else {
            actualPeriodId = applyMcm1006uResult(selection, user);
            if (repository.lockSiharai(actualPeriodId) == null) {
                var payment = new McmTkSiharaiEntity(); payment.setTkSiharaiId(repository.getMaxSiharaiId().add(BigDecimal.ONE));
                payment.setTkKikanId(actualPeriodId); payment.setKaisu(BigDecimal.ZERO); repository.insertSiharai(payment, user);
            }
        }
        BigDecimal contractId = contract.getTkKeiyakuId();
        var storedPeriods = repository.lockKikanList(contractId);
        Map<String,String> periodIds = new HashMap<>(), inspectionIds = new HashMap<>();
        List<McmTkKikanEntity> periods = new ArrayList<>();
        for (var input : periodInputs) {
            var p = copyBean(input, McmTkKikanEntity::new);
            if (sameId(p.getTkKikanId(), draft.periodId())) p.setTkKikanId(actualPeriodId);
            var stored = storedPeriods.stream().filter(k -> sameId(k.getTkKikanId(), p.getTkKikanId())).findFirst().orElseThrow(Mcm1005uService::conflict);
            periodIds.put(uiId(input.getTkKikanId()), uiId(stored.getTkKikanId()));
            // Selection may shorten the old tab. Keep its entered notes/method, but never undo those history dates.
            if (!sameId(p.getTkKikanId(),actualPeriodId) && originallyEditable.contains(p.getTkKikanId())) {
                p.setKaisiDt(stored.getKaisiDt()); p.setSyuryoDt(stored.getSyuryoDt());
            }
            p.setTkKeiyakuId(contractId); p.setLastupdateDt(stored.getLastupdateDt()); periods.add(p);
        }
        List<McmTkTenkenEntity> inspections = new ArrayList<>();
        for (var input : inspectionInputs) {
            var expected = expectedInspections.stream().filter(t -> sameId(t.getTkTenkenId(), input.getTkTenkenId())).findFirst().orElseThrow(Mcm1005uService::conflict);
            BigDecimal periodId = sameId(expected.getTkKikanId(), draft.periodId()) ? actualPeriodId : expected.getTkKikanId();
            var stored = repository.lockTenkenList(periodId).stream().filter(t ->
                    inspectionGroup(t.getKikikoseiId(), t.getOyakikibunruiCd()).equals(inspectionGroup(expected.getKikikoseiId(), expected.getOyakikibunruiCd())))
                    .findFirst().orElseThrow(Mcm1005uService::conflict);
            var t = copyBean(input, McmTkTenkenEntity::new); t.setTkTenkenId(stored.getTkTenkenId()); t.setLastupdateDt(stored.getLastupdateDt()); inspections.add(t);
            inspectionIds.put(uiId(expected.getTkTenkenId()), uiId(stored.getTkTenkenId()));
        }
        List<McmTkSiharaiEntity> payments = new ArrayList<>();
        for (var input : paymentInputs) {
            var expected = expectedPayments.stream().filter(p -> sameId(p.getTkSiharaiId(), input.getTkSiharaiId())).findFirst().orElseThrow(Mcm1005uService::conflict);
            if (!sameId(input.getTkKikanId(), expected.getTkKikanId())) throw conflict();
            BigDecimal periodId = sameId(expected.getTkKikanId(), draft.periodId()) ? actualPeriodId : expected.getTkKikanId();
            var stored = repository.lockSiharai(periodId); if (stored == null) throw conflict();
            var p = copyBean(input, McmTkSiharaiEntity::new); p.setTkSiharaiId(stored.getTkSiharaiId()); p.setTkKikanId(periodId); p.setLastupdateDt(stored.getLastupdateDt()); payments.add(p);
        }
        contract.setLastupdateDt(repository.lockKeiyaku(contractId).getLastupdateDt());
        BigDecimal selected = sameId(selectedId, draft.periodId()) ? actualPeriodId : selectedId;
        if ("report".equals(operation)) checkReport(contractId, selected, contract.getKaiyakuDt() != null, user);
        if ("sinsei".equals(operation)) saveAndSinseiInternal(contract, periods, inspections, payments, months, List.of(), selected, user, originallyEditable);
        else saveInternal(contract, periods, inspections, payments, months, List.of(), user, originallyEditable, false);
        if (fresh) initializeDeviceTerms(contract, user);
        return new SavedSelection(contractId, selected, Map.copyOf(periodIds), Map.copyOf(inspectionIds));
    }

    /** VB UpdButton: initialize missing individual dates only when creating a contract. */
    private void initializeDeviceTerms(McmTkKeiyakuEntity contract, String user) {
        LocalDateTime firstContract = contract.getShokaiKeiyakuDt() != null ? contract.getShokaiKeiyakuDt() : contract.getKeiyakuDt();
        for (var device : repository.lockNewContractDeviceTerms(contract.getTkKeiyakuId(), java.time.LocalDate.now())) {
            LocalDateTime base = switch (device.baseKind() == null ? "" : device.baseKind()) {
                case McmConstants.KISANBI_KBN_ITIJI -> device.firstDelivery();
                case McmConstants.KISANBI_KBN_NIJI -> device.secondDelivery();
                case McmConstants.KISANBI_KBN_KEIYAKU -> firstContract;
                default -> null;
            };
            java.time.LocalDate expiry = null;
            if (base != null && device.years() != null && device.years().signum() > 0) {
                try {
                    expiry = base.toLocalDate().plusYears(device.years().intValueExact()).minusDays(1);
                    if (expiry.getYear() > 9999) throw new java.time.DateTimeException("Out of range");
                } catch (ArithmeticException | java.time.DateTimeException ex) {
                    throw new InputCheckException("機器個体の契約可能期間を確認してください。契約期限を計算できません。");
                }
            }
            java.time.LocalDate makerEnd = device.makerEnd() == null ? null : device.makerEnd().toLocalDate();
            java.time.LocalDate deadline = makerEnd == null ? expiry : expiry == null || makerEnd.isBefore(expiry) ? makerEnd : expiry;
            if (deadline != null || expiry != null) repository.fillMissingDeviceTerms(device.id(), deadline, expiry, user);
        }
    }

    private static <T> T copyBean(T source, java.util.function.Supplier<T> factory) {
        T target = factory.get(); org.springframework.beans.BeanUtils.copyProperties(source, target); return target;
    }
    private static <T> List<T> copySelectionRows(List<T> source, java.util.function.Supplier<T> factory) { if (source == null) return null; List<T> rows = new ArrayList<>(); for (T row : source) rows.add(copyBean(row, factory)); return rows; }
    private LocalDateTime selectionDate(String value, String label) {
        String error = validation.validateField(label, value, true, com.daifuku.mcm.common.AppConstants.VALIDATE_TYPE_DATE, 10, 0);
        if (error != null) throw new InputCheckException(error);
        try { return java.time.LocalDate.parse(value, java.time.format.DateTimeFormatter.ofPattern("uuuu/MM/dd")
                .withResolverStyle(java.time.format.ResolverStyle.STRICT)).atStartOfDay(); }
        catch (java.time.format.DateTimeParseException ex) { throw new InputCheckException("契約期間を正しく入力してください。"); }
    }
    private static Mcm1006uForm copySelection(Mcm1006uForm source) { Mcm1006uForm target = copyBean(source, Mcm1006uForm::new); target.setTmKeiyakujikanIds(source.getTmKeiyakujikanIds() == null ? null : new ArrayList<>(source.getTmKeiyakujikanIds())); target.setTmKikanIds(source.getTmKikanIds() == null ? null : new ArrayList<>(source.getTmKikanIds())); target.setMitsumoriRows(copySelectionRows(source.getMitsumoriRows(), Mcm1006uForm.MitsumoriRowForm::new)); target.setKoseiRows(copySelectionRows(source.getKoseiRows(), Mcm1006uForm.KoseiRowForm::new)); target.setMeisaiRows(copySelectionRows(source.getMeisaiRows(), Mcm1006uForm.MeisaiRowForm::new)); target.setKotaiRows(copySelectionRows(source.getKotaiRows(), Mcm1006uForm.KotaiRowForm::new)); target.setTankaRows(copySelectionRows(source.getTankaRows(), Mcm1006uForm.TankaRowForm::new)); target.setTankaDelRows(copySelectionRows(source.getTankaDelRows(), Mcm1006uForm.TankaRowForm::new)); return target; }
    /** Construct the screen without writing or allocating database IDs (VB DataSet behaviour). */
    @Transactional(readOnly = true)
    public SelectionDraft previewSelection(Mcm1006uForm source, String user) {
        if (!canUpdate(user)) throw new McmBusinessException("権限がないため実行できません。");
        if (source == null) throw conflict();
        var result = copySelection(source);
        boolean fresh = result.getTkKeiyakuId() == null;
        var start = selectionDate(result.getKaisiDt(), "期間開始日");
        var end = selectionDate(result.getSyuryoDt(), "期間終了日");
        if (end.isBefore(start)) throw new InputCheckException("終了日が開始日より前になっています。");
        List<McmTkKikanEntity> periods = new ArrayList<>();
        McmTkKikanEntity target;
        boolean separate;
        if (fresh) {
            if (result.getSeniMotoKbn() != 1 || result.getPlantId() == null || result.getTorihikisakiId() == null)
                throw conflict();
            target = repository.newContractPeriod(result.getPlantId(), result.getTorihikisakiId());
            target.setTkKeiyakuId(BigDecimal.ZERO); target.setYukoFlg(BigDecimal.ZERO);
            repository.fillContractTerms(target, result.getTmKeiyakujikanIds());
            separate = true;
        } else {
            var current = repository.findKeiyakuById(result.getTkKeiyakuId());
            if (current == null) throw conflict();
            if (result.getSelectionVersion() == null || !result.getSelectionVersion().equals(repository.selectionVersion(result.getTkKeiyakuId(), false))) throw conflict();
            repository.findKikanByKeiyakuId(result.getTkKeiyakuId()).forEach(p -> periods.add(copyBean(p, McmTkKikanEntity::new)));
            var old = periods.stream().filter(p -> sameId(p.getTkKikanId(), result.getTkKikanId())).findFirst().orElseThrow(Mcm1005uService::conflict);
            if (!canChangePeriod(current, old, periods)) throw new McmBusinessException("現在の期間では変更できません。");
            var latest = periods.stream().filter(p -> p.getYukoFlg() == null || p.getYukoFlg().signum() == 0)
                    .max(java.util.Comparator.comparing(McmTkKikanEntity::getKaisiDt)).orElse(null);
            if (!isUnlocked(result.getTkKeiyakuId()) && (latest == null || !sameId(latest.getTkKikanId(), old.getTkKikanId())))
                throw new InputCheckException("最新の有効な期間から機器を変更してください。");
            validateChangeStart(old, periods, start.toLocalDate());
            boolean approved = repository.findTenpuByKikanId(old.getTkKikanId()).stream()
                    .anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai()));
            separate = !old.getKaisiDt().toLocalDate().isAfter(java.time.LocalDate.now()) || approved;
            target = separate ? copyBean(old, McmTkKikanEntity::new) : old;
            if (separate) {
                target.setYukoFlg(BigDecimal.ZERO);
                if (!start.isAfter(old.getKaisiDt())) old.setYukoFlg(BigDecimal.ONE);
                else if (!start.isAfter(old.getSyuryoDt())) old.setSyuryoDt(start.minusDays(1));
            }
        }
        if (separate) { target.setTkKikanId(BigDecimal.ONE.negate()); target.setLastupdateDt(null); periods.add(target); }
        target.setKaisiDt(start); target.setSyuryoDt(end);
        var equipment = selectionEquipment(result);
        target.setHyojungokeiKin(equipment.stream().map(r -> reportNumber(r.get("HYOJUN_KIN")).multiply(reportNumber(r.get("SURYO_NM")))).reduce(BigDecimal.ZERO, BigDecimal::add));
        target.setSikirisyokeiKin(equipment.stream().map(r -> reportNumber(r.get("SIKIRI_KIN")).multiply(reportNumber(r.get("SURYO_NM")))).reduce(BigDecimal.ZERO, BigDecimal::add));
        target.setSyusseinebikiKin(BigDecimal.valueOf(result.getSyusseinebikiGokei()));
        target.setSikirigokeiKin(target.getSikirisyokeiKin().subtract(target.getSyusseinebikiKin()));
        List<Map<String,Object>> data = new ArrayList<>();
        for (var period : periods) {
            var entry = new HashMap<String,Object>();
            boolean selected = sameId(period.getTkKikanId(), target.getTkKikanId());
            if (selected) {
                entry.put("kiki", equipment);
                entry.put("tenken", previewInspections(result, period.getTkKikanId(), separate ? List.of() : loadTenkenList(period.getTkKikanId())));
                var payment = separate ? new McmTkSiharaiEntity() : loadSiharai(period.getTkKikanId());
                if (payment == null) payment = new McmTkSiharaiEntity();
                if (payment.getTkSiharaiId() == null) { payment.setTkSiharaiId(BigDecimal.ONE.negate()); payment.setTkKikanId(period.getTkKikanId()); payment.setKaisu(BigDecimal.ZERO); }
                entry.put("siharai", payment); entry.put("tenpu", separate ? List.of() : loadTenpuList(period.getTkKikanId()));
            } else {
                entry.put("kiki", loadKikiList(period.getTkKikanId())); entry.put("tenken", loadTenkenList(period.getTkKikanId()));
                entry.put("siharai", loadSiharai(period.getTkKikanId())); entry.put("tenpu", loadTenpuList(period.getTkKikanId()));
            }
            data.add(entry);
        }
        return new SelectionDraft(java.util.UUID.randomUUID().toString(), user, result, periods, data, target.getTkKikanId());
    }

    /** Re-selection replaces the same unsaved tab, as VB TAB_FLG_NEW does. */
    @Transactional(readOnly = true)
    @SuppressWarnings("unchecked")
    public SelectionDraft previewSelection(Mcm1006uForm result, String user, SelectionDraft previous) {
        var next = previewSelection(result, user);
        if (previous == null) return next;
        if (!previous.user().equals(user)) throw conflict();
        var oldRows = (List<McmTkTenkenEntity>) previous.data().get(java.util.stream.IntStream.range(0,previous.periods().size()).filter(i -> sameId(previous.periods().get(i).getTkKikanId(),previous.periodId())).findFirst().orElseThrow()).get("tenken");
        var newRows = (List<McmTkTenkenEntity>) next.data().get(java.util.stream.IntStream.range(0,next.periods().size()).filter(i -> sameId(next.periods().get(i).getTkKikanId(),next.periodId())).findFirst().orElseThrow()).get("tenken");
        long min = oldRows.stream().map(McmTkTenkenEntity::getTkTenkenId).mapToLong(BigDecimal::longValue).min().orElse(0);
        for (var row : newRows) {
            var prior = oldRows.stream().filter(t -> inspectionGroup(t.getKikikoseiId(),t.getOyakikibunruiCd()).equals(inspectionGroup(row.getKikikoseiId(),row.getOyakikibunruiCd()))).findFirst().orElse(null);
            if (prior != null) row.setTkTenkenId(prior.getTkTenkenId());
            else if (row.getTkTenkenId().signum() < 0) row.setTkTenkenId(BigDecimal.valueOf(--min));
        }
        return next;
    }

    /** Own attachment changes advance the draft snapshot while the contract remains locked. */
    @Transactional(rollbackFor = Exception.class)
    public SelectionDraft changeDraftAttachment(SelectionDraft draft, String user, Runnable change) {
        if (!draft.user().equals(user) || draft.selection().getTkKeiyakuId() == null) throw conflict();
        var id = draft.selection().getTkKeiyakuId();
        repository.lockKeiyaku(id);
        if (!draft.selection().getSelectionVersion().equals(repository.selectionVersion(id,true))) throw conflict();
        change.run();
        var selection = copySelection(draft.selection());
        selection.setSelectionVersion(repository.selectionVersion(id,true));
        List<Map<String,Object>> data = new ArrayList<>();
        for (int i=0;i<draft.periods().size();i++) {
            var entry = new HashMap<>(draft.data().get(i));
            var periodId = draft.periods().get(i).getTkKikanId();
            if (periodId.signum()>0) entry.put("tenpu",loadTenpuList(periodId));
            data.add(entry);
        }
        return new SelectionDraft(draft.token(),user,selection,draft.periods(),data,draft.periodId());
    }

    private List<Map<String,Object>> selectionEquipment(Mcm1006uForm selection) {
        List<Map<String,Object>> rows = new ArrayList<>();
        for (var detail : selection.getMeisaiRows()) {
            if (detail.getCheckFlg() != 1) continue;
            var parent = selection.getKoseiRows().stream().filter(k -> k.getCheckFlg() == 1 && sameId(k.getKikikoseiId(), detail.getKikikoseiId())).findFirst().orElseThrow(Mcm1005uService::conflict);
            var prices = selection.getTankaRows().stream().filter(t -> t.getCheckFlg() == 1
                    && sameId(t.getKikikoseiId(), detail.getKikikoseiId()) && sameId(t.getKikimeisaiId(), detail.getKikimeisaiId()))
                    .sorted(java.util.Comparator.comparing(t -> t.getKaisiDt() == null ? "" : t.getKaisiDt())).toList();
            var price = prices.isEmpty() ? null : prices.get(0);
            Map<String,Object> row = new HashMap<>();
            row.put("KIKIKOSEI_NK", parent.getKikikoseiNk()); row.put("SEIZOMAKER_NK", detail.getSeizomakerNk());
            row.put("KIKIHINMEI_NK", detail.getKikihinmeiNk()); row.put("KIKIKATASHIKI", detail.getKikikatashiki());
            row.put("SURYO_NM", detail.getSuryoNm()); row.put("TEHAISEIBAN", parent.getTehaiseiban()); row.put("MAB_SUPPORT_ID", selection.getSupportId());
            row.put("HYOJUN_KIN", prices.stream().map(t -> reportNumber(t.getHyojunKin())).reduce(BigDecimal.ZERO, BigDecimal::add));
            row.put("SIKIRI_KIN", prices.stream().map(t -> reportNumber(t.getSikiriKin())).reduce(BigDecimal.ZERO, BigDecimal::add));
            row.put("TORIHOSYUJIKAN_ID", price == null ? null : price.getTorihosyujikanId()); row.put("TENKENUMU", price == null ? null : price.getTenkenumu());
            row.put("HOSYUHOHO", price == null ? null : price.getHosyuhoho()); row.put("SERVICEKEITAI", price == null ? null : price.getServicekeitai()); rows.add(row);
        }
        if (rows.isEmpty()) throw new InputCheckException("契約対象の機器が選択されていません。");
        return rows;
    }

    private List<McmTkTenkenEntity> previewInspections(Mcm1006uForm selection, BigDecimal periodId, List<McmTkTenkenEntity> existing) {
        var groups = new java.util.LinkedHashMap<InspectionGroup, Mcm1006uForm.MeisaiRowForm>();
        selection.getMeisaiRows().stream().filter(r -> r.getCheckFlg() == 1).forEach(r -> groups.putIfAbsent(inspectionGroup(r.getKikikoseiId(), r.getOyakikibunruiCd()), r));
        List<McmTkTenkenEntity> rows = new ArrayList<>();
        for (var group : groups.entrySet()) {
            var row = existing.stream().filter(t -> group.getKey().equals(inspectionGroup(t.getKikikoseiId(), t.getOyakikibunruiCd())))
                    .map(t -> copyBean(t, McmTkTenkenEntity::new)).findFirst().orElse(null);
            if (row == null) {
                row = newInspection(selection, group.getValue()); row.setTkTenkenId(BigDecimal.valueOf(-rows.size()-1)); row.setTkKikanId(periodId);
            }
            rows.add(row);
        }
        // synchronizeInspections は「チェックを変更したグループ」の点検しか削除しない。
        // 選定明細に対応しない既存点検（変更対象外）はDBに残るため、画面にも残して
        // 登録時の点検行照合（saveInternal）と一致させる。
        var changed = new java.util.HashSet<InspectionGroup>();
        selection.getMeisaiRows().stream().filter(r -> r.getCheckFlg() != r.getCheckFlgOld())
                .forEach(r -> changed.add(inspectionGroup(r.getKikikoseiId(), r.getOyakikibunruiCd())));
        for (var t : existing) {
            var key = inspectionGroup(t.getKikikoseiId(), t.getOyakikibunruiCd());
            if (!groups.containsKey(key) && !changed.contains(key)) rows.add(copyBean(t, McmTkTenkenEntity::new));
        }
        return rows;
    }

    private McmTkTenkenEntity newInspection(Mcm1006uForm selection, Mcm1006uForm.MeisaiRowForm source) {
        if (source.getOyakikibunruiCd() == null) throw conflict();
        var row = new McmTkTenkenEntity(); row.setKikikoseiId(source.getKikikoseiId()); row.setOyakikibunruiCd(source.getOyakikibunruiCd());
        row.setKikihinmeiNk(source.getOyakikihinmeiNk()); row.setKikikatashiki(source.getOyakikikatashiki());
        row.setKikikoseiNk(selection.getKoseiRows().stream().filter(k -> sameId(k.getKikikoseiId(), source.getKikikoseiId()))
                .map(Mcm1006uForm.KoseiRowForm::getKikikoseiNk).filter(java.util.Objects::nonNull).findFirst().orElse(""));
        row.setTenkenkaisu(BigDecimal.ZERO); row.setTenkenkanoyobi(""); row.setYakantaioumu("");
        if (!selection.getTmKikanIds().isEmpty() && source.getTmKikikoseiId() != null) {
            var settings = repository.findQuoteInspection(selection.getTmKikanIds().get(0), source.getTmKikikoseiId(), source.getOyakikibunruiCd());
            if (settings.size() == 1) { var values = settings.get(0); row.setTenkenkaisu(reportNumber(values.get("TENKENKAISU")));
                row.setTenkenkanoyobi(str(values.get("TENKENKANOYOBI"))); row.setYakantaioumu(str(values.get("YAKANTAIOUMU"))); }
        }
        return row;
    }
    private record EntryMode(BigDecimal contractId, String user, int mode) implements java.io.Serializable { }
    public void setEntryMode(jakarta.servlet.http.HttpSession session, BigDecimal id, String user, int mode) {
        session.setAttribute("MCM1005U_ENTRY", new EntryMode(id, user, mode));
    }
    private int entryMode(BigDecimal id) {
        if (!(org.springframework.web.context.request.RequestContextHolder.getRequestAttributes() instanceof org.springframework.web.context.request.ServletRequestAttributes attrs)) return 0;
        var session = attrs.getRequest().getSession(false);
        var auth = org.springframework.security.core.context.SecurityContextHolder.getContext().getAuthentication();
        if (session == null || auth == null || !(session.getAttribute("MCM1005U_ENTRY") instanceof EntryMode entry)
                || !entry.user().equals(auth.getName()) || !sameId(id, entry.contractId())) return 0;
        return entry.mode();
    }
    public boolean canSave(McmTkKeiyakuEntity contract) {
        return contract != null && (entryMode(contract.getTkKeiyakuId()) != 6 || isUnlocked(contract.getTkKeiyakuId()));
    }
    private BigDecimal validAmount(BigDecimal value) {
        if (value.stripTrailingZeros().scale() > 0)
            throw new InputCheckException("金額は整数で入力してください。");
        return value;
    }
    private static final String ADMIN_UNLOCK = "MCM1005U_ADMIN_UNLOCK";
    private record AdminUnlock(BigDecimal contractId, String user) implements java.io.Serializable { }

    public boolean canReleaseLock(String user) {
        var auth = org.springframework.security.core.context.SecurityContextHolder.getContext().getAuthentication();
        if (auth == null || !auth.isAuthenticated() || !auth.getName().equals(user)
                || auth.getAuthorities().stream().noneMatch(a -> "ROLE_UPDATE".equals(a.getAuthority()))) return false;
        return canUpdate(user) && repository.canReleaseLock(user);
    }

    public void clearUnlock(jakarta.servlet.http.HttpSession session) {
        session.removeAttribute(ADMIN_UNLOCK);
    }

    public void releaseLock(BigDecimal contractId, String user, jakarta.servlet.http.HttpSession session) {
        clearUnlock(session);
        if (!canReleaseLock(user)) throw new McmBusinessException("管理者用ロック解除の権限がありません。");
        if (contractId == null || repository.findKeiyakuById(contractId) == null)
            throw new InputCheckException("対象の契約が見つかりません。検索画面から開き直してください。");
        session.setAttribute(ADMIN_UNLOCK, new AdminUnlock(contractId, user));
    }

    /** VBの解除は当該画面の編集許可。DBの契約状態・承認状態は変更しない。 */
    public boolean isUnlocked(BigDecimal contractId) {
        if (!(org.springframework.web.context.request.RequestContextHolder.getRequestAttributes()
                instanceof org.springframework.web.context.request.ServletRequestAttributes attrs)) return false;
        var session = attrs.getRequest().getSession(false);
        if (session == null || !(session.getAttribute(ADMIN_UNLOCK) instanceof AdminUnlock grant)) return false;
        if (!canReleaseLock(grant.user())) { clearUnlock(session); return false; }
        return sameId(contractId, grant.contractId());
    }

    public boolean canEditWithUnlock(McmTkKeiyakuEntity contract) {
        return contract != null && (canEditDetails(contract) || isUnlocked(contract.getTkKeiyakuId()));
    }
    @Autowired private Mcm1003pExcelService reportExcel;

    /** VB KeiyakuHakoButton_Click: enforce period ownership and approved attachment restriction. */
    @Transactional(readOnly = true)
    public void checkReport(BigDecimal contractId, BigDecimal periodId, boolean cancellation, String user) {
        if (!canUpdate(user)) throw new McmBusinessException("権限がないため実行できません。");
        var current = repository.findKeiyakuById(contractId);
        if (current == null || repository.findKikanByKeiyakuId(contractId).stream().noneMatch(p -> sameId(p.getTkKikanId(), periodId))) throw conflict();
        if (!canEditWithUnlock(current)) throw new McmBusinessException("現在の状態では契約発行できません。");
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
        saveInternal(keiyaku,kikanList,tenkenList,siharaiList,siharaiTsukiList,tenpuList,loginUser,Set.of(),false);
    }

    /**
     * 登録ボタン（既存契約の通常更新）。表示時から実質的な変更がない場合は更新SQLを実行せず
     * DataNotChangedException（FWM_0010E「値が変更されていません。」）をスローする。
     * 【変換元】Mcm1005uScreen.vb UpdateButton_Click: checkChangedStatus=False → Throw CPDataNotChangedException
     */
    @Transactional(rollbackFor = Exception.class)
    public void saveIfChanged(McmTkKeiyakuEntity keiyaku,
                     List<McmTkKikanEntity> kikanList,
                     List<McmTkTenkenEntity> tenkenList,
                     List<McmTkSiharaiEntity> siharaiList,
                     List<Boolean[]> siharaiTsukiList,
                     List<McmTkTenpuEntity> tenpuList,
                     String loginUser) {
        saveInternal(keiyaku,kikanList,tenkenList,siharaiList,siharaiTsukiList,tenpuList,loginUser,Set.of(),true);
    }

    /** FWM_0010E（CPMessage.xml）: 値が変更されていません。 */
    public static final String FWM_0010E_NOT_CHANGED = "値が変更されていません。";

    /** 文字列は null と空文字を同一視し、数値はスケール差を無視して比較する。 */
    private static boolean sameValue(Object a, Object b) {
        if (a instanceof String s) a = s.isEmpty() ? null : s;
        if (b instanceof String s) b = s.isEmpty() ? null : s;
        if (a == null || b == null) return a == b;
        if (a instanceof BigDecimal x && b instanceof BigDecimal y) return x.compareTo(y) == 0;
        if (a instanceof Object[] x && b instanceof Object[] y) return java.util.Arrays.deepEquals(x, y);
        return a.equals(b);
    }

    /** 同一型のエンティティ同士を、無視するプロパティを除いて全プロパティで比較する。 */
    private static boolean sameValues(Object before, Object after, Set<String> ignore) {
        var left = new org.springframework.beans.BeanWrapperImpl(before);
        var right = new org.springframework.beans.BeanWrapperImpl(after);
        for (var pd : left.getPropertyDescriptors()) {
            String name = pd.getName();
            if ("class".equals(name) || ignore.contains(name)) continue;
            if (!left.isReadableProperty(name) || !right.isReadableProperty(name)) continue;
            if (!sameValue(left.getPropertyValue(name), right.getPropertyValue(name))) return false;
        }
        return true;
    }

    private static boolean sameMonths(List<Integer> stored, String[] posted) {
        var expected = new java.util.TreeSet<Integer>(stored);
        var actual = new java.util.TreeSet<Integer>();
        for (int i = 0; i < 12; i++) if (posted != null && i < posted.length && "1".equals(posted[i])) actual.add(i + 1);
        return expected.equals(actual);
    }

    private static boolean sameMonths(List<Integer> stored, Boolean[] posted) {
        String[] flags = new String[12];
        for (int i = 0; i < 12; i++) flags[i] = posted != null && i < posted.length && Boolean.TRUE.equals(posted[i]) ? "1" : "0";
        return sameMonths(stored, flags);
    }

    private void saveInternal(McmTkKeiyakuEntity keiyaku, List<McmTkKikanEntity> kikanList,
            List<McmTkTenkenEntity> tenkenList,List<McmTkSiharaiEntity> siharaiList,List<Boolean[]> siharaiTsukiList,
            List<McmTkTenpuEntity> tenpuList,String loginUser,Set<BigDecimal> originallyEditable,boolean requireChange) {
        if (!canUpdate(loginUser)) throw new McmBusinessException("権限がないため実行できません。");
        // Serialize writes to this contract, then compare the version originally displayed.
        McmTkKeiyakuEntity current = repository.lockKeiyaku(keiyaku.getTkKeiyakuId());
        if (current == null) throw new InputCheckException("対象の契約が見つかりません。検索画面から開き直してください。");
        checkVersion(keiyaku.getLastupdateDt(), current.getLastupdateDt());
        if (!canSave(current)) throw new McmBusinessException("参照画面からは登録できません。");
        boolean headerEditable = canEditWithUnlock(current);
        List<McmTkKikanEntity> periods = repository.lockKikanList(current.getTkKeiyakuId());
        // 変更有無の判定用に、画面入力を反映する前の値（DBの現在値）を保持する。
        boolean changed = false;
        McmTkKeiyakuEntity currentBefore = copyBean(current, McmTkKeiyakuEntity::new);
        List<McmTkKikanEntity> periodsBefore = new ArrayList<>();
        for (var p : periods) periodsBefore.add(copyBean(p, McmTkKikanEntity::new));
        if (kikanList.isEmpty()) throw conflict();
        requireSameIds(periods.stream().map(McmTkKikanEntity::getTkKikanId).toList(), kikanList.stream().map(McmTkKikanEntity::getTkKikanId).toList());
        for (var period : periods) {
            if (period.getKaisiDt() == null || period.getSyuryoDt() == null || !period.getSyuryoDt().isAfter(period.getKaisiDt()))
                throw new InputCheckException("期間の終了日は開始日より後の日付を指定してください。");
        }
        // Decide from stored dates, before accepting any posted period edits.
        var editableCostIds = new java.util.TreeSet<BigDecimal>();
        periods.stream().filter(p -> canEditPeriodCosts(current, p)).forEach(p -> editableCostIds.add(p.getTkKikanId()));
        var editablePeriodIds = new java.util.TreeSet<BigDecimal>();
        editablePeriodIds.addAll(originallyEditable);
        periods.stream().filter(p -> canEditPeriodDetails(current, p, periods)).forEach(p -> editablePeriodIds.add(p.getTkKikanId()));
        for (McmTkKikanEntity input : kikanList) {
            McmTkKikanEntity period = periods.stream().filter(p -> sameId(p.getTkKikanId(), input.getTkKikanId()))
                    .findFirst().orElseThrow(Mcm1005uService::conflict);
            checkVersion(input.getLastupdateDt(), period.getLastupdateDt());
            if (!editablePeriodIds.contains(period.getTkKikanId())) continue;
            // 期間は機器選定で確定済み。通常登録で金額・履歴と独立して書き換えない。
            validateContractHours(input.getKeiyakujikantai());
            period.setBiko(input.getBiko()); period.setKeiyakujikantai(input.getKeiyakujikantai());
            // 親契約の契約NOとは別の、期間ごとの契約NO。未送信時は既存値を保持する。
            if (input.getKeiyakuNo() != null) {
                if (input.getKeiyakuNo().length() > 50)
                    throw new InputCheckException("契約NOは50桁以下で入力してください。");
                period.setKeiyakuNo(input.getKeiyakuNo());
            }
            if (input.getHosyuhoho() != null) {
                if (!java.util.Set.of("", "F", "S", "C", "H", "I", "T").contains(input.getHosyuhoho()))
                    throw new InputCheckException("保守方法を選択してください。");
                period.setHosyuhoho(input.getHosyuhoho());
            }
            if (isUnlocked(current.getTkKeiyakuId())) {
                if (input.getHyojungokeiKin() != null) period.setHyojungokeiKin(validAmount(input.getHyojungokeiKin()));
                if (input.getSikirisyokeiKin() != null) period.setSikirisyokeiKin(validAmount(input.getSikirisyokeiKin()));
                if (input.getSyusseinebikiKin() != null) period.setSyusseinebikiKin(validAmount(input.getSyusseinebikiKin()));
                if (input.getSikirigokeiKin() != null) period.setSikirigokeiKin(validAmount(input.getSikirigokeiKin()));
            }
        }
        // Only fields editable in this screen may change. Preserve state, totals and parent IDs.
        if (headerEditable) {
        current.setKeiyakuNo(keiyaku.getKeiyakuNo()); current.setShokaiKeiyakuDt(keiyaku.getShokaiKeiyakuDt());
        current.setKeiyakuDt(keiyaku.getKeiyakuDt()); current.setJidokosinFlg(keiyaku.getJidokosinFlg());
        current.setKaiyakuDt(keiyaku.getKaiyakuDt()); current.setKeiyakumanryoDt(keiyaku.getKeiyakumanryoDt());
        current.setEntyokeiyakumanryoDt(keiyaku.getEntyokeiyakumanryoDt()); current.setPackFlg(keiyaku.getPackFlg());
        current.setPackkeiyakunaiyo(keiyaku.getPackkeiyakunaiyo()); current.setBiko(keiyaku.getBiko());
        current.setKosinnaiyo(keiyaku.getKosinnaiyo());
        // VB 2010/05/10 correction deliberately uses the last tab, not MAX(end date).
        current.setJikaikosinDt(periods.get(periods.size() - 1).getSyuryoDt().plusDays(1));
        }
        current.setBiko(keiyaku.getBiko()); current.setKosinnaiyo(keiyaku.getKosinnaiyo());

        List<McmTkTenkenEntity> inspectionUpdates = new ArrayList<>();
        int inspectionCount = 0;
        for (McmTkKikanEntity period : periods) {
            for (McmTkTenkenEntity stored : repository.lockTenkenList(period.getTkKikanId())) {
                McmTkTenkenEntity input = tenkenList.stream().filter(t -> sameId(t.getTkTenkenId(), stored.getTkTenkenId()))
                        .findFirst().orElse(null);
                if (input == null) {
                    // 【調査用】DBにあるが画面から送られていない点検行を特定する
                    org.slf4j.LoggerFactory.getLogger(Mcm1005uService.class).warn(
                            "[排他調査] 未送信の点検行: kikanId={} tenkenId={} kikikoseiId={} oyaCd={} / 送信tenkenIds={}",
                            period.getTkKikanId(), stored.getTkTenkenId(), stored.getKikikoseiId(), stored.getOyakikibunruiCd(),
                            tenkenList.stream().map(t -> String.valueOf(t.getTkTenkenId())).toList());
                    throw conflict();
                }
                checkVersion(input.getLastupdateDt(), stored.getLastupdateDt());
                inspectionCount++;
                if (!editableCostIds.contains(period.getTkKikanId())) continue;
                validateCount(input.getTenkenkaisu(), java.util.Arrays.stream(input.getTsuki()).filter("1"::equals).count(), "点検");
                String weekday = input.getTenkenkanoyobi();
                if (weekday != null && !java.util.Set.of("", "1", "2", "3", "4").contains(weekday) && !java.util.Objects.equals(weekday, stored.getTenkenkanoyobi()))
                    throw new InputCheckException("点検の可能曜日を一覧から選択してください。");
                McmTkTenkenEntity storedBefore = copyBean(stored, McmTkTenkenEntity::new);
                stored.setTenkenkaisu(input.getTenkenkaisu()); stored.setTenkenkanoyobi(weekday);
                stored.setYakantaioumu(input.getYakantaioumu()); stored.setBiko(input.getBiko()); stored.setTsuki(input.getTsuki());
                // Validate only changed selections. Existing retired choices remain displayable.
                for (BigDecimal company : java.util.Arrays.asList(input.getHoshugaishaId(),input.getYakanhoshugaishaId())) {
                    if (company != null && company.compareTo(BigDecimal.valueOf(-1)) != 0 && !sameId(company,stored.getHoshugaishaId()) && !sameId(company,stored.getYakanhoshugaishaId())
                            && !maintenanceCompanies(period.getTkKikanId(),null).containsKey(uiId(company)))
                        throw new InputCheckException("保守会社を一覧から選択してください。");
                }
                // null = field absent (older form); -1 = explicitly cleared dropdown.
                if (input.getHoshugaishaId() != null) stored.setHoshugaishaId(input.getHoshugaishaId().signum() < 0 ? null : input.getHoshugaishaId());
                if (input.getYakanhoshugaishaId() != null) stored.setYakanhoshugaishaId(input.getYakanhoshugaishaId().signum() < 0 ? null : input.getYakanhoshugaishaId());
                // 月チェックは点検明細に保持されるため、tsuki は別途DBの現在値と比較する。
                if (!sameValues(storedBefore, stored, Set.of("tsuki"))
                        || !sameMonths(repository.findTenkenMeisaiTsukiMonths(stored.getTkTenkenId()), input.getTsuki())) changed = true;
                inspectionUpdates.add(stored);
            }
        }
        if (inspectionCount != tenkenList.size()) throw conflict();
        for (int i = 0; i < siharaiList.size(); i++) {
            McmTkSiharaiEntity input = siharaiList.get(i);
            if (periods.stream().noneMatch(p -> sameId(p.getTkKikanId(), input.getTkKikanId()))) throw conflict();
            McmTkSiharaiEntity stored = repository.lockSiharai(input.getTkKikanId());
            if (stored == null || !sameId(stored.getTkSiharaiId(), input.getTkSiharaiId())) throw conflict();
            checkVersion(input.getLastupdateDt(), stored.getLastupdateDt());
            if (!editableCostIds.contains(input.getTkKikanId())) continue;
            validateCount(input.getKaisu(), java.util.Arrays.stream(siharaiTsukiList.get(i)).filter(Boolean.TRUE::equals).count(), "支払");
            if (!sameValue(stored.getKaisu(), input.getKaisu()) || !sameValue(stored.getBiko(), input.getBiko())
                    || !sameMonths(repository.findSiharaiMeisaiTsukiMonths(stored.getTkSiharaiId()), siharaiTsukiList.get(i))) changed = true;
        }
        // 契約ヘッダーと期間の変更有無（次回更新日は期間から導出される値のため比較対象外）。
        if (!sameValues(currentBefore, current, Set.of("jikaikosinDt"))) changed = true;
        for (int i = 0; i < periods.size(); i++) if (!sameValues(periodsBefore.get(i), periods.get(i), Set.of())) changed = true;
        // 【不具合修正 #513】実質的な変更がない場合は更新SQLを実行せず中止する（VB: CPDataNotChangedException）。
        if (requireChange && !changed) throw new DataNotChangedException(FWM_0010E_NOT_CHANGED);
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
            if (!editableCostIds.contains(payment.getTkKikanId())) continue;
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
        saveAndSinseiInternal(contract,periods,inspections,payments,months,attachments,selectedPeriod,user,Set.of());
    }

    private void saveAndSinseiInternal(McmTkKeiyakuEntity contract,List<McmTkKikanEntity> periods,
            List<McmTkTenkenEntity> inspections,List<McmTkSiharaiEntity> payments,List<Boolean[]> months,
            List<McmTkTenpuEntity> attachments,BigDecimal selectedPeriod,String user,Set<BigDecimal> originallyEditable) {
        // Check approval preconditions before the first write as well as inside sinsei.
        var current = repository.lockKeiyaku(contract.getTkKeiyakuId());
        if (!canUpdate(user) || !canEditDetails(current)) throw new McmBusinessException("現在の状態では申請できません。");
        var files = repository.findTenpuByKikanId(selectedPeriod);
        if (contract.getKaiyakuDt() == null && files.stream().anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("再申請はできません。期間タブを作成し直して下さい。");
        if (files.stream().noneMatch(a -> McmConstants.SHONINJOTAI_SAKUSEICHU_CD.equals(a.getShoninjotai()) || McmConstants.SHONINJOTAI_SASHIMODOSHI_CD.equals(a.getShoninjotai())))
            throw new McmBusinessException("契約依頼資料を添付してください。");
        saveInternal(contract, periods, inspections, payments, months, attachments, user, originallyEditable, false);
        sinsei(contract.getTkKeiyakuId(), selectedPeriod, user);
    }

    // CPValidate.xml: KEIYAKUJIKANTAI、任意入力、Decimal、min=0、max=24。
    private void validateContractHours(String value) {
        String error = validation.validateField("保守契約時間帯", value, false, com.daifuku.mcm.common.AppConstants.VALIDATE_TYPE_NUMERIC, 2, 0);
        if (error == null) error = validation.validateNumericRange("保守契約時間帯", value, BigDecimal.ZERO, new BigDecimal("24"));
        if (error != null) throw new InputCheckException(error);
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
        return contract != null && entryMode(contract.getTkKeiyakuId()) != 5 && entryMode(contract.getTkKeiyakuId()) != 6 && !McmConstants.JOTAI_HAKI.equals(contract.getJotai())
                && !McmConstants.JOTAI_KAIYAKU.equals(contract.getJotai())
                && !McmConstants.SHONINJOTAI_SHINSACHU_CD.equals(contract.getShoninjotai())
                && !McmConstants.SHONINJOTAI_SHONINCHU_CD.equals(contract.getShoninjotai());
    }

    public boolean canEditPeriodCosts(McmTkKeiyakuEntity contract, McmTkKikanEntity period) {
        return contract != null && canSave(contract) && (isUnlocked(contract.getTkKeiyakuId())
                || (period.getSyuryoDt() != null && !period.getSyuryoDt().toLocalDate().isBefore(java.time.LocalDate.now())));
    }

    /** VB View: only the latest current/unapproved future tab permits period and attachment edits. */
    public boolean canEditPeriodDetails(McmTkKeiyakuEntity contract, McmTkKikanEntity period,
            List<McmTkKikanEntity> periods) {
        if (!canEditWithUnlock(contract) || period == null || periods.isEmpty()) return false;
        if (isUnlocked(contract.getTkKeiyakuId())) return true;
        if (period.getTkKikanId().signum() < 0) return true; // Unsaved selection tab.
        if (!sameId(period.getTkKikanId(), periods.get(periods.size() - 1).getTkKikanId())) return false;
        var today = java.time.LocalDate.now();
        if (period.getKaisiDt() == null || period.getSyuryoDt() == null || period.getSyuryoDt().toLocalDate().isBefore(today)) return false;
        return !period.getKaisiDt().toLocalDate().isAfter(today)
                || repository.findTenpuByKikanId(period.getTkKikanId()).stream()
                    .noneMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai()));
    }

    public boolean canChangePeriod(McmTkKeiyakuEntity contract, McmTkKikanEntity period, List<McmTkKikanEntity> periods) {
        if (contract == null || period == null || periods.isEmpty()) return false;
        if (isUnlocked(contract.getTkKeiyakuId())) return true;
        if (entryMode(contract.getTkKeiyakuId()) == 6 || !sameId(period.getTkKikanId(),periods.get(periods.size()-1).getTkKikanId())) return false;
        if (canEditDetails(contract)) return true;
        var today=java.time.LocalDate.now();
        return period.getSyuryoDt()!=null && period.getSyuryoDt().toLocalDate().isBefore(today)
            || period.getKaisiDt()!=null && period.getKaisiDt().toLocalDate().isAfter(today)
               && repository.findTenpuByKikanId(period.getTkKikanId()).stream().anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai()));
    }

    public void checkAttachmentEditable(McmTkKeiyakuEntity contract, BigDecimal periodId) {
        var periods = repository.findKikanByKeiyakuId(contract.getTkKeiyakuId());
        var period = periods.stream().filter(p -> sameId(p.getTkKikanId(), periodId)).findFirst().orElseThrow(Mcm1005uService::conflict);
        if (!canEditPeriodDetails(contract, period, periods)) throw new McmBusinessException("現在の状態では添付ファイルを変更できません。");
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
        if (!canEditWithUnlock(contract)) throw new McmBusinessException("現在の状態では削除できません。");
        McmTkTenpuEntity attachment = repository.findKikanByKeiyakuId(contractId).stream()
                .flatMap(p -> repository.findTenpuByKikanId(p.getTkKikanId()).stream())
                .filter(a -> sameId(a.getTkTenpuId(), tkTenpuId)).findFirst().orElseThrow(Mcm1005uService::conflict);
        checkAttachmentEditable(contract, attachment.getTkKikanId());
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
        form.setSelectionVersion(repository.selectionVersion(keiyaku.getTkKeiyakuId(), false));
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
        if (!form.getSelectionVersion().equals(repository.selectionVersion(keiyaku.getTkKeiyakuId(), false))) throw conflict();
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
     * 元VB: MCM1005Uタブ内 TabHenkoButtonClick → SearchInsert/setKikiKoseiDataTable等
     *
     * VB版は CHECK_FLG_OLD と CHECK_FLG の比較により、行単位で
     *   OFF→ON: 新規追加, ON→OFF: 該当行削除, ON→ON: 値のみ更新（既存IDは維持）
     * を行い、全削除・再登録は行わない。既存機器（TK_KIKIKOSEI_ID等）や
     * その関連データ（点検情報等）を安全に維持するため、Java版もこの差分更新方式を踏襲する。
     *
     * 対象期間に既存機器が1件も無い場合（新規期間からの起動）は、
     * 全行が CHECK_FLG_OLD=OFF として扱われるため、結果的に全行が新規追加となる。
     */
    @Transactional(rollbackFor = Exception.class)
    public BigDecimal applyMcm1006uResult(Mcm1006uForm result, String loginUser) {
        if (!canUpdate(loginUser)) throw new McmBusinessException("権限がないため実行できません。");
        if (result == null || result.getTkKeiyakuId() == null || result.getTkKikanId() == null) throw conflict();
        BigDecimal tkKikanId = result.getTkKikanId();
        // 選定画面を開いている間の承認・契約状態変更を保存直前にも確認する。
        var current = repository.lockKeiyaku(result.getTkKeiyakuId());
        if (current == null) throw conflict();
        var periods = repository.lockKikanList(result.getTkKeiyakuId());
        var originalPeriodId = tkKikanId;
        var period = periods.stream().filter(p -> sameId(p.getTkKikanId(), originalPeriodId)).findFirst().orElseThrow(Mcm1005uService::conflict);
        if (!canChangePeriod(current, period, periods)) throw new McmBusinessException("現在の期間では変更できません。");
        if (result.getSelectionVersion() == null
                || !result.getSelectionVersion().equals(repository.selectionVersion(current.getTkKeiyakuId(), true))) throw conflict();
        // 失敗時にセッションの選定結果を壊さないよう作業用コピーを使う。
        result = copySelection(result);
        period = prepareSelectionPeriod(result, period, periods, loginUser);
        tkKikanId = period.getTkKikanId();

        persistSelectionRows(result, tkKikanId, loginUser);
        repository.refreshNextRenewal(current.getTkKeiyakuId(), loginUser);
        return tkKikanId;
    }

    private void persistSelectionRows(Mcm1006uForm result, BigDecimal tkKikanId, String loginUser) {

        // kikikoseiId → 新規追加したTK_KIKIKOSEI_ID マッピング（既存の場合は既存IDが入る）
        // JDBC decimals may have a different scale from the delivery (5 vs 5.0000).
        Map<BigDecimal, BigDecimal> koseiIdMap = new java.util.TreeMap<>();
        // kikimeisaiId → 新規追加/既存のTK_KIKIMEISAI_ID マッピング
        Map<BigDecimal, BigDecimal> meisaiIdMap = new java.util.TreeMap<>();

        diffApplyKosei(result, tkKikanId, loginUser, koseiIdMap);
        diffApplyMeisai(result, tkKikanId, loginUser, koseiIdMap, meisaiIdMap);
        synchronizeInspections(result, tkKikanId, loginUser, koseiIdMap);
        diffApplyKotai(result, tkKikanId, loginUser, koseiIdMap, meisaiIdMap);
        diffApplyTanka(result, tkKikanId, loginUser, koseiIdMap, meisaiIdMap);
        repository.refreshPeriodTotals(tkKikanId, BigDecimal.valueOf(result.getSyusseinebikiGokei()), loginUser);
    }

    /** VB CalledFromForwardScreen: current/past/approved periods retain their equipment history. */
    private McmTkKikanEntity prepareSelectionPeriod(Mcm1006uForm result, McmTkKikanEntity period,
            List<McmTkKikanEntity> periods, String user) {
        String error = validation.validateField("期間開始日", result.getKaisiDt(), true, com.daifuku.mcm.common.AppConstants.VALIDATE_TYPE_DATE, 10, 0);
        if (error == null) error = validation.validateField("期間終了日", result.getSyuryoDt(), true, com.daifuku.mcm.common.AppConstants.VALIDATE_TYPE_DATE, 10, 0);
        if (error != null) throw new InputCheckException(error);
        var fmt = java.time.format.DateTimeFormatter.ofPattern("uuuu/MM/dd").withResolverStyle(java.time.format.ResolverStyle.STRICT);
        java.time.LocalDate start, end;
        try {
            start = java.time.LocalDate.parse(result.getKaisiDt(), fmt);
            end = java.time.LocalDate.parse(result.getSyuryoDt(), fmt);
        } catch (RuntimeException ex) { throw new InputCheckException("契約期間を正しく入力してください。"); }
        if (end.isBefore(start)) throw new InputCheckException("終了日が開始日より前になっています。");
        if (period.getKaisiDt() == null || period.getSyuryoDt() == null) throw conflict();
        boolean unlocked = isUnlocked(result.getTkKeiyakuId());
        var latest = periods.stream().filter(p -> p.getYukoFlg() == null || p.getYukoFlg().signum() == 0)
                .max(java.util.Comparator.comparing(McmTkKikanEntity::getKaisiDt)).orElse(null);
        if (!unlocked && (latest == null || !sameId(latest.getTkKikanId(), period.getTkKikanId())))
            throw new InputCheckException("最新の有効な期間から機器を変更してください。");
        validateChangeStart(period, periods, start);
        boolean approved = repository.findTenpuByKikanId(period.getTkKikanId()).stream()
                .anyMatch(a -> McmConstants.SHONINJOTAI_SHONINZUMI_CD.equals(a.getShoninjotai()));
        boolean newPeriod = !period.getKaisiDt().toLocalDate().isAfter(java.time.LocalDate.now()) || approved;
        if (newPeriod) {
            var next = new McmTkKikanEntity();
            org.springframework.beans.BeanUtils.copyProperties(period, next);
            next.setTkKikanId(repository.getMaxKikanId().add(BigDecimal.ONE));
            next.setYukoFlg(BigDecimal.ZERO);
            next.setKaisiDt(start.atStartOfDay()); next.setSyuryoDt(end.atStartOfDay());
            repository.insertKikan(next, user);
            // 元期間と同日またはそれ以前への変更では元期間を無効とする。
            if (!start.isAfter(period.getKaisiDt().toLocalDate())) period.setYukoFlg(BigDecimal.ONE);
            else if (!start.isAfter(period.getSyuryoDt().toLocalDate())) period.setSyuryoDt(start.minusDays(1).atStartOfDay());
            repository.updateKikan(period, user);
            var payment = new McmTkSiharaiEntity();
            payment.setTkSiharaiId(repository.getMaxSiharaiId().add(BigDecimal.ONE));
            payment.setTkKikanId(next.getTkKikanId()); payment.setKaisu(BigDecimal.ZERO);
            repository.insertSiharai(payment, user);
            // 新しい期間には選定された内容のみを登録。旧期間の子レコードは変更しない。
            result.getKoseiRows().forEach(r -> r.setCheckFlgOld(0));
            result.getMeisaiRows().forEach(r -> r.setCheckFlgOld(0));
            result.getKotaiRows().forEach(r -> r.setCheckFlgOld(0));
            result.getTankaRows().forEach(r -> r.setCheckFlgOld(0));
            result.setTankaDelRows(new ArrayList<>());
            result.setTkKikanId(next.getTkKikanId());
            return next;
        }
        period.setKaisiDt(start.atStartOfDay()); period.setSyuryoDt(end.atStartOfDay());
        repository.updateKikan(period, user);
        return period;
    }

    /** 【変換元】Mcm1005uTabControl.vb setKikiKoseiDataTable() */
    private void diffApplyKosei(Mcm1006uForm result, BigDecimal tkKikanId, String loginUser,
            Map<BigDecimal, BigDecimal> koseiIdMap) {
        BigDecimal koseiMaxId = repository.getMaxKikikoseiId();
        for (Mcm1006uForm.KoseiRowForm r : result.getKoseiRows()) {
            boolean wasOn = r.getCheckFlgOld() == 1;
            boolean isOn  = r.getCheckFlg() == 1;
            if (!wasOn && isOn) {
                // 新規追加
                koseiMaxId = koseiMaxId.add(BigDecimal.ONE);
                koseiIdMap.put(r.getKikikoseiId(), koseiMaxId);
                repository.insertKikikosei(koseiMaxId, tkKikanId,
                    r.getKikikoseiId(), r.getKikikoseiNk(), r.getSetNm(),
                    r.getTehaiseiban(), r.getHyojijun(), loginUser);
            } else if (wasOn && !isOn) {
                // 既存行削除（配下の機器明細・個体明細・単価・点検も削除）
                for (Map<String, Object> row : repository.findExistingKoseiByKikikoseiId(tkKikanId, r.getKikikoseiId())) {
                    repository.deleteKikikoseiById(toBD(row.get("TK_KIKIKOSEI_ID")));
                }
            } else if (wasOn && isOn) {
                // 既存行を維持し、値のみ更新
                for (Map<String, Object> row : repository.findExistingKoseiByKikikoseiId(tkKikanId, r.getKikikoseiId())) {
                    BigDecimal tkKikikoseiId = toBD(row.get("TK_KIKIKOSEI_ID"));
                    koseiIdMap.put(r.getKikikoseiId(), tkKikikoseiId);
                    repository.updateKikikosei(tkKikikoseiId, r.getKikikoseiNk(), r.getSetNm(),
                        r.getTehaiseiban(), loginUser);
                }
            }
        }
    }

    /** 【変換元】Mcm1005uTabControl.vb setKikiMeisaiDataTable() */
    private void diffApplyMeisai(Mcm1006uForm result, BigDecimal tkKikanId, String loginUser,
            Map<BigDecimal, BigDecimal> koseiIdMap, Map<BigDecimal, BigDecimal> meisaiIdMap) {
        BigDecimal meisaiMaxId = repository.getMaxKikimeisaiId();
        for (Mcm1006uForm.MeisaiRowForm r : result.getMeisaiRows()) {
            boolean wasOn = r.getCheckFlgOld() == 1;
            boolean isOn  = r.getCheckFlg() == 1;
            if (!wasOn && isOn) {
                BigDecimal newKoseiId = koseiIdMap.get(r.getKikikoseiId());
                if (newKoseiId == null) throw conflict();
                meisaiMaxId = meisaiMaxId.add(BigDecimal.ONE);
                meisaiIdMap.put(r.getKikimeisaiId(), meisaiMaxId);
                repository.insertKikimeisai(meisaiMaxId, newKoseiId,
                    r.getKikikoseiId(), r.getKikimeisaiId(),
                    r.getSeizomakerId(), r.getSeizomakerNk(),
                    r.getKikihinmeiNk(), r.getKikikatashiki(),
                    r.getSuryoNm(), r.getAtsukaikikiId(), r.getHyojijun(), loginUser);
            } else if (wasOn && !isOn) {
                for (Map<String, Object> row : repository.findExistingMeisaiByKikimeisaiId(tkKikanId, r.getKikimeisaiId())) {
                    repository.deleteKikimeisaiById(toBD(row.get("TK_KIKIMEISAI_ID")));
                }
            } else if (wasOn && isOn) {
                for (Map<String, Object> row : repository.findExistingMeisaiByKikimeisaiId(tkKikanId, r.getKikimeisaiId())) {
                    BigDecimal tkKikimeisaiId = toBD(row.get("TK_KIKIMEISAI_ID"));
                    meisaiIdMap.put(r.getKikimeisaiId(), tkKikimeisaiId);
                    repository.updateKikimeisai(tkKikimeisaiId, r.getKikihinmeiNk(), r.getKikikatashiki(),
                        r.getSuryoNm(), loginUser);
                }
            }
        }
    }

    private record InspectionGroup(BigDecimal component, String classification) { }

    private static InspectionGroup inspectionGroup(BigDecimal component, String classification) {
        return new InspectionGroup(component == null ? null : component.stripTrailingZeros(), classification);
    }

    /** 同じ構成・親分類で1点検。既存の点検ID、入力値、実施月を保持する。 */
    private void synchronizeInspections(Mcm1006uForm result, BigDecimal periodId, String user,
            Map<BigDecimal, BigDecimal> componentIds) {
        if (result.getMeisaiRows().isEmpty()) return;
        var selected = new HashMap<InspectionGroup, Mcm1006uForm.MeisaiRowForm>();
        var changed = new java.util.HashSet<InspectionGroup>();
        for (var row : result.getMeisaiRows()) {
            var key = inspectionGroup(row.getKikikoseiId(), row.getOyakikibunruiCd());
            if (row.getCheckFlg() == 1) selected.putIfAbsent(key, row);
            if (row.getCheckFlg() != row.getCheckFlgOld()) changed.add(key);
        }
        var stored = repository.lockTenkenList(periodId);
        var existing = new java.util.HashSet<InspectionGroup>();
        for (var row : stored) {
            var key = inspectionGroup(row.getKikikoseiId(), row.getOyakikibunruiCd());
            existing.add(key);
            if (changed.contains(key) && !selected.containsKey(key)) repository.deleteInspection(row.getTkTenkenId());
        }
        BigDecimal nextId = null;
        for (var key : changed) {
            if (existing.contains(key) || !selected.containsKey(key)) continue;
            var source = selected.get(key);
            var parentId = componentIds.get(source.getKikikoseiId());
            if (parentId == null || key.classification() == null) throw conflict();
            if (nextId == null) nextId = repository.getMaxTenkenId();
            nextId = nextId.add(BigDecimal.ONE);
            var row = newInspection(result, source);
            row.setTkTenkenId(nextId); row.setTkKikokoseiId(parentId); row.setTkKikanId(periodId);
            repository.insertTenken(row, user);
        }
    }

    /** 【変換元】Mcm1005uTabControl.vb setKotaimeisaiDataTable() */
    private void diffApplyKotai(Mcm1006uForm result, BigDecimal tkKikanId, String loginUser,
            Map<BigDecimal, BigDecimal> koseiIdMap, Map<BigDecimal, BigDecimal> meisaiIdMap) {
        BigDecimal kotaiMaxId = repository.getMaxKotaimeisaiId();
        for (Mcm1006uForm.KotaiRowForm r : result.getKotaiRows()) {
            boolean wasOn = r.getCheckFlgOld() == 1;
            boolean isOn  = r.getCheckFlg() == 1;
            if (!wasOn && isOn) {
                BigDecimal newKoseiId  = koseiIdMap.get(r.getKikikoseiId());
                BigDecimal newMeisaiId = meisaiIdMap.get(r.getKikimeisaiId());
                if (newKoseiId == null || newMeisaiId == null) throw conflict();
                kotaiMaxId = kotaiMaxId.add(BigDecimal.ONE);
                repository.insertKotaimeisai(kotaiMaxId, newKoseiId, newMeisaiId,
                    r.getKikikoseiId(), r.getKikimeisaiId(),
                    r.getKotaikanriId(), r.getKotaiNk(), r.getSerialNo(),
                    r.getItijinonyugDt(), r.getSetchibasyo(),
                    r.getTekkyoDt(), r.getEnchokeiyakukigenDt(), loginUser);
            } else if (wasOn && !isOn) {
                for (Map<String, Object> row : repository.findExistingKotaiByKotaikanriId(
                        tkKikanId, r.getKikimeisaiId(), r.getKotaikanriId())) {
                    repository.deleteKotaimeisaiById(toBD(row.get("TK_KOTAIMEISAI_ID")));
                }
            } else if (wasOn && isOn) {
                for (Map<String, Object> row : repository.findExistingKotaiByKotaikanriId(
                        tkKikanId, r.getKikimeisaiId(), r.getKotaikanriId())) {
                    repository.updateKotaimeisai(toBD(row.get("TK_KOTAIMEISAI_ID")),
                        r.getKotaiNk(), r.getSerialNo(), r.getItijinonyugDt(), r.getSetchibasyo(),
                        r.getTekkyoDt(), r.getEnchokeiyakukigenDt(), loginUser);
                }
            }
        }
    }

    /** 【変換元】Mcm1005uTabControl.vb setTankaDataTable()（tankaDelRowList＝期間不一致行削除を含む） */
    private void diffApplyTanka(Mcm1006uForm result, BigDecimal tkKikanId, String loginUser,
            Map<BigDecimal, BigDecimal> koseiIdMap, Map<BigDecimal, BigDecimal> meisaiIdMap) {
        BigDecimal tankaMaxId = repository.getMaxTankaId();
        for (Mcm1006uForm.TankaRowForm r : result.getTankaRows()) {
            boolean wasOn = r.getCheckFlgOld() == 1;
            boolean isOn  = r.getCheckFlg() == 1;
            if (!wasOn && isOn) {
                BigDecimal newKoseiId  = koseiIdMap.get(r.getKikikoseiId());
                BigDecimal newMeisaiId = meisaiIdMap.get(r.getKikimeisaiId());
                if (newKoseiId == null || newMeisaiId == null) throw conflict();
                tankaMaxId = tankaMaxId.add(BigDecimal.ONE);
                repository.insertTanka(tankaMaxId, newKoseiId, newMeisaiId,
                    r.getKikikoseiId(), r.getKikimeisaiId(),
                    r.getTmKikimeisaiId(), r.getTmTankaId(),
                    r.getPackFlg(), r.getKeiyakuNo(), r.getKeiyakunaiyo(),
                    r.getTorihosyujikanId(), r.getTenkenumu(),
                    r.getHosyuhoho(), r.getServicekeitai(),
                    r.getHyojunKin(), r.getSikiriKin(), r.getSuryoNm(),
                    r.getKaisiDt(), r.getSyuryoDt(), loginUser);
            } else if (wasOn && !isOn) {
                repository.deleteTankaByTmTankaId(tkKikanId, r.getTmTankaId());
            } else if (wasOn && isOn) {
                for (Map<String, Object> row : repository.findExistingTankaByTmTankaId(tkKikanId, r.getTmTankaId())) {
                    repository.updateTanka(toBD(row.get("TK_TANKA_ID")), r.getKaisiDt(), r.getSyuryoDt(),
                        r.getSuryoNm(), r.getHyojunKin(), r.getSikiriKin(), loginUser);
                }
            }
        }

        // 契約期間の見直しにより対象外となった単価行を削除する
        // 【変換元】Mcm1006u2Screen_Load() tankaDelRowList → setTankaDataTable() 末尾ループ
        if (result.getTankaDelRows() != null) {
            for (Mcm1006uForm.TankaRowForm r : result.getTankaDelRows()) {
                repository.deleteTankaByTmTankaId(tkKikanId, r.getTmTankaId());
            }
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
