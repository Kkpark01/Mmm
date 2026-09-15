/**
 * 【変換元】Mcm0012uScreen.vb（611行）
 *   MCM0012U プラントマスタ画面のビジネスロジック
 *
 *   ★ PlantEntity.plantId / nonyusakiId が BigDecimal 型のため、
 *     Form(Long) ↔ Entity(BigDecimal) の変換を行っています。
 */
package com.daifuku.mcm.service;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.HashMap;      // ★追加（#300）
import java.util.HashSet;
import java.util.List;
import java.util.Map;          // ★追加（#300）
import java.util.Optional;
import java.util.Set;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.daifuku.mcm.common.ExclusiveLockKey;      // ★追加（#296）排他制御
import com.daifuku.mcm.common.Mcm0012uConstants;
import com.daifuku.mcm.common.MessageService;        // ★追加（#296）排他制御
import com.daifuku.mcm.entity.BrandKoseiEntity;
import com.daifuku.mcm.entity.NonyusakiEntity;
import com.daifuku.mcm.entity.PlantEntity;
import com.daifuku.mcm.entity.TantoEntity;              // ★追加（#300）担当者マスタ
import com.daifuku.mcm.exception.ExclusiveControlException; // ★追加（#296）排他制御
import com.daifuku.mcm.form.Mcm0012uForm;
import com.daifuku.mcm.form.Mcm0012uForm.BrandKoseiRow;
import com.daifuku.mcm.repository.BrandKoseiRepository;
import com.daifuku.mcm.repository.KikikotaikanriRepository;
import com.daifuku.mcm.repository.NonyusakiRepository;
import com.daifuku.mcm.repository.PlantRepository;
import com.daifuku.mcm.repository.TantoRepository;       // ★追加（#300）担当者マスタ

@Service
public class Mcm0012uService {

    private static final Logger log = LoggerFactory.getLogger(Mcm0012uService.class);
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    @Autowired
    private PlantRepository plantRepository;

    @Autowired
    private BrandKoseiRepository brandKoseiRepository;
    
	@Autowired
	private NonyusakiRepository nonyusakiRepository;

    @Autowired
    private KikikotaikanriRepository kikiKotaikanriRepository;

    // ★追加（#300）：担当者マスタ（MCM_MO_TANTO）参照用
    @Autowired
    private TantoRepository tantoRepository;

    /** ★追加（#296）：排他エラーメッセージ（mcm0012u.msg.exclusive）の外部化に使用 */
    @Autowired
    private MessageService messageService;

	@Autowired
	private JdbcTemplate jdbcTemplate;

    // ========================================
    // ユーティリティ（private）
    // ========================================

    private static boolean isEmpty(String s) {
        return s == null || s.trim().isEmpty();
    }

    private String fmtDate(LocalDateTime dt) {
        return dt == null ? "" : dt.format(DATE_FMT);
    }
    

	/**
	 * 日時フォーマット（更新履歴 / グリッド監査列用）
	 * 【変換元】VB Format="G" → "yyyy/MM/dd HH:mm:ss"
	 */
	private static final DateTimeFormatter DATETIME_FMT = 
	    DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
	
	private String fmtDateTime(LocalDateTime dt) {
	    return dt == null ? "" : dt.format(DATETIME_FMT);
	}

    // ========================================
    // ★追加（#300）：担当者マスタ変換ユーティリティ
    //   更新履歴枠・ブランド情報グリッドの登録者／更新者表示を、
    //   担当者マスタ（MCM_MO_TANTO）の「担当者名」に統一するための変換処理。
    // ========================================

    /**
     * ログインIDの集合から、担当者マスタを1クエリでまとめて取得し、
     * 「ログインID → 担当者名」のマップを作成する（N+1回避）。
     */
    private Map<String, String> buildTantoNkMap(Set<String> loginIds) {
        loginIds.remove(null);
        loginIds.remove("");
        Map<String, String> map = new HashMap<>();
        if (loginIds.isEmpty()) {
            return map;
        }
        for (TantoEntity t : tantoRepository.findByLoginIdIn(loginIds)) {
            map.put(t.getLoginId(), t.getTantoNk());
        }
        return map;
    }

    // ★追加（不具合修正）：システムが自動セットする擬似ログインID（Controller既定値
    //   "system" / AuditorAware既定値 "SYSTEM"）は担当者マスタに実在しないため、
    //   ログインIDそのままではなく固定の表示文言に変換する。
    private static final String SYSTEM_LOGIN_ID = "system";
    private static final String SYSTEM_DISPLAY_NAME = "システム";

    /**
     * ログインIDを担当者名に変換する。
     * ★担当者マスタに該当データが無い場合、
     *   ログインIDが システム用擬似ID（"system"/"SYSTEM"、大文字小文字を区別しない）であれば
     *   固定文言「システム」を表示する。
     *   それ以外（退職者ID等）は追跡可能性を優先し、ログインIDをそのまま表示する。
     */
    private String resolveTantoNk(String loginId, Map<String, String> tantoNkMap) {
        if (isEmpty(loginId)) {
            return "";
        }
        String tantoNk = tantoNkMap.get(loginId);
        if (!isEmpty(tantoNk)) {
            return tantoNk;
        }
        if (SYSTEM_LOGIN_ID.equalsIgnoreCase(loginId)) {
            return SYSTEM_DISPLAY_NAME;
        }
        return loginId;
    }


    private String bdToStr(BigDecimal bd) {
        return bd == null ? "" : bd.toPlainString();
    }

    private BigDecimal toBd(String s) {
        if (isEmpty(s)) return null;
        try {
            return new BigDecimal(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /** Long → BigDecimal 変換（null安全） */
    private BigDecimal longToBd(Long v) {
        return v == null ? null : BigDecimal.valueOf(v);
    }

    /** BigDecimal → Long 変換（null安全） */
    private Long bdToLong(BigDecimal bd) {
        return bd == null ? null : bd.longValue();
    }

    private LocalDateTime parseDate(String s) {
        if (isEmpty(s)) return null;
        try {
            return LocalDateTime.parse(s + " 00:00:00",
                    DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss"));
        } catch (DateTimeParseException e) { return null; }
    }

    private boolean isValidDate(String s) {
        try { java.time.LocalDate.parse(s, DATE_FMT); return true; }
        catch (DateTimeParseException e) { return false; }
    }

    private boolean isNum(String s) {
        if (isEmpty(s)) return true;
        try { new BigDecimal(s); return true; }
        catch (NumberFormatException e) { return false; }
    }

    // ★追加（#254）：実DBスキーマ（decimal精度）に合わせた桁あふれ防止チェック
    //   出典：INFORMATION_SCHEMA.COLUMNS 確認結果（2026-09-08確認）
    private static final int MAX_KIN_DIGITS      = 18; // SYSTEMSEKKEI_KIN / KIHONSEKKEI_KIN / PROGRAMSAKUSEI_KIN
    private static final int MAX_HYOJIJUN_DIGITS = 6;  // HYOJIJUN：実DBは decimal(6,0)
    private static final int MAX_DREMOSNM_DIGITS = 4;  // DREMOS_NM：実DBは decimal(4,0)　★今回のオーバーフロー原因カラム

    // ★追加：画面側のmaxlength制限撤廃に伴う、サーバー側桁数チェック用の最大桁数
    //   出典：PlantEntity / BrandKoseiEntity の @Column(length=…) 定義に一致
    private static final int MAX_SUPPORT_ID_LENGTH      = 7;  // SUPPORT_ID
    private static final int MAX_PLANT_NK_LENGTH        = 80; // PLANT_NK
    private static final int MAX_NONYUBUSYO_NK_LENGTH   = 60; // NONYUBUSYO_NK
    private static final int MAX_NONYUTANTOSYA_NK_LENGTH = 40; // NONYUTANTOSYA_NK
    private static final int MAX_NONYUTEL_NO_LENGTH     = 50; // NONYUTEL_NO
    private static final int MAX_BRANDSYOSAI_NK_LENGTH  = 80; // BRANDSYOSAI_NK
    private static final int MAX_BIKO_LENGTH            = 4000; // BIKO（PlantEntity#biko @Column length=4000）
    private static final int MAX_REMOTERENRAKUSAKI_LENGTH = 200; // REMOTERENRAKUSAKI（BrandKoseiEntity#remoterenrakusaki @Column length=200）

    /**
     * 数値文字列の整数部桁数がDBカラムの許容桁数以内か判定する。
     * @param s 入力文字列
     * @param maxIntegerDigits 許容する整数部の最大桁数
     */
    private boolean isWithinDigitLimit(String s, int maxIntegerDigits) {
        if (isEmpty(s)) return true;
        BigDecimal bd = toBd(s);
        if (bd == null) return false;
        String digitsOnly = bd.unscaledValue().abs().toString();
        int scale = Math.max(bd.scale(), 0);
        int integerDigits = Math.max(digitsOnly.length() - scale, 0);
        return integerDigits <= maxIntegerDigits;
    }
    
    /**
     * ★修正：削除フラグの立っていない行のうち、
     *   「実体の無い行（幽霊行＝brandkoseiIdも無く、ブランド・ブランド詳細名も未入力）」を
     *   除外してカウントする。validate()内のphantom row判定と同じ基準を用いる。
     *   これを除外しないと、グリッドの最後の1行を削除した直後に残る末尾の空行テンプレートが
     *   「有効な行」として誤カウントされ、プラント削除時の存在チェックが常に失敗してしまう。
     */
    private boolean isPhantomBrandRow(BrandKoseiRow row) {
        return row.getBrandkoseiId() == null
                && row.getBrandId() == null
                && isEmpty(row.getBrandsyosaiNk());
    }

    private int countAliveBrandRows(Mcm0012uForm form) {
        if (form == null || form.getBrandKoseiRows() == null) {
            return 0;
        }
        int count = 0;
        for (BrandKoseiRow row : form.getBrandKoseiRows()) {
            if (row != null && !row.isDeleteFlg() && !isPhantomBrandRow(row)) {
                count++;
            }
        }
        return count;
    }
    
    private long countMitsumoriByPlantId(Long plantId) {
        String sql = "SELECT COUNT(1) FROM MCM.MCM_TM_MITSUMORI WHERE PLANT_ID = ?";
        Long count = jdbcTemplate.queryForObject(sql, Long.class, plantId);
        return count == null ? 0L : count;
    }

    private long countKikanByPlantId(Long plantId) {
        String sql = "SELECT COUNT(1) FROM MCM.MCM_TK_KIKAN WHERE PLANT_ID = ?";
        Long count = jdbcTemplate.queryForObject(sql, Long.class, plantId);
        return count == null ? 0L : count;
    }

    private long countKikiKotaikanriByBrandkoseiId(Long brandkoseiId) {
        String sql = "SELECT COUNT(1) FROM MCM.MCM_MA_KIKIKOTAIKANRI WHERE BRANDKOSEI_ID = ?";
        Long count = jdbcTemplate.queryForObject(sql, Long.class, brandkoseiId);
        return count == null ? 0L : count;
    }


    // ========================================
    // データ読込
    // 【変換元】Mcm0012uScreen_Load() + Search()
    // ========================================

    public Mcm0012uForm loadPlantData(Long plantId, String nonyusakiCd,
            String nonyusakiNk, Long nonyusakiId) {
		Mcm0012uForm form = new Mcm0012uForm();
		form.setUpdateMode(Mcm0012uConstants.UPDATE_MODE_UPDATE);
		form.setPlantId(plantId);
		
		// まずは遷移元から受け取った値を仮設定
		form.setNonyusakiCd(nonyusakiCd);
		form.setNonyusakiNk(nonyusakiNk);
		form.setNonyusakiId(nonyusakiId);

		// ★追加（#300）：ヘッダー側の登録者/更新者（ログインID）を一時保持する変数。
		//   担当者マスタへの変換はブランド構成側のログインIDも合わせて
		//   1クエリでまとめて行うため、この時点ではformにセットしない。
		String plantCreatedBy = "";
		String plantLastupdateBy = "";
		
		// ★ Long → BigDecimal に変換して findById に渡す
		Optional<PlantEntity> plantOpt = plantRepository.findById(longToBd(plantId));
		if (plantOpt.isPresent()) {
			PlantEntity p = plantOpt.get();
			
			// plant から採れる情報を設定
			form.setSupportId(p.getSupportId());
			form.setPlantNk(p.getPlantNk());
			form.setTehaiseiban(p.getTehaiseiban());
			form.setNonyuDt(fmtDate(p.getNonyuDt()));
			form.setTekkyoDt(fmtDate(p.getTekkyoDt()));
			form.setHosyusyusokuDt(fmtDate(p.getHosyusyusokuDt()));
			form.setSystemkadojikan(bdToStr(p.getSystemkadojikan()));
			form.setSystemkadonissu(bdToStr(p.getSystemkadonissu()));
			form.setSystemkadoyoubi(p.getSystemkadoyoubi());
			form.setNonyubusyoNk(p.getNonyubusyoNk());
			form.setNonyutantosyaNk(p.getNonyutantosyaNk());
			form.setNonyutelNo(p.getNonyutelNo());
			form.setNonyufaxNo(p.getNonyufaxNo());
			form.setBiko(p.getBiko());
			form.setDtsrenkeiFlg(p.getDtsrenkeiFlg() != null && p.getDtsrenkeiFlg() == 1);

			// 【変換元】KOSINRIREKIGroupBox — 登録日時/登録者/更新日時/更新者（ReadOnly）
			// ★修正（#300）：登録者/更新者は生のログインIDをそのまま出さず、
			//   担当者マスタ変換後の値をform確定時（後段）にまとめてセットする。
			form.setCreatedDt(fmtDateTime(p.getCreatedDt()));
			form.setLastupdateDt(fmtDateTime(p.getLastupdateDt()));
			// 【不具合修正 #296】表示時点の LASTUPDATE_DT を排他キーとして画面へ渡す
			form.setLockLastupdateDt(ExclusiveLockKey.format(p.getLastupdateDt()));
			plantCreatedBy = p.getCreatedBy() != null ? p.getCreatedBy() : "";
			plantLastupdateBy = p.getLastupdateBy() != null ? p.getLastupdateBy() : "";

		
			// ★ URLから nonyusakiId が来ていなくても、Plant から補完する
			if (form.getNonyusakiId() == null && p.getNonyusakiId() != null) {
				form.setNonyusakiId(bdToLong(p.getNonyusakiId()));
			}
		}
		
		// ★ 納入先コード・納入先名を DB から補完する
		//    これにより、0010のサポートIDリンクのように
		//    URLに nonyusakiCd / nonyusakiNk が無い遷移でも右上表示が出る
		if (form.getNonyusakiId() != null) {
			Optional<NonyusakiEntity> nonyusakiOpt =
			nonyusakiRepository.findById(longToBd(form.getNonyusakiId()));
			
			if (nonyusakiOpt.isPresent()) {
				NonyusakiEntity n = nonyusakiOpt.get();
				
				// DBの正を優先して表示項目を設定
				form.setNonyusakiCd(n.getNonyusakiCd());
				form.setNonyusakiNk(n.getNonyusakiNk());
			}
		}
		
		// もしDBから取得できなかった場合のみ、遷移元パラメータを利用
		if (isEmpty(form.getNonyusakiCd())) {
			form.setNonyusakiCd(nonyusakiCd);
		}
		if (isEmpty(form.getNonyusakiNk())) {
			form.setNonyusakiNk(nonyusakiNk);
		}
		
		// ブランド構成読込
		List<BrandKoseiEntity> bkList = brandKoseiRepository.findByPlantIdOrderByHyojijun(plantId);
		
		// 有効契約チェック
		List<Long> activeIds;
		try {
			activeIds = brandKoseiRepository.findBrandkoseiIdsWithActiveContracts(plantId);
		} catch (Exception e) {
			log.warn("契約チェックスキップ（テーブル未作成の可能性）: {}", e.getMessage());
			activeIds = new ArrayList<>();
		}
		Set<Long> activeSet = new HashSet<>(activeIds);

		// ★追加（#300）：ヘッダー＋ブランド構成グリッド全行の登録者/更新者ログインIDを収集し、
		//   担当者マスタを1クエリでまとめて取得する（N+1回避）。
		Set<String> loginIds = new HashSet<>();
		loginIds.add(plantCreatedBy);
		loginIds.add(plantLastupdateBy);
		for (BrandKoseiEntity entity : bkList) {
			loginIds.add(entity.getCreatedBy());
			loginIds.add(entity.getLastupdateBy());
		}
		Map<String, String> tantoNkMap = buildTantoNkMap(loginIds);

		// ★追加（#300）：ヘッダー側の登録者/更新者を担当者名に変換して確定
		form.setCreatedBy(resolveTantoNk(plantCreatedBy, tantoNkMap));
		form.setLastupdateBy(resolveTantoNk(plantLastupdateBy, tantoNkMap));
		
		List<BrandKoseiRow> rows = new ArrayList<>();
		for (BrandKoseiEntity entity : bkList) {
			BrandKoseiRow r = new BrandKoseiRow();
			r.setBrandkoseiId(entity.getBrandkoseiId());
			r.setPlantId(entity.getPlantId());
			r.setHyojijun(bdToStr(entity.getHyojijun()));
			r.setBrandId(entity.getBrandId());
			r.setBrandsyosaiNk(entity.getBrandsyosaiNk());
			r.setNounyuKbn(entity.getNounyuKbn());
			r.setNonyuDt(fmtDate(entity.getNonyuDt()));
			r.setTekkyoDt(fmtDate(entity.getTekkyoDt()));
			r.setHosyusyusokuDt(fmtDate(entity.getHosyusyusokuDt()));
			r.setRemoteFlg(entity.getRemoteFlg() != null && entity.getRemoteFlg() == 1);
			r.setRemoterenrakusaki(entity.getRemoterenrakusaki());
			r.setDremosFlg(entity.getDremosFlg() != null && entity.getDremosFlg() == 1);
			r.setDremosNm(bdToStr(entity.getDremosNm()));
			r.setSystemsekkeiKin(bdToStr(entity.getSystemsekkeiKin()));
			r.setKihonsekkeiKin(bdToStr(entity.getKihonsekkeiKin()));
			r.setProgramsakuseiKin(bdToStr(entity.getProgramsakuseiKin()));
			r.setBiko(entity.getBiko());
			

			r.setCreatedDt(fmtDateTime(entity.getCreatedDt()));
			r.setLastupdateDt(fmtDateTime(entity.getLastupdateDt()));
			// 【不具合修正 #296】ブランド構成行ごとの排他キー
			r.setLockLastupdateDt(ExclusiveLockKey.format(entity.getLastupdateDt()));
			// ★修正（#300）：登録者/更新者を担当者マスタの担当者名に変換してセット
			r.setCreatedBy(resolveTantoNk(entity.getCreatedBy(), tantoNkMap));
			r.setLastupdateBy(resolveTantoNk(entity.getLastupdateBy(), tantoNkMap));

			
			r.setHasActiveContract(activeSet.contains(entity.getBrandkoseiId()));
			rows.add(r);
		}
		form.setBrandKoseiRows(rows);
		
        // 【変換元】VB DataGridView.AllowUserToAddRows = True
        //   VB.NETではDataGridViewが自動で空行を1行表示していた。
        //   Java版では新規時にデフォルト1行を手動で追加する。
		// 修正後
        // ★修正：既存ブランド構成が0件の場合でも、デフォルト空行は追加しない。
        //   理由はコントローラー側の修正コメントと同様（[0]/[1]二重生成の防止）。
        //   HTML側の末尾空行テンプレートが brandKoseiRows[0] として機能するため、
        //   ここで rows が空のままでも画面上は問題なく1行目の入力欄が表示される。
		
		return form;
	}


    // ========================================
    // バリデーション
    // ========================================

    public List<String> validate(Mcm0012uForm form) {
        List<String> errors = new ArrayList<>();

        if (isEmpty(form.getSupportId())) {
            errors.add("サポートIDは必ず入力してください。");
        }
        // ★追加：画面side maxlength撤廃対応。サーバー側で最大桁数チェックを行う。
        //   出典：PlantEntity#supportId（@Column length=7）
        if (form.getSupportId() != null && form.getSupportId().length() > MAX_SUPPORT_ID_LENGTH) {
            errors.add("サポートIDは" + MAX_SUPPORT_ID_LENGTH + "桁以下で入力してください。");
        }
        // ★プラント名は必須ではないため、空チェックは行わない。
        // ★追加：PlantEntity#plantNk（@Column length=80）
        if (form.getPlantNk() != null && form.getPlantNk().length() > MAX_PLANT_NK_LENGTH) {
            errors.add("プラント名は" + MAX_PLANT_NK_LENGTH + "桁以下で入力してください。");
        }
        if (!isEmpty(form.getNonyuDt()) && !isValidDate(form.getNonyuDt())) {
            errors.add("納入日の日付形式が不正です。（yyyy/MM/dd）");
        }
        if (!isEmpty(form.getTekkyoDt()) && !isValidDate(form.getTekkyoDt())) {
            errors.add("撤去日の日付形式が不正です。（yyyy/MM/dd）");
        }
        if (!isEmpty(form.getHosyusyusokuDt()) && !isValidDate(form.getHosyusyusokuDt())) {
            errors.add("保守終息日の日付形式が不正です。（yyyy/MM/dd）");
        }
        // ★追加：PlantEntity#nonyubusyoNk（@Column length=60）
        if (form.getNonyubusyoNk() != null && form.getNonyubusyoNk().length() > MAX_NONYUBUSYO_NK_LENGTH) {
            errors.add("納入部署は" + MAX_NONYUBUSYO_NK_LENGTH + "桁以下で入力してください。");
        }
        // ★追加：PlantEntity#nonyutantosyaNk（@Column length=40）
        if (form.getNonyutantosyaNk() != null && form.getNonyutantosyaNk().length() > MAX_NONYUTANTOSYA_NK_LENGTH) {
            errors.add("納入担当者は" + MAX_NONYUTANTOSYA_NK_LENGTH + "桁以下で入力してください。");
        }
        // ★変更：メッセージ形式をシステム共通形式（「○○は○○桁以下で入力してください。」）に統一
        if (form.getNonyutelNo() != null && form.getNonyutelNo().length() > MAX_NONYUTEL_NO_LENGTH) {
            errors.add("納入電話は" + MAX_NONYUTEL_NO_LENGTH + "桁以下で入力してください。");
        }
        if (form.getNonyufaxNo() != null && form.getNonyufaxNo().length() > 50) {
            errors.add("納入先FAX番号は50文字以内で入力してください。");
        }
        if (form.getTehaiseiban() != null && form.getTehaiseiban().length() > 50) {
            errors.add("手配製番は50文字以内で入力してください。");
        }
        // ★追加：備考は4000桁を超えるとDBの桁あふれエラーになるため、事前に入力チェックを行う。
        //   PlantEntity#biko（@Column length=4000）
        if (form.getBiko() != null && form.getBiko().length() > MAX_BIKO_LENGTH) {
            errors.add("備考は" + MAX_BIKO_LENGTH + "桁以下で入力してください。");
        }

     // 修正後
        if (form.getBrandKoseiRows() != null) {
            for (int i = 0; i < form.getBrandKoseiRows().size(); i++) {
                BrandKoseiRow row = form.getBrandKoseiRows().get(i);
                if (row == null || row.isDeleteFlg()) {
                    continue;
                }
                // ★追加：brandkoseiId も無く、ブランド・ブランド詳細名も両方未入力の
                //   「実体の無い行（幽霊行）」はバリデーション対象から除外する。
                //   画面側の空行昇格ロジックが正しく機能していれば通常は発生しないが、
                //   インデックス欠番等による意図しない空行混入への保険として設置。
                if (isPhantomBrandRow(row)) {
                    continue;
                }
                int n = i + 1;

                if (row.getBrandId() == null) {
                    errors.add("ブランド構成 " + n + "行目: ブランドは必須です。");
                }
                if (isEmpty(row.getBrandsyosaiNk())) {
                    errors.add("ブランド構成 " + n + "行目: ブランド詳細名は必須です。");
                }
                // ★追加：画面side maxlength撤廃対応。BrandKoseiEntity#brandsyosaiNk（@Column length=80）
                if (row.getBrandsyosaiNk() != null && row.getBrandsyosaiNk().length() > MAX_BRANDSYOSAI_NK_LENGTH) {
                    errors.add("ブランド構成 " + n + "行目: ブランド詳細名は" + MAX_BRANDSYOSAI_NK_LENGTH + "桁以下で入力してください。");
                }
                if (!isEmpty(row.getNonyuDt()) && !isValidDate(row.getNonyuDt())) {
                    errors.add("ブランド構成 " + n + "行目: 納入日の日付形式が不正です。");
                }
                if (!isEmpty(row.getTekkyoDt()) && !isValidDate(row.getTekkyoDt())) {
                    errors.add("ブランド構成 " + n + "行目: 撤去日の日付形式が不正です。");
                }
                if (!isEmpty(row.getHosyusyusokuDt()) && !isValidDate(row.getHosyusyusokuDt())) {
                    errors.add("ブランド構成 " + n + "行目: 保守終息日の日付形式が不正です。");
                }
                if (!isEmpty(row.getSystemsekkeiKin())) {
                    if (!isNum(row.getSystemsekkeiKin())) {
                        errors.add("ブランド構成 " + n + "行目: システム設計費は数値で入力してください。");
                    } else if (!isWithinDigitLimit(row.getSystemsekkeiKin(), MAX_KIN_DIGITS)) {
                        errors.add("ブランド構成 " + n + "行目: システム設計費は" + MAX_KIN_DIGITS + "桁以内で入力してください。");
                    }
                }
                if (!isEmpty(row.getKihonsekkeiKin())) {
                    if (!isNum(row.getKihonsekkeiKin())) {
                        errors.add("ブランド構成 " + n + "行目: 基本設計費は数値で入力してください。");
                    } else if (!isWithinDigitLimit(row.getKihonsekkeiKin(), MAX_KIN_DIGITS)) {
                        errors.add("ブランド構成 " + n + "行目: 基本設計費は" + MAX_KIN_DIGITS + "桁以内で入力してください。");
                    }
                }
                if (!isEmpty(row.getProgramsakuseiKin())) {
                    if (!isNum(row.getProgramsakuseiKin())) {
                        errors.add("ブランド構成 " + n + "行目: プログラム作成費は数値で入力してください。");
                    } else if (!isWithinDigitLimit(row.getProgramsakuseiKin(), MAX_KIN_DIGITS)) {
                        errors.add("ブランド構成 " + n + "行目: プログラム作成費は" + MAX_KIN_DIGITS + "桁以内で入力してください。");
                    }
                }
                // ★重要：#254の主原因カラム（実DB decimal(4,0)＝最大9999）
                // ★修正：負数チェックを追加。負数は桁数チェックより先に判定し、専用メッセージを表示する。
                //   また、桁数エラーメッセージも指定のシステム共通形式に統一。
                if (!isEmpty(row.getDremosNm())) {
                    if (!isNum(row.getDremosNm())) {
                        errors.add("ブランド構成 " + n + "行目: DREMOS契約台数は数値で入力してください。");
                    } else if (toBd(row.getDremosNm()).compareTo(BigDecimal.ZERO) < 0) {
                        errors.add("DREMOS契約台数は0より大きい値を入力してください。");
                    } else if (!isWithinDigitLimit(row.getDremosNm(), MAX_DREMOSNM_DIGITS)) {
                        errors.add("DREMOS契約台数は" + MAX_DREMOSNM_DIGITS + "桁以下で入力してください。");
                    }
                }
                // ★変更：メッセージを画面表示ラベル「No」に合わせたシステム共通形式に統一
                if (!isEmpty(row.getHyojijun())) {
                    if (!isNum(row.getHyojijun())) {
                        errors.add("ブランド構成 " + n + "行目: Noは数値で入力してください。");
                    } else if (!isWithinDigitLimit(row.getHyojijun(), MAX_HYOJIJUN_DIGITS)) {
                        errors.add("Noは" + MAX_HYOJIJUN_DIGITS + "桁以下で入力してください。");
                    }
                }
                // ★追加：画面側のmaxlength制限撤廃対応。BrandKoseiEntity#remoterenrakusaki（@Column length=200）
                if (row.getRemoterenrakusaki() != null && row.getRemoterenrakusaki().length() > MAX_REMOTERENRAKUSAKI_LENGTH) {
                    errors.add("リモート連絡先は" + MAX_REMOTERENRAKUSAKI_LENGTH + "桁以下で入力してください。");
                }
            }
        }
        return errors;
    }
    
    /**
     * 【変換元】MasterUpdate()
     * ブランド構成が1件以上存在することを保証して保存系前提条件を確認する。
     */
    public List<String> validateForSave(Mcm0012uForm form) {
        List<String> errors = validate(form);
        if (countAliveBrandRows(form) == 0) {
            errors.add("ブランド構成情報が存在しません。最低1件のブランド構成を登録してください。");
        }
        // ★追加：ブランド構成削除の参照整合性チェック
        //   画面の「行削除」ボタン押下時点ではdeleteFlgを立てて非表示にするだけで、
        //   実際の削除はUpdateButton押下時のsaveAll()内で一括実行される。
        //   そのため、削除対象行が機器個体管理で使用中でないかを保存前にここで検証する。
        errors.addAll(checkBrandKoseiRowsDeletable(form));
        return errors;
    }

    /**
     * 【変換元】BeforeDeleteBrandKoseiCheck()
     * フォーム上で削除フラグが立っている既存ブランド構成行について、
     * 機器個体管理（MCM_MA_KIKIKOTAIKANRI）で使用中でないかをチェックする。
     */
    private List<String> checkBrandKoseiRowsDeletable(Mcm0012uForm form) {
        List<String> errors = new ArrayList<>();
        if (form.getBrandKoseiRows() == null) {
            return errors;
        }
        for (BrandKoseiRow row : form.getBrandKoseiRows()) {
            if (row == null || !row.isDeleteFlg() || row.getBrandkoseiId() == null) {
                continue;
            }
            errors.addAll(checkBrandKoseiDeletable(row.getBrandkoseiId()));
        }
        return errors;
    }
    
    /**
     * 【変換元】NONYUSAKITSUIKAButton_Click()
     * 納入機器追加前専用チェック。
     */
    public List<String> validateForAddEquipment(Mcm0012uForm form) {
        return validateForSave(form);
    }

    // ========================================
    // プラント + ブランド構成 一括保存
    // ========================================

    /**
     * プラント＋ブランド構成の一括保存。
     *
     * 【不具合修正 #296】楽観的排他制御を追加。
     *   従来は PK のみを条件に更新していたため、同一レコードを2クライアントで
     *   同時編集すると後勝ちで上書きされ、先行ユーザーの更新内容が失われていた
     *   （試験項目 INT07-059 NG）。MCM0013U（機器構成マスタ）と同一の方針で、
     *   「画面表示時点の LASTUPDATE_DT」と「DB更新時点の LASTUPDATE_DT」を比較し、
     *   不一致なら排他エラーとして処理全体を中止する（@Transactional によりロールバック）。
     *
     *   判定は2段構え。
     *   (1) 画面が保持していた排他キーと DB の現在値を Java 側で比較する。
     *   (2) 更新前に排他条件付きの UPDATE を発行し、更新件数0件を競合とする。
     *       これにより該当行が排他ロックされ、(1) から実更新までの割り込みも塞げる。
     *
     *   ※ 新規登録（INSERT）は排他対象レコードが存在しないためチェック対象外。
     *
     * @throws ExclusiveControlException 排他競合を検知した場合（DBは更新しない）
     */
    @Transactional
    public Long saveAll(Mcm0012uForm form, String loginUser) {
        PlantEntity plant;
        LocalDateTime now = LocalDateTime.now();

        if (form.getUpdateMode() == Mcm0012uConstants.UPDATE_MODE_UPDATE && form.getPlantId() != null) {
            // 更新
            plant = plantRepository.findById(longToBd(form.getPlantId()))
                    // 【不具合修正 #296】他ユーザーが削除済みのケースも排他エラーとする
                    .orElseThrow(this::exclusiveError);

            // 【不具合修正 #296】排他チェック
            //   入力内容の妥当性より先に判定する。画面が古い時点の情報を前提に
            //   しているため、入力値が何であれ処理を続行してはならない。
            now = ExclusiveLockKey.nextVersion(plant.getLastupdateDt());
            checkPlantLock(plant, form.getLockLastupdateDt(), now, loginUser);

            // 更新日時／更新者
            plant.setLastupdateDt(now);
            plant.setLastupdateBy(loginUser);

        } else {
            // 新規
            plant = new PlantEntity();

            // ★ 手動採番
            BigDecimal nextPlantId = plantRepository.findNextPlantId();
            plant.setPlantId(nextPlantId);

            // form側にも反映しておくと、その後のリダイレクトや子明細保存で安全
            form.setPlantId(nextPlantId.longValue());

            // 作成日時／作成者
            plant.setCreatedDt(now);
            plant.setCreatedBy(loginUser);

            // 更新日時／更新者
            plant.setLastupdateDt(now);
            plant.setLastupdateBy(loginUser);
        }

        // 共通項目設定
        plant.setNonyusakiId(longToBd(form.getNonyusakiId()));
        plant.setSupportId(form.getSupportId());
        plant.setPlantNk(form.getPlantNk());
        plant.setTehaiseiban(form.getTehaiseiban());
        plant.setNonyuDt(parseDate(form.getNonyuDt()));
        plant.setTekkyoDt(parseDate(form.getTekkyoDt()));
        plant.setHosyusyusokuDt(parseDate(form.getHosyusyusokuDt()));
        plant.setSystemkadojikan(toBd(form.getSystemkadojikan()));
        plant.setSystemkadonissu(toBd(form.getSystemkadonissu()));
        plant.setSystemkadoyoubi(form.getSystemkadoyoubi());
        plant.setNonyubusyoNk(form.getNonyubusyoNk());
        plant.setNonyutantosyaNk(form.getNonyutantosyaNk());
        plant.setNonyutelNo(form.getNonyutelNo());
        plant.setNonyufaxNo(form.getNonyufaxNo());
        plant.setBiko(form.getBiko());
        plant.setDtsrenkeiFlg(form.isDtsrenkeiFlg() ? 1 : 0);

        plant = plantRepository.save(plant);

        // 保存後のPlantId
        Long savedPlantId = bdToLong(plant.getPlantId());

        // --- ブランド構成の保存 ---
        if (form.getBrandKoseiRows() != null) {
            Long maxId = brandKoseiRepository.findMaxBrandkoseiId();
            long nextId = (maxId != null ? maxId : 0) + 1;

            for (BrandKoseiRow row : form.getBrandKoseiRows()) {
                if (row == null) {
                    continue;
                }
                if (row.isDeleteFlg()) {
                    if (row.getBrandkoseiId() != null) {
                        // 【不具合修正 #296】削除も更新と同じ方式で排他判定する
                        deleteBrandKoseiWithLock(row.getBrandkoseiId(), row.getLockLastupdateDt());
                    }
                    continue;
                }

                BrandKoseiEntity entity;
                if (row.getBrandkoseiId() != null && row.getBrandkoseiId() > 0) {
                    // 【不具合修正 #296】既存行が他ユーザーに削除されていた場合、
                    //   従来は new BrandKoseiEntity() へフォールバックして
                    //   同一IDで復活INSERTされてしまうため排他エラーに変更する。
                    entity = brandKoseiRepository.findById(row.getBrandkoseiId())
                            .orElseThrow(this::exclusiveError);
                    // 【不具合修正 #296】排他チェック
                    checkBrandKoseiLock(entity, row.getLockLastupdateDt(), now, loginUser);
                } else {
                    entity = new BrandKoseiEntity();
                    entity.setBrandkoseiId(nextId++);
                }

                entity.setPlantId(savedPlantId);
                entity.setBrandId(row.getBrandId());
                entity.setBrandsyosaiNk(row.getBrandsyosaiNk());
                entity.setNounyuKbn(row.getNounyuKbn());
                entity.setNonyuDt(parseDate(row.getNonyuDt()));
                entity.setTekkyoDt(parseDate(row.getTekkyoDt()));
                entity.setHosyusyusokuDt(parseDate(row.getHosyusyusokuDt()));
                entity.setRemoteFlg(row.isRemoteFlg() ? 1 : 0);
                entity.setRemoterenrakusaki(row.getRemoterenrakusaki());
                entity.setDremosFlg(row.isDremosFlg() ? 1 : 0);
                entity.setDremosNm(toBd(row.getDremosNm()));
                entity.setHyojijun(toBd(row.getHyojijun()));
                entity.setSystemsekkeiKin(toBd(row.getSystemsekkeiKin()));
                entity.setKihonsekkeiKin(toBd(row.getKihonsekkeiKin()));
                entity.setProgramsakuseiKin(toBd(row.getProgramsakuseiKin()));
                entity.setBiko(row.getBiko());

                LocalDateTime brandVersion = entity.getLastupdateDt();
                brandKoseiRepository.saveAndFlush(entity);
                if (row.getBrandkoseiId() != null && row.getBrandkoseiId() > 0) {
                    // JPA監査リスナーの時刻設定後にも、確保した排他キーを保持する。
                    jdbcTemplate.update("UPDATE MCM.MCM_MA_BRAND_KOSEI SET LASTUPDATE_DT=?, LASTUPDATE_BY=? WHERE BRANDKOSEI_ID=?",
                            brandVersion, loginUser, entity.getBrandkoseiId());
                }
            }
        }

        return savedPlantId;
    }

    // ========================================
    // プラント削除
    // ========================================

    public List<String> checkPlantDeletable(Long plantId) {
        List<String> errors = new ArrayList<>();
        if (plantId == null) {
            errors.add("プラントIDが指定されていません。");
            return errors;
        }

		long tmPlantDataCount = countMitsumoriByPlantId(plantId);
		long tkPlantDataCount = countKikanByPlantId(plantId);
        if (tmPlantDataCount > 0 || tkPlantDataCount > 0) {
            errors.add("見積・契約関連情報のプラントとして使用されているため、削除できません。");
        }
        return errors;
    }

    /**
     * ★修正：プラント削除チェックボックス押下時の確認ダイアログでは判定せず、
     *   登録ボタン押下時（update）に入力エラーとして扱うためのブランド構成存在チェック。
     *   従来はcheckPlantDeletable()内で削除可否チェックの一部として実施していたが、
     *   確認メッセージ表示後にしか判明しない仕様となっていたため、
     *   登録ボタン押下時点で先に判定し「入力エラー」として画面に表示する形に変更した。
     *
     * ★修正：DBの実登録件数（brandKoseiRepository.countByPlantId）ではなく、
     *   フォーム上の有効行数（countAliveBrandRows）で判定する。
     *   「ブランド構成グリッドで行削除→プラント削除チェック→登録」という操作順の場合、
     *   グリッドの行削除はdeleteFlgを立てるだけで、実際のDB削除は登録処理時に行われる。
     *   そのためDBを直接見ると削除予定の行が「存在する」と誤判定されてしまうため、
     *   フォームの削除フラグを反映した有効行数で判定する。
     */
    public List<String> checkBrandKoseiExistsForPlantDelete(Mcm0012uForm form) {
        List<String> errors = new ArrayList<>();
        if (countAliveBrandRows(form) > 0) {
            errors.add("ブランド構成情報を削除して下さい。");
        }
        return errors;
    }

    /**
     * プラント削除。
     *
     * 【不具合修正 #296】削除も更新と同じ方式で排他判定する。
     *   元VBの DELETE 文も WHERE 句に Original_* 値を持つ楽観的同時実行制御が有効だった。
     *
     * @param screenLockKey 画面が保持していた排他キー（LASTUPDATE_DT）
     * @throws ExclusiveControlException 排他競合を検知した場合（DBは更新しない）
     */
    @Transactional
    public void deletePlant(Long plantId, String loginUser, String screenLockKey) {
        log.info("プラント削除: plantId={}, loginUser={}", plantId, loginUser);

        PlantEntity plant = plantRepository.findById(longToBd(plantId))
                .orElseThrow(this::exclusiveError);
        if (!ExclusiveLockKey.matches(plant.getLastupdateDt(),
                                      ExclusiveLockKey.parse(screenLockKey))) {
            log.info("プラント削除 排他競合検知（画面保持値とDB現在値の不一致）: plantId={}", plantId);
            throw exclusiveError();
        }

        int deleted = jdbcTemplate.update(
                "DELETE FROM MCM.MCM_MA_PLANT WHERE PLANT_ID = ?"
                + " AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
                plantId, plant.getLastupdateDt(), plant.getLastupdateDt());
        if (deleted == 0) {
            log.info("プラント削除 排他競合検知（削除件数0件）: plantId={}", plantId);
            throw exclusiveError();
        }
    }

    // ========================================
    // 【不具合修正 #296】排他制御ヘルパー
    // ========================================

    /** 排他エラー（FWM_0009 相当）を生成する */
    private ExclusiveControlException exclusiveError() {
        return new ExclusiveControlException(
                messageService.getMessage("mcm0012u.msg.exclusive"));
    }

    /**
     * プラント行の排他チェック。
     *
     * <p>1段目として画面保持値とDB現在値を Java 側で比較し、
     * 2段目として排他条件付き UPDATE（更新日時のみ更新）で更新件数0件を競合と判定する。
     * 2段目の UPDATE は対象行に排他ロックを掛けるため、以降の
     * {@code plantRepository.save()} までの間に他トランザクションが割り込めない。</p>
     *
     * @param plant         DBから読み取った現在のプラント
     * @param screenLockKey 画面が保持していた排他キー
     * @param now           今回の更新日時（後続の save() と同じ値を設定する）
     * @param loginUser     更新者
     */
    private void checkPlantLock(PlantEntity plant, String screenLockKey,
                                LocalDateTime now, String loginUser) {
        LocalDateTime dbValue = plant.getLastupdateDt();
        // 【#296 調査用ログ】排他判定に使った実値を必ず残す。
        //   「排他エラーにならない」事象の切り分けに必要な情報は次の3点。
        //     screenLockKey : 画面(hidden)から届いた生文字列。空/nullなら画面側で往復していない
        //     parsed        : 上記を復元した値。screenLockKeyが非空でnullなら書式不正
        //     dbValue       : DB現在値。parsed と一致していれば競合なしと判定される
        log.info("MCM0012U 排他判定[プラント] plantId={} screenLockKey='{}' parsed={} dbValue={}",
                 plant.getPlantId(), screenLockKey, ExclusiveLockKey.parse(screenLockKey), dbValue);
        if (!ExclusiveLockKey.matches(dbValue, ExclusiveLockKey.parse(screenLockKey))) {
            log.info("プラント更新 排他競合検知（画面保持値とDB現在値の不一致）: plantId={}",
                     plant.getPlantId());
            throw exclusiveError();
        }
        // WHERE 句へはDBから読み取った値を渡す。画面から往復した値ではなくDB由来の値を
        // 使うことで、LASTUPDATE_DT の型精度（SQL Server DATETIME=約3.33ms丸め等）による
        // 誤判定を回避する。画面保持値との突き合わせは上記1段目で完了している。
        int updated = jdbcTemplate.update(
                "UPDATE MCM.MCM_MA_PLANT SET LASTUPDATE_DT = ?, LASTUPDATE_BY = ?"
                + " WHERE PLANT_ID = ?"
                + " AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
                now, loginUser, plant.getPlantId(), dbValue, dbValue);
        if (updated == 0) {
            log.info("プラント更新 排他競合検知（更新件数0件）: plantId={}", plant.getPlantId());
            throw exclusiveError();
        }
    }

    /**
     * ブランド構成行の排他チェック。プラント行と同一方式。
     *
     * @param entity        DBから読み取った現在のブランド構成
     * @param screenLockKey 画面が保持していた行ごとの排他キー
     */
    private void checkBrandKoseiLock(BrandKoseiEntity entity, String screenLockKey,
                                     LocalDateTime now, String loginUser) {
        LocalDateTime dbValue = entity.getLastupdateDt();
        now = ExclusiveLockKey.nextVersion(dbValue);
        // 【#296 調査用ログ】判定に使った実値を残す（checkPlantLock と同じ意図）
        log.info("MCM0012U 排他判定[ブランド構成] brandkoseiId={} screenLockKey='{}' parsed={} dbValue={}",
                 entity.getBrandkoseiId(), screenLockKey,
                 ExclusiveLockKey.parse(screenLockKey), dbValue);
        if (!ExclusiveLockKey.matches(dbValue, ExclusiveLockKey.parse(screenLockKey))) {
            log.info("ブランド構成更新 排他競合検知（画面保持値とDB現在値の不一致）: brandkoseiId={}",
                     entity.getBrandkoseiId());
            throw exclusiveError();
        }
        int updated = jdbcTemplate.update(
                "UPDATE MCM.MCM_MA_BRAND_KOSEI SET LASTUPDATE_DT = ?, LASTUPDATE_BY = ?"
                + " WHERE BRANDKOSEI_ID = ?"
                + " AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
                now, loginUser, entity.getBrandkoseiId(), dbValue, dbValue);
        if (updated == 0) {
            log.info("ブランド構成更新 排他競合検知（更新件数0件）: brandkoseiId={}",
                     entity.getBrandkoseiId());
            throw exclusiveError();
        }
        // JDBCで確保した新しい排他キーをJPAの保存で旧値に戻さない。
        entity.setLastupdateDt(now);
        entity.setLastupdateBy(loginUser);
    }

    /**
     * 排他条件付きでブランド構成行を削除する。
     *
     * @param screenLockKey 画面が保持していた行ごとの排他キー
     */
    private void deleteBrandKoseiWithLock(Long brandkoseiId, String screenLockKey) {
        BrandKoseiEntity entity = brandKoseiRepository.findById(brandkoseiId).orElse(null);
        if (entity == null) {
            // 他ユーザーが既に削除済み。DB上の結果は同じだが、画面が古い情報を
            // 前提にしているため「やり直し」を促す（旧VBも更新0件でエラーとなる）。
            throw exclusiveError();
        }
        if (!ExclusiveLockKey.matches(entity.getLastupdateDt(),
                                      ExclusiveLockKey.parse(screenLockKey))) {
            log.info("ブランド構成削除 排他競合検知（画面保持値とDB現在値の不一致）: brandkoseiId={}",
                     brandkoseiId);
            throw exclusiveError();
        }
        int deleted = jdbcTemplate.update(
                "DELETE FROM MCM.MCM_MA_BRAND_KOSEI WHERE BRANDKOSEI_ID = ?"
                + " AND ((? IS NULL AND LASTUPDATE_DT IS NULL) OR LASTUPDATE_DT = ?)",
                brandkoseiId, entity.getLastupdateDt(), entity.getLastupdateDt());
        if (deleted == 0) {
            log.info("ブランド構成削除 排他競合検知（削除件数0件）: brandkoseiId={}", brandkoseiId);
            throw exclusiveError();
        }
    }

    /**
     * ★追加：プラント削除時、フォーム上で削除マーク（deleteFlg=true）された
     *   既存ブランド構成行を実際にDBから削除する。
     *   「グリッドで行削除→プラント削除チェック→登録」の操作順の場合、
     *   グリッドの行削除時点ではdeleteFlgが立つだけでDB上はまだ残っているため、
     *   プラント削除実行前にこれらを削除しておく（FK制約対策も兼ねる）。
     *
     * 【不具合修正 #296】ここでの削除も排他判定の対象とする。
     *   プラント削除は「削除マーク済みブランド構成の削除 → プラント削除」の順で
     *   同一トランザクション内で実行されるため、いずれかで競合を検知すれば
     *   全体がロールバックされ、DB更新件数は0件となる。
     */
    @Transactional
    public void deleteMarkedBrandKoseiRows(Mcm0012uForm form, String loginUser) {
        if (form == null || form.getBrandKoseiRows() == null) {
            return;
        }
        for (BrandKoseiRow row : form.getBrandKoseiRows()) {
            if (row != null && row.isDeleteFlg() && row.getBrandkoseiId() != null) {
                deleteBrandKoseiWithLock(row.getBrandkoseiId(), row.getLockLastupdateDt());
            }
        }
    }

    // ========================================
    // ブランド構成行 削除
    // ========================================

    public List<String> checkBrandKoseiDeletable(Long brandkoseiId) {
        List<String> errors = new ArrayList<>();
        if (brandkoseiId == null) {
            errors.add("ブランド構成IDが指定されていません。");
            return errors;
        }

        long count = countKikiKotaikanriByBrandkoseiId(brandkoseiId);
        if (count > 0) {
            errors.add("個体管理情報のブランド構成として、既に使用されている為、削除する事が出来ません。");
        }
        return errors;
    }

    @Transactional
    public void deleteBrandKosei(Long brandkoseiId, String loginUser) {
    	log.info("ブランド構成行削除: brandkoseiId={}, loginUser={}", brandkoseiId, loginUser);
        brandKoseiRepository.deleteByBrandkoseiId(brandkoseiId);
    }
}
